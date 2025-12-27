/**
 * Syncs multiple Google Sheets tabs to Firestore collections.
 * Uses batch operations (REST batchWrite) where possible for efficiency.
 */

/**
 * Main orchestrator to sync all configured sheets.
 * Run this function manually to start the sync.
 */
function syncAllDataToFirestore() {
  const tasks = [
    { sheetName: 'DD', collection: 'dd_beneficiaries', idField: 'Beneficiary' },
    { sheetName: 'CarT_P', collection: 'cartp_plan', idField: 'Ref' },
    { sheetName: 'Ops_P', collection: 'opsp_plan', idField: 'Ref' },
    { sheetName: 'Vehicle_InUse', collection: 'vehicle_in_use', idField: 'Vehicle Number' },
    { sheetName: 'Vehicle_Released', collection: 'vehicle_released', idField: '' }, // Auto-ID
    { sheetName: 'Vehicle_History', collection: 'vehicle_history', idField: '' }    // Auto-ID
  ];

  const results = [];
  tasks.forEach(task => {
    try {
      console.log(`Starting sync for ${task.sheetName}...`);
      const res = syncSheetToCollection(task.sheetName, task.collection, task.idField);
      results.push(`✅ ${task.sheetName}: ${res.count} docs`);
    } catch (e) {
      console.error(`Failed to sync ${task.sheetName}`, e);
      results.push(`❌ ${task.sheetName}: ${e.message}`);
    }
  });

  // Sync Submissions (Specialized)
  try {
     console.log('Starting specialized sync for Submissions...');
     const res = syncSubmissionsToFirestore();
     results.push(`✅ Submissions: ${res.count} docs`);
  } catch (e) {
     console.error('Failed to sync Submissions', e);
     results.push(`❌ Submissions: ${e.message}`);
  }

  console.log('Sync Complete:\n' + results.join('\n'));
  return results.join('\n');
}

/**
 * Syncs a single sheet to a Firestore collection.
 * Replaces column headers with camelCase keys.
 */
function syncSheetToCollection(sheetName, collectionName, idFieldHeader) {
  // 1. Resolve Sheet
  let sheetId = SHEET_ID; // default
  if (['CarT_P', 'Vehicle_InUse', 'Vehicle_Released', 'Vehicle_History'].includes(sheetName) && typeof CAR_SHEET_ID !== 'undefined') {
    sheetId = CAR_SHEET_ID;
  }
  
  // Handle DD alias
  const targetSheetName = (sheetName === 'DD' && typeof DATA_SHEET_NAME !== 'undefined') ? DATA_SHEET_NAME : sheetName;

  const ss = SpreadsheetApp.openById(sheetId);
  const sh = ss.getSheetByName(targetSheetName);
  if (!sh) throw new Error(`Sheet "${targetSheetName}" not found in spreadsheet ${sheetId}`);

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return { count: 0 }; // No data

  // 2. Process Headers
  const rawHeaders = data[0].map(String);
  const headers = rawHeaders.map(h => toCamelCase_(h));
  
  // Find mapped ID column index
  let idColIdx = -1;
  if (idFieldHeader) {
    const normIdHeader = toCamelCase_(idFieldHeader);
    idColIdx = headers.indexOf(normIdHeader);
    // fallback to loose match if not found
    if (idColIdx === -1) {
       idColIdx = rawHeaders.findIndex(h => h.trim().toLowerCase() === idFieldHeader.trim().toLowerCase());
    }
  }

  // 3. Process Rows
  const documents = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const doc = {};
    let isEmpty = true;
    
    // ID generation
    let docId = null;
    if (idColIdx !== -1) {
       const val = String(row[idColIdx]||'').trim();
       if (val) docId = sanitizeDocId_(val);
    }

    row.forEach((cell, idx) => {
      const key = headers[idx];
      if (!key) return; // skip empty headers
      const val = normalizeValueForFirestore_(cell);
      
      if (val !== null && val !== '' && val !== undefined) {
        doc[key] = val;
        isEmpty = false;
      }
    });

    if (!isEmpty) {
      documents.push({ 
        id: docId, 
        fields: doc 
      });
    }
  }

  console.log(`Prepared ${documents.length} documents for ${collectionName}. sending to batch write...`);
  
  // 4. Batch Write (Chunked)
  // Firestore REST batchWrite allows max 500 writes per request.
  const CHUNK_SIZE = 400; // conservative limit
  let processed = 0;
  
  for (let i = 0; i < documents.length; i += CHUNK_SIZE) {
    const chunk = documents.slice(i, i + CHUNK_SIZE);
    firestoreBatchWrite_(collectionName, chunk);
    processed += chunk.length;
    console.log(`Synced ${processed}/${documents.length} to ${collectionName}`);
  }

  return { count: processed };
}

/**
 * Perform a batch write to Firestore using REST API.
 * Uses `commit` endpoint.
 */
function firestoreBatchWrite_(collectionName, docs) {
  const projectId = getFirestoreProjectId_();
  const url = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents:commit`;
  const token = getFirestoreAccessToken_();

  const writes = docs.map(d => {
    const docPath = `projects/${projectId}/databases/(default)/documents/${collectionName}/${d.id || ''}`;
    
    // If we have a specific ID, use 'update' with upsert (if exists or not). 
    // If no ID (auto-id), checking logic needed? 
    // Actually, for auto-ID with REST batch, strictly we need to generate IDs client-side or use createDocument one by one.
    // To support "Auto-ID" in batch, we can assume we generate a random one if missing.
    
    const finalPath = d.id 
      ? docPath 
      // Generate a client-side auto-ID if missing to allow batching
      : `projects/${projectId}/databases/(default)/documents/${collectionName}/${Utilities.getUuid()}`;

    return {
      update: {
        name: finalPath,
        fields: firestoreFields_(d.fields)
      },
      // Upsert behavior: replace if exists, create if not
    };
  });

  const payload = { writes: writes };
  
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { Authorization: `Bearer ${token}` },
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() >= 300) {
    throw new Error(`Firestore batch commit failed: ${response.getContentText()}`);
  }
}

/**
 * Helper: toCamelCase
 * e.g. "Fuel Amount" -> "fuelAmount"
 */
function toCamelCase_(str) {
  return str
    .toString()
    .trim()
    .toLowerCase()
    .replace(/[^a-zA-Z0-9]+(.)/g, (m, chr) => chr.toUpperCase());
}

/**
 * Helper: Sanitize Document ID
 * Firestore IDs cannot contain forward slashes, cannot be . or .., and cannot match regex __.*__
 */
function sanitizeDocId_(str) {
  return str.replace(/\//g, '_').replace(/^\.+$/, '_').trim();
}

/**
 * Specialized sync for Submissions tab to ensure nested objects match
 * the schema created by submitToSubmissions().
 */
function syncSubmissionsToFirestore() {
  const collectionName = 'submissions';
  const sheetName = (typeof SUBMISSIONS_SHEET_NAME !== 'undefined') ? SUBMISSIONS_SHEET_NAME : 'submissions';
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  
  if (!sh) throw new Error(`Submissions sheet "${sheetName}" not found`);

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return { count: 0 };

  const rawHeaders = data[0].map(String);
  const normalizedHeaders = rawHeaders.map(h => toCamelCase_(h));

  // Helper to safely get value from row by header alias
  const idxMap = {};
  normalizedHeaders.forEach((h, i) => idxMap[h] = i);
  // Fallback map for known headers if aliases fail
  const findIdx = (keys) => {
    const validKeys = Array.isArray(keys) ? keys : [keys];
    for (const k of validKeys) {
      const camel = toCamelCase_(k);
      if (idxMap.hasOwnProperty(camel)) return idxMap[camel];
    }
    return -1;
  };
  
  // Map important columns
  const IDX = {
    id: findIdx(['Submission ID', 'ID']),
    timestamp: findIdx(['Timestamp', 'Date', 'Time']),
    beneficiary: findIdx(['Beneficiary', 'Name']),
    accountHolder: findIdx(['Account Holder']),
    team: findIdx(['Team Name', 'Team']),
    project: findIdx(['Project Name', 'Project']),
    total: findIdx(['Row Total', 'Total']),
    designation: findIdx(['Designation']),
    
    // Nested: Fuel
    fuelFrom: findIdx(['Fuel From']),
    fuelTo: findIdx(['Fuel To']),
    fuelAmt: findIdx(['Fuel Amt', 'Fuel Amount']),
    
    // Nested: DA
    daFrom: findIdx(['DA From']),
    daTo: findIdx(['DA To']),
    daAmt: findIdx(['DA Amt', 'DA Amount']),
    
    // Nested: Car
    carFrom: findIdx(['Car From', 'Vehicle From']),
    carTo: findIdx(['Car To', 'Vehicle To']),
    carNum: findIdx(['Vehicle Number', 'Car Number']),
    carAmt: findIdx(['Car Amt', 'Car Amount']),
    
    // Nested: Airtime
    airFrom: findIdx(['Airtime From']),
    airTo: findIdx(['Airtime To']),
    airAmt: findIdx(['Airtime Amt', 'Airtime Amount']),
    
    // Nested: Transport
    transFrom: findIdx(['Transport From']),
    transTo: findIdx(['Transport To']),
    transAmt: findIdx(['Transport Amt', 'Transport Amount']),
    
    // Nested: Misc
    miscFrom: findIdx(['Misc From']),
    miscTo: findIdx(['Misc To']),
    miscAmt: findIdx(['Misc Amt', 'Misc Amount']),
    
    // Others
    mob: findIdx(['Mob No', 'Mobile']),
    display: findIdx(['Display Name']),
    wh: findIdx(['W/H Charges', 'WH Charges']),
    remarks: findIdx(['Remarks']),
    submitter: findIdx(['Submitter']),
    approvalDate: findIdx(['Approval date']),
    approvedBy: findIdx(['Approved by']),
    paidAmt: findIdx(['Paid amt']),
    transferBy: findIdx(['Transfer By']),
    finRemarks: findIdx(['Finance Remarks'])
  };

  const documents = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Get ID or generate
    let docId = null;
    if (IDX.id !== -1) {
      const val = String(row[IDX.id]||'').trim();
      if (val) docId = sanitizeDocId_(val);
    }
    
    const val = (idx) => (idx !== -1 && row[idx] !== undefined) ? row[idx] : null;
    const num = (idx) => {
       const v = val(idx); 
       return (typeof v === 'number') ? v : Number(parseAmount(v));
    };
    const date = (idx) => {
       const v = val(idx);
       if (v instanceof Date) return normalizeValueForFirestore_(v); // keep as ISO string or timestamp
       if (!v) return null;
       // try parsing if string? using existing parseDate or logic?
       // For sync, let's trust row value if it's already a date object, else string.
       return v; 
    };
    const str = (idx) => {
      const v = val(idx);
      return (v === null || v === undefined) ? '' : String(v).trim();
    };

    const doc = {
       submissionId: docId || '',
       timestamp: date(IDX.timestamp),
       project: str(IDX.project),
       submitter: str(IDX.submitter),
       
       // Reconstruct row structure to match submitToSubmissions
       rows: [{
         beneficiary: str(IDX.beneficiary),
         accountHolder: str(IDX.accountHolder),
         teamName: str(IDX.team),
         projectName: str(IDX.project),
         total: num(IDX.total),
         designation: str(IDX.designation),
         
         fuel: {
           from: date(IDX.fuelFrom) || '',
           to: date(IDX.fuelTo) || '',
           amount: num(IDX.fuelAmt)
         },
         da: {
           from: date(IDX.daFrom) || '',
           to: date(IDX.daTo) || '',
           amount: num(IDX.daAmt)
         },
         car: {
           from: date(IDX.carFrom) || '',
           to: date(IDX.carTo) || '',
           vehicleNumber: str(IDX.carNum),
           amount: num(IDX.carAmt)
         },
         air: {
           from: date(IDX.airFrom) || '',
           to: date(IDX.airTo) || '',
           amount: num(IDX.airAmt)
         },
         transport: {
           from: date(IDX.transFrom) || '',
           to: date(IDX.transTo) || '',
           amount: num(IDX.transAmt)
         },
         misc: {
           from: date(IDX.miscFrom) || '',
           to: date(IDX.miscTo) || '',
           amount: num(IDX.miscAmt)
         },
         
         mob: str(IDX.mob),
         displayName: str(IDX.display),
         whCharges: num(IDX.wh),
         remarks: str(IDX.remarks),
         submitter: str(IDX.submitter),
         approvalDate: date(IDX.approvalDate) || '',
         approvedBy: str(IDX.approvedBy),
         paidAmt: num(IDX.paidAmt),
         transferBy: str(IDX.transferBy),
         financeRemarks: str(IDX.finRemarks)
       }],
       
       metadata: {
         rowCount: 1, // Submissions tab is flat, 1 row per doc usually in this legacy view? 
                      // actually submitToSubmissions writes 1 doc per submission which can have multiple rows.
                      // BUT, the sheet basically flattens it? 
                      // Wait, current sheet structure: 1 row per "item".
                      // If a submission had 5 rows, the sheet has 5 rows.
                      // Does the sheet share a 'Submission ID' across 5 rows? Yes.
                      // So we should GROUP by Submission ID?
                      // The current generic sync was treating each row as a doc. 
                      // `submitToSubmissions` creates ONE doc for multiple rows.
                      // If I want to match, I must Group By Submission ID.
          backfilled: true
       }
    };
    
    documents.push(doc);
  }
  
  // GROUP BY Submission ID
  const grouped = {};
  documents.forEach(d => {
    const key = d.submissionId;
    // If no ID, generate unique? legacy data might check uniqueness?
    // If key is empty, treat as individual doc? 
    if (!key) {
      // no grouping possible
      grouped[Utilities.getUuid()] = d; // single
    } else {
      if (!grouped[key]) {
        grouped[key] = d; // init
      } else {
        // merge rows
        grouped[key].rows.push(d.rows[0]);
        // Update totals?
        grouped[key].metadata.rowCount += 1;
      }
    }
  });
  
  const finalDocs = Object.entries(grouped).map(([id, doc]) => {
     return {
       id: id,
       fields: doc
     };
  });

  console.log(`Backfill: Prepared ${finalDocs.length} submission documents (from ${data.length-1} rows).`);

  // Batch Write
  const CHUNK_SIZE = 400;
  let processed = 0;
  for (let i = 0; i < finalDocs.length; i += CHUNK_SIZE) {
    const chunk = finalDocs.slice(i, i + CHUNK_SIZE);
    firestoreBatchWrite_(collectionName, chunk);
    processed += chunk.length;
    console.log(`Synced ${processed}/${finalDocs.length} to ${collectionName}`);
  }
  
  return { count: processed };
}
