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
 * Persist a single row from a sheet (or row-like object) to Firestore.
 * Used for real-time hooks.
 * @param {string} collectionName Target collection
 * @param {string} idFieldHeader Header name to use as ID (optional)
 * @param {number|Object} rowData Either an array of values (if header provided) or an object (map).
 * @param {Array<string>} [headerRow] Header array if rowData is an array.
 */
function persistRowToFirestore(collectionName, idFieldHeader, rowData, headerRow) {
  try {
    let doc = {};
    let docId = null;

    if (Array.isArray(rowData) && Array.isArray(headerRow)) {
       // Array mode
       const headers = headerRow.map(h => toCamelCase_(h));
       let idColIdx = -1;
       if (idFieldHeader) {
          const normId = toCamelCase_(idFieldHeader);
          idColIdx = headers.indexOf(normId);
          if (idColIdx === -1) {
            idColIdx = headerRow.findIndex(h => String(h).trim().toLowerCase() === String(idFieldHeader).trim().toLowerCase());
          }
       }
       if (idColIdx !== -1) {
         const val = String(rowData[idColIdx]||'').trim();
         if (val) docId = sanitizeDocId_(val);
       }

       rowData.forEach((cell, idx) => {
          const key = headers[idx];
          if (!key) return;
          const val = normalizeValueForFirestore_(cell);
          if (val !== null && val !== '' && val !== undefined) {
            doc[key] = val;
          }
       });
    } else if (typeof rowData === 'object') {
       // Object mode (already keyed) - we assume keys are close to desired, but maybe we should normalize?
       // If coming from `upsertVehicleSummaryRow` rowData is an object with headers as keys.
       // So we need to normalize keys.
       
       for (const [k, v] of Object.entries(rowData)) {
         const key = toCamelCase_(k);
         const val = normalizeValueForFirestore_(v);
         if (val !== null && val !== '' && val !== undefined) {
           doc[key] = val;
         }
         // ID check
         if (idFieldHeader && (k === idFieldHeader || key === toCamelCase_(idFieldHeader))) {
            const idVal = String(v||'').trim();
            if (idVal) docId = sanitizeDocId_(idVal);
         }
       }
    }

    if (Object.keys(doc).length === 0) return; // empty

    const projectId = getFirestoreProjectId_();
    
    // Construct Path
    const finalId = docId || Utilities.getUuid();
    const docPath = `projects/${projectId}/databases/(default)/documents/${collectionName}/${finalId}`;
    
    // REST API patch (upsert)
    // We use patch to allow merging or set fields.
    // To mimic "set" (overwrite), we don't use updateMask.
    
    const url = `https://firestore.googleapis.com/v1/${docPath}`;
    const token = getFirestoreAccessToken_();
    
    const payload = {
      name: docPath,
      fields: firestoreFields_(doc)
    };
    
    const response = UrlFetchApp.fetch(url, {
       method: 'patch',
       headers: { Authorization: `Bearer ${token}` },
       contentType: 'application/json',
       payload: JSON.stringify(payload),
       muteHttpExceptions: true
     });
     
     if (response.getResponseCode() >= 300) {
       console.error(`Persist row failed: ${response.getContentText()}`);
     } else {
       console.log(`Persisted doc to ${collectionName}/${finalId}`);
     }
  } catch (e) {
    console.warn('persistRowToFirestore error', e);
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
    paidAmt: findIdx(['Paid amt', 'Paid Amount']),
    transferBy: findIdx(['Transfer By']),
    finRemarks: findIdx(['Finance Remarks']),
    
    // [NEW] Finance Fields
    dateP: findIdx(['Date P']),
    advRef: findIdx(['Adv Ref', 'Ref', 'Reference']),
    claimBal: findIdx(['Claim/Balance amt', 'Claim/Balance Amount']),
    balAmt: findIdx(['Balance Amt', 'Balance Amount']),
    vchType: findIdx(['Vch Type-Pymt', 'Vch Type']),
    jv: findIdx(['JV']),
    transferTo: findIdx(['Transfer To']),
    mode: findIdx(['Mode'])
  };

  const groups = {}; // Map<SubmissionID, Array<Row>>

  // 1. Group Rows by Submission ID
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let subId = null;
    if (IDX.id !== -1) {
      const val = String(row[IDX.id]||'').trim();
      if (val) subId = sanitizeDocId_(val);
    }
    
    // If no ID, generate one for the group (assuming single-row submission if missing ID, or skip?)
    // Creating orphan groups for missing IDs might be messy, but safer than skipping.
    if (!subId) {
       subId = 'orphan_' + Utilities.getUuid();
    }
    
    if (!groups[subId]) groups[subId] = [];
    groups[subId].push(row);
  }

  // 2. Prepare Batch Writes
  const writes = [];
  const projectId = getFirestoreProjectId_(); // Assumes firestore.gs is loaded
  const databaseRoot = `projects/${projectId}/databases/(default)/documents`;

  Object.entries(groups).forEach(([subId, rows]) => {
    // A. Parse Parent Data (from first row)
    const firstRow = rows[0];
    
    const val = (r, idx) => (idx !== -1 && r[idx] !== undefined) ? r[idx] : null;
    const str = (r, idx) => {
      const v = val(r, idx);
      return (v === null || v === undefined) ? '' : String(v).trim();
    };
    const date = (r, idx) => {
        const v = val(r, idx);
        if (v instanceof Date) return normalizeValueForFirestore_(v); 
        if (!v) return null;
        return v; 
    };

    const parentPath = `${collectionName}/${subId}`;
    const parentData = {
        submissionId: subId,
        project: str(firstRow, IDX.project),
        submitter: str(firstRow, IDX.submitter),
        team: str(firstRow, IDX.team), // [NEW] Optimized Query Field
        totalAmount: rows.reduce((sum, r) => {
           const v = val(r, IDX.total);
           const amount = (typeof v === 'number') ? v : Number(parseAmount(v)) || 0;
           return sum + amount;
        }, 0), // [NEW] Optimized Query Field
        timestamp: date(firstRow, IDX.timestamp), // Use first row's timestamp
        rowCount: rows.length,
        metadata: {
          backfilled: true,
          syncedAt: new Date().toISOString()
        },
        createdAt: new Date().toISOString()
    };
    
    writes.push({
      update: {
        name: `${databaseRoot}/${parentPath}`,
        fields: firestoreFields_(parentData)
      }
    });
    
    // B. Parse & Write Line Items
    rows.forEach((r, idx) => {
      const num = (idx) => {
         const v = val(r, idx); 
         return (typeof v === 'number') ? v : Number(parseAmount(v));
      };
      // helper for dates in row
      const d = (idx) => date(r, idx) || '';

      const lineId = `row_${idx}`;
      const linePath = `${parentPath}/lines/${lineId}`;
      
      const lineData = {
         beneficiary: str(r, IDX.beneficiary),
         accountHolder: str(r, IDX.accountHolder),
         Ref: str(r, IDX.advRef), // Mapping Adv Ref to Ref
         project: str(r, IDX.project),
         team: str(r, IDX.team),
         submitter: str(r, IDX.submitter),
         total: num(IDX.total),
         designation: str(r, IDX.designation),
         
         fuel: { from: d(IDX.fuelFrom), to: d(IDX.fuelTo), amount: num(IDX.fuelAmt) },
         da: { from: d(IDX.daFrom), to: d(IDX.daTo), amount: num(IDX.daAmt) },
         car: { from: d(IDX.carFrom), to: d(IDX.carTo), vehicleNumber: str(r, IDX.carNum), amount: num(IDX.carAmt) },
         air: { from: d(IDX.airFrom), to: d(IDX.airTo), amount: num(IDX.airAmt) },
         transport: { from: d(IDX.transFrom), to: d(IDX.transTo), amount: num(IDX.transAmt) },
         misc: { from: d(IDX.miscFrom), to: d(IDX.miscTo), amount: num(IDX.miscAmt) },
         
         timestamp: parentData.timestamp, // Legacy/Consistency
         
         // Extra fields
         mob: str(r, IDX.mob),
         displayName: str(r, IDX.display),
         whCharges: num(IDX.wh),
         remarks: str(r, IDX.remarks),
         approvalDate: d(IDX.approvalDate),
         approvedBy: str(r, IDX.approvedBy),
         paidAmt: num(IDX.paidAmt),
         transferBy: str(r, IDX.transferBy),
         financeRemarks: str(r, IDX.finRemarks),
         dateP: d(IDX.dateP),
         claimBalAmt: num(IDX.claimBal),
         balanceAmt: num(IDX.balAmt),
         vchType: str(r, IDX.vchType),
         jv: str(r, IDX.jv),
         transferTo: str(r, IDX.transferTo),
         mode: str(r, IDX.mode)
      };
      
      writes.push({
        update: {
          name: `${databaseRoot}/${linePath}`,
          fields: firestoreFields_(lineData)
        }
      });
    });
  });

  console.log(`Backfill: Prepared ${writes.length} writes (Parents + Lines) for ${Object.keys(groups).length} submissions.`);

  // 3. Execute Batch
  // Check if executeFirestoreBatch is available (from firestore.gs)
  if (typeof executeFirestoreBatch === 'function') {
      executeFirestoreBatch(writes);
  } else {
      // Fallback if firestore.gs not updated or reachable? 
      // Should not happen if part of same project, but for safety:
      console.warn('executeFirestoreBatch not found, using local batch logic??');
      // If we are here, we can reuse firestore.gs helpers if exposed, else rewrite.
      // Assuming firestore.gs is loaded.
      throw new Error('executeFirestoreBatch helper missing. Ensure firestore.gs is updated.');
  }
  
  return { count: writes.length };
}
