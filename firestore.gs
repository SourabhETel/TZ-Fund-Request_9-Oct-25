/**
 * Firestore helpers for Apps Script submissions.
 * 
 * Setup:
 *   - Create a service account that has Firestore access and download the JSON key.
 *   - In the Apps Script project settings, add script properties (or use a `config.gs` helper):
 *       FIRESTORE_PROJECT_ID: your Firestore project ID (e.g. "my-project")
 *       FIRESTORE_COLLECTION_NAME: (optional) Firestore collection name (default: "submissions")
 *       FIRESTORE_SERVICE_ACCOUNT_JSON: raw JSON contents of the service account key
 *         OR FIRESTORE_SERVICE_ACCOUNT_JSON_BASE64: base64-encoded JSON to avoid newline issues.
 *   - If using FirestoreApp (https://github.com/googledrive/FirestoreApp), supply a `getFirestore()` helper
 *     (matching the `config.gs` snippet shared) and the helper below will reuse it automatically.
 *   - Enable the Firestore REST API for the project.
 *
 * Once configured, call `persistSubmissionToFirestore()` with your submission metadata.
 */

const FIRESTORE_PROJECT_ID_PROP = 'FIRESTORE_PROJECT_ID';
const FIRESTORE_COLLECTION_PROP = 'FIRESTORE_COLLECTION_NAME';
const FIRESTORE_SA_JSON_PROP = 'FIRESTORE_SERVICE_ACCOUNT_JSON';
const FIRESTORE_SA_JSON_BASE64_PROP = 'FIRESTORE_SERVICE_ACCOUNT_JSON_BASE64';
const FIRESTORE_TOKEN_CACHE_KEY = 'FIRESTORE_ACCESS_TOKEN';

function getFirestoreProjectId_() {
  const projectId = PropertiesService.getScriptProperties().getProperty(FIRESTORE_PROJECT_ID_PROP);
  if (!projectId) {
    throw new Error(`Please configure the ${FIRESTORE_PROJECT_ID_PROP} script property before writing to Firestore.`);
  }
  return projectId;
}

function getFirestoreCollectionName_() {
  const collection = PropertiesService.getScriptProperties().getProperty(FIRESTORE_COLLECTION_PROP);
  return collection ? collection.trim() : 'submissions';
}

function getFirestoreServiceAccount_() {
  const props = PropertiesService.getScriptProperties();
  let raw = props.getProperty(FIRESTORE_SA_JSON_PROP);
  if (!raw) {
    const base64 = props.getProperty(FIRESTORE_SA_JSON_BASE64_PROP);
    if (!base64) {
      throw new Error(`Supply service account credentials via ${FIRESTORE_SA_JSON_PROP} or ${FIRESTORE_SA_JSON_BASE64_PROP}.`);
    }
    raw = Utilities.newBlob(Utilities.base64Decode(base64)).getDataAsString('utf-8');
  }

  try {
    return JSON.parse(raw);
  } catch (err) {
    throw new Error('Unable to parse the Firestore service account JSON. Verify the script property value.');
  }
}

function getFirestoreAccessToken_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(FIRESTORE_TOKEN_CACHE_KEY);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (parsed.accessToken && parsed.expiresAt && Number(parsed.expiresAt) > Date.now()) {
        return parsed.accessToken;
      }
    } catch (_err) {
      // Fall through and fetch a new token.
    }
  }

  const credentials = getFirestoreServiceAccount_();
  const now = Math.floor(Date.now() / 1000);
  const payload = {
    iss: credentials.client_email,
    scope: 'https://www.googleapis.com/auth/datastore https://www.googleapis.com/auth/cloud-platform',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now
  };
  const jwt = buildJwt_({ alg: 'RS256', typ: 'JWT' }, payload, credentials.private_key);
  const response = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`Firestore token request failed (${response.getResponseCode()}): ${response.getContentText()}`);
  }

  const tokenData = JSON.parse(response.getContentText());
  if (!tokenData.access_token) {
    throw new Error('Firestore token response did not include an access_token.');
  }
  const expiresInMs = (tokenData.expires_in || 3600) * 1000;
  cache.put(FIRESTORE_TOKEN_CACHE_KEY, JSON.stringify({
    accessToken: tokenData.access_token,
    expiresAt: Date.now() + expiresInMs
  }), Math.floor((expiresInMs - 60 * 1000) / 1000));
  return tokenData.access_token;
}

function buildJwt_(header, payload, privateKey) {
  const encodedHeader = base64UrlEncode_(JSON.stringify(header));
  const encodedPayload = base64UrlEncode_(JSON.stringify(payload));
  const signingInput = `${encodedHeader}.${encodedPayload}`;
  const signatureBytes = Utilities.computeRsaSha256Signature(signingInput, privateKey);
  const signature = base64UrlEncode_(signatureBytes);
  return `${encodedHeader}.${encodedPayload}.${signature}`;
}

function base64UrlEncode_(value) {
  const bytes = typeof value === 'string' ? Utilities.newBlob(value).getBytes() : value;
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function firestoreFields_(obj) {
  const fields = {};
  Object.keys(obj || {}).forEach(key => {
    const value = obj[key];
    if (value === undefined) return;
    fields[key] = firestoreValue_(value);
  });
  return fields;
}

function firestoreValue_(value) {
  if (value === null || value === undefined) {
    return { nullValue: null };
  }
  if (value instanceof Date) {
    return { timestampValue: value.toISOString() };
  }
  if (Array.isArray(value)) {
    return { arrayValue: { values: value.map(v => firestoreValue_(v)) } };
  }
  if (typeof value === 'object') {
    return { mapValue: { fields: firestoreFields_(value) } };
  }
  if (typeof value === 'boolean') {
    return { booleanValue: value };
  }
  if (typeof value === 'number') {
    return Number.isInteger(value)
      ? { integerValue: String(value) }
      : { doubleValue: value };
  }
  return { stringValue: String(value) };
}

function createFirestoreDocument_(collectionPath, data, opts) {
  if (!collectionPath) {
    throw new Error('Firestore collection path is required.');
  }
  const projectId = getFirestoreProjectId_();
  const baseUrl = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/${collectionPath}`;
  const url = opts && opts.documentId ? `${baseUrl}?documentId=${encodeURIComponent(opts.documentId)}` : baseUrl;
  const payload = { fields: firestoreFields_(data) };
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      Authorization: `Bearer ${getFirestoreAccessToken_()}`
    },
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  if (response.getResponseCode() >= 300) {
    throw new Error(`Firestore document create failed (${response.getResponseCode()}): ${response.getContentText()}`);
  }
  return JSON.parse(response.getContentText());
}

function normalizeValueForFirestore_(value) {
  if (value === undefined) {
    return undefined;
  }
  if (value === null) {
    return null;
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (Array.isArray(value)) {
    return value.map(v => normalizeValueForFirestore_(v));
  }
  if (typeof value === 'object') {
    const normalized = {};
    Object.keys(value).forEach(key => {
      const normalizedValue = normalizeValueForFirestore_(value[key]);
      if (normalizedValue !== undefined) {
        normalized[key] = normalizedValue;
      }
    });
    return normalized;
  }
  return value;
}

function persistSubmissionToFirestore(submissionPayload, opts) {
  if (!submissionPayload || typeof submissionPayload !== 'object') {
    throw new Error('A submission payload object is required to persist to Firestore.');
  }

  const collection = (opts && opts.collection) || getFirestoreCollectionName_();
  const submissionId = submissionPayload.submissionId || (opts && opts.documentId);
  
  if (!submissionId) {
    throw new Error('Submission ID is required for hierarchical storage.');
  }
  
  const projectId = getFirestoreProjectId_();
  const databaseRoot = `projects/${projectId}/databases/(default)/documents`;
  
  // Prepare Line Documents (Subcollection: lines)
  const rows = Array.isArray(submissionPayload.rows) ? submissionPayload.rows : [];
  
  // Calculate Summaries for Parent
  const totalAmount = rows.reduce((sum, r) => sum + (Number(r.total) || 0), 0);
  const firstRow = rows[0] || {};
  // Check payload from code.gs; it maps teamName/team.
  const team = firstRow.team || firstRow.teamName || ''; 

  // 1. Prepare Parent Document (Submission)
  const parentPath = `${collection}/${submissionId}`;
  
  // Extract strictly parent-level metadata
  const parentData = {
    submissionId: submissionId,
    project: submissionPayload.project || '',
    submitter: submissionPayload.submitter || '',
    team: team, // [NEW] Optimized Query Field
    totalAmount: totalAmount, // [NEW] Optimized Query Field
    timestamp: submissionPayload.timestamp instanceof Date
      ? submissionPayload.timestamp.toISOString()
      : submissionPayload.timestamp || new Date().toISOString(),
    rowCount: rows.length, // Corrected to use hoisted rows
    metadata: normalizeValueForFirestore_(submissionPayload.metadata || {}),
    createdAt: new Date().toISOString()
  };
  
  const writes = [];
  
  // Write Parent
  writes.push({
    update: {
      name: `${databaseRoot}/${parentPath}`,
      fields: firestoreFields_(parentData)
    },
    // updateMask could be added if we want to merge, but overwrite seems appropriate for new submission
  });
  
  // 2. Prepare Line Documents (Subcollection: lines)
  // rows array is already defined above
  
  rows.forEach((row, index) => {
    // Generate Row ID: row_0, row_1, etc. or use a provided ID if available
    const rowId = `row_${index}`; 
    const linePath = `${parentPath}/lines/${rowId}`;
    
    // Normalize and Ensure Schema Fields
    // User Requirement: beneficiary, accountHolder, Ref, project, team, submitter, total, component maps
    
    const lineData = {
      beneficiary: row.beneficiary || '',
      accountHolder: row.accountHolder || '',
      Ref: row.Ref || row.advRef || '', // Mapping Ref
      project: row.project || row.projectName || '', // Handle varied aliases
      team: row.team || row.teamName || '',
      submitter: row.submitter || '',
      total: row.total || 0,
      designation: row.designation || '',
      
      // Component Maps
      fuel: normalizeValueForFirestore_(row.fuel || {}),
      da: normalizeValueForFirestore_(row.da || {}),
      car: normalizeValueForFirestore_(row.car || {}), // 'Vehicle' category as 'car'
      air: normalizeValueForFirestore_(row.air || {}),
      transport: normalizeValueForFirestore_(row.transport || {}),
      misc: normalizeValueForFirestore_(row.misc || {}),
      
      // Extra Fields
      mob: row.mob || '',
      displayName: row.displayName || '',
      whCharges: row.whCharges || 0,
      remarks: row.remarks || '',
      approvalDate: row.approvalDate || null,
      approvedBy: row.approvedBy || '',
      paidAmt: row.paidAmt || 0,
      transferBy: row.transferBy || '',
      financeRemarks: row.financeRemarks || '',
      
      timestamp: parentData.timestamp // Ensure document creation time/submission time is recorded
    };

    writes.push({
      update: {
        name: `${databaseRoot}/${linePath}`,
        fields: firestoreFields_(lineData)
      }
    });
  });

  // 3. Execute Batch Write
  return executeFirestoreBatch(writes);
}

/**
 * Executes a batch of writes using the Firestore REST API `commit` endpoint.
 */
function executeFirestoreBatch(writes) {
  if (!writes || writes.length === 0) return;
  
  const projectId = getFirestoreProjectId_();
  const url = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents:commit`;
  const token = getFirestoreAccessToken_();
  
  // Firestore limits batch to 500 writes. Simple chunking if needed, though submissions are unlikely to exceed 500 rows.
  // maximizing safety:
  const CHUNK_SIZE = 500;
  
  for (let i = 0; i < writes.length; i += CHUNK_SIZE) {
    const chunk = writes.slice(i, i + CHUNK_SIZE);
    const payload = { writes: chunk };
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        Authorization: `Bearer ${token}`
      },
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() >= 300) {
      throw new Error(`Firestore batch commit failed (${response.getResponseCode()}): ${response.getContentText()}`);
    }
  }
  
  return { count: writes.length };
}

function getConfigFirestoreClient_() {
  if (typeof getFirestore !== 'function') return null;
  if (typeof FirestoreApp === 'undefined') {
    console.warn('FirestoreApp library not available; skipping config.getFirestore()');
    return null;
  }
  try {
    const client = getFirestore();
    if (client && typeof client.createDocument === 'function') {
      return client;
    }
  } catch (err) {
    console.warn('Unable to initialize FirestoreApp client from config:', err);
  }
  return null;
}

/**
 * Persist a vehicle event (Assignment/Release) to the 'cartp_plan' Firestore collection.
 * This replaces writing to the CarT_P Google Sheet.
 * 
 * @param {Object} eventData - The flattened event object (similar to what was written to the sheet).
 */
/**
 * Fetch a specific vehicle document by Vehicle Number (ID).
 * Strict 1 Read Cost.
 */
function getVehicleDocument(vehicleNumber) {
  const projectId = getFirestoreProjectId_();
  const token = getFirestoreAccessToken_();
  const safeId = String(vehicleNumber).replace(/\//g, '_').trim();
  const url = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/cart_plan/${safeId}`;

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 404) return null;
    if (response.getResponseCode() !== 200) {
      throw new Error(`Firestore get failed: ${response.getContentText()}`);
    }
    
    const json = JSON.parse(response.getContentText());
    if (!json.fields) return null; // Should have fields
    
    return parseFirestoreMap_(json.fields);
  } catch (e) {
    console.error('getVehicleDocument error:', e);
    return null;
  }
}

/**
 * Assign crew to a vehicle.
 * Rule A: If AVAILABLE -> Status=IN_USE, Create Assignment.
 *         If IN_USE -> Append to activeCrew.
 */
/**
 * Assign crew to a vehicle.
 * Rule A: If AVAILABLE -> Status=IN_USE, Create Assignment.
 *         If IN_USE -> Append to activeCrew.
 * 
 * @param {string} vehicleNumber
 * @param {Object} assignmentData { ref, project, team, submitter }
 * @param {Array} newCrewList [{name, role, joinedAt}]
 * @param {Object} vehicleMetadata { make, model, category, owner, usageType } (Optional, for new cars)
 */
function assignCrewToVehicle(vehicleNumber, assignmentData, newCrewList, vehicleMetadata) {
  const currentDoc = getVehicleDocument(vehicleNumber);
  const projectId = getFirestoreProjectId_();
  const databaseRoot = `projects/${projectId}/databases/(default)/documents`;
  const safeId = String(vehicleNumber).replace(/\//g, '_').trim();
  const docPath = `cart_plan/${safeId}`;
  
  // Prepare base object if new (though docs should ideally exist)
  let docData = currentDoc || {
    vehicleNumber: vehicleNumber,
    status: 'AVAILABLE',
    currentAssignment: null,
    recentHistory: []
  };

  // If new doc or updating metadata
  if (vehicleMetadata) {
     if (vehicleMetadata.make) docData.make = vehicleMetadata.make;
     if (vehicleMetadata.model) docData.model = vehicleMetadata.model;
     if (vehicleMetadata.category) docData.category = vehicleMetadata.category;
     if (vehicleMetadata.owner) docData.owner = vehicleMetadata.owner;
     if (vehicleMetadata.usageType) docData.usageType = vehicleMetadata.usageType;
  }

  const nowStr = new Date().toISOString();

  // Logic: Merge Crew
  if (docData.status === 'IN_USE' && docData.currentAssignment) {
     // Append new crew members
     // Avoid duplicates based on name/role?
     const existingCrew = docData.currentAssignment.activeCrew || [];
     newCrewList.forEach(nc => {
        // Simple duplicate check by name
        if (!existingCrew.find(c => c.name === nc.name)) {
           existingCrew.push(nc);
        }
     });
     docData.currentAssignment.activeCrew = existingCrew;
  } else {
     // Status is AVAILABLE (or DISCARDED, but assuming we can re-assign)
     docData.status = 'IN_USE';
     docData.currentAssignment = {
        ref: assignmentData.ref,
        project: assignmentData.project,
        team: assignmentData.team,
        startTime: nowStr,
        submitter: assignmentData.submitter,
        activeCrew: newCrewList
     };
  }

  // Write back (Update/Overwrite)
  console.log('[FIRESTORE] assignCrewToVehicle - DocPath:', docPath);
  console.log('[FIRESTORE] assignCrewToVehicle - Is New/Metadata?', !!vehicleMetadata);
  console.log('[FIRESTORE] assignCrewToVehicle - Final DocData:', JSON.stringify(docData));

  const fields = firestoreFields_(docData);
  const write = {
    update: {
      name: `${databaseRoot}/${docPath}`,
      fields: fields
    }
  };
  
  return executeFirestoreBatch([write]);
}

/**
 * Handle vehicle release with "Last Man Standing" logic.
 * Rule B: Remove users. If crew empty -> Status=AVAILABLE.
 */
function updateVehicleAsLastManStanding(vehicleNumber, crewNamesToRemove, releaseData) {
  const docData = getVehicleDocument(vehicleNumber);
  if (!docData) throw new Error(`Vehicle ${vehicleNumber} not found in Firestore.`);
  
  if (docData.status !== 'IN_USE' || !docData.currentAssignment) {
     console.warn(`Vehicle ${vehicleNumber} is not IN_USE. Release ignored.`);
     return;
  }

  const activeCrew = docData.currentAssignment.activeCrew || [];
  const initialCount = activeCrew.length;
  
  // Filter out removed users
  const remainingCrew = activeCrew.filter(c => !crewNamesToRemove.includes(c.name));
  
  if (remainingCrew.length === 0) {
     // FULL RELEASE
     docData.status = 'AVAILABLE';
     
     // Move assignment to history
     const historyEntry = {
        action: 'RELEASE',
        date: new Date().toISOString(),
        ref: docData.currentAssignment.ref,
        project: docData.currentAssignment.project,
        team: docData.currentAssignment.team,
        remarks: releaseData.remarks || '',
        rating: releaseData.rating || 0,
        releasedCrew: activeCrew // The crew that was just cleared
     };
     
     const history = docData.recentHistory || [];
     history.unshift(historyEntry);
     docData.recentHistory = history.slice(0, 5); // Keep last 5
     
     docData.currentAssignment = null; // Clear assignment
     
  } else {
     // PARTIAL RELEASE
     // Status remains IN_USE
     docData.currentAssignment.activeCrew = remainingCrew;
     
     // Log partial release event? Optional. User said "Log a 'Partial Release' entry in recentHistory".
     const partialEntry = {
        action: 'PARTIAL_RELEASE',
        date: new Date().toISOString(),
        removedCrew: crewNamesToRemove,
        remainingCount: remainingCrew.length
     };
     const history = docData.recentHistory || [];
     history.unshift(partialEntry);
     docData.recentHistory = history.slice(0, 5);
  }

  // Write back
  const projectId = getFirestoreProjectId_();
  const databaseRoot = `projects/${projectId}/databases/(default)/documents`;
  const safeId = String(vehicleNumber).replace(/\//g, '_').trim();
  const docPath = `cart_plan/${safeId}`;

  const fields = firestoreFields_(docData);
  const write = {
    update: {
      name: `${databaseRoot}/${docPath}`,
      fields: fields
    }
  };
  
  return executeFirestoreBatch([write]);
}


/**
 * Helper: Run a cheap query to get available or active cars.
 * Limit is mandatory.
 */
function queryVehiclesByStatus(status, limit) {
  const projectId = getFirestoreProjectId_();
  const token = getFirestoreAccessToken_();
  const url = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents:runQuery`;

  const payload = {
    structuredQuery: {
      from: [{ collectionId: 'cart_plan' }],
      where: {
        fieldFilter: {
          field: { fieldPath: 'status' },
          op: 'EQUAL', 
          value: { stringValue: status }
        }
      },
      limit: limit || 20
    }
  };

  try {
     const response = UrlFetchApp.fetch(url, {
       method: 'post',
       headers: { Authorization: `Bearer ${token}` },
       contentType: 'application/json',
       payload: JSON.stringify(payload),
       muteHttpExceptions: true
     });
     
     if (response.getResponseCode() !== 200) return [];
     
     const json = JSON.parse(response.getContentText());
     // json is [{document: ...}]
     if (!Array.isArray(json)) return [];
     
     return json.map(item => item.document ? parseFirestoreMap_(item.document.fields) : null).filter(Boolean);
  } catch (e) {
     console.error('queryVehiclesByStatus error:', e);
     return [];
  }
}

function parseFirestoreValue_(valueObj) {
  if (!valueObj) return undefined;
  if ('nullValue' in valueObj) return null;
  if ('stringValue' in valueObj) return valueObj.stringValue;
  if ('integerValue' in valueObj) return Number(valueObj.integerValue);
  if ('doubleValue' in valueObj) return Number(valueObj.doubleValue);
  if ('booleanValue' in valueObj) return valueObj.booleanValue;
  if ('timestampValue' in valueObj) return new Date(valueObj.timestampValue);
  if ('mapValue' in valueObj) return parseFirestoreMap_(valueObj.mapValue.fields);
  if ('arrayValue' in valueObj) return (valueObj.arrayValue.values || []).map(parseFirestoreValue_);
  return undefined;
}

function parseFirestoreMap_(fields) {
    const obj = {};
    if (!fields) return obj;
    for (const key in fields) {
        obj[key] = parseFirestoreValue_(fields[key]);
    }
    return obj;
}

/**
 * Change the responsible beneficiary for a vehicle.
 * Rules:
 * 1. Downgrade any current 'Responsible' to 'Beneficiary'.
 * 2. Promote 'newRespName' to 'Responsible'.
 * 3. If 'newRespName' is not in crew, add them.
 */
function changeResponsibleBeneficiaryInFirestore(vehicleNumber, newRespName) {
  const docData = getVehicleDocument(vehicleNumber);
  if (!docData) throw new Error(`Vehicle ${vehicleNumber} not found.`);
  
  if (docData.status !== 'IN_USE' || !docData.currentAssignment) {
     throw new Error(`Vehicle ${vehicleNumber} is not IN_USE.`);
  }

  const activeCrew = docData.currentAssignment.activeCrew || [];
  let found = false;

  // Downgrade existing and Promote new
  activeCrew.forEach(member => {
     if (member.role === 'Responsible') {
        member.role = 'Beneficiary';
     }
     if (member.name === newRespName) {
        member.role = 'Responsible';
        found = true;
     }
  });

  // If not found, add
  if (!found) {
     activeCrew.push({
        name: newRespName,
        role: 'Responsible',
        joinedAt: new Date().toISOString()
     });
  }

  docData.currentAssignment.activeCrew = activeCrew;

  // Write back
  const projectId = getFirestoreProjectId_();
  const databaseRoot = `projects/${projectId}/databases/(default)/documents`;
  const safeId = String(vehicleNumber).replace(/\//g, '_').trim();
  const docPath = `cart_plan/${safeId}`;

  const fields = firestoreFields_(docData);
  const write = {
    update: {
      name: `${databaseRoot}/${docPath}`,
      fields: fields
    }
  };
  
  return executeFirestoreBatch([write]);
}

/**
 * Register a new vehicle in Firestore with status AVAILABLE (or RELEASE).
 */
function registerNewVehicleInFirestore(vehicleNumber, metadata, entryData) {
  const projectId = getFirestoreProjectId_();
  const databaseRoot = `projects/${projectId}/databases/(default)/documents`;
  const safeId = String(vehicleNumber).replace(/\//g, '_').trim();
  const docPath = `cart_plan/${safeId}`;

  const nowStr = new Date().toISOString();

  // Construct Base Doc
  const docData = {
    vehicleNumber: vehicleNumber,
    status: 'AVAILABLE', // Default to available
    make: metadata.make || '',
    model: metadata.model || '',
    category: metadata.category || '',
    owner: metadata.owner || '',
    usageType: metadata.usageType || '',
    currentAssignment: null,
    recentHistory: []
  };

  // Add initial history entry for creation/release
  const historyEntry = {
    action: 'NEW_REGISTRATION',
    date: nowStr,
    project: entryData.project || '',
    team: entryData.team || '',
    remarks: entryData.remarks || '',
    submitter: entryData.submitter || ''
  };
  docData.recentHistory.push(historyEntry);

  const fields = firestoreFields_(docData);
  const write = {
    update: { // Use update with upsert semantics (in this lib, update creates if missing often, or check?)
              // Actually executeFirestoreBatch 'update' usually requires existence?
              // Standard behavior: 'update' performs PATCH. 'transform' can operate.
              // To be safe, if we want to Create or Overwrite, usually we use 'currentDocument' constraints.
              // But here, let's assume we want to SET.
              // The library I'm using here seems to use "update" key for Writes.
      name: `${databaseRoot}/${docPath}`,
      fields: fields
    }
  };
  
  return executeFirestoreBatch([write]);
}

