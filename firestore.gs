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
 * Persist a vehicle transaction (and its line items) to Firestore hierarchically.
 * Structure: cartp_plan/{transactionId} -> cartp_lines/{itemId}
 * 
 * @param {Object} payload 
 *  {
 *    transactionId: string, // Optional, auto-generated if missing
 *    timestamp: Date|string,
 *    type: 'ASSIGN'|'RELEASE',
 *    submitter: string,
 *    rows: Array<Object> // The line items
 *  }
 */
function persistCarTransactionToFirestore(payload) {
  if (!payload || !Array.isArray(payload.rows)) {
    throw new Error('Invalid vehicle transaction payload.');
  }

  const projectId = getFirestoreProjectId_();
  const databaseRoot = `projects/${projectId}/databases/(default)/documents`;
  
  // 1. Prepare Parent Document
  const collection = 'cartp_plan';
  const transactionId = payload.transactionId || `tx_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  const parentPath = `${collection}/${transactionId}`;
  
  const parentData = {
     transactionId: transactionId,
     type: payload.type || 'UNKNOWN',
     timestamp: payload.timestamp instanceof Date ? payload.timestamp.toISOString() : (payload.timestamp || new Date().toISOString()),
     submitter: payload.submitter || '',
     rowCount: payload.rows.length,
     project: payload.project || '', // top-level summary
     team: payload.team || '',       // top-level summary
     createdAt: new Date().toISOString()
  };

  const writes = [];

  // Write Parent
  writes.push({
    update: {
      name: `${databaseRoot}/${parentPath}`,
      fields: firestoreFields_(parentData)
    }
  });

  // 2. Prepare Child Documents (cartp_lines)
  payload.rows.forEach((row, index) => {
     const lineId = row.ref || `row_${index}_${Math.random().toString(36).substr(2, 5)}`;
     // Ensure we don't have slashes in IDs
     const safeLineId = String(lineId).replace(/\//g, '_'); 
     const linePath = `${parentPath}/cartp_lines/${safeLineId}`;
     
     // Ensure timestamp is ready for Firestore
     const ts = row.timestamp instanceof Date ? row.timestamp.toISOString() : (row.timestamp || parentData.timestamp);
     
     const lineData = Object.assign({}, row, {
        timestamp: ts,
        parentId: transactionId // back-reference for convenience
     });

     writes.push({
       update: {
         name: `${databaseRoot}/${linePath}`,
         fields: firestoreFields_(lineData)
       }
     });
  });

  // 3. Execute Batch
  return executeFirestoreBatch(writes);
}

/**
 * Retrieve all car events using a Collection Group Query on 'cartp_lines'.
 * effectively flattens the hierarchy for the application view.
 */
function getCarEventsFromFirestore() {
  const projectId = getFirestoreProjectId_();
  const token = getFirestoreAccessToken_();
  
  // Use runQuery to perform a Collection Group Query
  const url = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents:runQuery`;
  
  const queryPayload = {
    structuredQuery: {
      from: [{ collectionId: 'cartp_lines', allDescendants: true }],
      orderBy: [{ field: { fieldPath: 'timestamp' }, direction: 'ASCENDING' }]
    }
  };

  // Note: runQuery returns a stream of results, potentially paginated or just a list
  // The REST API for runQuery usually returns a JSON list of objects, usually [{document:..., readTime:...}]
  
  let documents = [];
  
  try {
      const response = UrlFetchApp.fetch(url, {
        method: 'post',
        headers: {
          Authorization: `Bearer ${token}`
        },
        contentType: 'application/json',
        payload: JSON.stringify(queryPayload),
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() !== 200) {
        console.error('Firestore runQuery failed', response.getContentText());
        return [];
      }
      
      const json = JSON.parse(response.getContentText());
      // json is array of results. Each result has 'document'
      if (Array.isArray(json)) {
          json.forEach(item => {
             if (item.document) {
                 documents.push(item.document);
             }
          });
      }
  } catch (e) {
      console.error('getCarEventsFromFirestore error:', e);
      return [];
  }
  
  // Parse documents
  return documents.map(doc => {
    return parseFirestoreMap_(doc.fields);
  });
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

