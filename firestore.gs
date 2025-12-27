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
  const rows = Array.isArray(submissionPayload.rows)
    ? submissionPayload.rows.map(row => normalizeValueForFirestore_(row))
    : [];
  const payload = {
    submissionId: submissionPayload.submissionId || (opts && opts.documentId) || '',
    project: submissionPayload.project || '',
    submitter: submissionPayload.submitter || '',
    timestamp: submissionPayload.timestamp instanceof Date
      ? submissionPayload.timestamp.toISOString()
      : submissionPayload.timestamp || new Date().toISOString(),
    rowCount: rows.length,
    rows: rows,
    metadata: normalizeValueForFirestore_(submissionPayload.metadata || {}),
    createdAt: new Date().toISOString()
  };
  const documentId = submissionPayload.submissionId || (opts && opts.documentId);
  const firestoreApp = getConfigFirestoreClient_();
  if (firestoreApp) {
    const args = [collection, payload];
    if (documentId) args.push(documentId);
    return firestoreApp.createDocument.apply(firestoreApp, args);
  }
  return createFirestoreDocument_(collection, payload, documentId ? { documentId } : undefined);
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
