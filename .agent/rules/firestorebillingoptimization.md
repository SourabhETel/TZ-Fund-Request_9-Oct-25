---
trigger: always_on
---

# FIRESTORE BILLING OPTIMIZATION (STRICT)

## 1. THE GOLDEN RULE: No Client-Side Filtering
* **FORBIDDEN:** Never fetch a collection list (e.g., `getDocuments()`) and then use `.filter()` in JavaScript. This causes massive billing spikes (1 read per document scanned).
* **REQUIRED:** You must use **Server-Side Filtering** via the Firestore REST API (`runQuery`).
* **GOAL:** If we need 10 "Available" cars, we must pay for **10 reads**, not the 1000 cars in the database.

## 2. REQUIRED CODE PATTERN
When writing functions to query `cart_plan` or any collection, you must use this specific `runQuery` structure based on `firestore.gs`:

```javascript
function queryCheaply(collection, statusValue) {
  const projectId = getFirestoreProjectId_();
  const token = getFirestoreAccessToken_();
  const url = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents:runQuery`;

  const payload = {
    structuredQuery: {
      from: [{ collectionId: collection }],
      where: {
        fieldFilter: {
          field: { fieldPath: 'status' }, // Filter field
          op: 'EQUAL', 
          value: { stringValue: statusValue } // e.g. 'AVAILABLE'
        }
      },
      limit: 20 // ALWAYS include a limit for safety
    }
  };

  // ... standard UrlFetchApp call ...
}