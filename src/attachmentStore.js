const DATABASE_NAME = 'thunder-team-attachments';
const STORE_NAME = 'attachments';
const DATABASE_VERSION = 1;

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = window.indexedDB.open(DATABASE_NAME, DATABASE_VERSION);
    request.onupgradeneeded = () => {
      const database = request.result;
      if (!database.objectStoreNames.contains(STORE_NAME)) {
        database.createObjectStore(STORE_NAME, { keyPath: 'id' });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error('Unable to open local document storage.'));
  });
}

export async function listAttachments() {
  const database = await openDatabase();
  return new Promise((resolve, reject) => {
    const request = database.transaction(STORE_NAME, 'readonly').objectStore(STORE_NAME).getAll();
    request.onsuccess = () => {
      database.close();
      resolve(request.result.map(({ file, ...metadata }) => metadata));
    };
    request.onerror = () => {
      database.close();
      reject(request.error || new Error('Unable to read saved documents.'));
    };
  });
}

export async function saveAttachment(attachment) {
  const database = await openDatabase();
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(STORE_NAME, 'readwrite');
    transaction.objectStore(STORE_NAME).put(attachment);
    transaction.oncomplete = () => {
      database.close();
      const { file, ...metadata } = attachment;
      resolve(metadata);
    };
    transaction.onerror = () => {
      database.close();
      reject(transaction.error || new Error('Unable to save the document on this device.'));
    };
  });
}

export async function getAttachment(id) {
  const database = await openDatabase();
  return new Promise((resolve, reject) => {
    const request = database.transaction(STORE_NAME, 'readonly').objectStore(STORE_NAME).get(id);
    request.onsuccess = () => {
      database.close();
      resolve(request.result || null);
    };
    request.onerror = () => {
      database.close();
      reject(request.error || new Error('Unable to open the saved document.'));
    };
  });
}
