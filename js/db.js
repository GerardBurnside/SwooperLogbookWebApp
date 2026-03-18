// IndexedDB abstraction layer for Swooper Logbook
// Manages jumps, canopies, harnesses, and locations stores.
// Settings and sync metadata remain in localStorage.

/**
 * Ensure storage access when the Storage Access API is available (Safari/iOS).
 * Requests access if we don't have it. Does not reject so that callers can
 * still attempt IDB open; requestStorageAccess() may require a user gesture.
 */
function ensureStorageAccess() {
    if (typeof document.hasStorageAccess !== 'function' || typeof document.requestStorageAccess !== 'function') {
        return Promise.resolve();
    }
    return document.hasStorageAccess()
        .then((has) => (has ? undefined : document.requestStorageAccess()))
        .catch(() => { /* e.g. needs user gesture; caller will show banner and retry on tap */ });
}

const DB = (() => {
    const DB_NAME = 'swooper-logbook';
    const DB_VERSION = 1;
    let _db = null;

    async function open() {
        if (_db) return _db;

        await ensureStorageAccess();

        return new Promise((resolve, reject) => {
            const request = indexedDB.open(DB_NAME, DB_VERSION);

            request.onupgradeneeded = (event) => {
                const db = event.target.result;

                // Jumps store with indexes for common query patterns
                if (!db.objectStoreNames.contains('jumps')) {
                    const jumpsStore = db.createObjectStore('jumps', { keyPath: 'id' });
                    jumpsStore.createIndex('by-date', 'date', { unique: false });
                    jumpsStore.createIndex('by-timestamp', 'timestamp', { unique: false });
                    jumpsStore.createIndex('by-equipment', 'equipment', { unique: false });
                    jumpsStore.createIndex('by-jumpNumber', 'jumpNumber', { unique: false });
                }

                // Equipment stores — small collections, loaded in full
                if (!db.objectStoreNames.contains('canopies')) {
                    db.createObjectStore('canopies', { keyPath: 'id' });
                }
                if (!db.objectStoreNames.contains('harnesses')) {
                    db.createObjectStore('harnesses', { keyPath: 'id' });
                }
                if (!db.objectStoreNames.contains('locations')) {
                    db.createObjectStore('locations', { keyPath: 'id' });
                }
            };

            request.onsuccess = (event) => {
                _db = event.target.result;
                resolve(_db);
            };

            request.onerror = (event) => {
                console.error('IndexedDB open failed:', event.target.error);
                reject(event.target.error);
            };
        });
    }

    // ── Generic helpers ─────────────────────────────────────────────────

    function _getAll(storeName) {
        return new Promise((resolve, reject) => {
            const tx = _db.transaction(storeName, 'readonly');
            const store = tx.objectStore(storeName);
            const req = store.getAll();
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
        });
    }

    function _putAll(storeName, items) {
        return new Promise((resolve, reject) => {
            const tx = _db.transaction(storeName, 'readwrite');
            const store = tx.objectStore(storeName);
            for (const item of items) {
                store.put(item);
            }
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
        });
    }

    function _clear(storeName) {
        return new Promise((resolve, reject) => {
            const tx = _db.transaction(storeName, 'readwrite');
            const store = tx.objectStore(storeName);
            const req = store.clear();
            req.onsuccess = () => resolve();
            req.onerror = () => reject(req.error);
        });
    }

    // ── Jumps ───────────────────────────────────────────────────────────

    async function getAllJumps() {
        return _getAll('jumps');
    }

    async function getJumpsByDateRange(startDate, endDate) {
        return new Promise((resolve, reject) => {
            const tx = _db.transaction('jumps', 'readonly');
            const index = tx.objectStore('jumps').index('by-date');
            const range = IDBKeyRange.bound(startDate, endDate);
            const req = index.getAll(range);
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
        });
    }

    async function putJump(jump) {
        return new Promise((resolve, reject) => {
            const tx = _db.transaction('jumps', 'readwrite');
            tx.objectStore('jumps').put(jump);
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
        });
    }

    async function putAllJumps(jumps) {
        return _putAll('jumps', jumps);
    }

    async function deleteJump(id) {
        return new Promise((resolve, reject) => {
            const tx = _db.transaction('jumps', 'readwrite');
            tx.objectStore('jumps').delete(id);
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
        });
    }

    async function clearJumps() {
        return _clear('jumps');
    }

    // ── Equipment (canopies, harnesses, locations) ──────────────────────

    async function getAll(storeName) {
        return _getAll(storeName);
    }

    async function putAll(storeName, items) {
        return _putAll(storeName, items);
    }

    async function clear(storeName) {
        return _clear(storeName);
    }

    // ── Bulk clear all stores ───────────────────────────────────────────

    async function clearAll() {
        const storeNames = ['jumps', 'canopies', 'harnesses', 'locations'];
        return new Promise((resolve, reject) => {
            const tx = _db.transaction(storeNames, 'readwrite');
            for (const name of storeNames) {
                tx.objectStore(name).clear();
            }
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
        });
    }

    // ── Replace all items in a store (clear + put) ──────────────────────

    async function replaceAll(storeName, items) {
        return new Promise((resolve, reject) => {
            const tx = _db.transaction(storeName, 'readwrite');
            const store = tx.objectStore(storeName);
            store.clear();
            for (const item of items) {
                store.put(item);
            }
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
        });
    }

    async function replaceAllJumps(jumps) {
        return replaceAll('jumps', jumps);
    }

    // ── Migration from localStorage ─────────────────────────────────────

    async function migrateFromLocalStorage() {
        const jumpsKey     = 'skydiving-jumps';
        const canopiesKey  = 'skydiving-canopies';
        const harnessesKey = 'skydiving-harnesses';
        const locationsKey = 'skydiving-locations';

        const hasData = localStorage.getItem(jumpsKey)
            || localStorage.getItem(canopiesKey)
            || localStorage.getItem(harnessesKey)
            || localStorage.getItem(locationsKey);

        if (!hasData) return false; // nothing to migrate

        const parse = (key) => {
            try { return JSON.parse(localStorage.getItem(key)) || []; }
            catch (_) { return []; }
        };

        const jumps     = parse(jumpsKey);
        const canopies  = parse(canopiesKey);
        const harnesses = parse(harnessesKey);
        const locations = parse(locationsKey);

        // Write all collections into IDB in a single batch per store
        const storeNames = ['jumps', 'canopies', 'harnesses', 'locations'];
        await new Promise((resolve, reject) => {
            const tx = _db.transaction(storeNames, 'readwrite');
            for (const j of jumps)     tx.objectStore('jumps').put(j);
            for (const c of canopies)  tx.objectStore('canopies').put(c);
            for (const h of harnesses) tx.objectStore('harnesses').put(h);
            for (const l of locations) tx.objectStore('locations').put(l);
            tx.oncomplete = () => resolve();
            tx.onerror    = () => reject(tx.error);
        });

        // Remove migrated keys from localStorage
        localStorage.removeItem(jumpsKey);
        localStorage.removeItem(canopiesKey);
        localStorage.removeItem(harnessesKey);
        localStorage.removeItem(locationsKey);

        console.log(`[DB] Migrated localStorage → IndexedDB (${jumps.length} jumps, ${canopies.length} canopies, ${harnesses.length} harnesses, ${locations.length} locations)`);
        return true;
    }

    /**
     * Request storage access (Storage Access API). Call from a user gesture
     * (e.g. button click) when storage was blocked; then call open() and reload.
     */
    function requestStorageAccess() {
        if (typeof document.requestStorageAccess !== 'function') return Promise.resolve();
        return document.requestStorageAccess();
    }

    // ── Public API ──────────────────────────────────────────────────────

    return {
        open,
        requestStorageAccess,
        getAllJumps,
        getJumpsByDateRange,
        putJump,
        putAllJumps,
        deleteJump,
        clearJumps,
        getAll,
        putAll,
        clear,
        clearAll,
        replaceAll,
        replaceAllJumps,
        migrateFromLocalStorage
    };
})();
