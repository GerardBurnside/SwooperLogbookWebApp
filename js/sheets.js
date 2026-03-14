// Google Sheets API v4 Integration via OAuth (replaces Apps Script proxy)
// Requires js/auth.js (AuthManager) to be loaded first.

const SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';

class SheetsAPI {
    constructor() {
        this.spreadsheetId = '';
        this.initialized = false;
        this._pollTimer = null;
        this._syncInProgress = false;

        this.ready = this.setupAPI();
    }

    // ── Initialisation ──────────────────────────────────────────────────

    async setupAPI() {
        try {
            // Load OAuth-based spreadsheet ID
            this.spreadsheetId = localStorage.getItem('oauth-spreadsheet-id') || '';

            await window.AuthManager.ready;

            if (this.spreadsheetId && window.AuthManager.isSignedIn()) {
                this.initialized = true;
                console.log('[Sheets] OAuth API initialised, spreadsheet:', this.spreadsheetId);
                this.updateSyncStatus('Ready');
            } else if (this.spreadsheetId) {
                // Have a sheet but no active token — will try silent refresh on sync
                this.initialized = true;
                console.log('[Sheets] Spreadsheet configured, token will refresh on sync');
                this.updateSyncStatus('Ready');
            } else {
                console.log('[Sheets] Not configured — sign in to enable sync');
                this.updateSyncStatus('Not signed in');
            }
        } catch (error) {
            console.error('[Sheets] Setup failed:', error);
            this.updateSyncStatus('Configuration error');
        }
    }

    /** Re-initialise after OAuth sign-in or spreadsheet creation. */
    reinitialize(spreadsheetId) {
        this.spreadsheetId = spreadsheetId || '';
        if (spreadsheetId) {
            localStorage.setItem('oauth-spreadsheet-id', spreadsheetId);
        }

        if (this.spreadsheetId) {
            this.initialized = true;
            console.log('[Sheets] Re-initialised with spreadsheet:', this.spreadsheetId);
            this.updateSyncStatus('Ready');
        } else {
            this.initialized = false;
            this.updateSyncStatus('Not signed in');
        }
    }

    /** Generate a unique jump ID (stable across renumbers). */
    static generateJumpId() {
        return typeof crypto !== 'undefined' && crypto.randomUUID
            ? crypto.randomUUID()
            : 'jump-' + Date.now() + '-' + Math.random().toString(36).slice(2);
    }

    /**
     * Persistent device/browser identifier for sync. When the sheet's last write
     * was from this device, we can safely push only (no pull), avoiding data loss.
     */
    getDeviceId() {
        const key = 'skydiving-device-id';
        let id = localStorage.getItem(key);
        if (!id) {
            id = typeof crypto !== 'undefined' && crypto.randomUUID
                ? crypto.randomUUID()
                : 'browser-' + Math.random().toString(36).slice(2) + '-' + Date.now().toString(36);
            localStorage.setItem(key, id);
        }
        return id;
    }

    // ── Sheets API v4 transport layer ───────────────────────────────────

    /**
     * Make an authenticated request to the Google Sheets API v4.
     * Handles token refresh and 401 retry automatically.
     */
    async _apiCall(method, path, body = null, retry = true) {
        const token = await window.AuthManager.getValidToken();
        const url = `${SHEETS_API}/${this.spreadsheetId}${path}`;

        const opts = {
            method,
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json',
            },
        };
        if (body !== null) {
            opts.body = JSON.stringify(body);
        }

        const resp = await fetch(url, opts);

        if (resp.status === 401 && retry) {
            // Token was rejected — force refresh and retry once
            console.warn('[Sheets] 401 — refreshing token and retrying');
            await window.AuthManager.silentRefresh();
            return this._apiCall(method, path, body, false);
        }

        if (!resp.ok) {
            const text = await resp.text().catch(() => '');
            throw new Error(`Sheets API ${resp.status}: ${text}`);
        }

        // Some calls (e.g. clear) may return empty body
        const ct = resp.headers.get('content-type') || '';
        return ct.includes('application/json') ? resp.json() : {};
    }

    // ── Spreadsheet discovery & creation ─────────────────────────────────

    /**
     * Search the user's Drive for an existing Swooper Logbook spreadsheet
     * created by this app (drive.file scope). Returns the spreadsheetId or null.
     */
    async findExistingSpreadsheet() {
        const token = await window.AuthManager.getValidToken();
        const query = "name contains 'Swooper Logbook' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false";
        const url = `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=${encodeURIComponent('files(id,name)')}&orderBy=createdTime`;

        const resp = await fetch(url, {
            headers: { 'Authorization': `Bearer ${token}` },
        });

        if (!resp.ok) {
            console.warn('[Sheets] Drive search failed:', resp.status);
            return null;
        }

        const data = await resp.json();
        if (data.files && data.files.length > 0) {
            console.log('[Sheets] Found existing spreadsheet:', data.files[0].id, data.files[0].name);
            return data.files[0].id;
        }
        return null;
    }

    /**
     * Find an existing Swooper Logbook spreadsheet, or create a new one.
     * Returns the spreadsheetId.
     */
    async findOrCreateSpreadsheet() {
        const existingId = await this.findExistingSpreadsheet();
        if (existingId) {
            localStorage.setItem('oauth-spreadsheet-id', existingId);
            this.spreadsheetId = existingId;
            this.initialized = true;
            return existingId;
        }
        return this.createSpreadsheet();
    }

    /**
     * Create a new spreadsheet in the user's Drive with the required structure.
     * Returns the new spreadsheetId.
     */
    async createSpreadsheet() {
        const token = await window.AuthManager.getValidToken();
        const email = window.AuthManager.userEmail || 'User';

        const body = {
            properties: { title: `Swooper Logbook — ${email}` },
            sheets: [
                {
                    properties: { title: 'Jumps', index: 0 },
                    data: [{
                        startRow: 0, startColumn: 0,
                        rowData: [{
                            values: [
                                'Jump ID', 'Jump Number', 'Date', 'Location', 'Equipment',
                                'Notes', 'Timestamp', 'Equipment ID', 'Lineset Number'
                            ].map(v => ({ userEnteredValue: { stringValue: v } }))
                        }]
                    }]
                },
                {
                    properties: { title: 'deletedJumps', index: 1 },
                    data: [{
                        startRow: 0, startColumn: 0,
                        rowData: [{
                            values: [
                                { userEnteredValue: { stringValue: 'Jump ID' } },
                                { userEnteredValue: { stringValue: 'Date deleted' } }
                            ]
                        }]
                    }]
                },
                {
                    properties: { title: 'Equipment', index: 2 },
                    data: [{
                        startRow: 0, startColumn: 0,
                        rowData: [
                            { values: [{ userEnteredValue: { stringValue: 'harnesses' } }, { userEnteredValue: { stringValue: '[]' } }] },
                            { values: [{ userEnteredValue: { stringValue: 'canopies' } },  { userEnteredValue: { stringValue: '[]' } }] },
                            { values: [{ userEnteredValue: { stringValue: 'rigs' } },      { userEnteredValue: { stringValue: '[]' } }] },
                            { values: [{ userEnteredValue: { stringValue: 'settings' } },  { userEnteredValue: { stringValue: '{}' } }] },
                            { values: [{ userEnteredValue: { stringValue: 'locations' } }, { userEnteredValue: { stringValue: '[]' } }] },
                            { values: [{ userEnteredValue: { stringValue: '_syncMeta' } }, { userEnteredValue: { stringValue: '{}' } }] },
                        ]
                    }]
                }
            ]
        };

        const resp = await fetch(SHEETS_API, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(body),
        });

        if (!resp.ok) {
            const text = await resp.text().catch(() => '');
            throw new Error(`Create spreadsheet failed ${resp.status}: ${text}`);
        }

        const result = await resp.json();
        const newId = result.spreadsheetId;

        localStorage.setItem('oauth-spreadsheet-id', newId);
        this.spreadsheetId = newId;
        this.initialized = true;

        console.log('[Sheets] Created spreadsheet:', newId);
        return newId;
    }

    // ── Read operations (Sheets API v4) ─────────────────────────────────

    async getAllJumps() {
        if (!this.initialized) throw new Error('API not initialized');

        const result = await this._apiCall(
            'GET',
            '/values/Jumps!A2:I?majorDimension=ROWS'
        );

        const rows = result.values || [];
        if (rows.length === 0) return [];

        return rows.map((row, index) => {
            const hasJumpIdColumn = row && row.length >= 9;
            if (hasJumpIdColumn) {
                const jumpId = (row[0] && String(row[0]).trim()) || SheetsAPI.generateJumpId();
                const timestamp = row[6] || new Date().toISOString();
                const parsedTime = new Date(timestamp).getTime();
                const id = Number.isFinite(parsedTime) ? parsedTime : Date.now() + index;
                const equipment = (row[7] && row[7] !== '') ? row[7] : row[4] || '';
                let date = '';
                if (row[2]) {
                    const s = String(row[2]);
                    if (/^\d{4}-\d{2}-\d{2}/.test(s)) date = s.slice(0, 10);
                    else { const d = new Date(s); date = isNaN(d.getTime()) ? s : d.toISOString().slice(0, 10); }
                }
                return {
                    id,
                    jumpId,
                    jumpNumber: parseInt(row[1]) || 0,
                    date,
                    location: row[3] || '',
                    equipment,
                    linesetNumber: parseInt(row[8]) || 1,
                    notes: row[5] || '',
                    timestamp
                };
            }
            // Backward compat: 8 columns (no Jump ID)
            const timestamp = row[5] || new Date().toISOString();
            const parsedTime = new Date(timestamp).getTime();
            const id = Number.isFinite(parsedTime) ? parsedTime : Date.now() + index;
            const equipment = (row[6] && row[6] !== '') ? row[6] : row[3] || '';
            let date = '';
            if (row[1]) {
                const s = String(row[1]);
                if (/^\d{4}-\d{2}-\d{2}/.test(s)) date = s.slice(0, 10);
                else { const d = new Date(s); date = isNaN(d.getTime()) ? s : d.toISOString().slice(0, 10); }
            }
            return {
                id,
                jumpId: SheetsAPI.generateJumpId(),
                jumpNumber: parseInt(row[0]) || 0,
                date,
                location: row[2] || '',
                equipment,
                linesetNumber: parseInt(row[7]) || 1,
                notes: row[4] || '',
                timestamp
            };
        });
    }

    /** Read the Equipment sheet (6 key-value rows). Returns a parsed object. */
    async _getEquipment() {
        const result = await this._apiCall(
            'GET',
            '/values/Equipment!A1:B6?majorDimension=ROWS'
        );

        const rows = result.values || [];
        const data = {};
        for (const row of rows) {
            const key = (row[0] || '').trim();
            if (!key) continue;
            try { data[key] = JSON.parse(row[1] || '{}'); }
            catch { data[key] = row[1]; }
        }
        return data;
    }

    /** Ensure deletedJumps sheet exists (for existing spreadsheets created before this feature). */
    async _ensureDeletedJumpsSheet() {
        const meta = await this._apiCall('GET', '?fields=sheets(properties(title,sheetId))');
        const hasDeletedJumps = (meta.sheets || []).some(s => (s.properties && s.properties.title) === 'deletedJumps');
        if (hasDeletedJumps) return;
        await this._apiCall('POST', ':batchUpdate', {
            requests: [{
                addSheet: {
                    properties: { title: 'deletedJumps' },
                    data: [{
                        startRow: 0, startColumn: 0,
                        rowData: [{
                            values: [
                                { userEnteredValue: { stringValue: 'Jump ID' } },
                                { userEnteredValue: { stringValue: 'Date deleted' } }
                            ]
                        }]
                    }]
                }
            }]
        });
        console.log('[Sheets] Added deletedJumps sheet');
    }

    /** Read deletedJumps sheet and return a Set of jump IDs. */
    async getDeletedJumpIds() {
        try {
            const result = await this._apiCall('GET', '/values/deletedJumps!A2:A?majorDimension=ROWS');
            const rows = result.values || [];
            const ids = new Set();
            for (const row of rows) {
                const id = (row[0] && String(row[0]).trim()) || '';
                if (id) ids.add(id);
            }
            return ids;
        } catch (e) {
            if (e.message && e.message.includes('404')) return new Set();
            const meta = await this._apiCall('GET', '?fields=sheets(properties(title))');
            const hasSheet = (meta.sheets || []).some(s => (s.properties && s.properties.title) === 'deletedJumps');
            if (!hasSheet) return new Set();
            throw e;
        }
    }

    /** Append rows to deletedJumps sheet (one row per jumpId). Call _ensureDeletedJumpsSheet first if needed. */
    async appendDeletedJumps(jumpIds) {
        if (!jumpIds || jumpIds.length === 0) return;
        await this._ensureDeletedJumpsSheet();
        const now = new Date().toISOString();
        const rows = [...jumpIds].map(id => [id, now]);
        const result = await this._apiCall('GET', '/values/deletedJumps?majorDimension=ROWS');
        const existing = (result.values || []).length;
        const startRow = existing + 1;
        const endRow = existing + rows.length;
        const range = `deletedJumps!A${startRow}:B${endRow}`;
        await this._apiCall('PUT', `/values/${encodeURIComponent(range)}?valueInputOption=RAW`, { values: rows });
        console.log('[Sheets] Appended', jumpIds.length, 'deletion(s) to deletedJumps');
    }

    // ── Write operations (Sheets API v4) ────────────────────────────────

    async uploadAllJumps(jumps) {
        if (!this.initialized) throw new Error('API not initialized');
        this.updateSyncStatus('Uploading jumps...');

        const header = ['Jump ID', 'Jump Number', 'Date', 'Location', 'Equipment', 'Notes', 'Timestamp', 'Equipment ID', 'Lineset Number'];
        const sortedJumps = [...jumps].sort((a, b) => (a.jumpNumber || 0) - (b.jumpNumber || 0));

        const dataRows = sortedJumps.map(jump => {
            const jumpId = jump.jumpId || SheetsAPI.generateJumpId();
            if (!jump.jumpId) jump.jumpId = jumpId;
            let equipmentName = jump.equipment;
            if (window.logbook) {
                const canopy = window.logbook.canopies.find(c => c.id === jump.equipment);
                if (canopy) {
                    const ls = canopy.linesets?.find(l => l.number === jump.linesetNumber);
                    const hybridSuffix = ls?.hybrid ? ' (Hybrid)' : '';
                    equipmentName = `${canopy.name}-Lineset#${jump.linesetNumber || 1}${hybridSuffix}`;
                }
            }
            return [
                jumpId,
                jump.jumpNumber,
                jump.date,
                jump.location,
                equipmentName,
                jump.notes || '',
                jump.timestamp,
                jump.equipment,
                jump.linesetNumber || 1
            ];
        });

        await this._apiCall('POST', '/values/Jumps!A1:I:clear', {});
        await this._apiCall('PUT', '/values/Jumps!A1:I?valueInputOption=RAW', {
            values: [header, ...dataRows]
        });

        console.log(`[Sheets] Uploaded ${jumps.length} jumps`);
    }

    async syncEquipmentToSheet(dataModified = null) {
        if (!this.initialized) return;
        if (!window.AuthManager.isSignedIn()) {
            this.updateSyncStatus('Unsynced');
            return;
        }

        const logbook = window.logbook;
        if (!logbook) return;

        const rows = [
            ['harnesses',  JSON.stringify(logbook.harnesses || [])],
            ['canopies',   JSON.stringify(logbook.canopies || [])],
            ['rigs',       JSON.stringify([])],
            ['settings',   JSON.stringify(logbook.settings || {})],
            ['locations',  JSON.stringify(logbook.locations || [])],
            ['_syncMeta',  JSON.stringify(dataModified ? { dataModified, deviceId: this.getDeviceId() } : {})],
        ];

        try {
            await this._apiCall('PUT', '/values/Equipment!A1:B6?valueInputOption=RAW', {
                values: rows
            });
            console.log('[Sheets] Equipment synced');
        } catch (error) {
            console.error('[Sheets] Equipment sync failed:', error);
        }
    }

    // ── Sync logic (same as before, transport-agnostic) ─────────────────

    async doStartupSync() {
        if (!this.initialized) return;

        // Never trigger interactive sign-in at startup — bail silently so the
        // user can initiate auth manually by pressing the sync button.
        if (!window.AuthManager.isSignedIn()) {
            this.updateSyncStatus('Not signed in');
            return;
        }

        if (this._syncInProgress) {
            console.log('[Startup] Sync skipped — another sync in progress');
            this._schedulePoll();
            return;
        }
        this._syncInProgress = true;

        this._cancelPoll();
        this.updateSyncStatus('Syncing...');

        try {
            const d             = await this._getEquipment();
            const sheetTs       = (d._syncMeta && d._syncMeta.dataModified) || '';
            const sheetDeviceId = (d._syncMeta && d._syncMeta.deviceId) || null;
            const localSynced   = localStorage.getItem('skydiving-data-synced') || '';
            const localModified = localStorage.getItem('skydiving-data-modified') || '';

            const hasSheetData = !!(d.harnesses || d.canopies);
            const sheetIsNewer = (sheetTs && sheetTs > localSynced) ||
                                 (hasSheetData && !localSynced && !sheetTs);
            const hasPending   = !!(localModified && localModified > localSynced);
            const lastWriteFromThisDevice = sheetDeviceId && sheetDeviceId === this.getDeviceId();

            if (sheetIsNewer && !lastWriteFromThisDevice) {
                if (hasPending) {
                    console.warn('[Startup] Conflict — sheet is newer but local has changes, merging...');
                } else {
                    console.log('[Startup] Sheet is newer, pulling all data...');
                }
                await this._pullAllFromSheet(d, sheetTs);
                if (hasPending) {
                    const logbook = window.logbook;
                    if (logbook) logbook.showSyncConflictModal();
                }
            } else if (sheetIsNewer && lastWriteFromThisDevice) {
                console.log('[Startup] Sheet newer but last write from this device — pushing only (no pull)');
                const newTs   = new Date().toISOString();
                const logbook = window.logbook;
                await this.uploadAllJumps(logbook?.jumps || []);
                await this.syncEquipmentToSheet(newTs);
                localStorage.setItem('skydiving-data-synced', newTs);
                localStorage.setItem('skydiving-data-modified', newTs);
                console.log('[Sync] Startup push complete (same device), ts:', newTs);
            } else if (hasPending) {
                console.log('[Startup] Pending local changes, pushing...');
                const newTs   = new Date().toISOString();
                const logbook = window.logbook;
                await this.uploadAllJumps(logbook?.jumps || []);
                await this.syncEquipmentToSheet(newTs);
                localStorage.setItem('skydiving-data-synced', newTs);
                localStorage.setItem('skydiving-data-modified', newTs);
                console.log('[Sync] Startup push complete, ts:', newTs);
            }

            this.updateSyncStatus('Online');
        } catch (error) {
            // If the stored spreadsheet was deleted or is inaccessible (404),
            // clear the stale ID and try to find the real one on Drive.
            if (error.message && error.message.includes('404') && !this._recoveryAttempted) {
                console.warn('[Startup] Spreadsheet not found (404) — searching Drive for existing one...');
                this._recoveryAttempted = true;
                localStorage.removeItem('oauth-spreadsheet-id');
                this.spreadsheetId = '';
                this.initialized = false;
                this._syncInProgress = false;

                try {
                    const newId = await this.findOrCreateSpreadsheet();
                    if (newId) {
                        this.reinitialize(newId);
                        await this.doStartupSync();
                        return;
                    }
                } catch (recoveryError) {
                    console.error('[Startup] Recovery failed:', recoveryError);
                }
            }

            console.error('[Startup] Sync failed:', error);
            this.updateSyncStatus('Sync failed');
            setTimeout(() => this.updateSyncStatus('Unsynced'), 3000);
        } finally {
            this._recoveryAttempted = false;
            this._syncInProgress = false;
            this._schedulePoll();
        }
    }

    async syncWithSheet() {
        await this.ready;
        await this.doStartupSync();
    }

    async pushAllWithGuard() {
        if (!this.initialized || !navigator.onLine) return;
        if (!window.AuthManager.isSignedIn()) {
            this.updateSyncStatus('Unsynced');
            return;
        }
        if (this._syncInProgress) {
            console.log('[Sync] Push skipped — another sync in progress');
            return;
        }
        this._syncInProgress = true;

        this.updateSyncStatus('Syncing...');

        try {
            const d             = await this._getEquipment();
            const sheetTs       = (d._syncMeta && d._syncMeta.dataModified) || '';
            const sheetDeviceId = (d._syncMeta && d._syncMeta.deviceId) || null;
            const localSynced   = localStorage.getItem('skydiving-data-synced') || '';

            const lastWriteFromThisDevice = sheetDeviceId && sheetDeviceId === this.getDeviceId();

            if (sheetTs && sheetTs > localSynced && !lastWriteFromThisDevice) {
                console.warn('[Sync] Sheet is newer (other device) — pulling and merging');
                await this._pullAllFromSheet(d, sheetTs);
                this.updateSyncStatus('Online');
                return;
            }
            if (sheetTs && sheetTs > localSynced && lastWriteFromThisDevice) {
                console.log('[Sync] Sheet newer but last write from this device — pushing only');
            }

            const newTs   = new Date().toISOString();
            const logbook = window.logbook;
            await this.uploadAllJumps(logbook?.jumps || []);
            await this.syncEquipmentToSheet(newTs);

            localStorage.setItem('skydiving-data-synced', newTs);
            localStorage.setItem('skydiving-data-modified', newTs);

            this.updateSyncStatus('Synced');
            setTimeout(() => this.updateSyncStatus('Online'), 2000);
            console.log('[Sync] Push complete, ts:', newTs);
        } catch (error) {
            console.error('[Sync] pushAllWithGuard failed:', error);
            this.updateSyncStatus('Sync failed');
            setTimeout(() => this.updateSyncStatus('Unsynced'), 3000);
        } finally {
            this._syncInProgress = false;
        }
    }

    async _pullAllFromSheet(d, sheetDataModified) {
        const logbook = window.logbook;

        const localJumps = logbook ? [...logbook.jumps]
            : await DB.getAllJumps();

        if (d.harnesses)  DB.replaceAll('harnesses', d.harnesses).catch(err => console.error('[Sync] IDB harnesses write failed:', err));
        if (d.canopies)   DB.replaceAll('canopies',  d.canopies).catch(err => console.error('[Sync] IDB canopies write failed:', err));
        if (d.locations)   DB.replaceAll('locations',  d.locations).catch(err => console.error('[Sync] IDB locations write failed:', err));
        if (d.settings)    localStorage.setItem('skydiving-settings', JSON.stringify(d.settings));

        const deletedJumpIds = await this.getDeletedJumpIds();
        const sheetJumps = await this.getAllJumps();
        const mergedJumps = this._mergeJumps(localJumps, sheetJumps, deletedJumpIds);
        await DB.replaceAllJumps(mergedJumps);

        const ts = sheetDataModified || new Date().toISOString();
        localStorage.setItem('skydiving-data-synced', ts);

        if (mergedJumps.length > sheetJumps.length) {
            console.log(`[Sync] Merge recovered ${mergedJumps.length - sheetJumps.length} local-only jump(s)`);
            localStorage.setItem('skydiving-data-modified', new Date().toISOString());
        } else {
            localStorage.setItem('skydiving-data-modified', ts);
        }

        if (logbook) {
            if (d.harnesses)  logbook.harnesses  = d.harnesses;
            if (d.canopies)   logbook.canopies   = d.canopies;
            if (d.settings)   logbook.settings   = d.settings;
            if (d.locations) {
                logbook.locations = d.locations;
                logbook.locations.sort((a, b) => (a.sortOrder ?? Infinity) - (b.sortOrder ?? Infinity));
            }

            logbook.canopies.forEach(c => {
                if (!Array.isArray(c.linesets)) c.linesets = [];
                if (c.linesets.length === 0) c.linesets.push({ number: 1, hybrid: false, previousJumps: 0, jumpCount: 0, archived: false });
            });

            logbook.jumps = mergedJumps;
            logbook.renumberJumps();
            logbook.initializeCanopyLinesetJumpCounts();
            logbook.updateEquipmentOptions();
            logbook.updateLocationDatalist();
            logbook.updateStats();
            logbook.renderJumpsList();
            if (logbook.currentView === 'equipment') logbook.renderEquipmentView();
            if (logbook.currentView === 'stats') logbook.renderStats();
            logbook.preFillFormWithLastJump();
        }

        console.log('[Sync] Pulled and merged data from sheet, ts:', ts);
    }

    /**
     * True if two jumps have the same jumpNumber and same values for all fields except jumpId/id.
     * Used to avoid duplicating the same logical jump when local and sheet have different IDs.
     */
    _jumpContentEqual(a, b) {
        if (!a || !b) return false;
        const num = (n) => (typeof n === 'number' && !Number.isNaN(n)) ? n : parseInt(n, 10) || 0;
        if (num(a.jumpNumber) !== num(b.jumpNumber)) return false;
        const str = (s) => (s == null ? '' : String(s)).trim();
        const dateNorm = (d) => {
            const s = str(d);
            if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
            const t = new Date(s).getTime();
            return Number.isFinite(t) ? new Date(t).toISOString().slice(0, 10) : s;
        };
        return (
            dateNorm(a.date) === dateNorm(b.date) &&
            str(a.location) === str(b.location) &&
            str(a.equipment) === str(b.equipment) &&
            str(a.notes) === str(b.notes) &&
            num(a.linesetNumber) === num(b.linesetNumber) &&
            str(a.timestamp) === str(b.timestamp)
        );
    }

    /** Merge by jumpId. Only exclude jumps that are in deletedJumpIds. Local jumps not on sheet are kept (recovered).
     * If a local jump has the same jumpNumber and identical data as a sheet jump (only jumpId differs), we keep the
     * sheet row and do not duplicate — effectively adopting the sheet's jumpId for that jump. */
    _mergeJumps(localJumps, sheetJumps, deletedJumpIds) {
        const deletedSet = deletedJumpIds instanceof Set ? deletedJumpIds : new Set(deletedJumpIds);
        const sheetJumpIds = new Set(
            sheetJumps.map(j => j.jumpId).filter(Boolean)
        );
        const merged = sheetJumps.filter(j => !j.jumpId || !deletedSet.has(j.jumpId));
        for (const j of localJumps) {
            const jumpId = j.jumpId || SheetsAPI.generateJumpId();
            if (!j.jumpId) j.jumpId = jumpId;
            if (deletedSet.has(jumpId)) continue;
            if (sheetJumpIds.has(jumpId)) continue; // already in merged from sheet
            // Local-only by ID: check if any sheet jump is the same row (same jumpNumber + same content)
            const sameOnSheet = merged.some(sheetJump => this._jumpContentEqual(j, sheetJump));
            if (sameOnSheet) continue; // keep sheet version (with sheet's jumpId), do not duplicate
            merged.push(j);
        }
        return merged;
    }

    async doPendingPush() {
        if (!this.initialized || !navigator.onLine) return;
        if (!window.AuthManager.isSignedIn()) return; // bail silently — background poll must not show sign-in UI
        if (this._syncInProgress) return;
        this._syncInProgress = true;

        try {
            const localModified = localStorage.getItem('skydiving-data-modified') || '';
            const localSynced   = localStorage.getItem('skydiving-data-synced') || '';

            if (localModified && localModified > localSynced) {
                this.updateSyncStatus('Syncing...');
                const d             = await this._getEquipment();
                const sheetTs       = (d._syncMeta && d._syncMeta.dataModified) || '';
                const sheetDeviceId = (d._syncMeta && d._syncMeta.deviceId) || null;
                const lastWriteFromThisDevice = sheetDeviceId && sheetDeviceId === this.getDeviceId();

                if (sheetTs && sheetTs > localSynced && !lastWriteFromThisDevice) {
                    console.warn('[Poll] Sheet is newer (other device) — pulling and merging');
                    await this._pullAllFromSheet(d, sheetTs);
                } else {
                    const newTs   = new Date().toISOString();
                    const logbook = window.logbook;
                    await this.uploadAllJumps(logbook?.jumps || []);
                    await this.syncEquipmentToSheet(newTs);
                    localStorage.setItem('skydiving-data-synced', newTs);
                    localStorage.setItem('skydiving-data-modified', newTs);
                    console.log('[Poll] Push complete, ts:', newTs);
                }
                this.updateSyncStatus('Synced');
                setTimeout(() => this.updateSyncStatus('Online'), 2000);
                return;
            }

            // No pending changes — quietly check if sheet is newer (and from another device)
            const d             = await this._getEquipment();
            const sheetTs       = (d._syncMeta && d._syncMeta.dataModified) || '';
            const sheetDeviceId = (d._syncMeta && d._syncMeta.deviceId) || null;
            const lastWriteFromThisDevice = sheetDeviceId && sheetDeviceId === this.getDeviceId();
            if (sheetTs && sheetTs > localSynced && !lastWriteFromThisDevice) {
                console.log('[Poll] Sheet is newer (other device), pulling and merging...');
                this.updateSyncStatus('Syncing...');
                await this._pullAllFromSheet(d, sheetTs);
                this.updateSyncStatus('Synced');
                setTimeout(() => this.updateSyncStatus('Online'), 2000);
            }
        } catch (error) {
            console.warn('[Poll] doPendingPush failed:', error);
            this.updateSyncStatus('Sync failed');
            setTimeout(() => this.updateSyncStatus('Unsynced'), 3000);
        } finally {
            this._syncInProgress = false;
        }
    }

    /**
     * Called by the sync button. If not configured, triggers the full sign-in
     * and spreadsheet setup flow. If the token is expired, refreshes it first,
     * then performs a full sync and starts the background poll.
     */
    async userInitiatedSync() {
        if (!navigator.onLine) {
            this.updateSyncStatus('Offline');
            return;
        }

        // If not initialized (no spreadsheet configured), trigger the full
        // sign-in flow from app.js which handles OAuth + spreadsheet discovery.
        if (!this.initialized || !window.AuthManager.isSignedIn()) {
            if (window.logbook && typeof window.logbook.handleGoogleSignIn === 'function') {
                await window.logbook.handleGoogleSignIn();
                return;
            }
            this.updateSyncStatus('Not signed in');
            return;
        }

        // doStartupSync handles push/pull conflict detection and schedules the poll
        await this.doStartupSync();
    }

    // ── Polling ─────────────────────────────────────────────────────────

    _schedulePoll(intervalMs = 120000) {
        this._cancelPoll();
        if (!this.initialized) return;
        this._pollTimer = setTimeout(() => {
            if (navigator.onLine) {
                console.log('[Poll] Auto-sync triggered');
                this.doPendingPush().finally(() => this._schedulePoll(intervalMs));
            } else {
                this._schedulePoll(intervalMs);
            }
        }, intervalMs);
    }

    _cancelPoll() {
        if (this._pollTimer !== null) {
            clearTimeout(this._pollTimer);
            this._pollTimer = null;
        }
    }

    // ── UI helpers ──────────────────────────────────────────────────────

    updateSyncStatus(status) {
        const syncElement = document.getElementById('syncStatus');
        const syncBtn = document.getElementById('syncBtn');
        if (syncElement) {
            syncElement.textContent = status;

            syncElement.className = 'sync-status';
            if (status === 'Syncing...' || status === 'Uploading jumps...') {
                syncElement.classList.add('syncing');
            } else if (status === 'Synced' || status === 'Online' || status === 'Ready') {
                syncElement.classList.add('success');
            } else if (status === 'Unsynced' || status === 'Not signed in') {
                syncElement.classList.add('warning');
            } else if (status.includes('failed') || status.includes('error')) {
                syncElement.classList.add('error');
            }
        }
        if (syncBtn) {
            if (status === 'Syncing...' || status === 'Uploading jumps...') {
                syncBtn.classList.add('syncing');
                syncBtn.classList.remove('unsynced');
            } else if (status === 'Unsynced' || status === 'Not signed in') {
                syncBtn.classList.remove('syncing');
                syncBtn.classList.add('unsynced');
            } else {
                syncBtn.classList.remove('syncing');
                syncBtn.classList.remove('unsynced');
            }
        }
    }
}

// Initialise global instance
window.SheetsAPI = new SheetsAPI();

// Wire up the sync button
document.addEventListener('DOMContentLoaded', () => {
    const syncBtn = document.getElementById('syncBtn');
    if (syncBtn) {
        syncBtn.onclick = () => window.SheetsAPI.userInitiatedSync();
    }
});
