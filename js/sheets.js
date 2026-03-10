// Google Sheets API v4 Integration via OAuth (replaces Apps Script proxy)
// Requires js/auth.js (AuthManager) to be loaded first.

const SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';

class SheetsAPI {
    constructor() {
        this.spreadsheetId = '';
        this.initialized = false;
        this._pollTimer = null;
        this._syncInProgress = false;

        // Legacy Apps Script support for migration detection
        this.webAppUrl = '';

        this.ready = this.setupAPI();
    }

    // ── Initialisation ──────────────────────────────────────────────────

    async setupAPI() {
        try {
            // Detect legacy Apps Script config (for migration)
            const legacyCfg = JSON.parse(localStorage.getItem('sheets-config') || '{}');
            if (legacyCfg.webAppUrl) {
                this.webAppUrl = legacyCfg.webAppUrl;
            }

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

    // ── Spreadsheet creation ────────────────────────────────────────────

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
                                'Jump Number', 'Date', 'Location', 'Equipment',
                                'Notes', 'Timestamp', 'Equipment ID', 'Lineset Number'
                            ].map(v => ({ userEnteredValue: { stringValue: v } }))
                        }]
                    }]
                },
                {
                    properties: { title: 'Equipment', index: 1 },
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
            '/values/Jumps!A2:H?majorDimension=ROWS'
        );

        const rows = result.values || [];
        if (rows.length === 0) return [];

        return rows.map((row, index) => {
            const timestamp = row[5] || new Date().toISOString();
            const parsedTime = new Date(timestamp).getTime();
            const id = Number.isFinite(parsedTime) ? parsedTime : Date.now() + index;
            // Column 6 (index 6) = Equipment ID; fall back to col 3 (display name)
            const equipment = (row[6] && row[6] !== '') ? row[6] : row[3] || '';

            // Normalize date to YYYY-MM-DD
            let date = '';
            if (row[1]) {
                const s = String(row[1]);
                if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
                    date = s.slice(0, 10);
                } else {
                    const d = new Date(s);
                    date = isNaN(d.getTime()) ? s : d.toISOString().slice(0, 10);
                }
            }

            return {
                id,
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

    // ── Write operations (Sheets API v4) ────────────────────────────────

    async uploadAllJumps(jumps) {
        if (!this.initialized) throw new Error('API not initialized');
        this.updateSyncStatus('Uploading jumps...');

        // Build rows: header + data
        const header = ['Jump Number', 'Date', 'Location', 'Equipment', 'Notes', 'Timestamp', 'Equipment ID', 'Lineset Number'];

        const dataRows = jumps.map(jump => {
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
                jump.jumpNumber,
                jump.date,
                jump.location,
                equipmentName,
                jump.notes || '',
                jump.timestamp,
                jump.equipment,        // canopy ID
                jump.linesetNumber || 1
            ];
        });

        // Clear old data then write fresh
        await this._apiCall('POST', '/values/Jumps!A1:H:clear', {});
        await this._apiCall('PUT', '/values/Jumps!A1:H?valueInputOption=RAW', {
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
            ['_syncMeta',  JSON.stringify(dataModified ? { dataModified } : {})],
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
            const localSynced   = localStorage.getItem('skydiving-data-synced') || '';
            const localModified = localStorage.getItem('skydiving-data-modified') || '';

            const hasSheetData = !!(d.harnesses || d.canopies || d.rigs);
            const sheetIsNewer = (sheetTs && sheetTs > localSynced) ||
                                 (hasSheetData && !localSynced && !sheetTs);
            const hasPending   = !!(localModified && localModified > localSynced);

            if (sheetIsNewer) {
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
            console.error('[Startup] Sync failed:', error);
            this.updateSyncStatus('Sync failed');
            setTimeout(() => this.updateSyncStatus('Unsynced'), 3000);
        } finally {
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
            const d           = await this._getEquipment();
            const sheetTs     = (d._syncMeta && d._syncMeta.dataModified) || '';
            const localSynced = localStorage.getItem('skydiving-data-synced') || '';

            if (sheetTs && sheetTs > localSynced) {
                console.warn('[Sync] Sheet is newer — pulling and merging');
                await this._pullAllFromSheet(d, sheetTs);
                this.updateSyncStatus('Online');
                return;
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

        const localSettings = logbook ? logbook.settings
            : JSON.parse(localStorage.getItem('skydiving-settings') || '{}');
        const sheetSettings = d.settings || {};
        const deletedTimestamps = [...new Set([
            ...(localSettings.deletedJumpTimestamps || []),
            ...(sheetSettings.deletedJumpTimestamps || [])
        ])];

        if (d.settings) {
            d.settings.deletedJumpTimestamps = deletedTimestamps;
        }

        // Persist equipment to IndexedDB
        if (d.harnesses)  DB.replaceAll('harnesses', d.harnesses).catch(err => console.error('[Sync] IDB harnesses write failed:', err));
        if (d.canopies)   DB.replaceAll('canopies',  d.canopies).catch(err => console.error('[Sync] IDB canopies write failed:', err));
        if (d.locations)  DB.replaceAll('locations',  d.locations).catch(err => console.error('[Sync] IDB locations write failed:', err));
        if (d.rigs && d.rigs.length > 0) localStorage.setItem('skydiving-equipment-rigs', JSON.stringify(d.rigs));
        if (d.settings)   localStorage.setItem('skydiving-settings', JSON.stringify(d.settings));

        const sheetJumps = await this.getAllJumps();
        const mergedJumps = this._mergeJumps(localJumps, sheetJumps, deletedTimestamps);
        await DB.replaceAllJumps(mergedJumps);

        // Detect whether tombstones removed any jumps that were still on the sheet.
        const deletedSet = new Set(deletedTimestamps);
        const tombstonedFromSheet = sheetJumps.some(j => j.timestamp && deletedSet.has(j.timestamp));

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

            logbook.migrateFromRigsToCanopyLinesets();
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

        // If tombstones pruned jumps that were still in the sheet, re-upload the
        // clean list immediately so the sheet reflects the deletions right away.
        if (tombstonedFromSheet) {
            console.log('[Sync] Tombstones pruned sheet jumps — re-uploading clean list');
            const uploadTs = new Date().toISOString();
            await this.uploadAllJumps(logbook ? logbook.jumps : mergedJumps);
            localStorage.setItem('skydiving-data-synced', uploadTs);
            localStorage.setItem('skydiving-data-modified', uploadTs);
        }

        console.log('[Sync] Pulled and merged data from sheet, ts:', ts);
    }

    _mergeJumps(localJumps, sheetJumps, deletedTimestamps) {
        const deletedSet = new Set(deletedTimestamps);
        const sheetTimestamps = new Set(
            sheetJumps.map(j => j.timestamp).filter(Boolean)
        );
        const merged = sheetJumps.filter(j =>
            !j.timestamp || !deletedSet.has(j.timestamp)
        );
        for (const j of localJumps) {
            if (j.timestamp
                && !deletedSet.has(j.timestamp)
                && !sheetTimestamps.has(j.timestamp)) {
                merged.push(j);
            }
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
                const d       = await this._getEquipment();
                const sheetTs = (d._syncMeta && d._syncMeta.dataModified) || '';

                if (sheetTs && sheetTs > localSynced) {
                    console.warn('[Poll] Sheet is newer — pulling and merging');
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

            // No pending changes — quietly check if sheet is newer
            const d       = await this._getEquipment();
            const sheetTs = (d._syncMeta && d._syncMeta.dataModified) || '';
            if (sheetTs && sheetTs > localSynced) {
                console.log('[Poll] Sheet is newer, pulling and merging...');
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
     * Called by the sync button. If the token is expired, triggers interactive
     * sign-in first, then performs a full sync and starts the background poll.
     */
    async userInitiatedSync() {
        if (!this.initialized) return;
        if (!navigator.onLine) {
            this.updateSyncStatus('Offline');
            return;
        }
        if (!window.AuthManager.isSignedIn()) {
            try {
                this.updateSyncStatus('Signing in...');
                await window.AuthManager.getValidToken();
            } catch (e) {
                console.error('[Sync] Sign-in failed:', e);
                this.updateSyncStatus('Not signed in');
                return;
            }
        }
        // doStartupSync handles push/pull conflict detection and schedules the poll
        await this.doStartupSync();
    }

    // ── Backup rig sheet support ────────────────────────────────────────

    async hasBackupRigsSheet() {
        if (!this.initialized) return false;
        try {
            const meta = await this._apiCall('GET', '?fields=sheets.properties.title');
            const sheets = meta.sheets || [];
            return sheets.some(s => s.properties?.title === 'backupRigs');
        } catch { return false; }
    }

    async restoreEquipmentFromBackup() {
        if (!this.initialized) throw new Error('API not initialized');

        const result = await this._apiCall('GET', '/values/backupRigs!A1:B6?majorDimension=ROWS');
        const rows = result.values || [];
        const d = {};
        for (const row of rows) {
            const key = (row[0] || '').trim();
            if (!key) continue;
            try { d[key] = JSON.parse(row[1] || '{}'); }
            catch { d[key] = row[1]; }
        }

        const hasData = d.harnesses || d.canopies || d.rigs;
        if (!hasData) return false;

        if (d.harnesses)  await DB.replaceAll('harnesses', d.harnesses);
        if (d.canopies)   await DB.replaceAll('canopies',  d.canopies);
        if (d.locations)  await DB.replaceAll('locations', d.locations);
        if (d.rigs && d.rigs.length > 0) localStorage.setItem('skydiving-equipment-rigs', JSON.stringify(d.rigs));
        if (d.settings)   localStorage.setItem('skydiving-settings', JSON.stringify(d.settings));

        const logbook = window.logbook;
        if (logbook) {
            if (d.harnesses)  logbook.harnesses = d.harnesses;
            if (d.canopies)   logbook.canopies  = d.canopies;
            if (d.settings)   logbook.settings  = d.settings;
            if (d.locations)  logbook.locations  = d.locations;

            logbook.migrateFromRigsToCanopyLinesets();
            logbook.canopies.forEach(c => {
                if (!Array.isArray(c.linesets)) c.linesets = [];
                if (c.linesets.length === 0) c.linesets.push({ number: 1, hybrid: false, previousJumps: 0, jumpCount: 0, archived: false });
            });
            logbook.initializeCanopyLinesetJumpCounts();
            logbook.updateEquipmentOptions();
            logbook.updateLocationDatalist();
            if (logbook.currentView === 'equipment') logbook.renderEquipmentView();
            logbook.preFillFormWithLastJump();
        }

        const now = new Date().toISOString();
        await this.syncEquipmentToSheet(now);

        localStorage.setItem('skydiving-data-modified', now);
        localStorage.setItem('skydiving-data-synced', now);

        console.log('[Sheets] Equipment restored from backupRigs sheet');
        return true;
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
