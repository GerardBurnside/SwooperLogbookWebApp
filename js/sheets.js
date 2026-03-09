// Google Sheets API Integration via Apps Script
class SheetsAPI {
    constructor() {
        this.webAppUrl = ''; // Will be loaded from config
        this.spreadsheetId = ''; // Will be loaded from config
        this.initialized = false;
        this._pollTimer = null; // handle for the 2-minute background poll
        this._syncInProgress = false; // mutex to prevent overlapping sync operations
        
        // Expose a promise that resolves when setup is complete
        this.ready = this.setupAPI();
    }

    async setupAPI() {
        try {
            // Load config from local file
            const config = await this.loadConfig();
            this.webAppUrl = config.webAppUrl;
            this.spreadsheetId = config.spreadsheetId;
            
            if (this.webAppUrl) {
                this.initialized = true;
                console.log('Google Apps Script API initialized');
                this.updateSyncStatus('Ready');
            } else {
                console.log('Google Apps Script API not configured');
                this.updateSyncStatus('Not configured');
            }
        } catch (error) {
            console.error('Failed to initialize Google Apps Script API:', error);
            this.updateSyncStatus('Configuration error');
        }
    }

    async loadConfig() {
        const config = localStorage.getItem('sheets-config');
        return config ? JSON.parse(config) : {};
    }

    /**
     * Re-initialize the API with new credentials (called from Settings).
     */
    reinitialize(webAppUrl, spreadsheetId) {
        this.webAppUrl = webAppUrl || '';
        this.spreadsheetId = spreadsheetId || '';

        if (this.webAppUrl) {
            this.initialized = true;
            console.log('Google Apps Script API re-initialized with new config');
            this.updateSyncStatus('Ready');
        } else {
            this.initialized = false;
            this.updateSyncStatus('Not configured');
        }
    }

    async getAllJumps() {
        if (!this.initialized) {
            throw new Error('API not initialized');
        }

        const response = await fetch(this.webAppUrl + '?action=getJumps', {
            method: 'GET',
            redirect: 'follow'
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const result = await response.json();
        
        if (result.error) {
            throw new Error(result.error);
        }

        if (!result.data || result.data.length === 0) {
            return [];
        }

        // Convert to jump objects (skip header row)
        const rows = result.data.slice(1);
        return rows.map((row, index) => {
            const timestamp = row[5] || new Date().toISOString();
            // Restore numeric id from timestamp (originally Date.now()); fall back to
            // a unique value derived from index so id is never undefined.
            const parsedTime = new Date(timestamp).getTime();
            const id = Number.isFinite(parsedTime) ? parsedTime : Date.now() + index;
            // col 7 (row[6]) holds the rig ID written by the updated script;
            // fall back to row[3] (name string) for legacy rows without the ID column.
            const equipment = (row[6] && row[6] !== '') ? row[6] : row[3] || '';

            // Normalize date to YYYY-MM-DD. Google Sheets getValues() returns
            // Date objects for date cells which JSON.stringify serialises to
            // full ISO timestamps ("2026-03-08T00:00:00.000Z"); the app
            // expects plain "2026-03-08".
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

    /**
     * Called on startup: fetch the sheet's dataModified timestamp and decide
     * whether to pull (sheet is newer), push (pending local changes), or
     * do nothing (already in sync).
     */
    async doStartupSync() {
        if (!this.initialized) return;
        if (this._syncInProgress) {
            console.log('[Startup] Sync skipped — another sync in progress');
            this._schedulePoll();
            return;
        }
        this._syncInProgress = true;

        this._cancelPoll();
        this.updateSyncStatus('Syncing...');

        try {
            const response = await fetch(this.webAppUrl + '?action=getEquipment', {
                method: 'GET',
                redirect: 'follow'
            });

            if (!response.ok) throw new Error(`HTTP ${response.status}`);

            const result = await response.json();
            if (result.error) throw new Error(result.error);

            const d             = result.data || {};
            const sheetTs       = (d._syncMeta && d._syncMeta.dataModified) || '';
            const localSynced   = localStorage.getItem('skydiving-data-synced') || '';
            const localModified = localStorage.getItem('skydiving-data-modified') || '';

            const hasSheetData = !!(d.harnesses || d.canopies || d.rigs);
            // Sheet is newer if its timestamp is ahead, OR if it has data and we
            // have never synced before (migration / fresh device with no timestamps).
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
                // Push directly (we already checked timestamps — no conflict)
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

    /**
     * Backward-compatible entry point used by app.js.
     * Delegates to startup sync logic (pull/push/no-op + poll scheduling).
     */
    async syncWithSheet() {
        await this.ready;
        await this.doStartupSync();
    }

    /**
     * Push all local data to the sheet, but first verify the sheet has not
     * been updated by another device since our last sync. If it has, pull
     * instead and show the user a conflict notification.
     */
    async pushAllWithGuard() {
        if (!this.initialized || !navigator.onLine) return;
        if (this._syncInProgress) {
            console.log('[Sync] Push skipped — another sync in progress');
            return;
        }
        this._syncInProgress = true;

        this.updateSyncStatus('Syncing...');

        try {
            const response = await fetch(this.webAppUrl + '?action=getEquipment', {
                method: 'GET',
                redirect: 'follow'
            });

            if (!response.ok) throw new Error(`HTTP ${response.status}`);

            const result = await response.json();
            if (result.error) throw new Error(result.error);

            const d           = result.data || {};
            const sheetTs     = (d._syncMeta && d._syncMeta.dataModified) || '';
            const localSynced = localStorage.getItem('skydiving-data-synced') || '';

            if (sheetTs && sheetTs > localSynced) {
                // Sheet updated since our last sync — merge instead of blindly pushing
                console.warn('[Sync] Sheet is newer — pulling and merging');
                await this._pullAllFromSheet(d, sheetTs);
                this.updateSyncStatus('Online');
                return;
            }

            // No conflict — generate a new timestamp for this write
            const newTs   = new Date().toISOString();
            const logbook = window.logbook;
            await this.uploadAllJumps(logbook?.jumps || []);
            await this.syncEquipmentToSheet(newTs);

            // Record this push as the new sync baseline
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

    /**
     * Apply sheet equipment data + jumps to local storage and the live logbook.
     * Sets both timestamp keys so local state is clean and in sync with the sheet.
     */
    async _pullAllFromSheet(d, sheetDataModified) {
        const logbook = window.logbook;

        // Snapshot local jumps BEFORE the async getAllJumps fetch, so a concurrent
        // addJump that runs during the await is not silently discarded.
        const localJumps = logbook ? [...logbook.jumps]
            : await DB.getAllJumps();

        // Merge deletedJumpTimestamps from both local and sheet settings
        const localSettings = logbook ? logbook.settings
            : JSON.parse(localStorage.getItem('skydiving-settings') || '{}');
        const sheetSettings = d.settings || {};
        const deletedTimestamps = [...new Set([
            ...(localSettings.deletedJumpTimestamps || []),
            ...(sheetSettings.deletedJumpTimestamps || [])
        ])];

        // Write the merged deletions back into the sheet settings before persisting
        if (d.settings) {
            d.settings.deletedJumpTimestamps = deletedTimestamps;
        }

        // Persist equipment to IndexedDB
        if (d.harnesses)  DB.replaceAll('harnesses', d.harnesses).catch(err => console.error('[Sync] IDB harnesses write failed:', err));
        if (d.canopies)   DB.replaceAll('canopies',  d.canopies).catch(err => console.error('[Sync] IDB canopies write failed:', err));
        if (d.locations)  DB.replaceAll('locations',  d.locations).catch(err => console.error('[Sync] IDB locations write failed:', err));
        // If the sheet still has old-format rigs, store them so the migration runs on apply
        if (d.rigs && d.rigs.length > 0) localStorage.setItem('skydiving-equipment-rigs', JSON.stringify(d.rigs));
        if (d.settings)   localStorage.setItem('skydiving-settings',       JSON.stringify(d.settings));

        // Fetch jumps from sheet and MERGE with local jumps to prevent data loss
        const sheetJumps = await this.getAllJumps();
        const mergedJumps = this._mergeJumps(localJumps, sheetJumps, deletedTimestamps);
        await DB.replaceAllJumps(mergedJumps);

        // Stamp sync baseline
        const ts = sheetDataModified || new Date().toISOString();
        localStorage.setItem('skydiving-data-synced', ts);

        if (mergedJumps.length > sheetJumps.length) {
            // Local had jumps the sheet didn't — schedule a push on the next cycle
            console.log(`[Sync] Merge recovered ${mergedJumps.length - sheetJumps.length} local-only jump(s)`);
            localStorage.setItem('skydiving-data-modified', new Date().toISOString());
        } else {
            localStorage.setItem('skydiving-data-modified', ts);
        }

        // Apply to live logbook instance
        if (logbook) {
            if (d.harnesses)  logbook.harnesses     = d.harnesses;
            if (d.canopies)   logbook.canopies      = d.canopies;
            if (d.settings)   logbook.settings      = d.settings;
            if (d.locations)  logbook.locations     = d.locations;

            // Run migration if sheet had old-format rigs
            logbook.migrateFromRigsToCanopyLinesets();
            // Ensure all canopies have linesets
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
            logbook.preFillFormWithLastJump();
        }

        console.log('[Sync] Pulled and merged data from sheet, ts:', ts);
    }

    /**
     * Merge local jumps with sheet jumps by timestamp, preserving local-only
     * jumps that haven't been pushed to the sheet yet.
     */
    _mergeJumps(localJumps, sheetJumps, deletedTimestamps) {
        const deletedSet = new Set(deletedTimestamps);

        // Collect sheet jump timestamps for fast lookup
        const sheetTimestamps = new Set(
            sheetJumps.map(j => j.timestamp).filter(Boolean)
        );

        // Start with all sheet jumps (minus deleted)
        const merged = sheetJumps.filter(j =>
            !j.timestamp || !deletedSet.has(j.timestamp)
        );

        // Add local-only jumps (those with timestamps not on the sheet)
        for (const j of localJumps) {
            if (j.timestamp
                && !deletedSet.has(j.timestamp)
                && !sheetTimestamps.has(j.timestamp)) {
                merged.push(j);
            }
        }

        return merged;
    }

    /**
     * Called by the background poll and the online-event handler.
     * Pushes if there are pending local changes, otherwise checks quietly
     * whether the sheet is newer and pulls if so.
     */
    async doPendingPush() {
        if (!this.initialized || !navigator.onLine) return;
        if (this._syncInProgress) return;
        this._syncInProgress = true;

        try {
            const localModified = localStorage.getItem('skydiving-data-modified') || '';
            const localSynced   = localStorage.getItem('skydiving-data-synced') || '';

            if (localModified && localModified > localSynced) {
                // There are pending local changes — push them
                this.updateSyncStatus('Syncing...');
                const response = await fetch(this.webAppUrl + '?action=getEquipment', {
                    method: 'GET',
                    redirect: 'follow'
                });
                if (!response.ok) throw new Error(`HTTP ${response.status}`);
                const result = await response.json();
                if (result.error) throw new Error(result.error);

                const d       = result.data || {};
                const sheetTs = (d._syncMeta && d._syncMeta.dataModified) || '';

                if (sheetTs && sheetTs > localSynced) {
                    // Conflict: merge rather than blindly pushing
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

            // No pending changes — quietly check if the sheet is newer
            const response = await fetch(this.webAppUrl + '?action=getEquipment', {
                method: 'GET',
                redirect: 'follow'
            });
            if (!response.ok) return;

            const result = await response.json();
            if (result.error) return;

            const d       = result.data || {};
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
     * Push all equipment components + settings to the Equipment sheet tab.
     * Called whenever the user saves any equipment change.
     */
    async syncEquipmentToSheet(dataModified = null) {
        if (!this.initialized) return;

        const logbook = window.logbook;
        if (!logbook) return;

        const payload = {
            harnesses:    logbook.harnesses,
            canopies:     logbook.canopies,
            settings:     logbook.settings,
            locations:    logbook.locations
        };
        if (dataModified)  payload._syncMeta = { dataModified };

        try {
            const response = await fetch(this.webAppUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain' },
                redirect: 'follow',
                body: JSON.stringify({ action: 'saveEquipment', data: payload })
            });

            if (!response.ok) throw new Error(`HTTP ${response.status}`);

            const result = await response.json();
            if (result.error) throw new Error(result.error);

            console.log('Equipment synced to sheet');
        } catch (error) {
            console.error('Failed to sync equipment to sheet:', error);
        }
    }

    /**
     * Return true if the spreadsheet has a readable backupRigs sheet tab.
     */
    async hasBackupRigsSheet() {
        if (!this.initialized) return false;

        try {
            const response = await fetch(this.webAppUrl + '?action=getBackupEquipment', {
                method: 'GET',
                redirect: 'follow'
            });

            if (!response.ok) return false;

            const result = await response.json();
            if (result.error) return false;

            if (typeof result.hasBackupRigsSheet === 'boolean') {
                return result.hasBackupRigsSheet;
            }

            return !!(result.data && Object.keys(result.data).length > 0);
        } catch (error) {
            console.warn('Could not verify backupRigs sheet:', error);
            return false;
        }
    }

    /**
     * Restore all equipment from the "backupRigs" sheet tab.
     * Overwrites local equipment and pushes restored data to the main Equipment sheet.
     */
    async restoreEquipmentFromBackup() {
        if (!this.initialized) throw new Error('API not initialized');

        const response = await fetch(this.webAppUrl + '?action=getBackupEquipment', {
            method: 'GET',
            redirect: 'follow'
        });

        if (!response.ok) throw new Error(`HTTP ${response.status}`);

        const result = await response.json();
        if (result.error) throw new Error(result.error);

        const d = result.data || {};
        const hasData = d.harnesses || d.canopies || d.rigs;
        if (!hasData) return false;

        // Overwrite IndexedDB equipment stores (settings & legacy rigs stay in localStorage)
        if (d.harnesses)  await DB.replaceAll('harnesses', d.harnesses);
        if (d.canopies)   await DB.replaceAll('canopies',  d.canopies);
        if (d.locations)  await DB.replaceAll('locations', d.locations);
        if (d.rigs && d.rigs.length > 0) localStorage.setItem('skydiving-equipment-rigs', JSON.stringify(d.rigs));
        if (d.settings)   localStorage.setItem('skydiving-settings',        JSON.stringify(d.settings));

        // Apply to live logbook instance
        const logbook = window.logbook;
        if (logbook) {
            if (d.harnesses)  logbook.harnesses      = d.harnesses;
            if (d.canopies)   logbook.canopies       = d.canopies;
            if (d.settings)   logbook.settings       = d.settings;
            if (d.locations)  logbook.locations       = d.locations;

            // Run migration if backup had old-format rigs
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

        // Push restored data to the main Equipment sheet, stamped as authoritative now.
        const now = new Date().toISOString();
        await this.syncEquipmentToSheet(now);

        // Stamp both local timestamps — this restore is a clean baseline.
        localStorage.setItem('skydiving-data-modified', now);
        localStorage.setItem('skydiving-data-synced', now);

        console.log('Equipment restored from backupRigs sheet');
        return true;
    }

    /** Schedule a background sync poll after `intervalMs` ms (default 2 min). */
    _schedulePoll(intervalMs = 120000) {
        this._cancelPoll();
        if (!this.initialized) return;
        this._pollTimer = setTimeout(() => {
            if (navigator.onLine) {
                console.log('[Poll] Auto-sync triggered');
                this.doPendingPush(); // reschedules itself on completion
            } else {
                this._schedulePoll(intervalMs); // offline — try again later
            }
        }, intervalMs);
    }

    /** Cancel any pending background poll. */
    _cancelPoll() {
        if (this._pollTimer !== null) {
            clearTimeout(this._pollTimer);
            this._pollTimer = null;
        }
    }

    async uploadAllJumps(jumps) {
        if (!this.initialized) {
            throw new Error('API not initialized');
        }
        
        if (jumps.length === 0) {
            return;
        }
        
        this.updateSyncStatus('Uploading jumps...');
        
        // Prepare the data for upload
        const jumpData = jumps.map(jump => {
            // Resolve canopy name + lineset for human-readable display
            let equipmentName = jump.equipment;
            if (window.logbook) {
                const canopy = window.logbook.canopies.find(c => c.id === jump.equipment);
                if (canopy) {
                    const ls = canopy.linesets?.find(l => l.number === jump.linesetNumber);
                    const hybridSuffix = ls?.hybrid ? ' (Hybrid)' : '';
                    equipmentName = `${canopy.name}-Lineset#${jump.linesetNumber || 1}${hybridSuffix}`;
                }
            }
            
            return {
                jumpNumber: jump.jumpNumber,
                date: jump.date,
                location: jump.location,
                equipment: equipmentName,
                equipmentId: jump.equipment,  // preserve canopy ID for round-trip
                notes: jump.notes,
                timestamp: jump.timestamp,
                linesetNumber: jump.linesetNumber || 1
            };
        });
        
        const response = await fetch(this.webAppUrl, {
            method: 'POST',
            // Use text/plain to avoid CORS preflight (simple request)
            headers: {
                'Content-Type': 'text/plain',
            },
            redirect: 'follow',
            body: JSON.stringify({ action: 'uploadJumps', data: jumpData })
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const result = await response.json();
        if (result.error) {
            throw new Error(result.error);
        }
        
        console.log(`Uploaded ${jumps.length} jumps to Google Sheets`);
        return result;
    }

    updateSyncStatus(status) {
        const syncElement = document.getElementById('syncStatus');
        const syncBtn = document.getElementById('syncBtn');
        if (syncElement) {
            syncElement.textContent = status;
            
            // Update classes based on status
            syncElement.className = 'sync-status';
            if (status === 'Syncing...' || status === 'Uploading jumps...') {
                syncElement.classList.add('syncing');
            } else if (status === 'Synced' || status === 'Online' || status === 'Ready') {
                syncElement.classList.add('success');
            } else if (status === 'Unsynced' || status === 'Not configured') {
                syncElement.classList.add('warning');
            } else if (status.includes('failed') || status.includes('error')) {
                syncElement.classList.add('error');
            }
        }
        // Spin the sync button while syncing
        if (syncBtn) {
            if (status === 'Syncing...' || status === 'Uploading jumps...') {
                syncBtn.classList.add('syncing');
            } else {
                syncBtn.classList.remove('syncing');
            }
        }
    }

    // Helper method to create setup instructions
    generateSetupInstructions() {
        return `
Google Sheets Setup Instructions (Apps Script Method):

1. Create a new Google Spreadsheet
2. Rename the first sheet to "Jumps"
3. Add headers in row 1: Jump Number, Date, Location, Equipment, Notes, Timestamp
4. Get your Spreadsheet ID from the URL (the long string between /d/ and /edit)
5. Go to Extensions → Apps Script in your spreadsheet
6. Copy the code from config/apps-script.js into the script editor
7. Deploy as Web app with "Anyone" access
8. Create config/sheets-config.json with:
   {
     "webAppUrl": "your-apps-script-web-app-url",
     "spreadsheetId": "your-spreadsheet-id"
   }

For detailed instructions, see config/README.md
        `;
    }
}

// Initialize Sheets API
window.SheetsAPI = new SheetsAPI();

// Wire up the sync button in the header
document.addEventListener('DOMContentLoaded', () => {
    const syncBtn = document.getElementById('syncBtn');
    if (syncBtn) {
        syncBtn.onclick = () => window.SheetsAPI.doPendingPush();
    }
});
