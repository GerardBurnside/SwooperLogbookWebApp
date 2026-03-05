// Google Sheets API Integration via Apps Script
class SheetsAPI {
    constructor() {
        this.webAppUrl = ''; // Will be loaded from config
        this.spreadsheetId = ''; // Will be loaded from config
        this.initialized = false;
        this._pollTimer = null; // handle for the 2-minute background poll
        
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
        // Try to load from config file, fallback to localStorage
        try {
            const response = await fetch('./config/sheets-config.json');
            if (response.ok) {
                const fileConfig = await response.json();
                // Also persist to localStorage so it survives offline / missing file
                localStorage.setItem('sheets-config', JSON.stringify(fileConfig));
                return fileConfig;
            }
        } catch (error) {
            console.log('Config file not found, checking localStorage');
        }
        
        // Fallback to localStorage
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
            return {
                id,
                jumpNumber: parseInt(row[0]) || 0,
                date: row[1] || '',
                location: row[2] || '',
                equipment,
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
            const sheetTs       = d.dataModified || '';
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
                    console.warn('[Startup] Conflict — sheet is newer, pending local changes discarded');
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
                await this.pushAllWithGuard();
            }

            this.updateSyncStatus('Online');
        } catch (error) {
            console.error('[Startup] Sync failed:', error);
            this.updateSyncStatus('Sync failed');
            setTimeout(() => this.updateSyncStatus('Online'), 3000);
        } finally {
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
            const sheetTs     = d.dataModified || '';
            const localSynced = localStorage.getItem('skydiving-data-synced') || '';

            if (sheetTs && sheetTs > localSynced) {
                // Conflict: sheet was updated since our last sync — pull & notify
                console.warn('[Sync] Conflict detected — sheet is newer, aborting push');
                await this._pullAllFromSheet(d, sheetTs);
                const logbook = window.logbook;
                if (logbook) logbook.showSyncConflictModal();
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

            this.updateSyncStatus('Synced');
            setTimeout(() => this.updateSyncStatus('Online'), 2000);
            console.log('[Sync] Push complete, ts:', newTs);
        } catch (error) {
            console.error('[Sync] pushAllWithGuard failed:', error);
            this.updateSyncStatus('Sync failed');
            setTimeout(() => this.updateSyncStatus('Online'), 3000);
        }
    }

    /**
     * Apply sheet equipment data + jumps to local storage and the live logbook.
     * Sets both timestamp keys so local state is clean and in sync with the sheet.
     */
    async _pullAllFromSheet(d, sheetDataModified) {
        const logbook = window.logbook;

        // Persist equipment to localStorage
        if (d.harnesses)  localStorage.setItem('skydiving-harnesses',      JSON.stringify(d.harnesses));
        if (d.canopies)   localStorage.setItem('skydiving-canopies',       JSON.stringify(d.canopies));
        if (d.rigs)       localStorage.setItem('skydiving-equipment-rigs', JSON.stringify(d.rigs));
        if (d.locations)  localStorage.setItem('skydiving-locations',      JSON.stringify(d.locations));
        if (d.settings)   localStorage.setItem('skydiving-settings',       JSON.stringify(d.settings));

        // Fetch and persist jumps from sheet
        const sheetJumps = await this.getAllJumps();
        localStorage.setItem('skydiving-jumps', JSON.stringify(sheetJumps));

        // Stamp both timestamp keys — local is now in sync with the sheet
        const ts = sheetDataModified || new Date().toISOString();
        localStorage.setItem('skydiving-data-modified', ts);
        localStorage.setItem('skydiving-data-synced', ts);

        // Apply to live logbook instance
        if (logbook) {
            if (d.harnesses)  logbook.harnesses     = d.harnesses;
            if (d.canopies)   logbook.canopies      = d.canopies;
            if (d.rigs)       logbook.equipmentRigs = d.rigs;
            if (d.settings)   logbook.settings      = d.settings;
            if (d.locations)  logbook.locations     = d.locations;

            logbook.jumps = sheetJumps;
            logbook.initializeEquipmentJumpCounts();
            logbook.updateEquipmentOptions();
            logbook.updateLocationDatalist();
            logbook.updateStats();
            logbook.renderJumpsList();
            if (logbook.currentView === 'equipment') logbook.renderEquipmentView();
            logbook.preFillFormWithLastJump();
        }

        console.log('[Sync] Pulled all data from sheet, ts:', ts);
    }

    /**
     * Called by the background poll and the online-event handler.
     * Pushes if there are pending local changes, otherwise checks quietly
     * whether the sheet is newer and pulls if so.
     */
    async doPendingPush() {
        if (!this.initialized || !navigator.onLine) return;

        const localModified = localStorage.getItem('skydiving-data-modified') || '';
        const localSynced   = localStorage.getItem('skydiving-data-synced') || '';

        if (localModified && localModified > localSynced) {
            await this.pushAllWithGuard();
            return;
        }

        // No pending changes — quietly check if the sheet is newer
        try {
            const response = await fetch(this.webAppUrl + '?action=getEquipment', {
                method: 'GET',
                redirect: 'follow'
            });
            if (!response.ok) return;

            const result = await response.json();
            if (result.error) return;

            const d       = result.data || {};
            const sheetTs = d.dataModified || '';
            if (sheetTs && sheetTs > localSynced) {
                console.log('[Poll] Sheet is newer, pulling...');
                this.updateSyncStatus('Syncing...');
                await this._pullAllFromSheet(d, sheetTs);
                this.updateSyncStatus('Synced');
                setTimeout(() => this.updateSyncStatus('Online'), 2000);
            }
        } catch (error) {
            console.warn('[Poll] doPendingPush check failed:', error);
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

        // Safety: never push an empty rigs array – it would erase the sheet.
        // Omit the key so the Apps Script side leaves the existing data intact.
        const rigsToSend = (logbook.equipmentRigs && logbook.equipmentRigs.length > 0)
            ? logbook.equipmentRigs
            : undefined;

        const payload = {
            harnesses:    logbook.harnesses,
            canopies:     logbook.canopies,
            settings:     logbook.settings,
            locations:    logbook.locations
        };
        if (rigsToSend)    payload.rigs          = rigsToSend;
        if (dataModified)  payload.dataModified  = dataModified;

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

        // Overwrite localStorage
        if (d.harnesses)  localStorage.setItem('skydiving-harnesses',       JSON.stringify(d.harnesses));
        if (d.canopies)   localStorage.setItem('skydiving-canopies',        JSON.stringify(d.canopies));
        if (d.rigs)       localStorage.setItem('skydiving-equipment-rigs',  JSON.stringify(d.rigs));
        if (d.settings)   localStorage.setItem('skydiving-settings',        JSON.stringify(d.settings));
        if (d.locations)  localStorage.setItem('skydiving-locations',       JSON.stringify(d.locations));

        // Apply to live logbook instance
        const logbook = window.logbook;
        if (logbook) {
            if (d.harnesses)  logbook.harnesses     = d.harnesses;
            if (d.canopies)   logbook.canopies      = d.canopies;
            if (d.rigs)       logbook.equipmentRigs  = d.rigs;
            if (d.settings)   logbook.settings       = d.settings;
            if (d.locations)  logbook.locations       = d.locations;

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
            // Get equipment name from rig
            let equipmentName = jump.equipment;
            if (window.logbook) {
                const rig = window.logbook.equipmentRigs.find(eq => eq.id === jump.equipment);
                if (rig) {
                    equipmentName = rig.name;
                }
            }
            
            return {
                jumpNumber: jump.jumpNumber,
                date: jump.date,
                location: jump.location,
                equipment: equipmentName,
                equipmentId: jump.equipment,  // preserve rig ID for round-trip
                notes: jump.notes,
                timestamp: jump.timestamp
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
