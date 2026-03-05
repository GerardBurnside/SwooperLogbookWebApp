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

    async syncJump(jump) {
        if (!this.initialized) {
            console.log('Apps Script API not initialized, storing jump locally only');
            return;
        }

        this.updateSyncStatus('Syncing...');
        
        try {
            // Get equipment name from rig
            let equipmentName = jump.equipment;
            if (window.logbook) {
                const rig = window.logbook.equipmentRigs.find(eq => eq.id === jump.equipment);
                if (rig) {
                    equipmentName = rig.name;
                }
            }
            
            const jumpData = {
                jumpNumber: jump.jumpNumber,
                date: jump.date,
                location: jump.location,
                equipment: equipmentName,
                equipmentId: jump.equipment,  // preserve rig ID for round-trip
                notes: jump.notes,
                timestamp: jump.timestamp
            };
            
            const response = await fetch(this.webAppUrl, {
                method: 'POST',
                // Use text/plain to avoid CORS preflight (simple request)
                headers: {
                    'Content-Type': 'text/plain',
                },
                redirect: 'follow',
                body: JSON.stringify({ action: 'addJump', data: jumpData })
            });
            
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            
            const result = await response.json();
            if (result.error) {
                throw new Error(result.error);
            }
            
            this.updateSyncStatus('Synced');
            setTimeout(() => this.updateSyncStatus('Online'), 2000);
            
        } catch (error) {
            console.error('Failed to sync jump:', error);
            this.updateSyncStatus('Sync failed');
            setTimeout(() => this.updateSyncStatus('Online'), 3000);
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

    async syncWithSheet() {
        if (!this.initialized) {
            console.log('Cannot sync: API not initialized');
            return;
        }

        // Cancel any pending poll — will be rescheduled at the end of this sync
        this._cancelPoll();
        this.updateSyncStatus('Syncing...');

        try {
            // ── Equipment sync (lastModified-based direction) ─────────────────────────
            // syncEquipmentFromSheet compares timestamps internally and returns false
            // when local data is newer — in that case we push instead.
            const logbook = window.logbook;
            const localRigsEmpty = !logbook || !logbook.equipmentRigs || logbook.equipmentRigs.length === 0;

            if (localRigsEmpty) {
                // Fresh device — pull first to recover data, then push so sheet has IDs
                console.log('[Sync] Local rigs empty — pulling equipment from sheet...');
                const pulled = await this.syncEquipmentFromSheet();
                if (!pulled) await this.syncEquipmentToSheet();
            } else {
                // Let syncEquipmentFromSheet decide direction based on lastModified
                const pulled = await this.syncEquipmentFromSheet();
                if (!pulled) {
                    console.log('[Sync] Local equipment is newer — pushing to sheet...');
                    await this.syncEquipmentToSheet();
                }
            }

            // ── Jump sync (merge by jumpNumber) ──────────────────────────────────────
            const localJumps = JSON.parse(localStorage.getItem('skydiving-jumps')) || [];
            const sheetJumps = await this.getAllJumps();

            if (localJumps.length === 0 && sheetJumps.length === 0) {
                // Nothing to sync
            } else if (localJumps.length > 0 && sheetJumps.length === 0) {
                // Sheet is empty — push everything up
                await this.uploadAllJumps(localJumps);
            } else {
                // Merge: align by jumpNumber, surface conflicts
                const localMap  = new Map(localJumps.map(j => [j.jumpNumber, j]));
                const sheetMap  = new Map(sheetJumps.map(j => [j.jumpNumber, j]));
                const allNums   = new Set([...localMap.keys(), ...sheetMap.keys()]);

                const conflicts = [];
                const mergedMap = new Map();

                for (const num of allNums) {
                    const local = localMap.get(num);
                    const sheet = sheetMap.get(num);

                    if (!local) {
                        mergedMap.set(num, sheet);           // only on sheet
                    } else if (!sheet) {
                        mergedMap.set(num, local);           // only local
                    } else if (local.timestamp === sheet.timestamp) {
                        // Same creation timestamp.  The timestamp never changes after
                        // a jump is first logged, so two records with matching timestamps
                        // may still differ if notes/location/etc. were edited on another
                        // device.  In that case the sheet version is the authoritative
                        // one (it was explicitly synced from the editing device).
                        const sameContent = (local.notes    || '') === (sheet.notes    || '')
                                         && (local.location || '') === (sheet.location || '')
                                         && (local.date     || '') === (sheet.date     || '');
                        mergedMap.set(num, sameContent ? local : sheet);
                    } else {
                        // Same jump number, different content — ask user
                        conflicts.push({ jumpNumber: num, local, sheet });
                        mergedMap.set(num, local);           // tentative until dialog resolves
                    }
                }

                if (conflicts.length > 0) {
                    const resolutions = await this._showConflictDialog(conflicts);
                    resolutions.forEach(({ jumpNumber, chosen }) => mergedMap.set(jumpNumber, chosen));
                }

                // Sort and renumber sequentially from the base jump number
                const finalJumps = Array.from(mergedMap.values())
                    .sort((a, b) => a.jumpNumber - b.jumpNumber);
                const base = finalJumps.length > 0 ? finalJumps[0].jumpNumber : 1;
                finalJumps.forEach((j, i) => { j.jumpNumber = base + i; });

                // Persist merged list
                localStorage.setItem('skydiving-jumps', JSON.stringify(finalJumps));

                // Apply to live app state
                if (window.logbook) {
                    window.logbook.jumps = finalJumps;
                    if (finalJumps.length > 0) {
                        window.logbook.settings.startingJumpNumber = finalJumps[0].jumpNumber;
                        localStorage.setItem('skydiving-settings', JSON.stringify(window.logbook.settings));
                    }
                    window.logbook.initializeEquipmentJumpCounts();
                    window.logbook.updateStats();
                    window.logbook.renderJumpsList();
                    window.logbook.preFillFormWithLastJump();
                }

                // Upload the merged result if anything changed
                const uploadNeeded = conflicts.length > 0
                    || finalJumps.length !== sheetJumps.length
                    || localStorage.getItem('skydiving-needs-sync') === '1';
                if (uploadNeeded) await this.uploadAllJumps(finalJumps);
            }

            // Push equipment once more to capture any jump-count recalculations
            await this.syncEquipmentToSheet();

            // Clear the dirty flag
            localStorage.removeItem('skydiving-needs-sync');

            this.updateSyncStatus('Synced');
            setTimeout(() => this.updateSyncStatus('Online'), 2000);

            // Schedule the next automatic background poll
            this._schedulePoll();

        } catch (error) {
            console.error('Failed to sync with sheet:', error);
            this.updateSyncStatus('Sync failed');
            setTimeout(() => this.updateSyncStatus('Online'), 3000);
            this._schedulePoll(); // retry even on failure
        }
    }

    /**
     * Push all equipment components + settings to the Equipment sheet tab.
     * Called whenever the user saves any equipment change.
     */
    async syncEquipmentToSheet() {
        if (!this.initialized) return;

        const logbook = window.logbook;
        if (!logbook) return;

        // Safety: never push an empty rigs array – it would erase the sheet.
        // Omit the key so the Apps Script side leaves the existing data intact.
        const rigsToSend = (logbook.equipmentRigs && logbook.equipmentRigs.length > 0)
            ? logbook.equipmentRigs
            : undefined;

        // Embed a lastModified timestamp in the settings row so other devices
        // can compare on pull — no Apps Script changes required.
        const equipModified = new Date().toISOString();
        localStorage.setItem('skydiving-equipment-modified', equipModified);

        const payload = {
            harnesses:    logbook.harnesses,
            canopies:     logbook.canopies,
            settings:     { ...logbook.settings, _lastModified: equipModified },
            locations:    logbook.locations
        };
        if (rigsToSend) payload.rigs = rigsToSend;

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
     * Pull equipment components + settings from the Equipment sheet tab.
     * Returns true if usable data was fetched, false otherwise.
     */
    async syncEquipmentFromSheet() {
        if (!this.initialized) return false;

        try {
            const response = await fetch(this.webAppUrl + '?action=getEquipment', {
                method: 'GET',
                redirect: 'follow'
            });

            if (!response.ok) throw new Error(`HTTP ${response.status}`);

            const result = await response.json();
            if (result.error) throw new Error(result.error);

            const d = result.data || {};

            // Only apply if the sheet actually contains data
            const hasData = d.harnesses || d.canopies || d.rigs;
            if (!hasData) return false;

            // Compare lastModified timestamps (embedded in the settings row by
            // syncEquipmentToSheet). If local is newer, skip the pull so we don't
            // overwrite the user's unsynchronised changes.
            const sheetModified = d.settings?._lastModified || '';
            const localModified = localStorage.getItem('skydiving-equipment-modified') || '';
            if (sheetModified && localModified && localModified > sheetModified) {
                console.log('[Sync] Local equipment is newer than sheet — skipping pull');
                return false;
            }

            if (d.harnesses)  localStorage.setItem('skydiving-harnesses',       JSON.stringify(d.harnesses));
            if (d.canopies)   localStorage.setItem('skydiving-canopies',        JSON.stringify(d.canopies));
            if (d.rigs)       localStorage.setItem('skydiving-equipment-rigs',  JSON.stringify(d.rigs));
            if (d.locations)  localStorage.setItem('skydiving-locations',       JSON.stringify(d.locations));

            // Strip the internal _lastModified key before persisting settings locally
            if (d.settings) {
                const { _lastModified: sheetTs, ...cleanSettings } = d.settings;
                localStorage.setItem('skydiving-settings', JSON.stringify(cleanSettings));
                if (sheetTs) localStorage.setItem('skydiving-equipment-modified', sheetTs);
            }

            // Apply to live logbook instance
            const logbook = window.logbook;
            if (logbook) {
                if (d.harnesses)  logbook.harnesses    = d.harnesses;
                if (d.canopies)   logbook.canopies     = d.canopies;
                if (d.rigs)       logbook.equipmentRigs = d.rigs;
                if (d.settings) {
                    // Strip internal _lastModified before applying to the live object
                    const { _lastModified: _ts, ...cleanSettings } = d.settings;
                    logbook.settings = cleanSettings;
                }
                if (d.locations)  logbook.locations    = d.locations;

                logbook.updateEquipmentOptions();
                logbook.updateLocationDatalist();
                if (logbook.currentView === 'equipment') logbook.renderEquipmentView();
                // Re-fill the form since the dropdown was rebuilt
                logbook.preFillFormWithLastJump();
            }

            console.log('Equipment loaded from sheet');
            return true;
        } catch (error) {
            console.error('Failed to load equipment from sheet:', error);
            return false;
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

        // Push the restored data to the main Equipment sheet so it stays in sync
        await this.syncEquipmentToSheet();

        console.log('Equipment restored from backupRigs sheet');
        return true;
    }

    // Called after a local jump delete + renumber.  Because all jump numbers
    // shift, we re-upload the full current array as the source of truth.
    async syncAfterDelete(jumps) {
        if (!this.initialized || !navigator.onLine) return;

        this._cancelPoll();
        try {
            this.updateSyncStatus('Syncing...');
            localStorage.removeItem('skydiving-needs-sync');
            await this.uploadAllJumps(jumps);
            this.updateSyncStatus('Synced');
            setTimeout(() => this.updateSyncStatus('Online'), 2000);
        } catch (error) {
            console.error('Failed to sync delete with sheet:', error);
            this.updateSyncStatus('Sync failed');
            setTimeout(() => this.updateSyncStatus('Online'), 3000);
        } finally {
            this._schedulePoll();
        }
    }

    /** Schedule a background sync poll after `intervalMs` ms (default 2 min). */
    _schedulePoll(intervalMs = 120000) {
        this._cancelPoll();
        if (!this.initialized) return;
        this._pollTimer = setTimeout(() => {
            if (navigator.onLine) {
                console.log('[Poll] Auto-sync triggered');
                this.syncWithSheet(); // reschedules itself on completion
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

    /**
     * Show the conflict-resolution modal for a list of jump conflicts.
     * Returns a Promise that resolves with [{ jumpNumber, chosen }] after the
     * user selects which version to keep for each conflicting jump.
     */
    _showConflictDialog(conflicts) {
        return new Promise((resolve) => {
            const modal = document.getElementById('conflictModal');
            const list  = document.getElementById('conflictList');
            const btn   = document.getElementById('resolveConflictsBtn');

            list.innerHTML = '';

            conflicts.forEach(({ jumpNumber, local, sheet }) => {
                const localTs = new Date(local.timestamp).toLocaleString();
                const sheetTs = new Date(sheet.timestamp).toLocaleString();

                const item = document.createElement('div');
                item.className = 'conflict-item';
                item.dataset.jumpNumber = String(jumpNumber);
                item.innerHTML = `
                    <h4>Jump #${jumpNumber}</h4>
                    <div class="conflict-options">
                        <div class="conflict-option selected" data-side="local">
                            <label>This device</label>
                            <div class="conflict-detail">${local.date} &middot; ${local.location || '&mdash;'}</div>
                            ${local.notes ? `<div class="conflict-detail">${local.notes}</div>` : ''}
                            <div class="conflict-ts">Saved ${localTs}</div>
                        </div>
                        <div class="conflict-option" data-side="sheet">
                            <label>Other device</label>
                            <div class="conflict-detail">${sheet.date} &middot; ${sheet.location || '&mdash;'}</div>
                            ${sheet.notes ? `<div class="conflict-detail">${sheet.notes}</div>` : ''}
                            <div class="conflict-ts">Saved ${sheetTs}</div>
                        </div>
                    </div>`;

                item.querySelectorAll('.conflict-option').forEach(opt => {
                    opt.addEventListener('click', () => {
                        item.querySelectorAll('.conflict-option').forEach(o => o.classList.remove('selected'));
                        opt.classList.add('selected');
                    });
                });

                list.appendChild(item);
            });

            modal.style.display = 'block';

            const onResolve = () => {
                btn.removeEventListener('click', onResolve);
                modal.style.display = 'none';

                const resolutions = conflicts.map(({ jumpNumber, local, sheet }) => {
                    const item     = list.querySelector(`[data-jump-number="${jumpNumber}"]`);
                    const selected = item?.querySelector('.conflict-option.selected');
                    const side     = selected?.dataset.side || 'local';
                    return { jumpNumber, chosen: side === 'sheet' ? sheet : local };
                });

                resolve(resolutions);
            };

            btn.addEventListener('click', onResolve);
        });
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
        syncBtn.onclick = () => window.SheetsAPI.syncWithSheet();
    }
});