// Google Sheets API Integration via Apps Script
class SheetsAPI {
    constructor() {
        this.webAppUrl = ''; // Will be loaded from config
        this.spreadsheetId = ''; // Will be loaded from config
        this.initialized = false;
        
        this.setupAPI();
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
            // Get equipment name from combination
            let equipmentName = jump.equipment;
            if (window.logbook) {
                const combination = window.logbook.equipmentCombinations.find(eq => eq.id === jump.equipment);
                if (combination) {
                    equipmentName = combination.name;
                }
            }
            
            const jumpData = {
                jumpNumber: jump.jumpNumber,
                date: jump.date,
                location: jump.location,
                equipment: equipmentName,
                equipmentId: jump.equipment,  // preserve combination ID for round-trip
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
            // col 7 (row[6]) holds the combination ID written by the updated script;
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

        this.updateSyncStatus('Syncing...');
        
        try {
            // --- Sync equipment first so jump display names resolve correctly ---
            const equipmentSynced = await this.syncEquipmentFromSheet();

            // Get local jumps
            const localJumps = JSON.parse(localStorage.getItem('skydiving-jumps')) || [];
            
            // Get data from sheet
            const sheetJumps = await this.getAllJumps();
            
            // If we have local jumps but sheet is empty, upload local jumps
            if (localJumps.length > 0 && sheetJumps.length === 0) {
                await this.uploadAllJumps(localJumps);
            }
            // If sheet has data, use it as source of truth (download)
            else if (sheetJumps.length > 0) {
                localStorage.setItem('skydiving-jumps', JSON.stringify(sheetJumps));
                
                if (window.logbook) {
                    window.logbook.jumps = sheetJumps;

                    // Seed startingJumpNumber from the sheet's lowest jump number so
                    // renumbering after future adds/deletes continues from the right base.
                    const minFromSheet = Math.min(...sheetJumps.map(j => j.jumpNumber));
                    if (Number.isFinite(minFromSheet) && minFromSheet > 0) {
                        window.logbook.settings.startingJumpNumber = minFromSheet;
                        localStorage.setItem('skydiving-settings', JSON.stringify(window.logbook.settings));
                    }

                    // Recalculate equipment jump counts since jump data was replaced
                    window.logbook.initializeEquipmentJumpCounts();
                    window.logbook.updateStats();
                    window.logbook.renderJumpsList();
                }
            }
            // If both are empty, nothing to sync
            
            this.updateSyncStatus('Synced');
            setTimeout(() => this.updateSyncStatus('Online'), 2000);
            
        } catch (error) {
            console.error('Failed to sync with sheet:', error);
            this.updateSyncStatus('Sync failed');
            setTimeout(() => this.updateSyncStatus('Online'), 3000);
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

        const payload = {
            rigs:         logbook.rigs,
            canopies:     logbook.canopies,
            linesets:     logbook.linesets,
            combinations: logbook.equipmentCombinations,
            settings:     logbook.settings,
            locations:    logbook.locations
        };

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
            const hasData = d.rigs || d.canopies || d.linesets || d.combinations;
            if (!hasData) return false;

            if (d.rigs)         localStorage.setItem('skydiving-rigs',                   JSON.stringify(d.rigs));
            if (d.canopies)     localStorage.setItem('skydiving-canopies',               JSON.stringify(d.canopies));
            if (d.linesets)     localStorage.setItem('skydiving-linesets',               JSON.stringify(d.linesets));
            if (d.combinations) localStorage.setItem('skydiving-equipment-combinations', JSON.stringify(d.combinations));
            if (d.settings)     localStorage.setItem('skydiving-settings',               JSON.stringify(d.settings));
            if (d.locations)    localStorage.setItem('skydiving-locations',              JSON.stringify(d.locations));

            // Apply to live logbook instance
            const logbook = window.logbook;
            if (logbook) {
                if (d.rigs)         logbook.rigs                  = d.rigs;
                if (d.canopies)     logbook.canopies              = d.canopies;
                if (d.linesets)     logbook.linesets              = d.linesets;
                if (d.combinations) logbook.equipmentCombinations = d.combinations;
                if (d.settings)     logbook.settings              = d.settings;
                if (d.locations)    logbook.locations             = d.locations;

                logbook.updateEquipmentOptions();
                logbook.updateLocationDatalist();
                if (logbook.currentView === 'equipment') logbook.renderEquipmentView();
            }

            console.log('Equipment loaded from sheet');
            return true;
        } catch (error) {
            console.error('Failed to load equipment from sheet:', error);
            return false;
        }
    }
    
    // Called after a local jump delete + renumber.  Because all jump numbers
    // shift, we re-upload the full current array as the source of truth.
    async syncAfterDelete(jumps) {
        if (!this.initialized || !navigator.onLine) return;

        try {
            this.updateSyncStatus('Syncing...');
            await this.uploadAllJumps(jumps);
            this.updateSyncStatus('Synced');
            setTimeout(() => this.updateSyncStatus('Online'), 2000);
        } catch (error) {
            console.error('Failed to sync delete with sheet:', error);
            this.updateSyncStatus('Sync failed');
            setTimeout(() => this.updateSyncStatus('Online'), 3000);
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
            // Get equipment name from combination
            let equipmentName = jump.equipment;
            if (window.logbook) {
                const combination = window.logbook.equipmentCombinations.find(eq => eq.id === jump.equipment);
                if (combination) {
                    equipmentName = combination.name;
                }
            }
            
            return {
                jumpNumber: jump.jumpNumber,
                date: jump.date,
                location: jump.location,
                equipment: equipmentName,
                equipmentId: jump.equipment,  // preserve combination ID for round-trip
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
        if (syncElement) {
            syncElement.textContent = status;
            
            // Update classes based on status
            syncElement.className = 'sync-status';
            if (status === 'Syncing...') {
                syncElement.classList.add('syncing');
            } else if (status === 'Synced' || status === 'Online') {
                syncElement.classList.add('success');
            } else if (status.includes('failed') || status.includes('error')) {
                syncElement.classList.add('error');
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

// Add manual sync button functionality
document.addEventListener('DOMContentLoaded', () => {
    // Add sync button to footer if not exists
    if (!document.getElementById('syncBtn')) {
        const footer = document.querySelector('footer');
        const syncBtn = document.createElement('button');
        syncBtn.id = 'syncBtn';
        syncBtn.className = 'btn-secondary';
        syncBtn.textContent = 'Sync';
        syncBtn.onclick = () => window.SheetsAPI.syncWithSheet();
        footer.insertBefore(syncBtn, footer.firstChild);
    }
});