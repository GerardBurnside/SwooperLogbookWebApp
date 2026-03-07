// Swooper Logbook App - Main Application Logic
class SkydivingLogbook {
    constructor() {
        this.jumps = JSON.parse(localStorage.getItem('skydiving-jumps')) || [];
        this.settings = JSON.parse(localStorage.getItem('skydiving-settings')) || {
            startingJumpNumber: 1,
            recentJumpsDays: 3,
            standardRedThreshold: 160,
            standardOrangeThreshold: 140,
            hybridRedThreshold: 80,
            hybridOrangeThreshold: 60
        };
        // Backfill for existing saved settings that predate these fields
        if (this.settings.recentJumpsDays === undefined) {
            this.settings.recentJumpsDays = 3;
        }
        if (this.settings.standardRedThreshold === undefined) {
            this.settings.standardRedThreshold = 160;
        }
        if (this.settings.standardOrangeThreshold === undefined) {
            this.settings.standardOrangeThreshold = 140;
        }
        if (this.settings.hybridRedThreshold === undefined) {
            this.settings.hybridRedThreshold = 80;
        }
        if (this.settings.hybridOrangeThreshold === undefined) {
            this.settings.hybridOrangeThreshold = 60;
        }
        
        // Component-based equipment system
        this.harnesses = JSON.parse(localStorage.getItem('skydiving-harnesses')) || [
            { id: 'javelin', name: 'Javelin' },
            { id: 'mutant', name: 'Mutant' }
        ];
        this.canopies = JSON.parse(localStorage.getItem('skydiving-canopies')) || [
            { id: 'petra64', name: 'Petra64' },
            { id: 'petra68', name: 'Petra68' }
        ];
        this.equipmentRigs = JSON.parse(localStorage.getItem('skydiving-equipment-rigs')) || [];
        this.locations = JSON.parse(localStorage.getItem('skydiving-locations')) || [
            { id: 'loc_lfoz',    name: 'LFOZ Epcol',               lat: null, lng: null },
            { id: 'loc_klatovy', name: 'Klatovy',                   lat: null, lng: null },
            { id: 'loc_ravenna', name: 'Ravenna',                   lat: null, lng: null },
            { id: 'loc_palm',    name: 'The Palm Skydive Dubai',    lat: null, lng: null },
            { id: 'loc_desert',  name: 'Desert Skydive Dubai',      lat: null, lng: null }
        ];
        // Backfill lat/lng on locations loaded from older data
        this.locations.forEach(loc => {
            if (loc.lat === undefined) loc.lat = null;
            if (loc.lng === undefined) loc.lng = null;
        });
        
        // Initialize jump counts for equipment rigs
        this.initializeEquipmentJumpCounts();
        
        this.currentView = 'jumps'; // 'jumps', 'equipment', 'stats'
        this.equipmentSubView = 'rigs'; // 'rigs', 'harnesses', 'canopies', 'linesets', 'locations'
        this.showArchivedStats = false;
        
        this.init();
    }

    init() {
        this.setupEventListeners();
        this.updateStats();
        this.renderJumpsList();
        this.setCurrentDate();
        this.updateEquipmentOptions();
        this.setupLocationAutocomplete();
        this.preFillFormWithLastJump();
        this.showView('jumps');
        
        // Kick off background geocoding for any location missing coordinates
        this.geocodeAllLocations();
        
        // Check if we're online/offline
        this.updateOnlineStatus();
        window.addEventListener('online', () => {
            this.updateOnlineStatus();
            // Flush any pending writes as soon as connectivity returns
            if (window.SheetsAPI?.initialized) {
                window.SheetsAPI.syncWithSheet();
            }
        });
        window.addEventListener('offline', () => this.updateOnlineStatus());
        
        // Auto-sync equipment from Google Sheets on startup
        // Wait for SheetsAPI.ready (promise) instead of a fragile timeout
        this.autoSyncEquipmentOnStartup();
    }

    async autoSyncEquipmentOnStartup() {
        if (!navigator.onLine) return;
        if (!window.SheetsAPI) return;

        // Wait until the API has finished loading its config
        await window.SheetsAPI.ready;

        if (!window.SheetsAPI.initialized) return;

        // syncWithSheet handles equipment direction, jump merge, dirty flag, and
        // schedules the 2-minute background poll on completion.
        await window.SheetsAPI.syncWithSheet();
    }

    setupEventListeners() {
        const getValidMultiplier = () => {
            const multiplierInput = document.getElementById('jumpMultiplier');
            const parsed = parseInt(multiplierInput.value, 10);
            if (Number.isNaN(parsed)) return 1;
            return Math.max(1, Math.min(99, parsed));
        };

        // Jump form submission
        document.getElementById('jumpForm').addEventListener('submit', (e) => {
            e.preventDefault();
            const multiplierInput = document.getElementById('jumpMultiplier');
            const multiplier = getValidMultiplier();
            multiplierInput.value = multiplier;
            const form = document.getElementById('jumpForm');
            const jumpData = {
                date: form.elements['date'].value,
                location: form.elements['location'].value,
                equipment: form.elements['equipment'].value,
                notes: form.elements['notes'].value || ''
            };
            for (let i = 0; i < multiplier; i++) {
                this.addJump(jumpData, multiplier > 1);
            }
            // Reset multiplier back to 1
            document.getElementById('jumpMultiplier').value = 1;
            // Reset form after all jumps logged
            form.reset();
            this.setCurrentDate();
            this.preFillFormWithLastJump();
            if (multiplier > 1) {
                this.showMessage(`${multiplier} jumps logged successfully!`, 'success');
                // Sync all multiplier jumps at once
                if (navigator.onLine && window.SheetsAPI?.initialized) {
                    window.SheetsAPI.pushAllWithGuard();
                }
            }
        });

        // Multiplier widget buttons
        document.getElementById('multiplierUp').addEventListener('click', () => {
            const input = document.getElementById('jumpMultiplier');
            const val = getValidMultiplier();
            if (val < 99) input.value = val + 1;
        });
        document.getElementById('multiplierDown').addEventListener('click', () => {
            const input = document.getElementById('jumpMultiplier');
            const val = getValidMultiplier();
            if (val > 1) input.value = val - 1;
        });

        // Allow manual entry while enforcing numeric bounds.
        document.getElementById('jumpMultiplier').addEventListener('blur', () => {
            document.getElementById('jumpMultiplier').value = getValidMultiplier();
        });

        // Settings modal
        document.getElementById('settingsBtn').addEventListener('click', () => {
            this.openSettingsModal();
        });

        document.getElementById('saveSettings').addEventListener('click', () => {
            this.saveSettings();
        });

        document.getElementById('restoreFromBackupBtn').addEventListener('click', () => {
            this.restoreEquipmentFromBackup();
        });

        document.getElementById('useCurrentLocationBtn').addEventListener('click', () => {
            this.setComponentCoordsFromGPS();
        });

        // Google Sheets Integration modal
        document.getElementById('googleSheetsIntegrationBtn').addEventListener('click', () => {
            this.openSheetsModal();
        });

        document.getElementById('sheetsClose').addEventListener('click', () => {
            this.closeSheetsModal();
        });

        document.getElementById('saveSheetsConfig').addEventListener('click', () => {
            this.saveSheetsConfig();
        });

        // Modal close
        document.getElementById('settingsClose').addEventListener('click', () => {
            this.closeModal();
        });

        // Export data
        document.getElementById('exportBtn').addEventListener('click', () => {
            this.exportData();
        });

        // Auto-detect DZ checkbox in the log jump form
        const autoDetectChk = document.getElementById('autoDetectDZForm');
        if (autoDetectChk) {
            autoDetectChk.addEventListener('change', () => {
                if (autoDetectChk.checked) this.detectNearestLocation(true);
            });
        }

        // Share backup via native share sheet (e.g. Gmail on Android)
        document.getElementById('shareBtn').addEventListener('click', () => {
            this.shareDataViaEmail();
        });
        const shareBtn = document.getElementById('shareBtn');
        const canShareFiles = !!navigator.share && (
            !navigator.canShare || navigator.canShare({
                files: [new File(['{}'], 'share-test.json', { type: 'application/json' })]
            })
        );
        if (!canShareFiles) {
            shareBtn.style.display = 'none';
        }

        // Import data
        document.getElementById('importBtn').addEventListener('click', () => {
            document.getElementById('importFileInput').click();
        });
        document.getElementById('importFileInput').addEventListener('change', (e) => {
            this.importData(e);
        });

        // Navigation buttons
        document.getElementById('jumpsViewBtn').addEventListener('click', () => {
            this.showView('jumps');
        });

        document.getElementById('equipmentViewBtn').addEventListener('click', () => {
            this.showView('equipment');
        });

        document.getElementById('statsViewBtn').addEventListener('click', () => {
            this.showView('stats');
        });

        // Equipment management
        document.getElementById('addEquipmentBtn').addEventListener('click', () => {
            this.addEquipment();
        });

        document.getElementById('equipmentForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.saveEquipment();
        });
        
        // Component management
        document.getElementById('addHarnessBtn').addEventListener('click', () => {
            this.addComponent('harness');
        });
        
        document.getElementById('addCanopyBtn').addEventListener('click', () => {
            this.addComponent('canopy');
        });
        
        document.getElementById('addLocationBtn').addEventListener('click', () => {
            this.addComponent('location');
        });
        
        document.getElementById('componentForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.saveComponent();
        });

        // Lineset stepper buttons
        document.getElementById('linesetUp').addEventListener('click', () => {
            const input = document.getElementById('equipmentLinesetNumber');
            input.value = (parseInt(input.value) || 1) + 1;
        });
        document.getElementById('linesetDown').addEventListener('click', () => {
            const input = document.getElementById('equipmentLinesetNumber');
            const val = parseInt(input.value) || 1;
            if (val > 1) input.value = val - 1;
        });

        // Auto-fill lineset number and rig notes when canopy changes
        document.getElementById('equipmentCanopy').addEventListener('change', () => {
            if (!document.getElementById('equipmentId').value) {
                this.autoFillLinesetNumber();
            }
            this.autoFillRigNotes();
        });

        // Auto-fill rig notes when harness changes
        document.getElementById('equipmentHarness').addEventListener('change', () => {
            this.autoFillRigNotes();
        });

        // Equipment sub-navigation
        document.querySelectorAll('.equipment-nav-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                this.showEquipmentSubView(e.target.dataset.view);
            });
        });

        // Close modal when clicking outside
        window.addEventListener('click', (e) => {
            const settingsModal = document.getElementById('settingsModal');
            const equipmentModal = document.getElementById('equipmentModal');
            const componentModal = document.getElementById('componentModal');
            const sheetsModal = document.getElementById('sheetsModal');
            if (e.target === settingsModal) {
                this.closeModal();
            }
            if (e.target === equipmentModal) {
                this.closeEquipmentModal();
            }
            if (e.target === componentModal) {
                this.closeComponentModal();
            }
            if (e.target === sheetsModal) {
                this.closeSheetsModal();
            }
        });
    }

    setCurrentDate() {
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('date').value = today;
    }

    getLastJumpData() {
        if (this.jumps.length === 0) {
            return null;
        }
        
        // Get the most recent jump (last in sorted array)
        return this.jumps[this.jumps.length - 1];
    }

    preFillFormWithLastJump() {
        const lastJump = this.getLastJumpData();
        if (!lastJump) {
            return;
        }
        
        // Pre-fill location
        const locationInput = document.getElementById('location');
        if (lastJump.location && locationInput) {
            locationInput.value = lastJump.location;
        }
        
        // Pre-fill equipment
        const equipmentSelect = document.getElementById('equipment');
        if (lastJump.equipment && equipmentSelect) {
            // Check if the equipment rig still exists and is not archived
            const equipment = this.equipmentRigs.find(eq => eq.id === lastJump.equipment && !eq.archived);
            if (equipment) {
                equipmentSelect.value = lastJump.equipment;
            }
        }
    }

    updateNextJumpNumber() { /* field removed — kept as no-op for safety */ }

    getNextJumpNumber() {
        if (this.jumps.length === 0) return this.settings.startingJumpNumber;
        return Math.max(...this.jumps.map(j => j.jumpNumber)) + 1;
    }

    addJump(jumpData = null, silent = false) {
        const form = document.getElementById('jumpForm');
        
        // Use passed data or read from form
        const data = jumpData || {
            date: form.elements['date'].value,
            location: form.elements['location'].value,
            equipment: form.elements['equipment'].value,
            notes: form.elements['notes'].value || ''
        };
        
        // Remember the highest jump number before insertion to detect a past-date entry
        const maxBefore = this.jumps.length > 0
            ? Math.max(...this.jumps.map(j => j.jumpNumber))
            : this.settings.startingJumpNumber - 1;

        const jump = {
            id: Date.now() + Math.random(),
            jumpNumber: 0,          // assigned below by renumberJumps()
            date: data.date,
            location: data.location,
            equipment: data.equipment,
            notes: data.notes,
            timestamp: new Date().toISOString()
        };

        // Auto-add new location if it doesn't exist yet
        if (jump.location) {
            const locationExists = this.locations.some(
                loc => loc.name.toLowerCase() === jump.location.toLowerCase()
            );
            if (!locationExists) {
                const newId = 'loc_' + Date.now();
                const newLoc = { id: newId, name: jump.location, lat: null, lng: null };
                this.locations.push(newLoc);
                this.saveComponentsToLocalStorage();
                this.updateLocationDatalist();
                this.geocodeLocation(newLoc);
            }
        }

        // Update equipment jump count if equipment is selected
        if (jump.equipment) {
            const equipment = this.equipmentRigs.find(eq => eq.id === jump.equipment);
            if (equipment) {
                equipment.jumpCount = (equipment.jumpCount || 0) + 1;
                this.saveComponentsToLocalStorage();
            }
        }
        
        // Insert then renumber everything chronologically by date
        this.jumps.push(jump);
        this.renumberJumps(); // sorts by date, assigns numbers from startingJumpNumber
        
        // Save to localStorage
        this.saveToLocalStorage();
        
        // Update UI
        this.updateStats();
        this.renderJumpsList();
        
        // Re-render equipment view if currently displayed
        if (this.currentView === 'equipment') {
            this.renderEquipmentView();
        }
        
        // Reset form only for single jumps (multiplier handles its own reset)
        if (!silent) {
            form.reset();
            this.setCurrentDate();
            this.preFillFormWithLastJump();
        }
        
        // Inform user; note if past-date renumbering happened
        const isPastJump = jump.jumpNumber <= maxBefore;
        if (!silent) {
            this.showMessage(
                isPastJump
                    ? `Jump #${jump.jumpNumber} logged — subsequent jumps renumbered`
                    : 'Jump logged successfully!',
                'success'
            );
        }
        
        // Sync to Google Sheets (skip during multiplier batch — caller will sync once)
        if (!silent && navigator.onLine && window.SheetsAPI?.initialized) {
            window.SheetsAPI.pushAllWithGuard();
        }
    }

    updateStats() {
        const latestJumpNumber = this.jumps.length > 0 ? Math.max(...this.jumps.map(jump => jump.jumpNumber)) : 0;
        document.getElementById('totalJumps').textContent = `Latest Jump: #${latestJumpNumber}`;
    }

    renderJumpsList() {
        const jumpsList = document.getElementById('jumpsList');
        
        if (this.jumps.length === 0) {
            jumpsList.innerHTML = '<p class="no-jumps">No jumps logged yet. Add your first jump above!</p>';
            return;
        }

        // Show most recent jumps first
        const sortedJumps = [...this.jumps].sort((a, b) => b.jumpNumber - a.jumpNumber);

        // Cutoff: configurable number of days ago at midnight
        const cutoff = new Date();
        cutoff.setHours(0, 0, 0, 0);
        cutoff.setDate(cutoff.getDate() - (this.settings.recentJumpsDays || 3));

        const recentJumps = [];
        const olderJumps = [];

        sortedJumps.forEach(jump => {
            const jumpDate = new Date(jump.date);
            if (jumpDate >= cutoff) {
                recentJumps.push(jump);
            } else {
                olderJumps.push(jump);
            }
        });

        let html = '';

        // Render recent jumps individually
        if (recentJumps.length > 0) {
            html += recentJumps.map(jump => this.createJumpHTML(jump)).join('');
        }

        // Group older jumps by month
        if (olderJumps.length > 0) {
            const monthGroups = new Map();
            olderJumps.forEach(jump => {
                const d = new Date(jump.date);
                const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
                if (!monthGroups.has(key)) {
                    monthGroups.set(key, { label: d.toLocaleDateString(undefined, { year: 'numeric', month: 'long' }), jumps: [] });
                }
                monthGroups.get(key).jumps.push(jump);
            });

            // Keys are already in descending order (sorted input)
            for (const [key, group] of monthGroups) {
                const jumpCount = group.jumps.length;
                html += `
                    <div class="month-group" data-month="${key}">
                        <div class="month-group-header" onclick="logbook.toggleMonthGroup('${key}')">
                            <span class="month-group-arrow" id="arrow-${key}">&#9654;</span>
                            <span class="month-group-label">${group.label}</span>
                            <span class="month-group-count">${jumpCount} jump${jumpCount !== 1 ? 's' : ''}</span>
                        </div>
                        <div class="month-group-body" id="month-${key}" style="display:none;">
                            ${group.jumps.map(jump => this.createJumpHTML(jump)).join('')}
                        </div>
                    </div>
                `;
            }
        }

        jumpsList.innerHTML = html;
    }

    toggleMonthGroup(key) {
        const body = document.getElementById('month-' + key);
        const arrow = document.getElementById('arrow-' + key);
        if (!body) return;
        const open = body.style.display !== 'none';
        body.style.display = open ? 'none' : 'block';
        if (arrow) arrow.innerHTML = open ? '&#9654;' : '&#9660;';
    }

    createJumpHTML(jump) {
        const date = new Date(jump.date).toLocaleDateString();
        let equipmentName = jump.equipment;
        
        // Try to find the equipment rig for better display
        const rig = this.equipmentRigs.find(eq => eq.id === jump.equipment);
        if (rig) {
            const harness = this.harnesses.find(r => r.id === rig.harnessId);
            const canopy = this.canopies.find(c => c.id === rig.canopyId);
            
            if (harness && canopy && rig.linesetNumber) {
                const hybridSuffix = /-Hybrid$/i.test(rig.name || '') ? ' (Hybrid)' : '';
                equipmentName = `${harness.name} + ${canopy.name} + Lineset#${rig.linesetNumber}${hybridSuffix}`;
            } else {
                equipmentName = rig.name;
            }
        }

        return `
            <div class="jump-item">
                <div class="jump-header">
                    <span class="jump-number">#${jump.jumpNumber}</span>
                    <span class="jump-date">${date}</span>
                    <button class="delete-jump-btn" onclick="logbook.deleteJump('${jump.id}')" title="Delete jump">❌</button>
                </div>
                <div class="jump-details">
                    <div class="jump-location">📍 ${jump.location}</div>
                    <div class="jump-equipment">🎒 ${equipmentName}</div>
                    ${jump.notes ? `<div class="jump-notes">💭 ${jump.notes}</div>` : ''}
                </div>
            </div>
        `;
    }

    deleteJump(jumpId) {
        if (!confirm('Are you sure you want to delete this jump? This action cannot be undone.')) {
            return;
        }
        
        // Find the jump to delete
        const jumpIndex = this.jumps.findIndex(jump => jump.id.toString() === jumpId.toString());
        if (jumpIndex === -1) {
            this.showMessage('Jump not found', 'error');
            return;
        }
        
        const deletedJump = this.jumps[jumpIndex];
        
        // Update equipment jump count if equipment was selected
        if (deletedJump.equipment) {
            const equipment = this.equipmentRigs.find(eq => eq.id === deletedJump.equipment);
            if (equipment && equipment.jumpCount > 0) {
                equipment.jumpCount = equipment.jumpCount - 1;
                this.saveComponentsToLocalStorage();
            }
        }

        // Record deletion as a tombstone so other devices can sync the deletion.
        // Tombstones are stored in settings (which sync bidirectionally via Equipment sheet).
        if (deletedJump.timestamp) {
            if (!Array.isArray(this.settings.deletedJumpTimestamps)) {
                this.settings.deletedJumpTimestamps = [];
            }
            if (!this.settings.deletedJumpTimestamps.includes(deletedJump.timestamp)) {
                this.settings.deletedJumpTimestamps.push(deletedJump.timestamp);
            }
            localStorage.setItem('skydiving-settings', JSON.stringify(this.settings));
            this.markEquipmentModified();
        }
        
        // Remove the jump
        this.jumps.splice(jumpIndex, 1);
        
        // Renumber all jumps
        this.renumberJumps();
        
        // Save to localStorage
        this.saveToLocalStorage();
        
        // Update UI
        this.updateStats();
        this.renderJumpsList();
        
        // Re-render equipment view if currently displayed
        if (this.currentView === 'equipment') {
            this.renderEquipmentView();
        }
        
        this.showMessage('Jump deleted successfully', 'success');
        
        // Sync the updated (renumbered) jump list to Google Sheets
        if (navigator.onLine && window.SheetsAPI?.initialized) {
            window.SheetsAPI.pushAllWithGuard();
        }
    }
    
    renumberJumps() {
        // Sort jumps chronologically, handling invalid dates and same-day ties
        this.jumps.sort((a, b) => {
            const da = Date.parse(a.date), db = Date.parse(b.date);
            if (isNaN(da) && isNaN(db)) return 0;
            if (isNaN(da)) return 1;  // invalid dates go to end
            if (isNaN(db)) return -1;
            if (da !== db) return da - db;
            // Same date: secondary sort by creation timestamp
            return Date.parse(a.timestamp) - Date.parse(b.timestamp);
        });

        // Renumber jumps starting from the configured starting number
        this.jumps.forEach((jump, index) => {
            jump.jumpNumber = this.settings.startingJumpNumber + index;
        });
    }

    async openSettingsModal() {
        document.getElementById('startingJumpNumber').value = this.settings.startingJumpNumber;
        document.getElementById('recentJumpsDays').value = this.settings.recentJumpsDays ?? 3;
        document.getElementById('standardRedThreshold').value = this.settings.standardRedThreshold ?? 160;
        document.getElementById('standardOrangeThreshold').value = this.settings.standardOrangeThreshold ?? 140;
        document.getElementById('hybridRedThreshold').value = this.settings.hybridRedThreshold ?? 80;
        document.getElementById('hybridOrangeThreshold').value = this.settings.hybridOrangeThreshold ?? 60;

        const restoreBtn   = document.getElementById('restoreFromBackupBtn');
        const restoreDesc  = document.getElementById('restoreFromBackupDesc');
        const restoreTitle = document.getElementById('restoreFromBackupTitle');
        restoreBtn.style.display   = 'none';
        restoreDesc.style.display  = 'none';
        restoreTitle.style.display = 'none';

        if (navigator.onLine && window.SheetsAPI?.initialized) {
            const hasBackupRigsSheet = await window.SheetsAPI.hasBackupRigsSheet();
            const vis = hasBackupRigsSheet ? 'block' : 'none';
            restoreBtn.style.display   = vis;
            restoreDesc.style.display  = vis;
            restoreTitle.style.display = vis;
        }

        document.getElementById('settingsModal').style.display = 'block';
    }

    closeModal() {
        document.getElementById('settingsModal').style.display = 'none';
    }

    openSheetsModal() {
        // Populate Google Sheets config fields from localStorage (or current API state)
        const savedConfig = JSON.parse(localStorage.getItem('sheets-config') || '{}');
        const webAppUrl = (window.SheetsAPI && window.SheetsAPI.webAppUrl) || savedConfig.webAppUrl || '';
        const spreadsheetId = (window.SheetsAPI && window.SheetsAPI.spreadsheetId) || savedConfig.spreadsheetId || '';
        document.getElementById('cfgWebAppUrl').value = webAppUrl;
        document.getElementById('cfgSpreadsheetId').value = spreadsheetId;

        // Show connection status
        const statusEl = document.getElementById('sheetsConfigStatus');
        if (window.SheetsAPI && window.SheetsAPI.initialized) {
            statusEl.textContent = '✅ Connected to Google Sheets';
            statusEl.style.color = '#2e7d32';
        } else if (webAppUrl) {
            statusEl.textContent = '⚠️ Configured but not connected — check values';
            statusEl.style.color = '#e65100';
        } else {
            statusEl.textContent = 'ℹ️ Enter your Apps Script URL and Spreadsheet ID to enable sync';
            statusEl.style.color = '#666';
        }

        document.getElementById('sheetsModal').style.display = 'block';
    }

    closeSheetsModal() {
        document.getElementById('sheetsModal').style.display = 'none';
    }

    saveSheetsConfig() {
        const webAppUrl = document.getElementById('cfgWebAppUrl').value.trim();
        const spreadsheetId = document.getElementById('cfgSpreadsheetId').value.trim();

        if (webAppUrl || spreadsheetId) {
            const sheetsConfig = { webAppUrl, spreadsheetId };
            localStorage.setItem('sheets-config', JSON.stringify(sheetsConfig));

            // Re-initialize the Sheets API with the new values
            if (window.SheetsAPI) {
                window.SheetsAPI.reinitialize(webAppUrl, spreadsheetId);
            }
        }

        this.closeSheetsModal();
        this.showMessage('Google Sheets configuration saved!', 'success');
    }

    saveSettings() {
        const startingJumpNumber = parseInt(document.getElementById('startingJumpNumber').value);
        
        if (!startingJumpNumber || startingJumpNumber < 1) {
            this.showMessage('Please enter a valid starting jump number (1 or higher)', 'error');
            return;
        }

        const recentJumpsDays = parseInt(document.getElementById('recentJumpsDays').value);
        if (!recentJumpsDays || recentJumpsDays < 0) {
            this.showMessage('Please enter a valid number of days (0 or higher)', 'error');
            return;
        }

        const standardRedThreshold = parseInt(document.getElementById('standardRedThreshold').value);
        const standardOrangeThreshold = parseInt(document.getElementById('standardOrangeThreshold').value);
        const hybridRedThreshold = parseInt(document.getElementById('hybridRedThreshold').value);
        const hybridOrangeThreshold = parseInt(document.getElementById('hybridOrangeThreshold').value);

        if (!standardRedThreshold || standardRedThreshold < 1) {
            this.showMessage('Please enter a valid standard red threshold (1 or higher)', 'error');
            return;
        }
        if (!standardOrangeThreshold || standardOrangeThreshold < 1) {
            this.showMessage('Please enter a valid standard orange threshold (1 or higher)', 'error');
            return;
        }
        if (!hybridRedThreshold || hybridRedThreshold < 1) {
            this.showMessage('Please enter a valid hybrid red threshold (1 or higher)', 'error');
            return;
        }
        if (!hybridOrangeThreshold || hybridOrangeThreshold < 1) {
            this.showMessage('Please enter a valid hybrid orange threshold (1 or higher)', 'error');
            return;
        }

        this.settings.startingJumpNumber = startingJumpNumber;
        this.settings.recentJumpsDays = recentJumpsDays;
        this.settings.standardRedThreshold = standardRedThreshold;
        this.settings.standardOrangeThreshold = standardOrangeThreshold;
        this.settings.hybridRedThreshold = hybridRedThreshold;
        this.settings.hybridOrangeThreshold = hybridOrangeThreshold;
        localStorage.setItem('skydiving-settings', JSON.stringify(this.settings));
        this.markEquipmentModified();

        // Mark equipment dirty so an offline save is not overwritten on next sync.
        localStorage.setItem('skydiving-equipment-dirty', '1');

        // Push settings to Google Sheets if online
        if (navigator.onLine && window.SheetsAPI?.initialized) {
            window.SheetsAPI.syncEquipmentToSheet();
        }
        
        this.closeModal();
        this.showMessage('Settings saved successfully!', 'success');
    }

    async restoreEquipmentFromBackup() {
        if (!confirm('This will overwrite ALL local equipment data (harnesses, canopies, linesets, rigs, locations) with the data from the backupRigs sheet. Continue?')) {
            return;
        }

        if (!window.SheetsAPI || !window.SheetsAPI.initialized) {
            this.showMessage('Google Sheets is not connected. Configure it first.', 'error');
            return;
        }

        this.showMessage('Restoring equipment from backup...', 'success');

        try {
            const success = await window.SheetsAPI.restoreEquipmentFromBackup();
            if (success) {
                this.showMessage('Equipment restored from backup successfully!', 'success');
                this.closeModal();
            } else {
                this.showMessage('No data found in backupRigs sheet.', 'error');
            }
        } catch (err) {
            console.error('Restore from backup failed:', err);
            this.showMessage('Restore failed: ' + err.message, 'error');
        }
    }

    getCurrentUtcTimestamp() {
        return new Date().toISOString();
    }

    markJumpsModified() {
        localStorage.setItem('skydiving-jumps-last-modified-utc', this.getCurrentUtcTimestamp());
    }

    markEquipmentModified() {
        localStorage.setItem('skydiving-equipment-last-modified-utc', this.getCurrentUtcTimestamp());
    }

    saveToLocalStorage() {
        localStorage.setItem('skydiving-jumps', JSON.stringify(this.jumps));
        this.markJumpsModified();
        // Mark that there are local changes not yet pushed to the sheet
        localStorage.setItem('skydiving-needs-sync', '1');
        localStorage.setItem('skydiving-data-modified', new Date().toISOString());
    }

    saveComponentsToLocalStorage() {
        localStorage.setItem('skydiving-harnesses', JSON.stringify(this.harnesses));
        localStorage.setItem('skydiving-canopies', JSON.stringify(this.canopies));
        localStorage.setItem('skydiving-equipment-rigs', JSON.stringify(this.equipmentRigs));
        localStorage.setItem('skydiving-locations', JSON.stringify(this.locations));
        this.markEquipmentModified();
        // Mark equipment as locally modified so the next sync pushes instead of pulls.
        // Startup code (jump-count init) must NOT call this method — it
        // should write directly to localStorage to avoid falsely setting the dirty flag.
        localStorage.setItem('skydiving-equipment-dirty', '1');
        
        // Push to Google Sheets if online
        if (navigator.onLine && window.SheetsAPI?.initialized) {
            window.SheetsAPI.syncEquipmentToSheet();
        }
    }
    
    
    initializeEquipmentJumpCounts() {
        // Count all logged jumps for each equipment rig (no jump-number filter;
        // pre-app jumps are accounted for separately via eq.previousJumps).
        let needsSave = false;
        this.equipmentRigs.forEach(eq => {
            const actualJumpCount = this.jumps.filter(jump => jump.equipment === eq.id).length;
            
            if (eq.jumpCount !== actualJumpCount) {
                eq.jumpCount = actualJumpCount;
                needsSave = true;
            }
        });
        
        if (needsSave) {
            // Save only the rigs array.  Calling saveComponentsToLocalStorage() here
            // would update skydiving-equipment-modified to "now", making the laptop
            // appear newer than a phone that just edited equipment notes, which would
            // cause the subsequent syncWithSheet() to skip the pull and overwrite the
            // phone's changes with stale local data.
            localStorage.setItem('skydiving-equipment-rigs', JSON.stringify(this.equipmentRigs));
        }
    }

    showView(viewName) {
        this.currentView = viewName;
        
        // Hide all views
        document.querySelectorAll('.view').forEach(view => {
            view.style.display = 'none';
        });
        
        // Show selected view
        document.getElementById(`${viewName}View`).style.display = 'block';
        
        // Update navigation buttons
        document.querySelectorAll('.nav-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        document.getElementById(`${viewName}ViewBtn`).classList.add('active');
        
        // Load view-specific content
        if (viewName === 'equipment') {
            this.renderEquipmentView();
        } else if (viewName === 'stats') {
            this.renderStats();
        }
    }
    
    showEquipmentSubView(subView) {
        this.equipmentSubView = subView;
        
        // Update navigation buttons
        document.querySelectorAll('.equipment-nav-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        document.querySelector(`[data-view="${subView}"]`).classList.add('active');
        
        // Update section title and buttons
        const titleMap = {
            'rigs': 'Equipment Rigs',
            'harnesses':         'Harnesses',
            'canopies':     'Canopies',
            'locations':    'Drop Zones / Locations'
        };
        
        document.getElementById('equipmentSectionTitle').textContent = titleMap[subView];
        
        // Show/hide appropriate buttons
        document.getElementById('addEquipmentBtn').style.display  = subView === 'rigs' ? 'block' : 'none';
        document.getElementById('addHarnessBtn').style.display        = subView === 'harnesses'         ? 'block' : 'none';
        document.getElementById('addCanopyBtn').style.display     = subView === 'canopies'     ? 'block' : 'none';
        document.getElementById('addLocationBtn').style.display   = subView === 'locations'    ? 'block' : 'none';
        
        this.renderEquipmentView();
    }
    
    renderEquipmentView() {
        switch(this.equipmentSubView) {
            case 'rigs': this.renderEquipmentRigs(); break;
            case 'harnesses':         this.renderComponents('harnesses');      break;
            case 'canopies':     this.renderComponents('canopies');  break;
            case 'locations':    this.renderComponents('locations'); break;
        }
    }

    updateEquipmentOptions() {
        const select = document.getElementById('equipment');
        select.innerHTML = '<option value="">Select Equipment</option>';
        
        // Only show non-archived rigs
        const activeRigs = this.equipmentRigs.filter(eq => !eq.archived);
        
        activeRigs.forEach(eq => {
            const option = document.createElement('option');
            option.value = eq.id;
            option.textContent = eq.name; // Use the auto-generated name
            select.appendChild(option);
        });
    }

    renderEquipmentRigs() {
        const container = document.getElementById('equipmentList');
        
        if (this.equipmentRigs.length === 0) {
            container.innerHTML = '<p class="no-items">No equipment rigs created yet.</p>';
            return;
        }
        
        const sorted = [...this.equipmentRigs].sort((a, b) => !!a.archived - !!b.archived);

        container.innerHTML = sorted.map(eq => {
            const harness = this.harnesses.find(r => r.id === eq.harnessId);
            const canopy = this.canopies.find(c => c.id === eq.canopyId);
            
            let displayName = eq.name;
            let components = '';
            let jumpInfo = '';
            if (harness && canopy) {
                // Use the stored name (auto-generated)
                displayName = eq.name;
                components = `<div class="equipment-components">${harness.name} | ${canopy.name} | Lineset#${eq.linesetNumber || 1}</div>`;
                const logged = eq.jumpCount || 0;
                const preApp = eq.previousJumps || 0;
                const total = logged + preApp;
                jumpInfo = `<div class="jump-info">Logged: ${logged} | Pre-app: ${preApp} | Total: ${total}</div>`;
            }
            const rigNotes = eq.notes ? `<div class="component-notes">\uD83D\uDCDD ${eq.notes}</div>` : '';
            
            return `
                <div class="equipment-item ${eq.archived ? 'archived' : ''}">
                    <div class="equipment-info">
                        <span class="equipment-name">${displayName}</span>
                        ${components}
                        ${jumpInfo}
                        ${rigNotes}
                        ${eq.archived ? '<span class="archived-badge">Archived</span>' : ''}
                    </div>
                    <div class="equipment-actions">
                        <button onclick="window.logbook.editEquipment('${eq.id}')" class="btn-edit">Edit</button>
                        <button onclick="window.logbook.toggleArchiveEquipment('${eq.id}')" class="btn-toggle">
                            ${eq.archived ? 'Unarchive' : 'Archive'}
                        </button>
                        <button onclick="window.logbook.deleteEquipment('${eq.id}')" class="btn-delete">Delete</button>
                    </div>
                </div>
            `;
        }).join('');
    }

    autoFillRigNotes() {
        const harnessId = document.getElementById('equipmentHarness').value;
        const canopyId  = document.getElementById('equipmentCanopy').value;
        document.getElementById('equipmentNotes').value = this.composeRigNotes(harnessId, canopyId);
    }

    composeRigNotes(harnessId, canopyId) {
        const harness = this.harnesses.find(h => h.id === harnessId);
        const canopy  = this.canopies.find(c => c.id === canopyId);
        const parts = [
            harness?.notes ? `H: ${harness.notes}` : '',
            canopy?.notes  ? `C: ${canopy.notes}`  : ''
        ].filter(Boolean);
        return parts.join('\n');
    }

    /**
     * After a harness or canopy is saved, rebuild the stored notes field
     * of every rig that uses it from the (now updated) component notes.
     */
    propagateComponentNotesToRigs() {
        this.equipmentRigs.forEach(eq => {
            eq.notes = this.composeRigNotes(eq.harnessId, eq.canopyId);
        });
    }

    autoFillLinesetNumber() {
        const canopyId = document.getElementById('equipmentCanopy').value;
        if (!canopyId) {
            document.getElementById('equipmentLinesetNumber').value = 1;
            return;
        }
        const rigsWithCanopy = this.equipmentRigs.filter(eq => eq.canopyId === canopyId);
        const maxLineset = rigsWithCanopy.length > 0
            ? Math.max(...rigsWithCanopy.map(eq => eq.linesetNumber || 1))
            : 0;
        document.getElementById('equipmentLinesetNumber').value = maxLineset + 1;
    }

    addEquipment() {
        document.getElementById('equipmentForm').reset();
        document.getElementById('equipmentId').value = '';
        document.getElementById('equipmentStartingJumpNumber').value = 0;
        document.getElementById('equipmentLinesetNumber').value = 1;
        document.getElementById('equipmentNotes').value = '';
        this.populateComponentSelects();
        this.autoFillRigNotes();
        document.getElementById('equipmentModal').style.display = 'block';
    }
    
    populateComponentSelects() {
        // Populate harness select (active only)
        const harnessSelect = document.getElementById('equipmentHarness');
        harnessSelect.innerHTML = '<option value="">Select Harness</option>';
        this.harnesses.filter(h => !h.archived).forEach(harness => {
            const option = document.createElement('option');
            option.value = harness.id;
            option.textContent = harness.name;
            harnessSelect.appendChild(option);
        });
        
        // Populate canopy select (active only)
        const canopySelect = document.getElementById('equipmentCanopy');
        canopySelect.innerHTML = '<option value="">Select Canopy</option>';
        this.canopies.filter(c => !c.archived).forEach(canopy => {
            const option = document.createElement('option');
            option.value = canopy.id;
            option.textContent = canopy.name;
            canopySelect.appendChild(option);
        });
    }

    editEquipment(id) {
        const equipment = this.equipmentRigs.find(eq => eq.id === id);
        if (equipment) {
            document.getElementById('equipmentId').value = equipment.id;
            document.getElementById('equipmentStartingJumpNumber').value = equipment.previousJumps || 0;
            this.populateComponentSelects();
            
            // Set selected values
            document.getElementById('equipmentHarness').value = equipment.harnessId || '';
            document.getElementById('equipmentCanopy').value = equipment.canopyId || '';
            document.getElementById('equipmentLinesetNumber').value = equipment.linesetNumber || 1;
            document.getElementById('equipmentHybridCheck').checked = /-Hybrid$/i.test(equipment.name || '');
            this.autoFillRigNotes();
            
            document.getElementById('equipmentModal').style.display = 'block';
        }
    }

    saveEquipment() {
        const id = document.getElementById('equipmentId').value;
        const harnessId = document.getElementById('equipmentHarness').value;
        const canopyId = document.getElementById('equipmentCanopy').value;
        const linesetNumber = Math.max(1, parseInt(document.getElementById('equipmentLinesetNumber').value) || 1);
        const previousJumps = Math.max(0, parseInt(document.getElementById('equipmentStartingJumpNumber').value) || 0);
        const notes = this.composeRigNotes(harnessId, canopyId);
        
        if (!harnessId || !canopyId) {
            this.showMessage('Please select harness and canopy', 'error');
            return;
        }
        
        // Auto-generate name from components
        const harness = this.harnesses.find(r => r.id === harnessId);
        const canopy = this.canopies.find(c => c.id === canopyId);
        const isHybrid = document.getElementById('equipmentHybridCheck').checked;
        let name = `${harness.name}-${canopy.name}-Lineset#${linesetNumber}`;
        if (isHybrid) name += '-Hybrid';
        
        if (id) {
            // Edit existing
            const equipment = this.equipmentRigs.find(eq => eq.id === id);
            if (equipment) {
                equipment.name = name;
                equipment.harnessId = harnessId;
                equipment.canopyId = canopyId;
                equipment.linesetNumber = linesetNumber;
                equipment.previousJumps = previousJumps;
                equipment.notes = notes;
            }
        } else {
            // Add new
            const newId = 'eq_' + Date.now();
            this.equipmentRigs.push({
                id: newId,
                name: name,
                harnessId: harnessId,
                canopyId: canopyId,
                linesetNumber: linesetNumber,
                previousJumps: previousJumps,
                jumpCount: 0,
                archived: false,
                notes: notes
            });
        }
        
        this.saveComponentsToLocalStorage();
        this.updateEquipmentOptions();
        this.renderEquipmentView();
        this.closeEquipmentModal();
        this.showMessage('Equipment saved successfully!', 'success');
    }

    deleteEquipment(id) {
        if (confirm('Are you sure you want to delete this equipment rig?')) {
            // Check if equipment is used in any jumps
            const usedInJumps = this.jumps.some(jump => jump.equipment === id);
            if (usedInJumps) {
                this.showMessage('Cannot delete equipment that has been used in jumps. Archive it instead.', 'error');
                return;
            }
            
            this.equipmentRigs = this.equipmentRigs.filter(eq => eq.id !== id);
            this.saveComponentsToLocalStorage();
            this.updateEquipmentOptions();
            this.renderEquipmentView();
            this.showMessage('Equipment deleted successfully!', 'success');
        }
    }
    
    toggleArchiveEquipment(id) {
        const equipment = this.equipmentRigs.find(eq => eq.id === id);
        if (equipment) {
            equipment.archived = !equipment.archived;
            this.saveComponentsToLocalStorage();
            this.updateEquipmentOptions();
            this.renderEquipmentView();
            this.showMessage(`Equipment ${equipment.archived ? 'archived' : 'unarchived'} successfully!`, 'success');
        }
    }

    _singularize(plural) {
        const map = { harnesses: 'harness', canopies: 'canopy', linesets: 'lineset', locations: 'location' };
        return map[plural] || plural.slice(0, -1);
    }

    addComponent(type) {
        document.getElementById('componentForm').reset();
        document.getElementById('componentId').value = '';
        document.getElementById('componentType').value = type;
        document.getElementById('componentNotes').value = '';
        document.getElementById('componentModalTitle').textContent = `Add ${type.charAt(0).toUpperCase() + type.slice(1)}`;
        // Show/hide GPS coords section for locations
        const isLocation = type === 'location';
        document.getElementById('locationCoordsSection').style.display = isLocation ? 'block' : 'none';
        if (isLocation) {
            document.getElementById('componentLat').value = '';
            document.getElementById('componentLng').value = '';
            const hint = document.getElementById('coordsHint');
            hint.textContent = 'Leave blank to auto-geocode from the name.';
            hint.style.color = '#888';
        }
        document.getElementById('componentModal').style.display = 'block';
    }

    saveComponent() {
        const id = document.getElementById('componentId').value;
        const name = document.getElementById('componentName').value.trim();
        const type = document.getElementById('componentType').value;
        
        if (!name) {
            this.showMessage('Please enter component name', 'error');
            return;
        }
        const notes = document.getElementById('componentNotes').value.trim();

        // Read manual GPS coords if this is a location and the fields have values
        let manualLat = null, manualLng = null;
        if (type === 'location') {
            const latVal = document.getElementById('componentLat').value;
            const lngVal = document.getElementById('componentLng').value;
            if (latVal !== '' && lngVal !== '') {
                const parsedLat = parseFloat(latVal);
                const parsedLng = parseFloat(lngVal);
                if (!isNaN(parsedLat) && !isNaN(parsedLng)
                        && parsedLat >= -90 && parsedLat <= 90
                        && parsedLng >= -180 && parsedLng <= 180) {
                    manualLat = parsedLat;
                    manualLng = parsedLng;
                } else {
                    this.showMessage('Invalid coordinates — latitude must be −90…90, longitude −180…180', 'error');
                    return;
                }
            }
        }
        
        // Get the correct collection name for each type
        let collectionName;
        switch(type) {
            case 'harness':      collectionName = 'harnesses';      break;
            case 'canopy':   collectionName = 'canopies';  break;
            case 'location': collectionName = 'locations'; break;
            default:
                this.showMessage('Invalid component type', 'error');
                return;
        }
        
        const collection = this[collectionName];
        
        if (id) {
            // Edit existing
            const component = collection.find(c => c.id === id);
            if (component) {
                const nameChanged = type === 'location' && component.name !== name;
                component.name = name;
                component.notes = notes;
                if (type === 'location') {
                    if (manualLat !== null) {
                        // Manual coords override everything
                        component.lat = manualLat;
                        component.lng = manualLng;
                    } else {
                        if (nameChanged) { component.lat = null; component.lng = null; }
                        if (component.lat == null) this.geocodeLocation(component);
                    }
                }
            }
        } else {
            // Add new
            const newId = type + '_' + Date.now();
            const newComponent = { id: newId, name: name, notes: notes };
            if (type === 'location') {
                if (manualLat !== null) {
                    newComponent.lat = manualLat;
                    newComponent.lng = manualLng;
                } else {
                    newComponent.lat = null;
                    newComponent.lng = null;
                    this.geocodeLocation(newComponent);
                }
                collection.push(newComponent);
            } else {
                collection.push(newComponent);
            }
        }
        
        // Propagate updated component notes to all affected rigs
        if (type === 'harness' || type === 'canopy') {
            this.propagateComponentNotesToRigs();
        }
        this.saveComponentsToLocalStorage();
        this.renderEquipmentView();
        this.closeComponentModal();
        // Refresh autocomplete if a location was saved
        if (type === 'location') this.updateLocationDatalist();
        if (navigator.onLine && window.SheetsAPI) window.SheetsAPI.syncEquipmentToSheet();
        this.showMessage(`${type.charAt(0).toUpperCase() + type.slice(1)} saved successfully!`, 'success');
    }
    
    updateLocationDatalist() {
        // No-op: replaced by custom autocomplete dropdown
    }

    setupLocationAutocomplete() {
        const input = document.getElementById('location');
        const dropdown = document.getElementById('locationDropdown');
        if (!input || !dropdown) return;

        let activeIndex = -1;

        const showDropdown = () => {
            const query = input.value.trim().toLowerCase();
            const matches = this.locations.filter(loc =>
                loc.name.toLowerCase().includes(query)
            );

            if (matches.length === 0) {
                dropdown.classList.remove('open');
                return;
            }

            dropdown.innerHTML = matches.map((loc, i) => {
                let display = loc.name;
                if (query) {
                    const idx = loc.name.toLowerCase().indexOf(query);
                    if (idx !== -1) {
                        display = loc.name.slice(0, idx)
                            + '<span class="match">' + loc.name.slice(idx, idx + query.length) + '</span>'
                            + loc.name.slice(idx + query.length);
                    }
                }
                return `<div class="autocomplete-option" data-index="${i}" data-value="${loc.name}">${display}</div>`;
            }).join('');

            activeIndex = -1;
            dropdown.classList.add('open');
        };

        const selectOption = (value) => {
            input.value = value;
            dropdown.classList.remove('open');
            activeIndex = -1;
        };

        input.addEventListener('focus', showDropdown);
        input.addEventListener('input', showDropdown);

        input.addEventListener('keydown', (e) => {
            const options = dropdown.querySelectorAll('.autocomplete-option');
            if (!dropdown.classList.contains('open') || options.length === 0) return;

            if (e.key === 'ArrowDown') {
                e.preventDefault();
                activeIndex = Math.min(activeIndex + 1, options.length - 1);
                options.forEach((o, i) => o.classList.toggle('active', i === activeIndex));
                options[activeIndex].scrollIntoView({ block: 'nearest' });
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                activeIndex = Math.max(activeIndex - 1, 0);
                options.forEach((o, i) => o.classList.toggle('active', i === activeIndex));
                options[activeIndex].scrollIntoView({ block: 'nearest' });
            } else if (e.key === 'Enter' && activeIndex >= 0) {
                e.preventDefault();
                selectOption(options[activeIndex].dataset.value);
            } else if (e.key === 'Escape') {
                dropdown.classList.remove('open');
            }
        });

        dropdown.addEventListener('click', (e) => {
            const option = e.target.closest('.autocomplete-option');
            if (option) selectOption(option.dataset.value);
        });

        // Close dropdown when clicking outside
        document.addEventListener('click', (e) => {
            if (!e.target.closest('.location-autocomplete')) {
                dropdown.classList.remove('open');
            }
        });
    }

    // ── Geolocation helpers ─────────────────────────────────────────────────

    /** Haversine distance between two lat/lng points, returns km. */
    haversineKm(lat1, lng1, lat2, lng2) {
        const R = 6371;
        const dLat = (lat2 - lat1) * Math.PI / 180;
        const dLng = (lng2 - lng1) * Math.PI / 180;
        const a = Math.sin(dLat / 2) ** 2
            + Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * Math.sin(dLng / 2) ** 2;
        return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    }

    /**
     * Geocode a single location by name using OpenStreetMap Nominatim.
     * Updates location.lat / .lng in-place and persists to storage.
     */
    async geocodeLocation(location) {
        if (!navigator.onLine) return;
        try {
            const url = `https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(location.name)}&format=json&limit=1`;
            const res = await fetch(url, { headers: { 'Accept-Language': 'en' } });
            if (!res.ok) return;
            const data = await res.json();
            if (!data.length) return;
            location.lat = parseFloat(data[0].lat);
            location.lng = parseFloat(data[0].lon);
            this.saveComponentsToLocalStorage();
        } catch (_) {
            // Network error or rate-limited — silent fail
        }
    }

    /**
     * Background-geocodes all locations that are missing coordinates.
     * Staggered at 1.2 s per request to respect Nominatim's usage policy.
     */
    geocodeAllLocations() {
        if (!navigator.onLine) return;
        const ungeocoded = this.locations.filter(l => l.lat == null);
        ungeocoded.forEach((loc, i) => {
            setTimeout(() => this.geocodeLocation(loc), i * 1200);
        });
    }

    /**
     * Ask for the user's current position and pre-fill the location field
     * with the nearest known dropzone that has stored coordinates.
     *
     * @param {boolean} forceOverwrite – if true, overwrite even if the field
     *   already has a value (used by the manual button); if false (auto-mode)
     *   only fill when the field is empty or holds the last-jump value.
     */
    async detectNearestLocation(forceOverwrite = false) {
        if (!navigator.geolocation) return;
        const locationsWithCoords = this.locations.filter(l => l.lat != null && l.lng != null);
        if (locationsWithCoords.length === 0) return;

        const input = document.getElementById('location');
        const hint  = document.getElementById('locationGeoHint');
        if (!input) return;

        // In auto-mode only proceed if field is empty
        if (!forceOverwrite && input.value.trim() !== '') return;

        try {
            const pos = await new Promise((resolve, reject) =>
                navigator.geolocation.getCurrentPosition(resolve, reject, { timeout: 8000, maximumAge: 60000 })
            );
            const { latitude, longitude } = pos.coords;

            let nearest = null, minDist = Infinity;
            for (const loc of locationsWithCoords) {
                const d = this.haversineKm(latitude, longitude, loc.lat, loc.lng);
                if (d < minDist) { minDist = d; nearest = loc; }
            }

            if (nearest) {
                input.value = nearest.name;
                if (hint) {
                    hint.textContent = `📍 Nearest: ${nearest.name} (${Math.round(minDist)} km)`;
                    hint.style.display = 'block';
                    // Fade out the hint after 6 seconds
                    clearTimeout(this._geoHintTimer);
                    this._geoHintTimer = setTimeout(() => { if (hint) hint.style.display = 'none'; }, 6000);
                }
            }
        } catch (_) {
            // Permission denied or timeout — silent fail
        }
    }

    /**
     * Fires navigator.geolocation and fills the lat/lng inputs in the
     * component modal (used by the "Use current GPS position" button).
     */
    async setComponentCoordsFromGPS() {
        if (!navigator.geolocation) {
            this.showMessage('Geolocation is not supported by your browser', 'error');
            return;
        }
        const hint = document.getElementById('coordsHint');
        const btn  = document.getElementById('useCurrentLocationBtn');
        hint.textContent = 'Getting position…';
        hint.style.color = '#888';
        btn.disabled = true;
        try {
            const pos = await new Promise((resolve, reject) =>
                navigator.geolocation.getCurrentPosition(resolve, reject, { timeout: 10000 })
            );
            document.getElementById('componentLat').value = pos.coords.latitude.toFixed(6);
            document.getElementById('componentLng').value = pos.coords.longitude.toFixed(6);
            hint.textContent = `✅ Captured (±${Math.round(pos.coords.accuracy)} m accuracy)`;
            hint.style.color = '#2e7d32';
        } catch (err) {
            hint.textContent = err.code === 1 ? '⚠️ Location permission denied' : '⚠️ Could not get position — try again';
            hint.style.color = '#c62828';
        } finally {
            btn.disabled = false;
        }
    }

    renderComponents(type) {
        const container = document.getElementById('equipmentList');
        const collection = this[type];
        
        if (collection.length === 0) {
            container.innerHTML = `<p class="no-items">No ${type} added yet.</p>`;
            return;
        }

        // Active first, then archived
        const sorted = [...collection].sort((a, b) => !!a.archived - !!b.archived);
        
        container.innerHTML = sorted.map(component => `
            <div class="equipment-item ${component.archived ? 'archived' : ''}">
                <div class="equipment-info">
                    <span class="equipment-name">${component.name}</span>
                    ${component.notes ? `<div class="component-notes">\uD83D\uDCDD ${component.notes}</div>` : ''}
                    ${component.archived ? '<span class="archived-badge">Archived</span>' : ''}
                </div>
                <div class="equipment-actions">
                    <button onclick="window.logbook.editComponent('${component.id}', '${type}')" class="btn-edit">Edit</button>
                    <button onclick="window.logbook.toggleArchiveComponent('${component.id}', '${type}')" class="btn-toggle">
                        ${component.archived ? 'Unarchive' : 'Archive'}
                    </button>
                    <button onclick="window.logbook.deleteComponent('${component.id}', '${type}')" class="btn-delete">Delete</button>
                </div>
            </div>
        `).join('');
    }

    toggleArchiveComponent(id, type) {
        const collection = this[type];
        const component = collection.find(c => c.id === id);
        if (component) {
            component.archived = !component.archived;
            this.saveComponentsToLocalStorage();
            this.renderEquipmentView();
            const typeSingular = this._singularize(type);
            this.showMessage(`${typeSingular.charAt(0).toUpperCase() + typeSingular.slice(1)} ${component.archived ? 'archived' : 'unarchived'} successfully!`, 'success');
        }
    }
    
    editComponent(id, type) {
        const collection = this[type];
        const component = collection.find(c => c.id === id);
        if (component) {
            const singular = this._singularize(type);
            document.getElementById('componentId').value = component.id;
            document.getElementById('componentName').value = component.name;
            document.getElementById('componentNotes').value = component.notes || '';
            document.getElementById('componentType').value = singular;
            document.getElementById('componentModalTitle').textContent = `Edit ${singular.charAt(0).toUpperCase() + singular.slice(1)}`;
            // Show/hide GPS coords section for locations
            const isLocation = singular === 'location';
            document.getElementById('locationCoordsSection').style.display = isLocation ? 'block' : 'none';
            if (isLocation) {
                document.getElementById('componentLat').value = component.lat != null ? component.lat : '';
                document.getElementById('componentLng').value = component.lng != null ? component.lng : '';
                const hint = document.getElementById('coordsHint');
                if (component.lat != null) {
                    hint.textContent = `Saved: ${component.lat.toFixed(5)}, ${component.lng.toFixed(5)}`;
                    hint.style.color = '#2e7d32';
                } else {
                    hint.textContent = 'No coordinates saved yet — leave blank to auto-geocode from name.';
                    hint.style.color = '#888';
                }
            }
            document.getElementById('componentModal').style.display = 'block';
        }
    }
    
    deleteComponent(id, type) {
        const typeSingular = this._singularize(type);
        if (confirm(`Are you sure you want to delete this ${typeSingular}?`)) {
            // Check if component is used in any equipment rigs
            const usedInEquipment = this.equipmentRigs.some(eq => 
                eq.harnessId === id || eq.canopyId === id
            );
            if (usedInEquipment) {
                this.showMessage(`Cannot delete ${typeSingular} that is used in equipment rigs`, 'error');
                return;
            }
            
            const collection = this[type];
            const index = collection.findIndex(c => c.id === id);
            if (index !== -1) {
                collection.splice(index, 1);
                this.saveComponentsToLocalStorage();
                this.renderEquipmentView();
                if (type === 'locations') this.updateLocationDatalist();
                if (navigator.onLine && window.SheetsAPI) window.SheetsAPI.syncEquipmentToSheet();
                this.showMessage(`${typeSingular.charAt(0).toUpperCase() + typeSingular.slice(1)} deleted successfully!`, 'success');
            }
        }
    }
    
    closeComponentModal() {
        document.getElementById('componentModal').style.display = 'none';
    }
    
    closeEquipmentModal() {
        document.getElementById('equipmentModal').style.display = 'none';
    }

    renderStats() {
        const container = document.getElementById('statsContent');
        
        if (this.jumps.length === 0) {
            container.innerHTML = '<p class="no-items">No jumps logged yet.</p>';
            return;
        }
        
        // Calculate equipment rig statistics
        const equipmentStats = this.equipmentRigs.map(eq => {
            const loggedJumpsCount = this.jumps.filter(jump => jump.equipment === eq.id).length;
            const totalCount = loggedJumpsCount + (eq.previousJumps || 0);
            let name = eq.name;
            const harness = this.harnesses.find(r => r.id === eq.harnessId);
            const canopy = this.canopies.find(c => c.id === eq.canopyId);
            if (harness && canopy) {
                const hybridSuffix = /-Hybrid$/i.test(eq.name || '') ? ' (Hybrid)' : '';
                name = `${harness.name} + ${canopy.name} + Lineset#${eq.linesetNumber || 1}${hybridSuffix}`;
            }
            return { name, count: totalCount, logged: loggedJumpsCount, preApp: eq.previousJumps || 0, archived: eq.archived, hybrid: /-Hybrid$/i.test(eq.name || '') };
        });
        
        // Separate active and archived equipment, then sort each group
        const activeStats = equipmentStats
            .filter(stat => !stat.archived && stat.count > 0)
            .sort((a, b) => b.count - a.count);
            
        const archivedStats = equipmentStats
            .filter(stat => stat.archived)
            .sort((a, b) => b.count - a.count);
            
        // Combine: active first, then archived only if toggled on
        const sortedEquipmentStats = this.showArchivedStats
            ? [...activeStats, ...archivedStats]
            : activeStats;
        
        // Calculate component-level statistics
        const harnessStats = {};
        const canopyStats = {};
        
        // Count logged jumps for each component
        this.jumps.forEach(jump => {
            const rig = this.equipmentRigs.find(eq => eq.id === jump.equipment);
            if (rig) {
                const harness = this.harnesses.find(r => r.id === rig.harnessId);
                if (harness) harnessStats[harness.name] = (harnessStats[harness.name] || 0) + 1;
                
                const canopy = this.canopies.find(c => c.id === rig.canopyId);
                if (canopy) canopyStats[canopy.name] = (canopyStats[canopy.name] || 0) + 1;
            }
        });
        
        // Add pre-app jump counts for each equipment rig
        this.equipmentRigs.forEach(eq => {
            const preApp = eq.previousJumps || 0;
            if (preApp > 0) {
                const harness = this.harnesses.find(r => r.id === eq.harnessId);
                if (harness) harnessStats[harness.name] = (harnessStats[harness.name] || 0) + preApp;
                
                const canopy = this.canopies.find(c => c.id === eq.canopyId);
                if (canopy) canopyStats[canopy.name] = (canopyStats[canopy.name] || 0) + preApp;
            }
        });
        
        const hasArchived = archivedStats.length > 0;
        const archivedBtnLabel = this.showArchivedStats ? 'Hide Archived' : `Show Archived (${archivedStats.length})`;
        const archivedToggleBtn = hasArchived
            ? `<button class="btn-secondary btn-sm" onclick="window.logbook.toggleArchivedStats()">${archivedBtnLabel}</button>`
            : '';

        let html = `
            <div class="stats-section">
                <div class="stats-section-header">
                    <h3>Equipment Rigs</h3>
                    ${archivedToggleBtn}
                </div>
                <div class="stats-list">
        `;
        
        if (sortedEquipmentStats.length > 0) {
            sortedEquipmentStats.forEach(stat => {
                const redThreshold = stat.hybrid ? this.settings.hybridRedThreshold : this.settings.standardRedThreshold;
                const orangeThreshold = stat.hybrid ? this.settings.hybridOrangeThreshold : this.settings.standardOrangeThreshold;
                const percentage = Math.min((stat.count / redThreshold) * 100, 100);
                let barColorClass = '';
                if (stat.count >= redThreshold) {
                    barColorClass = 'stat-fill-red';
                } else if (stat.count >= orangeThreshold) {
                    barColorClass = 'stat-fill-orange';
                }
                const breakdown = stat.preApp > 0
                    ? `${stat.count} total (${stat.logged} logged + ${stat.preApp} pre-app)`
                    : `${stat.count} jumps`;
                html += `
                    <div class="stat-item${stat.archived ? ' archived' : ''}">
                        <div class="stat-info">
                            <span class="stat-name">${stat.name} ${stat.archived ? '(Archived)' : ''}</span>
                            <span class="stat-count">${breakdown}</span>
                        </div>
                        <div class="stat-bar">
                            <div class="stat-fill ${barColorClass}" style="width: ${percentage}%"></div>
                        </div>
                    </div>
                `;
            });
        } else {
            html += '<p class="no-items">No equipment statistics available.</p>';
        }
        
        html += '</div></div>';
        
        // Add canopy statistics
        html += this.renderComponentStats('Canopies', canopyStats);
        
        // Add harness statistics
        html += this.renderComponentStats('Harnesses', harnessStats);
        
        container.innerHTML = html;
    }
    
    toggleArchivedStats() {
        this.showArchivedStats = !this.showArchivedStats;
        this.renderStats();
    }

    renderComponentStats(title, statsObject) {
        const statsArray = Object.entries(statsObject)
            .map(([name, count]) => ({ name, count }))
            .sort((a, b) => b.count - a.count);
            
        let html = `
            <div class="stats-section">
                <h3>${title}</h3>
                <div class="stats-list">
        `;
        
        if (statsArray.length > 0) {
            const maxCount = Math.max(...statsArray.map(s => s.count));
            statsArray.forEach(stat => {
                const percentage = (stat.count / maxCount) * 100;
                html += `
                    <div class="stat-item">
                        <div class="stat-info">
                            <span class="stat-name">${stat.name}</span>
                            <span class="stat-count">${stat.count} jumps</span>
                        </div>
                        <div class="stat-bar">
                            <div class="stat-fill" style="width: ${percentage}%"></div>
                        </div>
                    </div>
                `;
            });
        } else {
            html += `<p class="no-items">No ${title.toLowerCase()} statistics available.</p>`;
        }
        
        html += '</div></div>';
        return html;
    }

    hasExportableData() {
        return this.jumps.length > 0
            || this.equipmentRigs.length > 0
            || this.harnesses.length > 0
            || this.canopies.length > 0;
    }

    buildExportPayload() {
        return {
            exportedAt: new Date().toISOString(),
            version: 1,
            data: {
                jumps: this.jumps,
                equipmentRigs: this.equipmentRigs,
                harnesses: this.harnesses,
                canopies: this.canopies,
                locations: this.locations,
                settings: this.settings
            }
        };
    }

    buildExportFilename() {
        return `skydiving-logbook-backup-${new Date().toISOString().split('T')[0]}.json`;
    }

    exportData() {
        if (!this.hasExportableData()) {
            this.showMessage('No data to export', 'error');
            return;
        }

        const exportPayload = this.buildExportPayload();

        const jsonContent = JSON.stringify(exportPayload, null, 2);
        const blob = new Blob([jsonContent], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = this.buildExportFilename();
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        this.showMessage('Data exported successfully!', 'success');
    }

    async shareDataViaEmail() {
        if (!this.hasExportableData()) {
            this.showMessage('No data to share', 'error');
            return;
        }

        const exportPayload = this.buildExportPayload();
        const file = new File(
            [JSON.stringify(exportPayload, null, 2)],
            this.buildExportFilename(),
            { type: 'application/json' }
        );

        if (!navigator.share) {
            this.exportData();
            this.showMessage('Sharing not supported on this device. Backup downloaded instead.', 'info');
            return;
        }

        if (navigator.canShare && !navigator.canShare({ files: [file] })) {
            this.exportData();
            this.showMessage('File sharing not supported here. Backup downloaded instead.', 'info');
            return;
        }

        try {
            await navigator.share({
                title: 'Skydiving Logbook Backup',
                text: 'Backup JSON attached',
                files: [file]
            });
            this.showMessage('Share sheet opened.', 'success');
        } catch (error) {
            if (error?.name === 'AbortError') {
                return;
            }
            console.error('Share failed:', error);
            this.exportData();
            this.showMessage('Could not share file. Backup downloaded instead.', 'error');
        }
    }

    async importData(event) {
        const file = event.target.files?.[0];
        if (!file) return;

        try {
            const text = await file.text();
            const parsed = JSON.parse(text);
            const payload = parsed?.data ? parsed.data : parsed;

            if (!payload || typeof payload !== 'object') {
                this.showMessage('Invalid import file format', 'error');
                return;
            }

            const hasAnySupportedData =
                Array.isArray(payload.jumps)
                || Array.isArray(payload.equipmentRigs)
                || Array.isArray(payload.harnesses)
                || Array.isArray(payload.canopies)
                || Array.isArray(payload.locations)
                || (payload.settings && typeof payload.settings === 'object');

            if (!hasAnySupportedData) {
                this.showMessage('Import file has no supported data', 'error');
                return;
            }

            const confirmed = confirm('Importing will overwrite local jumps and equipment data. Continue?');
            if (!confirmed) return;

            this.jumps = Array.isArray(payload.jumps) ? payload.jumps : [];
            this.equipmentRigs = Array.isArray(payload.equipmentRigs) ? payload.equipmentRigs : [];
            this.harnesses = Array.isArray(payload.harnesses) ? payload.harnesses : [];
            this.canopies = Array.isArray(payload.canopies) ? payload.canopies : [];
            this.locations = Array.isArray(payload.locations) ? payload.locations : [];

            if (payload.settings && typeof payload.settings === 'object') {
                this.settings = {
                    ...this.settings,
                    ...payload.settings
                };
            }

            if (this.settings.recentJumpsDays === undefined) {
                this.settings.recentJumpsDays = 3;
            }

            this.initializeEquipmentJumpCounts();
            this.saveToLocalStorage();
            this.saveComponentsToLocalStorage();
            localStorage.setItem('skydiving-settings', JSON.stringify(this.settings));
            this.markEquipmentModified();

            this.updateEquipmentOptions();
            this.renderJumpsList();
            this.updateStats();
            this.renderEquipmentView();
            this.renderStats();

            this.showMessage('Data imported successfully!', 'success');
        } catch (error) {
            console.error('Import failed:', error);
            this.showMessage('Import failed: invalid JSON file', 'error');
        } finally {
            event.target.value = '';
        }
    }

    updateOnlineStatus() {
        const syncStatus = document.getElementById('syncStatus');
        if (navigator.onLine) {
            syncStatus.textContent = 'Online';
            syncStatus.className = 'sync-status success';
            this.hideOfflineIndicator();
        } else {
            syncStatus.textContent = 'Offline';
            syncStatus.className = 'sync-status error';
            this.showOfflineIndicator();
        }
    }

    showOfflineIndicator() {
        let indicator = document.querySelector('.offline-indicator');
        if (!indicator) {
            indicator = document.createElement('div');
            indicator.className = 'offline-indicator';
            indicator.textContent = '📡 You are offline. Data will sync when connection is restored.';
            document.body.insertBefore(indicator, document.querySelector('.container'));
        }
        indicator.classList.remove('hidden');
    }

    hideOfflineIndicator() {
        const indicator = document.querySelector('.offline-indicator');
        if (indicator) {
            indicator.classList.add('hidden');
        }
    }

    showMessage(message, type = 'info') {
        const toast = document.createElement('div');
        toast.className = `toast toast-${type}`;
        toast.textContent = message;

        document.body.appendChild(toast);
        
        // Fade in (next frame so the transition fires)
        requestAnimationFrame(() => toast.classList.add('toast-visible'));
        
        // Fade out and remove
        setTimeout(() => {
            toast.classList.remove('toast-visible');
            toast.addEventListener('transitionend', () => toast.remove(), { once: true });
        }, 3000);
    }
}

// Initialize the app when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    window.logbook = new SkydivingLogbook();
});

// Service Worker registration for offline functionality
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('./sw.js')
            .then((registration) => {
                console.log('SW registered: ', registration);
                // Check for SW updates periodically (every 30 min)
                setInterval(() => registration.update(), 30 * 60 * 1000);
            })
            .catch((registrationError) => {
                console.log('SW registration failed: ', registrationError);
            });
    });
}
