// Skydiving Logbook App - Main Application Logic
class SkydivingLogbook {
    constructor() {
        this.jumps = JSON.parse(localStorage.getItem('skydiving-jumps')) || [];
        this.settings = JSON.parse(localStorage.getItem('skydiving-settings')) || {
            startingJumpNumber: 1
        };
        
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
            { id: 'loc_lfoz',    name: 'LFOZ Epcol' },
            { id: 'loc_klatovy', name: 'Klatovy' },
            { id: 'loc_ravenna', name: 'Ravenna' },
            { id: 'loc_palm',    name: 'The Palm Skydive Dubai' },
            { id: 'loc_desert',  name: 'Desert Skydive Dubai' }
        ];
        
        // Migrate old equipment data if needed
        this.migrateOldEquipmentData();
        
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
        // Jump form submission
        document.getElementById('jumpForm').addEventListener('submit', (e) => {
            e.preventDefault();
            const multiplier = parseInt(document.getElementById('jumpMultiplier').value) || 1;
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
            }
        });

        // Multiplier widget buttons
        document.getElementById('multiplierUp').addEventListener('click', () => {
            const input = document.getElementById('jumpMultiplier');
            const val = parseInt(input.value) || 1;
            if (val < 99) input.value = val + 1;
        });
        document.getElementById('multiplierDown').addEventListener('click', () => {
            const input = document.getElementById('jumpMultiplier');
            const val = parseInt(input.value) || 1;
            if (val > 1) input.value = val - 1;
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

        // Modal close
        document.querySelector('.close').addEventListener('click', () => {
            this.closeModal();
        });

        // Export data
        document.getElementById('exportBtn').addEventListener('click', () => {
            this.exportData();
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
            if (e.target === settingsModal) {
                this.closeModal();
            }
            if (e.target === equipmentModal) {
                this.closeEquipmentModal();
            }
            if (e.target === componentModal) {
                this.closeComponentModal();
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
                this.locations.push({ id: newId, name: jump.location });
                this.saveComponentsToLocalStorage();
                this.updateLocationDatalist();
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
        
        // Sync to Google Sheets
        if (navigator.onLine && window.SheetsAPI) {
            if (isPastJump) {
                // Existing jumps were renumbered — overwrite the whole sheet
                window.SheetsAPI.syncAfterDelete(this.jumps);
            } else {
                window.SheetsAPI.syncJump(jump);
            }
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

        // Cutoff: 3 days ago at midnight
        const cutoff = new Date();
        cutoff.setHours(0, 0, 0, 0);
        cutoff.setDate(cutoff.getDate() - 3);

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
        if (navigator.onLine && window.SheetsAPI) {
            window.SheetsAPI.syncAfterDelete(this.jumps);
        }
    }
    
    renumberJumps() {
        // Sort jumps by date to maintain chronological order when renumbering
        this.jumps.sort((a, b) => new Date(a.date) - new Date(b.date));
        
        // Renumber jumps starting from the configured starting number
        this.jumps.forEach((jump, index) => {
            jump.jumpNumber = this.settings.startingJumpNumber + index;
        });
    }

    openSettingsModal() {
        document.getElementById('startingJumpNumber').value = this.settings.startingJumpNumber;

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

        document.getElementById('settingsModal').style.display = 'block';
    }

    closeModal() {
        document.getElementById('settingsModal').style.display = 'none';
    }

    saveSettings() {
        const startingJumpNumber = parseInt(document.getElementById('startingJumpNumber').value);
        
        if (!startingJumpNumber || startingJumpNumber < 1) {
            this.showMessage('Please enter a valid starting jump number (1 or higher)', 'error');
            return;
        }

        this.settings.startingJumpNumber = startingJumpNumber;
        localStorage.setItem('skydiving-settings', JSON.stringify(this.settings));

        // ── Save Google Sheets config to localStorage ──
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

    saveToLocalStorage() {
        localStorage.setItem('skydiving-jumps', JSON.stringify(this.jumps));
        // Mark that there are local changes not yet pushed to the sheet
        localStorage.setItem('skydiving-needs-sync', '1');
    }

    saveComponentsToLocalStorage() {
        localStorage.setItem('skydiving-harnesses', JSON.stringify(this.harnesses));
        localStorage.setItem('skydiving-canopies', JSON.stringify(this.canopies));
        localStorage.setItem('skydiving-equipment-rigs', JSON.stringify(this.equipmentRigs));
        localStorage.setItem('skydiving-locations', JSON.stringify(this.locations));
        // Stamp the modification time so the other device knows to pull this version
        localStorage.setItem('skydiving-equipment-modified', new Date().toISOString());
        
        // Push to Google Sheets if online
        if (navigator.onLine && window.SheetsAPI?.initialized) {
            window.SheetsAPI.syncEquipmentToSheet();
        }
    }
    
    migrateOldEquipmentData() {
        const oldEquipment = JSON.parse(localStorage.getItem('skydiving-equipment'));
        if (oldEquipment && oldEquipment.length > 0 && this.equipmentRigs.length === 0) {
            // Migrate old simple equipment to new rig format
            oldEquipment.forEach(eq => {
                this.equipmentRigs.push({
                    id: eq.id,
                    name: eq.name,
                    harnessId: 'legacy',
                    canopyId: 'legacy',
                    linesetNumber: 1,
                    previousJumps: 0,
                    jumpCount: 0,
                    archived: false
                });
            });
            
            // Add legacy components
            if (!this.harnesses.find(r => r.id === 'legacy')) {
                this.harnesses.push({ id: 'legacy', name: 'Legacy Harness' });
            }
            if (!this.canopies.find(c => c.id === 'legacy')) {
                this.canopies.push({ id: 'legacy', name: 'Legacy Canopy' });
            }
            
            this.saveComponentsToLocalStorage();
            localStorage.removeItem('skydiving-equipment'); // Remove old data
        }
        
        // Migrate existing equipment rigs to new format
        let needsSave = false;
        this.equipmentRigs.forEach(eq => {
            // Migrate old startingJumpNumber → previousJumps (plain count of pre-app jumps)
            if (eq.previousJumps === undefined) {
                // startingJumpNumber was stored as the first jump number tracked, so
                // pre-app count = startingJumpNumber - 1.  Default was 1 → 0 pre-app jumps.
                const old = eq.startingJumpNumber;
                eq.previousJumps = (old && old > 1) ? old - 1 : 0;
                delete eq.startingJumpNumber;
                needsSave = true;
            }
            if (eq.jumpCount === undefined) {
                eq.jumpCount = 0;
                needsSave = true;
            }
            // Migrate linesetId → linesetNumber
            if (eq.linesetId && eq.linesetNumber === undefined) {
                const oldLinesets = JSON.parse(localStorage.getItem('skydiving-linesets') || '[]');
                const ls = oldLinesets.find(l => l.id === eq.linesetId);
                const nameMatch = ls && ls.name.match(/Lineset\s*#?(\d+)/i);
                eq.linesetNumber = nameMatch ? parseInt(nameMatch[1], 10) : 1;
                delete eq.linesetId;
                needsSave = true;
            }
            // Update names to auto-generated format if components exist
            if (eq.harnessId && eq.canopyId && eq.linesetNumber) {
                const harness = this.harnesses.find(r => r.id === eq.harnessId);
                const canopy = this.canopies.find(c => c.id === eq.canopyId);
                if (harness && canopy) {
                    const hybridSuffix = /-Hybrid$/i.test(eq.name || '') ? '-Hybrid' : '';
                    const autoName = `${harness.name}-${canopy.name}-Lineset#${eq.linesetNumber}${hybridSuffix}`;
                    if (eq.name !== autoName) {
                        eq.name = autoName;
                        needsSave = true;
                    }
                }
            }
        });
        
        if (needsSave) {
            this.saveComponentsToLocalStorage();
            // Clean up old linesets localStorage after migration
            localStorage.removeItem('skydiving-linesets');
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
            this.saveComponentsToLocalStorage();
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
            
            return `
                <div class="equipment-item ${eq.archived ? 'archived' : ''}">
                    <div class="equipment-info">
                        <span class="equipment-name">${displayName}</span>
                        ${components}
                        ${jumpInfo}
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

    addEquipment() {
        document.getElementById('equipmentForm').reset();
        document.getElementById('equipmentId').value = '';
        document.getElementById('equipmentStartingJumpNumber').value = 0;
        this.populateComponentSelects();
        document.getElementById('equipmentModal').style.display = 'block';
    }
    
    populateComponentSelects() {
        // Populate harness select
        const harnessSelect = document.getElementById('equipmentHarness');
        harnessSelect.innerHTML = '<option value="">Select Harness</option>';
        this.harnesses.forEach(harness => {
            const option = document.createElement('option');
            option.value = harness.id;
            option.textContent = harness.name;
            harnessSelect.appendChild(option);
        });
        
        // Populate canopy select
        const canopySelect = document.getElementById('equipmentCanopy');
        canopySelect.innerHTML = '<option value="">Select Canopy</option>';
        this.canopies.forEach(canopy => {
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
            
            document.getElementById('equipmentModal').style.display = 'block';
        }
    }

    saveEquipment() {
        const id = document.getElementById('equipmentId').value;
        const harnessId = document.getElementById('equipmentHarness').value;
        const canopyId = document.getElementById('equipmentCanopy').value;
        const linesetNumber = Math.max(1, parseInt(document.getElementById('equipmentLinesetNumber').value) || 1);
        const previousJumps = Math.max(0, parseInt(document.getElementById('equipmentStartingJumpNumber').value) || 0);
        
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
                archived: false
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
        document.getElementById('componentModalTitle').textContent = `Add ${type.charAt(0).toUpperCase() + type.slice(1)}`;
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
                component.name = name;
            }
        } else {
            // Add new
            const newId = type + '_' + Date.now();
            collection.push({ id: newId, name: name });
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
    
    renderComponents(type) {
        const container = document.getElementById('equipmentList');
        const collection = this[type];
        
        if (collection.length === 0) {
            container.innerHTML = `<p class="no-items">No ${type} added yet.</p>`;
            return;
        }
        
        container.innerHTML = collection.map(component => `
            <div class="equipment-item">
                <div class="equipment-info">
                    <span class="equipment-name">${component.name}</span>
                </div>
                <div class="equipment-actions">
                    <button onclick="window.logbook.editComponent('${component.id}', '${type}')" class="btn-edit">Edit</button>
                    <button onclick="window.logbook.deleteComponent('${component.id}', '${type}')" class="btn-delete">Delete</button>
                </div>
            </div>
        `).join('');
    }
    
    editComponent(id, type) {
        const collection = this[type];
        const component = collection.find(c => c.id === id);
        if (component) {
            const singular = this._singularize(type);
            document.getElementById('componentId').value = component.id;
            document.getElementById('componentName').value = component.name;
            document.getElementById('componentType').value = singular;
            document.getElementById('componentModalTitle').textContent = `Edit ${singular.charAt(0).toUpperCase() + singular.slice(1)}`;
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
                const redThreshold = stat.hybrid ? 90 : 180;
                const orangeThreshold = stat.hybrid ? 60 : 140;
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

    exportData() {
        if (this.jumps.length === 0) {
            this.showMessage('No jumps to export', 'error');
            return;
        }

        const csvHeader = 'Jump Number,Date,Location,Equipment,Notes,Timestamp\n';
        const csvData = this.jumps.map(jump => {
            // Resolve rig ID to a readable name for the export
            let equipmentDisplay = jump.equipment;
            const rig = this.equipmentRigs.find(eq => eq.id === jump.equipment);
            if (rig) {
                const harness = this.harnesses.find(r => r.id === rig.harnessId);
                const canopy = this.canopies.find(c => c.id === rig.canopyId);
                equipmentDisplay = (harness && canopy && rig.linesetNumber)
                    ? `${harness.name} + ${canopy.name} + Lineset#${rig.linesetNumber}`
                    : rig.name;
            }
            return `${jump.jumpNumber},"${jump.date}","${jump.location}","${equipmentDisplay}","${jump.notes}","${jump.timestamp}"`;
        }).join('\n');

        const csvContent = csvHeader + csvData;
        const blob = new Blob([csvContent], { type: 'text/csv' });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = `skydiving-logbook-${new Date().toISOString().split('T')[0]}.csv`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        this.showMessage('Data exported successfully!', 'success');
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