// Swooper Logbook App - Main Application Logic
class SkydivingLogbook {
    constructor() {
        // Data arrays — populated asynchronously from IndexedDB in init()
        this.jumps = [];
        this.harnesses = [];
        this.canopies = [];
        this.locations = [];

        this.settings = JSON.parse(localStorage.getItem('skydiving-settings')) || {
            startingJumpNumber: 1,
            recentJumpsDays: 7,
            standardRedThreshold: 160,
            standardOrangeThreshold: 140,
            hybridRedThreshold: 80,
            hybridOrangeThreshold: 60
        };
        // Backfill for existing saved settings that predate these fields
        if (this.settings.recentJumpsDays === undefined) {
            this.settings.recentJumpsDays = 7;
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
        
        this.currentView = 'jumps'; // 'jumps', 'equipment', 'stats'
        this.equipmentSubView = 'canopies'; // 'canopies', 'harnesses', 'locations'
        this.showArchivedStats = false;
        this.activeJumpNoteId = null;
        this._olderJumpsCache = []; // cached older jumps for lazy rendering
        this._renderedOlderCount = 0;
        
        this.init();
    }

    async init() {
        // Open IndexedDB and migrate from localStorage if needed
        try {
            await DB.open();
            await DB.migrateFromLocalStorage();
        } catch (err) {
            console.error('[DB] IndexedDB unavailable, running in memory-only mode:', err);
            // On Safari/iOS, storage may be blocked; offer Storage Access API flow
            if (typeof document.requestStorageAccess === 'function') {
                this.showStorageBlockedBanner();
            }
        }

        // Load data from IndexedDB (falls back to defaults if empty / IDB failed)
        try {
            const [jumps, canopies, harnesses, locations] = await Promise.all([
                DB.getAllJumps(),
                DB.getAll('canopies'),
                DB.getAll('harnesses'),
                DB.getAll('locations')
            ]);
            this.jumps     = jumps.length     ? jumps     : [];
            this.ensureJumpIds();
            this.canopies  = canopies.length  ? canopies  : [];
            this.canopies.sort((a, b) => (a.sortOrder ?? Infinity) - (b.sortOrder ?? Infinity));
            this.harnesses = harnesses.length ? harnesses : [
                { id: 'javelin', name: 'Javelin' },
                { id: 'mutant', name: 'Mutant' }
            ];
            this.locations = locations.length ? locations : [
                { id: 'loc_lfoz',    name: 'LFOZ Epcol',             lat: 47.90293836878678, lng: 2.168939324589268 },
                { id: 'loc_klatovy', name: 'Klatovy LKKT',                lat: 49.4181547, lng: 13.321609 },
                { id: 'loc_ravenna', name: 'Ravenna LIDR',                lat: 44.362208, lng: 12.202889 },
                { id: 'loc_eloy',    name: 'Eloy Skydive Arizona',   lat: 32.7555, lng: -111.8467 },
                { id: 'loc_palm',    name: 'The Palm Skydive Dubai', lat: 25.112176, lng: 55.153587 },
                { id: 'loc_desert',  name: 'Desert Skydive Dubai',   lat: 24.985654, lng: 55.146111 }
            ];
            this.locations.sort((a, b) => (a.sortOrder ?? Infinity) - (b.sortOrder ?? Infinity));
        } catch (err) {
            console.error('[DB] Failed to load data from IndexedDB:', err);
        }

        // Backfill lat/lng on locations loaded from older data
        this.locations.forEach(loc => {
            if (loc.lat === undefined) loc.lat = null;
            if (loc.lng === undefined) loc.lng = null;
        });

        // Ensure every canopy has a linesets array with at least one lineset
        this.canopies.forEach(canopy => {
            if (!Array.isArray(canopy.linesets)) canopy.linesets = [];
            if (canopy.linesets.length === 0) {
                canopy.linesets.push({ number: 1, hybrid: false, previousJumps: 0, archived: false });
            }
        });

        // Initialize jump counts for canopy linesets
        this.initializeCanopyLinesetJumpCounts();

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
        
        // Resume post-login flow when returning from an OAuth redirect (mobile)
        await window.AuthManager.ready;
        await this._resumeOAuthRedirectIfNeeded();

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
            return Math.max(1, Math.min(150, parsed));
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
            if (val < 150) input.value = val + 1;
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

        document.getElementById('resetAppBtn').addEventListener('click', () => {
            this.resetAppToFirstLaunch();
        });

        document.getElementById('useCurrentLocationBtn').addEventListener('click', () => {
            this.setComponentCoordsFromGPS();
        });

        // Google Sheets Integration modal (OAuth)
        document.getElementById('googleSheetsIntegrationBtn').addEventListener('click', () => {
            this.openSheetsModal();
        });

        document.getElementById('sheetsClose').addEventListener('click', () => {
            this.closeSheetsModal();
        });

        const jumpNoteClose = document.getElementById('jumpNoteClose');
        if (jumpNoteClose) {
            jumpNoteClose.addEventListener('click', () => {
                this.closeJumpNotePopup();
            });
        }

        const jumpNoteSave = document.getElementById('jumpNoteSave');
        if (jumpNoteSave) {
            jumpNoteSave.addEventListener('click', () => {
                this.saveJumpNote();
            });
        }

        // Search Notes modal
        document.getElementById('searchJumpsBtn').addEventListener('click', () => {
            this.openSearchNotesModal();
        });
        document.getElementById('searchNotesClose').addEventListener('click', () => {
            this.closeSearchNotesModal();
        });
        document.getElementById('searchNotesGoBtn').addEventListener('click', () => {
            this.executeNoteSearch();
        });
        document.getElementById('searchNotesInput').addEventListener('keydown', (e) => {
            if (e.key === 'Enter') this.executeNoteSearch();
        });

        document.getElementById('googleSignInBtn').addEventListener('click', () => {
            this.handleGoogleSignIn();
        });

        document.getElementById('googleSignOutBtn').addEventListener('click', () => {
            this.handleGoogleSignOut();
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
        document.getElementById('importChoiceMergeBtn').addEventListener('click', () => {
            this.applyImportChoice('merge');
        });
        document.getElementById('importChoiceReplaceBtn').addEventListener('click', () => {
            this.applyImportChoice('replace');
        });
        document.getElementById('importChoiceModalClose').addEventListener('click', () => {
            this.closeImportChoiceModal();
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
        document.getElementById('addCanopyBtn').addEventListener('click', () => {
            this.addComponent('canopy');
        });

        document.getElementById('linesetForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.saveLineset();
        });
        
        // Component management
        document.getElementById('addHarnessBtn').addEventListener('click', () => {
            this.addComponent('harness');
        });
        
        document.getElementById('addLocationBtn').addEventListener('click', () => {
            this.addComponent('location');
        });
        
        document.getElementById('componentForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.saveComponent();
        });

        // Update lineset hint when canopy selection changes in jump form
        document.getElementById('equipment').addEventListener('change', () => {
            this.updateLinesetHint();
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
            const linesetModal = document.getElementById('linesetModal');
            const componentModal = document.getElementById('componentModal');
            const sheetsModal = document.getElementById('sheetsModal');
            const jumpNoteModal = document.getElementById('jumpNoteModal');
            if (e.target === settingsModal) {
                this.closeModal();
            }
            if (e.target === linesetModal) {
                this.closeLinesetModal();
            }
            if (e.target === componentModal) {
                this.closeComponentModal();
            }
            if (e.target === sheetsModal) {
                this.closeSheetsModal();
            }
            if (e.target === jumpNoteModal) {
                this.closeJumpNotePopup();
            }
            const importChoiceModal = document.getElementById('importChoiceModal');
            if (e.target === importChoiceModal) {
                this.closeImportChoiceModal();
            }
            const searchNotesModal = document.getElementById('searchNotesModal');
            if (e.target === searchNotesModal) {
                this.closeSearchNotesModal();
            }
        });
    }

    setCurrentDate() {
        const now = new Date();
        const yyyy = now.getFullYear();
        const mm = String(now.getMonth() + 1).padStart(2, '0');
        const dd = String(now.getDate()).padStart(2, '0');
        document.getElementById('date').value = `${yyyy}-${mm}-${dd}`;
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
        
        // Pre-fill canopy only if it is still selectable in the dropdown
        const equipmentSelect = document.getElementById('equipment');
        if (lastJump.equipment && equipmentSelect) {
            const canopy = this.canopies.find(c => c.id === lastJump.equipment && !c.archived);
            const activeLineset = this.getActiveLineset(lastJump.equipment);
            if (canopy && activeLineset) {
                equipmentSelect.value = lastJump.equipment;
                this.updateLinesetHint();
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
        
        // Determine the active lineset for the selected canopy
        let linesetNumber = data.linesetNumber; // may be pre-set by caller
        if (!linesetNumber && data.equipment) {
            linesetNumber = this.getActiveLinesetNumber(data.equipment);
        }
        
        // Remember the highest jump number before insertion to detect a past-date entry
        const maxBefore = this.jumps.length > 0
            ? Math.max(...this.jumps.map(j => j.jumpNumber))
            : this.settings.startingJumpNumber - 1;

        const jump = {
            id: Date.now() + Math.random(),
            jumpId: window.SheetsAPI ? SheetsAPI.generateJumpId() : ('jump-' + Date.now() + '-' + Math.random().toString(36).slice(2)),
            jumpNumber: 0,          // assigned below by renumberJumps()
            date: data.date,
            location: data.location,
            equipment: data.equipment,  // canopy ID
            linesetNumber: linesetNumber || 1,
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
                DB.putAll('locations', this.locations).catch(err => console.error('[DB] Failed to save locations:', err));
                this.updateLocationDatalist();
                this.geocodeLocation(newLoc);
            }
        }

        // Update canopy lineset jump count
        if (jump.equipment) {
            const canopy = this.canopies.find(c => c.id === jump.equipment);
            if (canopy) {
                const lineset = canopy.linesets?.find(ls => ls.number === jump.linesetNumber);
                if (lineset) {
                    lineset.jumpCount = (lineset.jumpCount || 0) + 1;
                }
                DB.replaceAll('canopies', this.canopies).catch(err => console.error('[DB] Failed to save canopies:', err));
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
        const totalEl = document.getElementById('recentJumpsTotal');
        const allJumpsByMonth = this.settings.recentJumpsDays === 0;

        const updateRecentTotal = (count) => {
            if (!totalEl) return;
            if (allJumpsByMonth) {
                totalEl.textContent = '';
                totalEl.style.display = 'none';
            } else {
                totalEl.style.display = '';
                totalEl.textContent = `Total: ${count}`;
            }
        };

        if (this.jumps.length === 0) {
            updateRecentTotal(0);
            jumpsList.innerHTML = '<p class="no-jumps">No jumps logged yet. Add your first jump above!</p>';
            return;
        }

        // Show most recent jumps first
        const sortedJumps = [...this.jumps].sort((a, b) => b.jumpNumber - a.jumpNumber);

        const recentJumps = [];
        const olderJumps = [];

        if (allJumpsByMonth) {
            sortedJumps.forEach(jump => olderJumps.push(jump));
        } else {
            const cutoff = new Date();
            cutoff.setHours(0, 0, 0, 0);
            cutoff.setDate(cutoff.getDate() - (this.settings.recentJumpsDays || 3));

            sortedJumps.forEach(jump => {
                const jumpDate = new Date(jump.date);
                if (jumpDate >= cutoff) {
                    recentJumps.push(jump);
                } else {
                    olderJumps.push(jump);
                }
            });
        }

        // Cache older jumps for lazy "load more" rendering
        this._olderJumpsCache = olderJumps;
        const PAGE_SIZE = 100;
        const endIndex = this._findMonthCompleteIndex(olderJumps, PAGE_SIZE);
        const initialOlder = olderJumps.slice(0, endIndex);
        this._renderedOlderCount = initialOlder.length;

        let html = '';

        // Render recent jumps grouped by day + location (latest day expanded)
        if (recentJumps.length > 0) {
            html += this.renderDayLocationGroups(recentJumps, { expandFirst: true });
        }

        // Render initial page of older jumps grouped by month
        if (initialOlder.length > 0) {
            html += this._renderOlderMonthGroups(initialOlder);
        }

        // "Load more" button if there are remaining older jumps
        const remaining = olderJumps.length - this._renderedOlderCount;
        if (remaining > 0) {
            html += `<button class="btn-secondary load-more-btn" id="loadMoreJumpsBtn" onclick="logbook.loadMoreJumps()">Load more (${remaining} remaining)</button>`;
        }

        updateRecentTotal(recentJumps.length);
        jumpsList.innerHTML = html;
    }

    /** Render a batch of older jumps as collapsed month groups. */
    _renderOlderMonthGroups(jumps) {
        const monthGroups = new Map();
        jumps.forEach(jump => {
            const d = new Date(jump.date);
            const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
            if (!monthGroups.has(key)) {
                monthGroups.set(key, { label: d.toLocaleDateString(undefined, { year: 'numeric', month: 'long' }), jumps: [] });
            }
            monthGroups.get(key).jumps.push(jump);
        });

        let html = '';
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
                        ${this.renderDayLocationGroups(group.jumps)}
                    </div>
                </div>
            `;
        }
        return html;
    }

    /** Given an array of jumps sorted by jumpNumber desc, find the smallest index >= targetCount that sits on a month boundary. */
    _findMonthCompleteIndex(jumps, targetCount) {
        if (targetCount >= jumps.length) return jumps.length;
        const lastDate = new Date(jumps[targetCount - 1].date);
        const lastMonthKey = `${lastDate.getFullYear()}-${String(lastDate.getMonth() + 1).padStart(2, '0')}`;
        let endIndex = targetCount;
        while (endIndex < jumps.length) {
            const d = new Date(jumps[endIndex].date);
            const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
            if (key !== lastMonthKey) break;
            endIndex++;
        }
        return endIndex;
    }

    /** Append the next page of older jumps to the list. */
    loadMoreJumps() {
        const PAGE_SIZE = 100;
        const endIndex = this._findMonthCompleteIndex(
            this._olderJumpsCache,
            this._renderedOlderCount + PAGE_SIZE
        );
        const nextBatch = this._olderJumpsCache.slice(
            this._renderedOlderCount,
            endIndex
        );
        if (nextBatch.length === 0) return;
        this._renderedOlderCount = endIndex;

        // Remove existing load-more button
        const btn = document.getElementById('loadMoreJumpsBtn');
        if (btn) btn.remove();

        // Append new month groups
        const jumpsList = document.getElementById('jumpsList');
        const fragment = document.createElement('div');
        fragment.innerHTML = this._renderOlderMonthGroups(nextBatch);
        while (fragment.firstChild) jumpsList.appendChild(fragment.firstChild);

        // Re-add button if more remain
        const remaining = this._olderJumpsCache.length - this._renderedOlderCount;
        if (remaining > 0) {
            const newBtn = document.createElement('button');
            newBtn.className = 'btn-secondary load-more-btn';
            newBtn.id = 'loadMoreJumpsBtn';
            newBtn.textContent = `Load more (${remaining} remaining)`;
            newBtn.onclick = () => this.loadMoreJumps();
            jumpsList.appendChild(newBtn);
        }
    }

    toggleMonthGroup(key) {
        const body = document.getElementById('month-' + key);
        const arrow = document.getElementById('arrow-' + key);
        if (!body) return;
        const open = body.style.display !== 'none';
        body.style.display = open ? 'none' : 'block';
        if (arrow) arrow.innerHTML = open ? '&#9654;' : '&#9660;';
    }

    renderDayLocationGroups(jumps, { expandFirst = false } = {}) {
        const groups = [];
        const groupMap = new Map();
        jumps.forEach(jump => {
            const key = `${jump.date}|${jump.location || ''}`;
            if (!groupMap.has(key)) {
                const d = new Date(jump.date + 'T00:00:00');
                const entry = {
                    key,
                    dateLabel: d.toLocaleDateString(undefined, { weekday: 'short', month: 'short', day: 'numeric' }),
                    location: jump.location || '',
                    jumps: []
                };
                groupMap.set(key, entry);
                groups.push(entry);
            }
            groupMap.get(key).jumps.push(jump);
        });

        return groups.map((group, i) => {
            const expanded = expandFirst && i === 0;
            const dayId = 'day-' + group.key.replace(/[^a-zA-Z0-9]/g, '_');
            return `
            <div class="day-location-group">
                <div class="day-location-header" onclick="logbook.toggleDayGroup('${dayId}')">
                    <span class="day-group-arrow" id="arrow-${dayId}">${expanded ? '&#9660;' : '&#9654;'}</span>
                    <span class="day-date">${group.dateLabel}</span>
                    ${group.location ? `<span class="day-location">📍 ${group.location}</span>` : ''}
                    <span class="day-jump-range">${this.getJumpCountLabel(group.jumps)}</span>
                </div>
                <div class="day-group-body" id="${dayId}" style="display:${expanded ? 'block' : 'none'};">
                    ${group.jumps.map(j => this.createJumpRowHTML(j)).join('')}
                </div>
            </div>
        `;
        }).join('');
    }

    toggleDayGroup(dayId) {
        const body = document.getElementById(dayId);
        const arrow = document.getElementById('arrow-' + dayId);
        if (!body) return;
        const open = body.style.display !== 'none';
        body.style.display = open ? 'none' : 'block';
        if (arrow) arrow.innerHTML = open ? '&#9654;' : '&#9660;';
    }

    getJumpCountLabel(jumps) {
        if (!Array.isArray(jumps) || jumps.length === 0) return '';
        return `${jumps.length} jump${jumps.length === 1 ? '' : 's'}`;
    }

    escapeHtml(text) {
        return String(text)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    createJumpRowHTML(jump) {
        let canopyName = jump.equipment;
        const canopy = this.canopies.find(c => c.id === jump.equipment);
        if (canopy) {
            const ls = canopy.linesets?.find(l => l.number === jump.linesetNumber);
            const hybridSuffix = ls?.hybrid ? ' (H)' : '';
            canopyName = `${canopy.name} L#${jump.linesetNumber || 1}${hybridSuffix}`;
        }

        const noteText = typeof jump.notes === 'string' ? jump.notes : '';
        const compactNote = noteText.replace(/\s+/g, ' ').trim();
        const hasNote = compactNote.length > 0;
        const notePreview = hasNote
            ? `${compactNote.slice(0, 15)}${compactNote.length > 15 ? '...' : ''}`
            : '';
        const encodedJumpId = encodeURIComponent(String(jump.id));
        const encodedFullNote = hasNote ? encodeURIComponent(noteText) : '';

        const canopyNameHtml = this.escapeHtml(canopyName).replace(/\d{2,}/g, '<b>$&</b>');

        return `
            <div class="jump-row">
                <span class="jump-number">#${jump.jumpNumber}</span>
                <span class="jump-canopy">🪂 ${canopyNameHtml}</span>
                ${hasNote ? `<button type="button" class="jump-note-preview" onclick="logbook.openJumpNotePopup('${encodedJumpId}', '${encodedFullNote}')" title="View or edit note">${this.escapeHtml(notePreview)}</button>` : ''}
                <button class="delete-jump-btn" onclick="logbook.deleteJump('${jump.id}')" title="Delete jump">❌</button>
            </div>
        `;
    }

    openJumpNotePopup(encodedJumpId, encodedNote) {
        const modal = document.getElementById('jumpNoteModal');
        const content = document.getElementById('jumpNoteContent');
        if (!modal || !content) return;

        this.activeJumpNoteId = decodeURIComponent(encodedJumpId || '');
        const note = decodeURIComponent(encodedNote || '');
        content.value = note;
        modal.style.display = 'block';
        content.focus();
        content.setSelectionRange(content.value.length, content.value.length);
    }

    closeJumpNotePopup() {
        const modal = document.getElementById('jumpNoteModal');
        const content = document.getElementById('jumpNoteContent');
        if (!modal) return;
        this.activeJumpNoteId = null;
        if (content) content.value = '';
        modal.style.display = 'none';
    }

    saveJumpNote() {
        if (!this.activeJumpNoteId) {
            this.showMessage('Jump not found', 'error');
            return;
        }

        const content = document.getElementById('jumpNoteContent');
        if (!content) return;

        const jump = this.jumps.find(item => item.id.toString() === this.activeJumpNoteId.toString());
        if (!jump) {
            this.showMessage('Jump not found', 'error');
            this.closeJumpNotePopup();
            return;
        }

        jump.notes = content.value;
        this.saveToLocalStorage();
        this.renderJumpsList();
        this.closeJumpNotePopup();
        this.showMessage('Jump note saved', 'success');

        if (navigator.onLine && window.SheetsAPI?.initialized) {
            window.SheetsAPI.pushAllWithGuard();
        }
    }

    openSearchNotesModal() {
        const modal = document.getElementById('searchNotesModal');
        const input = document.getElementById('searchNotesInput');
        if (!modal) return;
        modal.style.display = 'block';
        if (input) {
            input.value = '';
            input.focus();
        }
        document.getElementById('searchNotesResults').innerHTML =
            '<p class="search-notes-placeholder">Enter a search term to find jumps by note text.</p>';
    }

    closeSearchNotesModal() {
        const modal = document.getElementById('searchNotesModal');
        if (modal) modal.style.display = 'none';
    }

    executeNoteSearch() {
        const input = document.getElementById('searchNotesInput');
        const resultsContainer = document.getElementById('searchNotesResults');
        if (!input || !resultsContainer) return;

        const query = input.value.trim();
        if (!query) {
            resultsContainer.innerHTML =
                '<p class="search-notes-placeholder">Enter a search term to find jumps by note text.</p>';
            return;
        }

        const lowerQuery = query.toLowerCase();
        const matches = this.jumps
            .filter(j => typeof j.notes === 'string' && j.notes.toLowerCase().includes(lowerQuery))
            .sort((a, b) => b.jumpNumber - a.jumpNumber);

        if (matches.length === 0) {
            resultsContainer.innerHTML = '<p class="search-notes-placeholder">No jumps found matching that text.</p>';
            return;
        }

        const escapedQuery = this.escapeHtml(query);
        const queryRegex = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');

        let html = `<p class="search-notes-count">${matches.length} result${matches.length !== 1 ? 's' : ''} found</p>`;
        matches.forEach(jump => {
            const dateStr = new Date(jump.date + 'T00:00:00')
                .toLocaleDateString(undefined, { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' });

            let canopyName = jump.equipment || '';
            const canopy = this.canopies.find(c => c.id === jump.equipment);
            if (canopy) {
                const ls = canopy.linesets?.find(l => l.number === jump.linesetNumber);
                const hybridSuffix = ls?.hybrid ? ' (H)' : '';
                canopyName = `${canopy.name} L#${jump.linesetNumber || 1}${hybridSuffix}`;
            }

            const noteHtml = this.escapeHtml(jump.notes).replace(queryRegex, m => `<mark>${m}</mark>`);

            const encodedJumpId = encodeURIComponent(String(jump.id));
            const encodedNote = encodeURIComponent(jump.notes);

            html += `
                <div class="search-result-item" onclick="logbook.closeSearchNotesModal(); logbook.openJumpNotePopup('${encodedJumpId}', '${encodedNote}')">
                    <div class="search-result-header">
                        <span class="search-result-jump-num">#${jump.jumpNumber}</span>
                        <span class="search-result-date">${dateStr}</span>
                    </div>
                    ${canopyName ? `<div class="search-result-canopy">🪂 ${this.escapeHtml(canopyName)}</div>` : ''}
                    <div class="search-result-note">${noteHtml}</div>
                </div>`;
        });

        resultsContainer.innerHTML = html;
    }

    deleteJump(jumpId) {
        if (!confirm('Are you sure you want to delete this jump? This action cannot be undone.')) {
            return;
        }
        
        const jumpIndex = this.jumps.findIndex(jump => jump.id.toString() === jumpId.toString());
        if (jumpIndex === -1) {
            this.showMessage('Jump not found', 'error');
            return;
        }
        
        const deletedJump = this.jumps[jumpIndex];
        const stableJumpId = deletedJump.jumpId || (window.SheetsAPI ? SheetsAPI.generateJumpId() : ('jump-' + Date.now() + '-' + Math.random().toString(36).slice(2)));
        if (!deletedJump.jumpId) deletedJump.jumpId = stableJumpId;

        if (deletedJump.equipment) {
            const canopy = this.canopies.find(c => c.id === deletedJump.equipment);
            if (canopy) {
                const lineset = canopy.linesets?.find(ls => ls.number === deletedJump.linesetNumber);
                if (lineset && lineset.jumpCount > 0) {
                    lineset.jumpCount = lineset.jumpCount - 1;
                }
                this.saveComponentsToLocalStorage();
            }
        }

        this.jumps.splice(jumpIndex, 1);
        this.renumberJumps();
        this.saveToLocalStorage();
        this.updateStats();
        this.renderJumpsList();
        if (this.currentView === 'equipment') this.renderEquipmentView();
        this.showMessage('Jump deleted successfully', 'success');

        if (navigator.onLine && window.SheetsAPI?.initialized) {
            window.SheetsAPI.appendDeletedJumps([stableJumpId]).catch(err => console.error('[Sync] appendDeletedJumps failed:', err));
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

    openSettingsModal() {
        document.getElementById('startingJumpNumber').value = this.settings.startingJumpNumber;
        const prev = this.settings.previousStartingJump;
        const current = this.settings.startingJumpNumber;
        const labelEl = document.getElementById('startingJumpNumberLabel');
        const showPrevious = prev != null && prev !== 1 && prev !== current;
        labelEl.textContent = showPrevious ? `Starting Jump Number (previous=${prev})` : 'Starting Jump Number';
        document.getElementById('recentJumpsDays').value = this.settings.recentJumpsDays ?? 3;
        document.getElementById('standardRedThreshold').value = this.settings.standardRedThreshold ?? 160;
        document.getElementById('standardOrangeThreshold').value = this.settings.standardOrangeThreshold ?? 140;
        document.getElementById('hybridRedThreshold').value = this.settings.hybridRedThreshold ?? 80;
        document.getElementById('hybridOrangeThreshold').value = this.settings.hybridOrangeThreshold ?? 60;

        document.getElementById('settingsModal').style.display = 'block';
    }

    closeModal() {
        document.getElementById('settingsModal').style.display = 'none';
    }

    openSheetsModal() {
        const statusEl = document.getElementById('sheetsConfigStatus');
        const signedOutEl = document.getElementById('oauthSignedOut');
        const signedInEl = document.getElementById('oauthSignedIn');

        const isSignedIn = window.AuthManager?.isSignedIn();
        const spreadsheetId = localStorage.getItem('oauth-spreadsheet-id') || '';

        if (isSignedIn && spreadsheetId) {
            signedOutEl.style.display = 'none';
            signedInEl.style.display = 'block';
            document.getElementById('oauthUserEmail').textContent = window.AuthManager.userEmail || 'Signed in';
            document.getElementById('oauthSheetLink').href = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(spreadsheetId)}/edit`;
            statusEl.textContent = '✅ Connected to Google Sheets';
            statusEl.style.color = '#2e7d32';
        } else {
            signedOutEl.style.display = 'block';
            signedInEl.style.display = 'none';

            // Only show Client ID input if no hardcoded ID (self-hosting scenario)
            const hasHardcodedId = typeof OAUTH_CLIENT_ID !== 'undefined' && OAUTH_CLIENT_ID;
            const clientIdGroup = document.getElementById('oauthClientIdGroup');
            if (!hasHardcodedId) {
                clientIdGroup.style.display = 'block';
                document.getElementById('cfgOAuthClientId').value = localStorage.getItem('oauth-client-id') || '';
                statusEl.textContent = 'ℹ️ Enter your OAuth Client ID and sign in to enable sync';
            } else {
                clientIdGroup.style.display = 'none';
                statusEl.textContent = 'ℹ️ Sign in with your Google account to enable sync';
            }
            statusEl.style.color = '#666';
        }

        document.getElementById('sheetsModal').style.display = 'block';
    }

    closeSheetsModal() {
        document.getElementById('sheetsModal').style.display = 'none';
    }

    /**
     * After an OAuth2 redirect sign-in on mobile, the page reloads fresh.
     * If sessionStorage has the pending flag and the token was recovered from
     * the URL hash, continue with the spreadsheet setup that was interrupted.
     */
    async _resumeOAuthRedirectIfNeeded() {
        if (!sessionStorage.getItem('oauth-redirect-pending')) return;
        sessionStorage.removeItem('oauth-redirect-pending');

        if (!window.AuthManager?.isSignedIn()) {
            // User cancelled or Google returned an error
            this.showMessage('Sign-in was cancelled or failed', 'error');
            return;
        }

        // Mirror the post-sign-in logic from handleGoogleSignIn()
        try {
            let spreadsheetId = localStorage.getItem('oauth-spreadsheet-id') || '';
            if (!spreadsheetId) {
                spreadsheetId = await window.SheetsAPI.findOrCreateSpreadsheet();
            }
            window.SheetsAPI.reinitialize(spreadsheetId);
            this.showMessage('Connected to Google Sheets!', 'success');
            window.SheetsAPI.syncWithSheet().catch(err =>
                console.error('[Sheets] Post-redirect sync failed:', err)
            );
        } catch (error) {
            console.error('[Auth] Post-redirect setup failed:', error);
            this.showMessage('Sign-in succeeded but setup failed: ' + error.message, 'error');
        }
    }

    async handleGoogleSignIn() {
        // Use hardcoded constant, or fall back to manual input for self-hosters
        const hasHardcodedId = typeof OAUTH_CLIENT_ID !== 'undefined' && OAUTH_CLIENT_ID;
        const clientId = hasHardcodedId
            ? OAUTH_CLIENT_ID
            : document.getElementById('cfgOAuthClientId').value.trim();

        if (!clientId) {
            this.showMessage('Please enter your OAuth Client ID first', 'error');
            return;
        }

        const signInBtn = document.getElementById('googleSignInBtn');
        try {
            // Disable button and show loading state
            if (signInBtn) {
                signInBtn.disabled = true;
                signInBtn.textContent = 'Signing in…';
            }

            // Configure AuthManager with the client ID if changed
            if (clientId !== window.AuthManager.clientId) {
                await window.AuthManager.configure(clientId);
            }

            await window.AuthManager.signIn();

            // If no spreadsheet exists yet, find existing or create one
            let spreadsheetId = localStorage.getItem('oauth-spreadsheet-id') || '';
            if (!spreadsheetId) {
                if (signInBtn) signInBtn.textContent = 'Looking for spreadsheet…';
                spreadsheetId = await window.SheetsAPI.findOrCreateSpreadsheet();
                window.SheetsAPI.reinitialize(spreadsheetId);
            } else {
                window.SheetsAPI.reinitialize(spreadsheetId);
            }

            this.showMessage('Connected to Google Sheets!', 'success');
            this.closeSheetsModal(); // close settings and return to main view

            // Trigger a sync (fire-and-forget; errors logged inside)
            window.SheetsAPI.syncWithSheet().catch(err =>
                console.error('[Sheets] Post-connect sync failed:', err)
            );
        } catch (error) {
            console.error('[Auth] Sign-in failed:', error);
            this.showMessage('Sign-in failed: ' + error.message, 'error');
        } finally {
            // Always restore button state
            if (signInBtn) {
                signInBtn.disabled = false;
                signInBtn.textContent = 'Sign in with Google';
            }
        }
    }

    async handleGoogleSignOut() {
        if (!confirm('Sign out and disconnect Google Sheets sync?')) return;

        await window.AuthManager.signOut();
        window.SheetsAPI.initialized = false;
        window.SheetsAPI.spreadsheetId = '';
        window.SheetsAPI._cancelPoll();
        window.SheetsAPI.updateSyncStatus('Not signed in');

        this.showMessage('Signed out from Google Sheets', 'success');
        this.openSheetsModal(); // refresh modal state
    }

    async resetAppToFirstLaunch() {
        const confirmed = confirm('⚠️ This will permanently erase ALL app data (jumps, equipment, settings) and reset to first-launch state.\n\nThis action CANNOT be undone.\n\nContinue?');
        if (!confirmed) return;

        localStorage.clear();
        try { await DB.clearAll(); } catch (_) { /* IDB may not be open */ }
        this.showMessage('App data erased. Reloading...', 'success');
        setTimeout(() => window.location.reload(), 300);
    }

    // saveSheetsConfig is no longer needed — OAuth sign-in handles everything.

    async saveSettings() {
        const startingJumpNumber = parseInt(document.getElementById('startingJumpNumber').value);
        
        if (!startingJumpNumber || startingJumpNumber < 1) {
            this.showMessage('Please enter a valid starting jump number (1 or higher)', 'error');
            return;
        }

        const recentJumpsDays = parseInt(document.getElementById('recentJumpsDays').value, 10);
        if (Number.isNaN(recentJumpsDays) || recentJumpsDays < 0) {
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

        const previousStartingJumpNumber = this.settings.startingJumpNumber;
        this.settings.startingJumpNumber = startingJumpNumber;
        // Always persist the value we're leaving, so we can show (previous=XX) when opening settings; display hides it when 1 or when equal to current
        this.settings.previousStartingJump = previousStartingJumpNumber;
        this.settings.recentJumpsDays = recentJumpsDays;
        this.settings.standardRedThreshold = standardRedThreshold;
        this.settings.standardOrangeThreshold = standardOrangeThreshold;
        this.settings.hybridRedThreshold = hybridRedThreshold;
        this.settings.hybridOrangeThreshold = hybridOrangeThreshold;
        localStorage.setItem('skydiving-settings', JSON.stringify(this.settings));
        this.markEquipmentModified();

        // Mark data as locally modified so the background poller detects pending changes.
        localStorage.setItem('skydiving-data-modified', new Date().toISOString());

        // If starting jump number changed: renumber all jumps, persist, then refresh
        if (previousStartingJumpNumber !== startingJumpNumber) {
            this.renumberJumps();
            await DB.replaceAllJumps(this.jumps).catch(err => console.error('[DB] Failed to save jumps after renumber:', err));
            this.markJumpsModified();
            localStorage.setItem('skydiving-needs-sync', '1');
            localStorage.setItem('skydiving-data-modified', new Date().toISOString());
        }

        // Push settings to Google Sheets if online (with timestamp so _syncMeta is updated)
        if (navigator.onLine && window.SheetsAPI?.initialized) {
            window.SheetsAPI.syncEquipmentToSheet(new Date().toISOString());
        }
        
        this.closeModal();
        if (previousStartingJumpNumber !== startingJumpNumber) {
            this.showMessage('Settings saved. Jump numbers updated. Reloading...', 'success');
            setTimeout(() => window.location.reload(), 300);
        } else {
            this.showMessage('Settings saved successfully!', 'success');
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

    /** Ensure every jump has a stable jumpId (for sync/deletedJumps). Persists if any were added. */
    ensureJumpIds() {
        const generator = () => (typeof crypto !== 'undefined' && crypto.randomUUID)
            ? crypto.randomUUID()
            : 'jump-' + Date.now() + '-' + Math.random().toString(36).slice(2);
        let needsSave = false;
        this.jumps.forEach(jump => {
            if (!jump.jumpId) {
                jump.jumpId = generator();
                needsSave = true;
            }
        });
        if (needsSave) {
            DB.replaceAllJumps(this.jumps).catch(err => console.error('[DB] Failed to save jumps after jumpId migration:', err));
        }
    }

    saveToLocalStorage() {
        DB.replaceAllJumps(this.jumps).catch(err => console.error('[DB] Failed to save jumps:', err));
        this.markJumpsModified();
        // Mark that there are local changes not yet pushed to the sheet
        localStorage.setItem('skydiving-needs-sync', '1');
        localStorage.setItem('skydiving-data-modified', new Date().toISOString());
    }

    saveComponentsToLocalStorage() {
        Promise.all([
            DB.replaceAll('harnesses', this.harnesses),
            DB.replaceAll('canopies', this.canopies),
            DB.replaceAll('locations', this.locations)
        ]).catch(err => console.error('[DB] Failed to save equipment:', err));
        this.markEquipmentModified();
        // Mark data as locally modified so the background poller detects pending changes.
        localStorage.setItem('skydiving-data-modified', new Date().toISOString());
        
        // Push to Google Sheets if online (with timestamp so _syncMeta is updated)
        if (navigator.onLine && window.SheetsAPI?.initialized) {
            window.SheetsAPI.syncEquipmentToSheet(new Date().toISOString());
        }
    }
    
    
    initializeCanopyLinesetJumpCounts() {
        let needsSave = false;
        this.canopies.forEach(canopy => {
            if (!Array.isArray(canopy.linesets)) return;
            canopy.linesets.forEach(ls => {
                const count = this.jumps.filter(j =>
                    j.equipment === canopy.id && j.linesetNumber === ls.number
                ).length;
                if (ls.jumpCount !== count) {
                    ls.jumpCount = count;
                    needsSave = true;
                }
            });
        });
        
        if (needsSave) {
            DB.replaceAll('canopies', this.canopies).catch(err => console.error('[DB] Failed to save canopy counts:', err));
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
        if (viewName === 'jumps') {
            this.renderJumpsList();
        } else if (viewName === 'equipment') {
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
            'canopies':     'Canopies',
            'harnesses':    'Harnesses',
            'locations':    'Drop Zones / Locations'
        };
        
        document.getElementById('equipmentSectionTitle').textContent = titleMap[subView];
        
        // Show/hide appropriate buttons
        document.getElementById('addHarnessBtn').style.display    = subView === 'harnesses' ? 'block' : 'none';
        document.getElementById('addCanopyBtn').style.display     = subView === 'canopies'  ? 'block' : 'none';
        document.getElementById('addLocationBtn').style.display   = subView === 'locations' ? 'block' : 'none';
        
        this.renderEquipmentView();
    }
    
    renderEquipmentView() {
        switch(this.equipmentSubView) {
            case 'canopies':     this.renderCanopiesWithLinesets(); break;
            case 'harnesses':    this.renderComponents('harnesses');    break;
            case 'locations':    this.renderLocations();               break;
        }
    }

    updateEquipmentOptions() {
        const select = document.getElementById('equipment');
        select.innerHTML = '<option value="">Select Canopy</option>';
        
        // Show active canopies that still have at least one active lineset
        const activeCanopies = this.canopies.filter(c =>
            !c.archived && this.getActiveLineset(c.id)
        );
        
        activeCanopies.forEach(canopy => {
            const ls = this.getActiveLineset(canopy.id);
            if (!ls) return;
            const hybridTag = ls?.hybrid ? ' (Hybrid)' : '';
            const option = document.createElement('option');
            option.value = canopy.id;
            option.textContent = `${canopy.name} — Lineset #${ls.number}${hybridTag}`;
            select.appendChild(option);
        });
    }

    getActiveLineset(canopyId) {
        const canopy = this.canopies.find(c => c.id === canopyId);
        if (!canopy || !Array.isArray(canopy.linesets)) return null;

        const active = canopy.linesets.filter(ls => !ls.archived);
        if (active.length === 0) return null;

        const activeNumber = Math.max(...active.map(ls => ls.number));
        return active.find(ls => ls.number === activeNumber) || null;
    }

    /**
     * Get the highest non-archived lineset number for a canopy.
     * Returns the highest active lineset number, or 1 if none exist.
     */
    getActiveLinesetNumber(canopyId) {
        return this.getActiveLineset(canopyId)?.number || 1;
    }

    /**
     * Update the lineset hint below the canopy selector in the jump form.
     */
    updateLinesetHint() {
        const hint = document.getElementById('linesetHint');
        const canopyId = document.getElementById('equipment').value;
        if (!hint) return;
        if (!canopyId) {
            hint.style.display = 'none';
            return;
        }
        const ls = this.getActiveLineset(canopyId);
        if (!ls) {
            hint.style.display = 'none';
            return;
        }

        const lsNum = ls.number;
        const hybridTag = ls?.hybrid ? ' (Hybrid)' : '';
        const total = (ls?.jumpCount || 0) + (ls?.previousJumps || 0);
        hint.textContent = `→ Lineset #${lsNum}${hybridTag} · ${total} total jumps`;
        hint.style.display = 'block';
    }

    /**
     * Render the canopies list with embedded lineset information.
     */
    renderCanopiesWithLinesets() {
        const container = document.getElementById('equipmentList');
        
        if (this.canopies.length === 0) {
            container.innerHTML = '<p class="no-items">No canopies added yet.</p>';
            return;
        }

        const sorted = [...this.canopies].sort((a, b) => !!a.archived - !!b.archived);

        container.innerHTML = sorted.map(canopy => {
            const allLinesets = (canopy.linesets || []).sort((a, b) => b.number - a.number);
            const activeLinesets = allLinesets.filter(ls => !ls.archived);
            const archivedLinesets = allLinesets.filter(ls => ls.archived);
            const hasArchived = archivedLinesets.length > 0;

            const renderLinesetRow = (ls) => {
                const logged = ls.jumpCount || 0;
                const preApp = ls.previousJumps || 0;
                const total = logged + preApp;
                const hybridBadge = ls.hybrid ? '<span class="hybrid-badge">Hybrid</span>' : '';
                const archivedBadge = ls.archived ? '<span class="archived-badge">Archived</span>' : '';
                return `
                    <div class="lineset-row ${ls.archived ? 'archived' : ''}">
                        <span class="lineset-info">
                            Lineset #${ls.number} ${hybridBadge} ${archivedBadge}
                            <span class="lineset-jumps">${total} jumps${preApp > 0 ? ` (${logged} logged + ${preApp} pre-app)` : ''}</span>
                        </span>
                        <span class="lineset-actions">
                            <button onclick="window.logbook.editLineset('${canopy.id}', ${ls.number})" class="btn-edit btn-sm">Edit</button>
                            <button onclick="window.logbook.toggleArchiveLineset('${canopy.id}', ${ls.number})" class="btn-toggle btn-sm">
                                ${ls.archived ? 'Unarchive' : 'Archive'}
                            </button>
                        </span>
                    </div>
                `;
            };

            const activeLinesetsHtml = activeLinesets.map(renderLinesetRow).join('');
            const archivedLinesetsHtml = archivedLinesets.map(renderLinesetRow).join('');
            const archivedId = 'archived-linesets-' + canopy.id.replace(/[^a-zA-Z0-9_-]/g, '_');

            const draggable = !canopy.archived;
            return `
                <div class="equipment-item ${canopy.archived ? 'archived' : ''}" data-canopy-id="${canopy.id}" ${draggable ? 'draggable="true"' : ''}>
                    <div class="equipment-info">
                        <div class="equipment-name-row">
                            ${draggable ? '<span class="drag-handle" title="Drag to reorder">⠿</span>' : ''}
                            <span class="equipment-name">${canopy.name}</span>
                            ${canopy.notes ? `<span class="component-notes-inline">\uD83D\uDCDD ${canopy.notes}</span>` : ''}
                        </div>
                        ${canopy.archived ? '<span class="archived-badge">Archived</span>' : ''}
                        <div class="linesets-container">
                            ${activeLinesetsHtml || '<p class="no-items" style="margin:4px 0;">No active linesets</p>'}
                            ${hasArchived ? `
                                <label class="show-archived-linesets-label">
                                    <input type="checkbox" onchange="window.logbook.toggleArchivedLinesets('${archivedId}', this.checked)">
                                    Show archived linesets (${archivedLinesets.length})
                                </label>
                                <div id="${archivedId}" class="archived-linesets-group" style="display:none;">
                                    ${archivedLinesetsHtml}
                                </div>
                            ` : ''}
                        </div>
                    </div>
                    <div class="equipment-actions">
                        <button onclick="window.logbook.editComponent('${canopy.id}', 'canopies')" class="btn-edit">Edit</button>
                        <button onclick="window.logbook.openAddLineset('${canopy.id}')" class="btn-secondary btn-sm">+ Lineset</button>
                        <button onclick="window.logbook.toggleArchiveComponent('${canopy.id}', 'canopies')" class="btn-toggle">
                            ${canopy.archived ? 'Unarchive' : 'Archive'}
                        </button>
                        <button onclick="window.logbook.deleteComponent('${canopy.id}', 'canopies')" class="btn-delete">Delete</button>
                    </div>
                </div>
            `;
        }).join('');

        this._initCanopyDragAndDrop(container);
    }

    toggleArchivedLinesets(id, show) {
        const el = document.getElementById(id);
        if (el) el.style.display = show ? 'block' : 'none';
    }

    _initCanopyDragAndDrop(container) {
        let dragSrcId = null;

        const clearOver = () =>
            container.querySelectorAll('.equipment-item').forEach(i => i.classList.remove('drag-over'));

        container.querySelectorAll('.equipment-item[draggable="true"]').forEach(item => {
            // ── Desktop HTML5 drag ──────────────────────────────────────
            item.addEventListener('dragstart', e => {
                dragSrcId = item.dataset.canopyId;
                e.dataTransfer.effectAllowed = 'move';
                setTimeout(() => item.classList.add('dragging'), 0);
            });

            item.addEventListener('dragend', () => {
                item.classList.remove('dragging');
                clearOver();
            });

            item.addEventListener('dragover', e => {
                e.preventDefault();
                e.dataTransfer.dropEffect = 'move';
                clearOver();
                item.classList.add('drag-over');
            });

            item.addEventListener('dragleave', () => item.classList.remove('drag-over'));

            item.addEventListener('drop', e => {
                e.preventDefault();
                clearOver();
                const targetId = item.dataset.canopyId;
                if (dragSrcId && dragSrcId !== targetId) {
                    this._reorderCanopy(dragSrcId, targetId);
                }
            });

            // ── Mobile touch drag (via handle) ──────────────────────────
            const handle = item.querySelector('.drag-handle');
            if (!handle) return;

            handle.addEventListener('touchstart', e => {
                dragSrcId = item.dataset.canopyId;
                item.classList.add('dragging');
                e.preventDefault();
            }, { passive: false });

            handle.addEventListener('touchmove', e => {
                e.preventDefault();
                const touch = e.touches[0];
                const hit = document.elementFromPoint(touch.clientX, touch.clientY);
                clearOver();
                const target = hit && hit.closest('.equipment-item[draggable="true"]');
                if (target && target !== item) target.classList.add('drag-over');
            }, { passive: false });

            handle.addEventListener('touchend', () => {
                item.classList.remove('dragging');
                const overEl = container.querySelector('.equipment-item.drag-over');
                clearOver();
                if (overEl && dragSrcId) {
                    const targetId = overEl.dataset.canopyId;
                    if (dragSrcId !== targetId) this._reorderCanopy(dragSrcId, targetId);
                }
            });
        });
    }

    _reorderCanopy(srcId, targetId) {
        const srcIdx = this.canopies.findIndex(c => c.id === srcId);
        const tgtIdx = this.canopies.findIndex(c => c.id === targetId);
        if (srcIdx === -1 || tgtIdx === -1) return;
        const [moved] = this.canopies.splice(srcIdx, 1);
        this.canopies.splice(tgtIdx, 0, moved);
        this.canopies.forEach((c, i) => { c.sortOrder = i; });
        this.saveComponentsToLocalStorage();
        this.renderEquipmentView();
        this.updateEquipmentOptions();
    }

    // ── Lineset management ──────────────────────────────────────────────────

    openAddLineset(canopyId) {
        const canopy = this.canopies.find(c => c.id === canopyId);
        if (!canopy) return;
        
        document.getElementById('linesetForm').reset();
        document.getElementById('linesetCanopyId').value = canopyId;
        document.getElementById('linesetEditNumber').value = '';
        document.getElementById('linesetCanopyName').value = canopy.name;
        document.getElementById('linesetModalTitle').textContent = 'Add Lineset';
        document.getElementById('linesetPreviousJumps').value = 0;
        document.getElementById('linesetHybridCheck').checked = false;
        
        // Auto-fill next lineset number
        const maxNum = (canopy.linesets || []).length > 0
            ? Math.max(...canopy.linesets.map(ls => ls.number))
            : 0;
        document.getElementById('linesetNumber').value = maxNum + 1;
        
        document.getElementById('linesetModal').style.display = 'block';
    }

    editLineset(canopyId, linesetNumber) {
        const canopy = this.canopies.find(c => c.id === canopyId);
        if (!canopy) return;
        const lineset = canopy.linesets?.find(ls => ls.number === linesetNumber);
        if (!lineset) return;
        
        document.getElementById('linesetCanopyId').value = canopyId;
        document.getElementById('linesetEditNumber').value = linesetNumber;
        document.getElementById('linesetCanopyName').value = canopy.name;
        document.getElementById('linesetModalTitle').textContent = `Edit Lineset #${linesetNumber}`;
        document.getElementById('linesetNumber').value = linesetNumber;
        document.getElementById('linesetHybridCheck').checked = lineset.hybrid || false;
        document.getElementById('linesetPreviousJumps').value = lineset.previousJumps || 0;
        
        document.getElementById('linesetModal').style.display = 'block';
    }

    saveLineset() {
        const canopyId = document.getElementById('linesetCanopyId').value;
        const editNumber = document.getElementById('linesetEditNumber').value;
        const linesetNumber = parseInt(document.getElementById('linesetNumber').value) || 1;
        const hybrid = document.getElementById('linesetHybridCheck').checked;
        const previousJumps = Math.max(0, parseInt(document.getElementById('linesetPreviousJumps').value) || 0);
        
        const canopy = this.canopies.find(c => c.id === canopyId);
        if (!canopy) return;
        if (!Array.isArray(canopy.linesets)) canopy.linesets = [];
        
        if (editNumber) {
            // Edit existing lineset
            const lineset = canopy.linesets.find(ls => ls.number === parseInt(editNumber));
            if (lineset) {
                lineset.hybrid = hybrid;
                lineset.previousJumps = previousJumps;
            }
        } else {
            // Add new lineset
            const existing = canopy.linesets.find(ls => ls.number === linesetNumber);
            if (existing) {
                this.showMessage(`Lineset #${linesetNumber} already exists for this canopy`, 'error');
                return;
            }
            // Archive all existing non-archived linesets
            canopy.linesets.forEach(ls => { if (!ls.archived) ls.archived = true; });
            canopy.linesets.push({
                number: linesetNumber,
                hybrid: hybrid,
                previousJumps: previousJumps,
                jumpCount: 0,
                archived: false
            });
            canopy.linesets.sort((a, b) => a.number - b.number);
        }
        
        this.saveComponentsToLocalStorage();
        this.updateEquipmentOptions();
        this.renderEquipmentView();
        this.closeLinesetModal();
        this.showMessage('Lineset saved successfully!', 'success');
    }

    toggleArchiveLineset(canopyId, linesetNumber) {
        const canopy = this.canopies.find(c => c.id === canopyId);
        if (!canopy) return;
        const lineset = canopy.linesets?.find(ls => ls.number === linesetNumber);
        if (!lineset) return;
        
        lineset.archived = !lineset.archived;
        this.saveComponentsToLocalStorage();
        this.updateEquipmentOptions();
        this.renderEquipmentView();
        this.showMessage(`Lineset #${linesetNumber} ${lineset.archived ? 'archived' : 'unarchived'} successfully!`, 'success');
    }

    closeLinesetModal() {
        document.getElementById('linesetModal').style.display = 'none';
    }

    _singularize(plural) {
        const map = { harnesses: 'harness', canopies: 'canopy', locations: 'location' };
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
        // Show/hide initial lineset section for new canopies
        const isCanopy = type === 'canopy';
        document.getElementById('canopyLinesetSection').style.display = isCanopy ? 'block' : 'none';
        if (isCanopy) {
            document.getElementById('newCanopyHybridCheck').checked = false;
            document.getElementById('newCanopyPreviousJumps').value = 0;
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
        
        // Give new canopies a lineset #1 using the values from the creation form
        if (type === 'canopy' && !id) {
            const canopy = collection[collection.length - 1];
            if (!Array.isArray(canopy.linesets)) {
                const hybrid = document.getElementById('newCanopyHybridCheck').checked;
                const previousJumps = Math.max(0, parseInt(document.getElementById('newCanopyPreviousJumps').value) || 0);
                canopy.linesets = [{ number: 1, hybrid: hybrid, previousJumps: previousJumps, jumpCount: 0, archived: false }];
            }
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
            const matches = this.locations
                .filter(loc => !loc.archived && loc.name.toLowerCase().includes(query))
                .sort((a, b) => (a.sortOrder ?? Infinity) - (b.sortOrder ?? Infinity));

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

    renderLocations() {
        const container = document.getElementById('equipmentList');

        if (this.locations.length === 0) {
            container.innerHTML = '<p class="no-items">No locations added yet.</p>';
            return;
        }

        // Active locations in user-defined order, archived appended at the end
        const active   = this.locations.filter(l => !l.archived);
        const archived = this.locations.filter(l =>  l.archived);
        const sorted   = [...active, ...archived];

        container.innerHTML = sorted.map(loc => {
            const draggable = !loc.archived;
            return `
                <div class="equipment-item ${loc.archived ? 'archived' : ''}" data-location-id="${loc.id}" ${draggable ? 'draggable="true"' : ''}>
                    <div class="equipment-info">
                        <div class="equipment-name-row">
                            ${draggable ? '<span class="drag-handle" title="Drag to reorder">&#x283f;</span>' : ''}
                            <span class="equipment-name">${loc.name}</span>
                        </div>
                        ${loc.notes ? `<div class="component-notes">\uD83D\uDCDD ${loc.notes}</div>` : ''}
                        ${loc.archived ? '<span class="archived-badge">Archived</span>' : ''}
                    </div>
                    <div class="equipment-actions">
                        <button onclick="window.logbook.editComponent('${loc.id}', 'locations')" class="btn-edit">Edit</button>
                        <button onclick="window.logbook.toggleArchiveComponent('${loc.id}', 'locations')" class="btn-toggle">
                            ${loc.archived ? 'Unarchive' : 'Archive'}
                        </button>
                        <button onclick="window.logbook.deleteComponent('${loc.id}', 'locations')" class="btn-delete">Delete</button>
                    </div>
                </div>
            `;
        }).join('');

        this._initLocationDragAndDrop(container);
    }

    _initLocationDragAndDrop(container) {
        let dragSrcId = null;

        const clearOver = () =>
            container.querySelectorAll('.equipment-item').forEach(i => i.classList.remove('drag-over'));

        container.querySelectorAll('.equipment-item[draggable="true"]').forEach(item => {
            // ── Desktop HTML5 drag ──────────────────────────────────────
            item.addEventListener('dragstart', e => {
                dragSrcId = item.dataset.locationId;
                e.dataTransfer.effectAllowed = 'move';
                setTimeout(() => item.classList.add('dragging'), 0);
            });

            item.addEventListener('dragend', () => {
                item.classList.remove('dragging');
                clearOver();
            });

            item.addEventListener('dragover', e => {
                e.preventDefault();
                e.dataTransfer.dropEffect = 'move';
                clearOver();
                item.classList.add('drag-over');
            });

            item.addEventListener('dragleave', () => item.classList.remove('drag-over'));

            item.addEventListener('drop', e => {
                e.preventDefault();
                clearOver();
                const targetId = item.dataset.locationId;
                if (dragSrcId && dragSrcId !== targetId) {
                    this._reorderLocation(dragSrcId, targetId);
                }
            });

            // ── Mobile touch drag (via handle) ──────────────────────────
            const handle = item.querySelector('.drag-handle');
            if (!handle) return;

            handle.addEventListener('touchstart', e => {
                dragSrcId = item.dataset.locationId;
                item.classList.add('dragging');
                e.preventDefault();
            }, { passive: false });

            handle.addEventListener('touchmove', e => {
                e.preventDefault();
                const touch = e.touches[0];
                const hit = document.elementFromPoint(touch.clientX, touch.clientY);
                clearOver();
                const target = hit && hit.closest('.equipment-item[draggable="true"]');
                if (target && target !== item) target.classList.add('drag-over');
            }, { passive: false });

            handle.addEventListener('touchend', () => {
                item.classList.remove('dragging');
                const overEl = container.querySelector('.equipment-item.drag-over');
                clearOver();
                if (overEl && dragSrcId) {
                    const targetId = overEl.dataset.locationId;
                    if (dragSrcId !== targetId) this._reorderLocation(dragSrcId, targetId);
                }
            });
        });
    }

    _reorderLocation(srcId, targetId) {
        // Only reorder within active (non-archived) locations
        const active   = this.locations.filter(l => !l.archived);
        const archived = this.locations.filter(l =>  l.archived);
        const srcIdx   = active.findIndex(l => l.id === srcId);
        const tgtIdx   = active.findIndex(l => l.id === targetId);
        if (srcIdx === -1 || tgtIdx === -1) return;
        const [moved] = active.splice(srcIdx, 1);
        active.splice(tgtIdx, 0, moved);
        active.forEach((l, i) => { l.sortOrder = i; });
        this.locations = [...active, ...archived];
        this.saveComponentsToLocalStorage();
        this.renderLocations();
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
            // Hide initial lineset section when editing (only shown for new canopies)
            document.getElementById('canopyLinesetSection').style.display = 'none';
            document.getElementById('componentModal').style.display = 'block';
        }
    }
    
    deleteComponent(id, type) {
        const typeSingular = this._singularize(type);
        if (confirm(`Are you sure you want to delete this ${typeSingular}?`)) {
            // Check if canopy is used in any jumps
            if (type === 'canopies') {
                const usedInJumps = this.jumps.some(j => j.equipment === id);
                if (usedInJumps) {
                    this.showMessage(`Cannot delete ${typeSingular} that has been used in jumps. Archive it instead.`, 'error');
                    return;
                }
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

    renderStats() {
        const container = document.getElementById('statsContent');
        
        if (this.jumps.length === 0) {
            container.innerHTML = '<p class="no-items">No jumps logged yet.</p>';
            return;
        }
        
        // Build canopy/lineset stats (replaces old rig stats)
        const linesetStats = [];
        this.canopies.forEach(canopy => {
            (canopy.linesets || []).forEach(ls => {
                const logged = this.jumps.filter(j => j.equipment === canopy.id && j.linesetNumber === ls.number).length;
                const preApp = ls.previousJumps || 0;
                const total = logged + preApp;
                const hybridSuffix = ls.hybrid ? ' (Hybrid)' : '';
                linesetStats.push({
                    name: `${canopy.name} — Lineset #${ls.number}${hybridSuffix}`,
                    count: total,
                    logged,
                    preApp,
                    archived: canopy.archived || ls.archived,
                    hybrid: ls.hybrid || false
                });
            });
        });
        
        const activeStats = linesetStats.filter(s => !s.archived && s.count > 0);
        const archivedStats = linesetStats.filter(s => s.archived);
        const sortedStats = this.showArchivedStats ? [...activeStats, ...archivedStats] : activeStats;
        
        const hasArchived = archivedStats.length > 0;
        const archivedBtnLabel = this.showArchivedStats ? 'Hide Archived' : `Show Archived (${archivedStats.length})`;
        const archivedToggleBtn = hasArchived
            ? `<button class="btn-secondary btn-sm" onclick="window.logbook.toggleArchivedStats()">${archivedBtnLabel}</button>`
            : '';

        let html = `
            <div class="stats-section">
                <div class="stats-section-header">
                    <h3>Canopy / Lineset</h3>
                    ${archivedToggleBtn}
                </div>
                <div class="stats-list">
        `;
        
        if (sortedStats.length > 0) {
            sortedStats.forEach(stat => {
                const redThreshold = stat.hybrid ? this.settings.hybridRedThreshold : this.settings.standardRedThreshold;
                const orangeThreshold = stat.hybrid ? this.settings.hybridOrangeThreshold : this.settings.standardOrangeThreshold;
                const percentage = Math.min((stat.count / redThreshold) * 100, 100);
                let barColorClass = '';
                if (stat.count >= redThreshold) barColorClass = 'stat-fill-red';
                else if (stat.count >= orangeThreshold) barColorClass = 'stat-fill-orange';
                const breakdown = stat.preApp > 0
                    ? `${stat.count} total (${stat.logged} logged + ${stat.preApp} pre-app)`
                    : `${stat.count} jumps`;
                html += `
                    <div class="stat-item${stat.archived ? ' archived' : ''}">
                        <div class="stat-info stat-info-stacked">
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
            html += '<p class="no-items">No canopy/lineset statistics available.</p>';
        }
        
        html += '</div></div>';
        
        // Add canopy aggregate statistics (preserve canopy order)
        const canopyTotalsArray = this.canopies.map(canopy => {
            const logged = this.jumps.filter(j => j.equipment === canopy.id).length;
            const preApp = (canopy.linesets || []).reduce((sum, ls) => sum + (ls.previousJumps || 0), 0);
            return { name: canopy.name, count: logged + preApp, archived: !!canopy.archived };
        }).filter(s => s.count > 0);
        html += this.renderOrderedComponentStats('Canopy Totals', canopyTotalsArray);
        
        container.innerHTML = html;
    }
    
    toggleArchivedStats() {
        this.showArchivedStats = !this.showArchivedStats;
        this.renderStats();
    }

    renderComponentStats(title, statsObject, useStackedLayout = false) {
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
                        <div class="stat-info${useStackedLayout ? ' stat-info-stacked' : ''}">
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

    // Like renderComponentStats but accepts a pre-ordered array { name, count }
    // so the display order is controlled by the caller (not sorted by count).
    renderOrderedComponentStats(title, statsArray) {
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
                    <div class="stat-item${stat.archived ? ' archived' : ''}">
                        <div class="stat-info">
                            <span class="stat-name">${stat.name}${stat.archived ? ' (Archived)' : ''}</span>
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
            || this.harnesses.length > 0
            || this.canopies.length > 0;
    }

    buildExportPayload() {
        return {
            exportedAt: new Date().toISOString(),
            version: 2,
            data: {
                jumps: this.jumps,
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
                || Array.isArray(payload.harnesses)
                || Array.isArray(payload.canopies)
                || Array.isArray(payload.locations)
                || (payload.settings && typeof payload.settings === 'object');

            if (!hasAnySupportedData) {
                this.showMessage('Import file has no supported data', 'error');
                return;
            }

            const importJumps = Array.isArray(payload.jumps) ? payload.jumps : [];
            const numH = Array.isArray(payload.harnesses) ? payload.harnesses.length : 0;
            const numC = Array.isArray(payload.canopies) ? payload.canopies.length : 0;
            const numL = Array.isArray(payload.locations) ? payload.locations.length : 0;
            const msg = document.getElementById('importChoiceModalMessage');
            if (msg) {
                msg.textContent = `The import file contains ${importJumps.length} jump(s), ${numH} harness(es), ${numC} canopy/canopies, and ${numL} location(s). Choose how to import.`;
            }
            this._pendingImportPayload = payload;
            this.showImportChoiceModal();
        } catch (error) {
            console.error('Import failed:', error);
            this.showMessage('Import failed: invalid JSON file', 'error');
        } finally {
            event.target.value = '';
        }
    }

    showImportChoiceModal() {
        const modal = document.getElementById('importChoiceModal');
        if (modal) modal.style.display = 'block';
    }

    closeImportChoiceModal() {
        const modal = document.getElementById('importChoiceModal');
        if (modal) modal.style.display = 'none';
        this._pendingImportPayload = null;
    }

    applyImportChoice(mode) {
        const payload = this._pendingImportPayload;
        this.closeImportChoiceModal();
        if (!payload) return;
        if (mode === 'merge') {
            this.applyImportMerge(payload);
        } else {
            this.applyImportReplace(payload);
        }
    }

    /** Merge import: keep local jumps and equipment; add/update from import (by jumpId / id). No deletions. */
    applyImportMerge(payload) {
        const importJumps = Array.isArray(payload.jumps) ? payload.jumps : [];
        const importHarnesses = Array.isArray(payload.harnesses) ? payload.harnesses : [];
        const importCanopies = Array.isArray(payload.canopies) ? payload.canopies : [];
        const importLocations = Array.isArray(payload.locations) ? payload.locations : [];

        const importJumpIds = new Set(importJumps.map(j => (j.jumpId || j.id || '').toString()).filter(Boolean));
        const localOnlyJumps = this.jumps.filter(j => {
            const id = (j.jumpId || j.id || '').toString();
            return id && !importJumpIds.has(id);
        });
        const importJumpIdsSeen = new Set();
        const mergedJumps = [...localOnlyJumps];
        for (const j of importJumps) {
            const id = (j.jumpId || j.id || '').toString();
            if (id && !importJumpIdsSeen.has(id)) {
                importJumpIdsSeen.add(id);
                mergedJumps.push(j);
            } else if (!id) {
                mergedJumps.push(j);
            }
        }
        this.jumps = mergedJumps;

        const byId = (arr, idKey) => new Map((arr || []).map(x => [x[idKey] || x.id, x]));
        const localH = byId(this.harnesses, 'id');
        importHarnesses.forEach(h => { if (h.id) localH.set(h.id, h); });
        this.harnesses = Array.from(localH.values());

        const localC = byId(this.canopies, 'id');
        importCanopies.forEach(c => { if (c.id) localC.set(c.id, c); });
        this.canopies = Array.from(localC.values());

        const localL = byId(this.locations, 'id');
        importLocations.forEach(l => { if (l && (l.id || l.name)) localL.set(l.id || l.name, l); });
        this.locations = Array.from(localL.values());

        if (payload.settings && typeof payload.settings === 'object') {
            this.settings = { ...this.settings, ...payload.settings };
        }
        if (this.settings.recentJumpsDays === undefined) this.settings.recentJumpsDays = 16;

        this.canopies.forEach(canopy => {
            if (!Array.isArray(canopy.linesets)) canopy.linesets = [];
            if (canopy.linesets.length === 0) {
                canopy.linesets.push({ number: 1, hybrid: false, previousJumps: 0, jumpCount: 0, archived: false });
            }
        });
        this.ensureJumpIds();
        this.initializeCanopyLinesetJumpCounts();
        this.saveToLocalStorage();
        this.saveComponentsToLocalStorage();
        localStorage.setItem('skydiving-settings', JSON.stringify(this.settings));
        this.markEquipmentModified();
        this.updateEquipmentOptions();
        this.renderJumpsList();
        this.updateStats();
        this.renderEquipmentView();
        this.renderStats();
        this.showMessage('Data merged successfully!', 'success');
    }

    /** Replace all: replace jumps and equipment with import file; merge settings. Local-only data is removed. */
    applyImportReplace(payload) {
        this.jumps = Array.isArray(payload.jumps) ? payload.jumps : [];
        this.harnesses = Array.isArray(payload.harnesses) ? payload.harnesses : [];
        this.canopies = Array.isArray(payload.canopies) ? payload.canopies : [];
        this.locations = Array.isArray(payload.locations) ? payload.locations : [];

        if (payload.settings && typeof payload.settings === 'object') {
            this.settings = { ...this.settings, ...payload.settings };
        }
        if (this.settings.recentJumpsDays === undefined) this.settings.recentJumpsDays = 16;

        this.canopies.forEach(canopy => {
            if (!Array.isArray(canopy.linesets)) canopy.linesets = [];
            if (canopy.linesets.length === 0) {
                canopy.linesets.push({ number: 1, hybrid: false, previousJumps: 0, jumpCount: 0, archived: false });
            }
        });
        this.ensureJumpIds();
        this.initializeCanopyLinesetJumpCounts();
        this.saveToLocalStorage();
        this.saveComponentsToLocalStorage();
        localStorage.setItem('skydiving-settings', JSON.stringify(this.settings));
        this.markEquipmentModified();
        this.updateEquipmentOptions();
        this.renderJumpsList();
        this.updateStats();
        this.renderEquipmentView();
        this.renderStats();
        this.showMessage('Data imported successfully! (Replace all)', 'success');
    }

    updateOnlineStatus() {
        const syncStatus = document.getElementById('syncStatus');
        if (navigator.onLine) {
            syncStatus.textContent = 'Unsynced';
            syncStatus.className = 'sync-status warning';
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

    /**
     * Show banner when storage (IndexedDB) is blocked (e.g. Safari/iOS).
     * The "Enable storage" button requests access via Storage Access API (user gesture)
     * then reopens the DB and reloads the page.
     */
    showStorageBlockedBanner() {
        const banner = document.getElementById('storageAccessBanner');
        const btn = document.getElementById('storageAccessBtn');
        if (!banner || !btn) return;
        banner.style.display = 'flex';
        const once = async () => {
            btn.disabled = true;
            try {
                await DB.requestStorageAccess();
                await DB.open();
                await DB.migrateFromLocalStorage();
                window.location.reload();
            } catch (e) {
                console.error('[DB] Storage access request failed:', e);
                this.showMessage('Could not enable storage. Try allowing storage for this site in Safari settings.', 'error');
                btn.disabled = false;
            }
        };
        btn.replaceWith(btn.cloneNode(true));
        document.getElementById('storageAccessBtn').addEventListener('click', once);
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

    showSyncConflictModal() {
        this.showMessage(
            'Sync conflict: local changes were merged with data from Google Sheets.',
            'info'
        );
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
