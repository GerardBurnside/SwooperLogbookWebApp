// Swooper Logbook App - Main Application Logic

/** Optional per-lineset overrides for statistics bar colors (normal linesets only; hybrid uses settings). */
const LINESET_STAT_THRESHOLD_PROPS = ['standardOrangeThreshold', 'standardRedThreshold'];
const LINESET_MODAL_STAT_THRESHOLD_INPUT_IDS = ['linesetStatStandardOrange', 'linesetStatStandardRed'];
const NEW_CANOPY_STAT_THRESHOLD_INPUT_IDS = ['newCanopyStatStandardOrange', 'newCanopyStatStandardRed'];

/** Header tabs users can show or hide from Settings. */
const MAIN_NAV_VIEWS = [
    { id: 'jumps', label: 'Jumps' },
    { id: 'equipment', label: 'Equipment' },
    { id: 'stats', label: 'Statistics' },
    { id: 'flysight', label: 'Flysight' },
    { id: 'todos', label: 'TODO' }
];

class SkydivingLogbook {
    constructor() {
        // Data arrays — populated asynchronously from IndexedDB in init()
        this.jumps = [];
        this.harnesses = [];
        this.canopies = [];
        this.locations = [];

        this.settings = JSON.parse(localStorage.getItem('skydiving-settings')) || {
            startingJumpNumber: 1,
            /** When true, add/edit/delete renumbers all jumps from startingJumpNumber in date order. When false, stored jump #s are kept; new jumps use max+1. */
            resequenceJumpsFromStartingNumber: true,
            recentJumpsDays: 7,
            recentJumpsGroupByMonth: false,
            standardRedThreshold: 160,
            standardOrangeThreshold: 140,
            hybridRedThreshold: 80,
            hybridOrangeThreshold: 60,
            autoDetectDropZone: true,
            statsPastMonthsWindow: 3
        };
        // Backfill for existing saved settings that predate these fields
        if (this.settings.recentJumpsDays === undefined) {
            this.settings.recentJumpsDays = 7;
        }
        if (this.settings.recentJumpsGroupByMonth === undefined) {
            this.settings.recentJumpsGroupByMonth = false;
        }
        if (this.settings.statsPastMonthsWindow === undefined) {
            this.settings.statsPastMonthsWindow = 3;
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
        if (this.settings.autoDetectDropZone === undefined) {
            this.settings.autoDetectDropZone = true;
        }
        if (this.settings.resequenceJumpsFromStartingNumber === undefined) {
            this.settings.resequenceJumpsFromStartingNumber = true;
        }
        this.normalizeNavSettings();

        try {
            this.todos = this.loadTodos();
        } catch (_) {
            this.todos = [];
        }
        try {
            this.deletedTodos = this.loadDeletedTodos();
        } catch (_) {
            this.deletedTodos = [];
        }
        this._applyingTodoSync = false;
        this._todoLongPressTimer = null;
        this._todoLongPressId = null;
        this._todoSuppressClick = false;
        this._editingTodoId = null;

        this.currentView = 'jumps'; // 'jumps', 'equipment', 'stats', 'flysight', 'todos'
        this.equipmentSubView = 'canopies'; // 'canopies', 'harnesses', 'locations'
        this.showArchivedStats = false;
        /** Statistics view: show archived canopies in the Canopy Totals block only. */
        this.showArchivedCanopyTotals = false;
        /** Statistics view: show archived harnesses in the Harness block only. */
        this.showArchivedHarnessStats = false;
        /** Equipment view: canopy id → whether "show older linesets" is checked. */
        this.showOlderLinesetsByCanopyId = Object.create(null);
        this.activeJumpNoteId = null;
        this.activeEditJumpId = null;
        /** Trimmed location when edit jump modal was opened (for bulk same-day option). */
        this.editJumpLocationAtOpen = null;
        /** `YYYY-MM-DD` when edit jump modal was opened (for bulk date-shift option). */
        this.editJumpDateAtOpen = null;
        this._olderJumpsCache = []; // cached older jumps for lazy rendering
        this._renderedOlderCount = 0;
        this._mergedJumpsCache = []; // recent + older when "by month" merges both
        this._renderedMergedCount = 0;
        this._useMergedListCache = false;
        this._monthLocationPieGroups = new Map();
        this._dayLocationPieGroups = new Map();
        this.flysightFiles = [];
        this._flysightBrowser = null;
        this._flysightGraph = null;
        this._flysightGraphRo = null;
        this._flysightScrollY = 0;
        const savedFlysightAvg = parseInt(localStorage.getItem('flysight-avg-points'), 10);
        this.flysightAvgPoints = Number.isFinite(savedFlysightAvg) ? Math.min(20, Math.max(1, savedFlysightAvg)) : 3;
        this.flysightMaxHeightSliderMaxM = this._parseFlysightMaxHeightSliderMax(
            localStorage.getItem('flysight-max-height-slider-max')
        );
        const savedFlysightMaxHeight = parseInt(localStorage.getItem('flysight-max-height'), 10);
        this.flysightMaxHeightM = Number.isFinite(savedFlysightMaxHeight)
            ? this._parseFlysightMaxHeight(savedFlysightMaxHeight)
            : Math.min(this.flysightMaxHeightSliderMaxM, this._flysightHeightLimits().defaultM);
        const savedFlysightSpeedMetric = localStorage.getItem('flysight-speed-metric');
        this.flysightSpeedMetric = ['vertical', 'total', 'both'].includes(savedFlysightSpeedMetric)
            ? savedFlysightSpeedMetric
            : 'vertical';
        const defaultCursorBAngle = (typeof Flysight !== 'undefined' && Number.isFinite(Flysight.CURSOR_B_DIVE_ANGLE_DEG))
            ? Flysight.CURSOR_B_DIVE_ANGLE_DEG
            : 5.5;
        const savedCursorBAngle = parseFloat(localStorage.getItem('flysight-cursor-b-dive-angle'));
        this.flysightCursorBDiveAngleDeg = Number.isFinite(savedCursorBAngle)
            ? Math.min(90, Math.max(0, savedCursorBAngle))
            : defaultCursorBAngle;
        const defaultCursorBTicks = (typeof Flysight !== 'undefined' && Number.isFinite(Flysight.CURSOR_B_ALT_TICKS))
            ? Flysight.CURSOR_B_ALT_TICKS
            : 2;
        const savedCursorBTicks = parseInt(localStorage.getItem('flysight-cursor-b-alt-ticks'), 10);
        this.flysightCursorBAltTicks = Number.isFinite(savedCursorBTicks)
            ? Math.min(100, Math.max(0, savedCursorBTicks))
            : defaultCursorBTicks;
        
        this.init();
    }

    async init() {
        // Hide deselected tabs before IndexedDB work so they never flash on startup
        this.applyNavVisibility();

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
            if (!Number.isFinite(Number(canopy.previousJumps))) canopy.previousJumps = 0;
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
        this.setupNotesAutocomplete();
        this.preFillFormWithLastJump();
        this.applyAutoDetectDropZoneUi(true);
        this.applyNavVisibility();
        this.showView(this.settings.startView);
        
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
                location: (form.elements['location'].value || '').trim(),
                equipment: form.elements['equipment'].value,
                notes: form.elements['notes'].value || ''
            };
            if (!jumpData.location) {
                this.showMessage('Please enter a location.', 'error');
                return;
            }
            if (!jumpData.equipment) {
                this.showMessage('Please select a canopy.', 'error');
                return;
            }
            for (let i = 0; i < multiplier; i++) {
                this.addJump(jumpData, multiplier > 1);
            }
            // Reset multiplier back to 1
            document.getElementById('jumpMultiplier').value = 1;
            // Reset form after all jumps logged
            form.reset();
            this.setCurrentDate();
            this.preFillFormWithLastJump();
            this.applyAutoDetectDropZoneUi(true);
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

        // const repairJumpIdsBtn = document.getElementById('repairJumpLocalIdsBtn');
        // if (repairJumpIdsBtn) {
        //     repairJumpIdsBtn.addEventListener('click', () => this.repairJumpLocalIdsFromSettings());
        // }

        const reseqChk = document.getElementById('settingsResequenceJumpsCheckbox');
        if (reseqChk) {
            reseqChk.addEventListener('change', () => this._updateStartingJumpUiState());
        }

        document.getElementById('navVisibilityOptions')?.addEventListener('change', (e) => {
            if (e.target.classList.contains('nav-visibility-chk')) {
                this._onNavVisibilityCheckboxChange(e.target);
            }
        });

        // document.getElementById('resetAppBtn').addEventListener('click', () => {
        //     this.resetAppToFirstLaunch();
        // });

        document.getElementById('useCurrentLocationBtn').addEventListener('click', () => {
            this.setComponentCoordsFromGPS();
        });

        // Google Sheets Integration modal (OAuth)
        document.getElementById('googleSheetsIntegrationBtn').addEventListener('click', () => {
            this.openSheetsModal();
        });

        // Reset DB: use capture on settings modal so taps work inside overflow-scroll (Android),
        // and keep logic off a lone button listener that can be flaky on some WebViews.
        const settingsModalEl = document.getElementById('settingsModal');
        if (settingsModalEl) {
            settingsModalEl.addEventListener(
                'click',
                (e) => {
                    if (e.target.closest('#resetDbBtn')) this.openResetDbConfirmModal();
                },
                true
            );
        }

        const resetDbConfirmModalEl = document.getElementById('resetDbConfirmModal');
        if (resetDbConfirmModalEl) {
            resetDbConfirmModalEl.addEventListener(
                'click',
                (e) => {
                    if (e.target.closest('#resetDbConfirmCancel')) this.closeResetDbConfirmModal();
                    else if (e.target.closest('#resetDbConfirmProceed')) this.confirmResetLocalDb();
                },
                true
            );
        }

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

        const editJumpClose = document.getElementById('editJumpClose');
        if (editJumpClose) {
            editJumpClose.addEventListener('click', () => this.closeEditJumpModal());
        }
        const editJumpCancel = document.getElementById('editJumpCancel');
        if (editJumpCancel) {
            editJumpCancel.addEventListener('click', () => this.closeEditJumpModal());
        }
        const editJumpSave = document.getElementById('editJumpSave');
        if (editJumpSave) {
            editJumpSave.addEventListener('click', () => this.saveEditedJump());
        }
        const editJumpDateEl = document.getElementById('editJumpDate');
        const editJumpLocEl = document.getElementById('editJumpLocation');
        const syncEditBulk = () => this.syncEditJumpModalBulkOptions();
        if (editJumpDateEl) {
            editJumpDateEl.addEventListener('input', syncEditBulk);
            editJumpDateEl.addEventListener('change', syncEditBulk);
        }
        if (editJumpLocEl) {
            editJumpLocEl.addEventListener('input', syncEditBulk);
            editJumpLocEl.addEventListener('change', syncEditBulk);
        }
        const editJumpShiftChk = document.getElementById('editJumpShiftFollowingChk');
        const editJumpShiftCnt = document.getElementById('editJumpShiftFollowingCount');
        if (editJumpShiftChk) editJumpShiftChk.addEventListener('change', syncEditBulk);
        if (editJumpShiftCnt) {
            editJumpShiftCnt.addEventListener('input', syncEditBulk);
            editJumpShiftCnt.addEventListener('change', syncEditBulk);
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
        document.getElementById('searchNotesCutawaysBtn').addEventListener('click', () => {
            this.executeNoteSearch({
                terms: ['cutaway', 'cut-away', 'libé', 'libe'],
                label: 'cutaways'
            });
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

        // Export data (download, or a download/share choice when sharing is available)
        document.getElementById('exportBtn').addEventListener('click', () => {
            this.handleExportClick();
        });
        document.getElementById('exportChoiceDownloadBtn')?.addEventListener('click', () => {
            this.closeExportChoiceModal();
            this.exportData();
        });
        document.getElementById('exportChoiceShareBtn')?.addEventListener('click', () => {
            this.closeExportChoiceModal();
            this.shareDataViaEmail();
        });
        document.getElementById('exportChoiceModalClose')?.addEventListener('click', () => {
            this.closeExportChoiceModal();
        });

        // Auto-detect DZ checkbox in the log jump form
        const autoDetectChk = document.getElementById('autoDetectDZForm');
        if (autoDetectChk) {
            autoDetectChk.addEventListener('change', () => {
                this.settings.autoDetectDropZone = autoDetectChk.checked;
                localStorage.setItem('skydiving-settings', JSON.stringify(this.settings));
                if (autoDetectChk.checked) this.detectNearestLocation(true);
            });
        }

        const recentByMonthTog = document.getElementById('recentJumpsGroupByMonthToggle');
        if (recentByMonthTog) {
            recentByMonthTog.addEventListener('change', () => {
                this.settings.recentJumpsGroupByMonth = recentByMonthTog.checked;
                const modalChk = document.getElementById('recentJumpsGroupByMonthSettings');
                if (modalChk) modalChk.checked = recentByMonthTog.checked;
                localStorage.setItem('skydiving-settings', JSON.stringify(this.settings));
                this.renderJumpsList();
            });
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

        document.getElementById('syncConflictOverwriteBtn')?.addEventListener('click', () => {
            this.applySyncConflictOverwrite();
        });
        document.getElementById('resolveConflictsBtn')?.addEventListener('click', () => {
            this.applySyncConflictMerge();
        });

        document.getElementById('importExternalCsvCancelBtn')?.addEventListener('click', () => {
            this._rejectExternalCsvEquipmentStep?.();
        });
        document.getElementById('importExternalCsvModalClose')?.addEventListener('click', () => {
            this._rejectExternalCsvEquipmentStep?.();
        });
        document.getElementById('importExternalCsvNextBtn')?.addEventListener('click', () => {
            this._resolveExternalCsvEquipmentStep?.();
        });

        // Navigation buttons
        document.getElementById('jumpsViewBtn')?.addEventListener('click', () => {
            this.showView('jumps');
        });

        document.getElementById('equipmentViewBtn')?.addEventListener('click', () => {
            this.showView('equipment');
        });

        document.getElementById('statsViewBtn')?.addEventListener('click', () => {
            this.showView('stats');
        });

        document.getElementById('flysightViewBtn')?.addEventListener('click', () => {
            this.showView('flysight');
        });

        document.getElementById('todosViewBtn')?.addEventListener('click', () => {
            this.showView('todos');
        });

        try {
            this._bindTodoEvents();
        } catch (err) {
            console.error('[TODOs] Failed to bind events:', err);
        }
        this._bindFlysightEvents();

        // Equipment management
        document.getElementById('addCanopyBtn').addEventListener('click', () => {
            this.addComponent('canopy');
        });

        document.getElementById('linesetForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.saveLineset();
        });
        document.getElementById('linesetHybridCheck').addEventListener('change', () => {
            this._syncLinesetModalStatThresholdSectionVisibility();
        });
        document.getElementById('newCanopyHybridCheck').addEventListener('change', () => {
            this._syncNewCanopyStatThresholdSectionVisibility();
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
            const exportChoiceModal = document.getElementById('exportChoiceModal');
            if (e.target === exportChoiceModal) {
                this.closeExportChoiceModal();
            }
            const conflictModal = document.getElementById('conflictModal');
            if (e.target === conflictModal && window.SheetsAPI?._syncConflictPending) {
                // Keep modal open until the user resolves the conflict.
            }
            const importExternalCsvModal = document.getElementById('importExternalCsvModal');
            if (e.target === importExternalCsvModal) {
                this._rejectExternalCsvEquipmentStep?.();
            }
            const searchNotesModal = document.getElementById('searchNotesModal');
            if (e.target === searchNotesModal) {
                this.closeSearchNotesModal();
            }
            const yearStatsModal = document.getElementById('yearStatsModal');
            if (e.target === yearStatsModal) {
                this.closeYearStatisticsModal();
            }
            const harnessCanopyPieModal = document.getElementById('harnessCanopyPieModal');
            if (e.target === harnessCanopyPieModal) {
                this.closeHarnessCanopyPieModal();
            }
            const resetDbConfirmModal = document.getElementById('resetDbConfirmModal');
            if (e.target === resetDbConfirmModal) {
                this.closeResetDbConfirmModal();
            }
            const todoItemModal = document.getElementById('todoItemModal');
            if (e.target === todoItemModal) {
                this.closeTodoItemModal();
            }
            const flysightGraphModal = document.getElementById('flysightGraphModal');
            if (e.target === flysightGraphModal) {
                this.closeFlysightGraphModal();
            }
            const flysightSettingsModal = document.getElementById('flysightSettingsModal');
            if (e.target === flysightSettingsModal) {
                this.closeFlysightSettingsModal();
            }
            const flysightBrowserModal = document.getElementById('flysightBrowserModal');
            if (e.target === flysightBrowserModal) {
                this.closeFlysightBrowserModal();
            }
        });

        document.getElementById('yearStatsClose')?.addEventListener('click', () => {
            this.closeYearStatisticsModal();
        });

        document.getElementById('yearStatsModal')?.addEventListener('click', (e) => {
            const btn = e.target.closest('.year-stats-loc-legend-btn');
            if (!btn) return;
            const modal = document.getElementById('yearStatsModal');
            if (!modal || !modal.contains(btn)) return;
            const yr = btn.dataset.year != null ? parseInt(btn.dataset.year, 10) : NaN;
            const locKey = btn.dataset.locKey;
            if (!Number.isFinite(yr) || locKey == null || locKey === '') return;
            this.openYearLocationCanopyPieModal(yr, locKey);
        });

        document.getElementById('harnessCanopyPieClose')?.addEventListener('click', () => {
            this.closeHarnessCanopyPieModal();
        });

        const jumpsYearSummary = document.getElementById('jumpsYearSummary');
        if (jumpsYearSummary) {
            const openSummaryStats = () => {
                if (!jumpsYearSummary.classList.contains('jumps-year-summary--clickable')) return;
                const raw = jumpsYearSummary.getAttribute('data-stats-year');
                const yr = raw != null ? parseInt(raw, 10) : NaN;
                if (!Number.isFinite(yr)) return;
                this.openYearStatisticsModal(yr);
            };
            jumpsYearSummary.addEventListener('click', openSummaryStats);
            jumpsYearSummary.addEventListener('keydown', (e) => {
                if (e.key !== 'Enter' && e.key !== ' ') return;
                if (!jumpsYearSummary.classList.contains('jumps-year-summary--clickable')) return;
                e.preventDefault();
                openSummaryStats();
            });
        }

        const statsPastMonthsInput = document.getElementById('statsPastMonthsWindow');
        if (statsPastMonthsInput) {
            statsPastMonthsInput.value = String(this.settings.statsPastMonthsWindow ?? 3);
            statsPastMonthsInput.addEventListener('input', () => this._updateJumpsPastMonthsSummary());
            statsPastMonthsInput.addEventListener('change', () => {
                let n = parseInt(statsPastMonthsInput.value, 10);
                if (!Number.isFinite(n) || n < 1) n = this.settings.statsPastMonthsWindow ?? 3;
                n = Math.min(240, Math.max(1, n));
                statsPastMonthsInput.value = String(n);
                this.settings.statsPastMonthsWindow = n;
                localStorage.setItem('skydiving-settings', JSON.stringify(this.settings));
                this._updateJumpsPastMonthsSummary();
            });
        }

        this.setupCanopyPicker();
    }

    setCurrentDate() {
        const now = new Date();
        const yyyy = now.getFullYear();
        const mm = String(now.getMonth() + 1).padStart(2, '0');
        const dd = String(now.getDate()).padStart(2, '0');
        document.getElementById('date').value = `${yyyy}-${mm}-${dd}`;
    }

    /**
     * Most recently *logged* jump (latest `timestamp`), not necessarily the latest calendar date.
     * `this.jumps` is sorted by date for numbering, so the array tail is not the last submission.
     */
    getLastJumpData() {
        if (this.jumps.length === 0) {
            return null;
        }
        const tsOf = (j) => {
            const t = Date.parse(j.timestamp);
            return Number.isFinite(t) ? t : 0;
        };
        const idOf = (j) => {
            const n = typeof j.id === 'number' ? j.id : Number(j.id);
            return Number.isFinite(n) ? n : 0;
        };
        let best = this.jumps[0];
        for (let i = 1; i < this.jumps.length; i++) {
            const j = this.jumps[i];
            const tj = tsOf(j);
            const tb = tsOf(best);
            if (tj > tb || (tj === tb && idOf(j) > idOf(best))) {
                best = j;
            }
        }
        return best;
    }

    preFillFormWithLastJump() {
        const lastJump = this.getLastJumpData();
        if (lastJump) {
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
        this.syncCanopyPickerDisplay();
    }

    /**
     * Restore the "Auto" drop zone checkbox from saved settings.
     * `form.reset()` clears the checkbox unless the default HTML has `checked`,
     * so call this after any jump-form reset / pre-fill.
     * @param {boolean} runDetect – when true and Auto is on, run GPS nearest-DZ (same as checking the box).
     */
    applyAutoDetectDropZoneUi(runDetect = false) {
        const chk = document.getElementById('autoDetectDZForm');
        if (!chk) return;
        chk.checked = !!this.settings.autoDetectDropZone;
        if (runDetect && chk.checked) {
            this.detectNearestLocation(true);
        }
    }

    updateNextJumpNumber() { /* field removed — kept as no-op for safety */ }

    getNextJumpNumber() {
        if (this.jumps.length === 0) {
            return this.settings.resequenceJumpsFromStartingNumber !== false
                ? this.settings.startingJumpNumber
                : 1;
        }
        return Math.max(...this.jumps.map(j => j.jumpNumber)) + 1;
    }

    /**
     * True when every jump in the import list has a positive finite jump # (typical app / backup JSON).
     * Used to turn off chronological renumbering from "starting jump" after import.
     */
    static importJumpsHaveExplicitNumbers(jumps) {
        if (!Array.isArray(jumps) || jumps.length === 0) return false;
        return jumps.every(j => {
            const n = typeof j.jumpNumber === 'number' ? j.jumpNumber : parseInt(j.jumpNumber, 10);
            return Number.isFinite(n) && n > 0;
        });
    }

    _updateStartingJumpUiState() {
        const chk = document.getElementById('settingsResequenceJumpsCheckbox');
        const input = document.getElementById('startingJumpNumber');
        const row = document.getElementById('startingJumpNumberRow');
        if (!chk || !input) return;
        const on = chk.checked;
        input.disabled = !on;
        input.style.opacity = on ? '' : '0.5';
        if (row) row.style.opacity = on ? '' : '0.72';
    }

    addJump(jumpData = null, silent = false) {
        const form = document.getElementById('jumpForm');
        
        // Use passed data or read from form
        const data = jumpData || {
            date: form.elements['date'].value,
            location: (form.elements['location'].value || '').trim(),
            equipment: form.elements['equipment'].value,
            notes: form.elements['notes'].value || ''
        };
        data.location = String(data.location ?? '').trim();
        if (!data.location) {
            this.showMessage('Please enter a location.', 'error');
            return;
        }

        // Determine the active lineset for the selected canopy
        let linesetNumber = data.linesetNumber; // may be pre-set by caller
        if (!linesetNumber && data.equipment) {
            linesetNumber = this.getActiveLinesetNumber(data.equipment);
        }
        
        // Remember the highest jump number before insertion to detect a past-date entry
        const reseq = this.settings.resequenceJumpsFromStartingNumber !== false;
        const maxBefore = this.jumps.length > 0
            ? Math.max(...this.jumps.map(j => j.jumpNumber))
            : (reseq ? this.settings.startingJumpNumber - 1 : 0);

        const jump = {
            id: Date.now() + Math.random(),
            jumpId: window.SheetsAPI ? SheetsAPI.generateJumpId() : ('jump-' + Date.now() + '-' + Math.random().toString(36).slice(2)),
            jumpNumber: reseq ? 0 : this.getNextJumpNumber(),
            date: data.date,
            location: data.location,
            equipment: data.equipment,  // canopy ID
            linesetNumber: linesetNumber || 1,
            notes: data.notes,
            timestamp: new Date().toISOString()
        };
        const harnessSnap = this._harnessIdSnapshotForJump(data.equipment);
        if (harnessSnap) jump.harnessId = harnessSnap;

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
        
        // Insert then sort by date; optionally renumber everything from startingJumpNumber
        this.jumps.push(jump);
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

    /**
     * Calendar year from a jump object or a raw date string (YYYY-MM-DD, ISO, etc.).
     * @param {{ date?: string }|string|null|undefined} jumpOrDate
     * @returns {number|null}
     */
    _jumpCalendarYear(jumpOrDate) {
        const raw = jumpOrDate != null && typeof jumpOrDate === 'object' && 'date' in jumpOrDate
            ? jumpOrDate.date
            : jumpOrDate;
        if (raw == null) return null;
        const s = typeof raw === 'string' ? raw.trim() : String(raw);
        if (s.length < 4) return null;
        if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
            const dt = new Date(s.includes('T') ? s : `${s.slice(0, 10)}T12:00:00`);
            if (!Number.isNaN(dt.getTime())) return dt.getFullYear();
        }
        if (/^\d{4}/.test(s)) {
            const y = parseInt(s.slice(0, 4), 10);
            if (Number.isFinite(y)) return y;
        }
        const dt = new Date(s);
        return Number.isNaN(dt.getTime()) ? null : dt.getFullYear();
    }

    /** Local midnight for the jump's calendar day (for range comparisons). */
    _jumpDateAtLocalMidnight(jump) {
        const s = jump.date;
        if (typeof s === 'string' && /^\d{4}-\d{2}-\d{2}/.test(s)) {
            const y = parseInt(s.slice(0, 4), 10);
            const m = parseInt(s.slice(5, 7), 10) - 1;
            const d = parseInt(s.slice(8, 10), 10);
            const dt = new Date(y, m, d);
            return Number.isNaN(dt.getTime()) ? null : dt;
        }
        const t = new Date(jump.date);
        if (Number.isNaN(t.getTime())) return null;
        return new Date(t.getFullYear(), t.getMonth(), t.getDate());
    }

    _updateJumpsPastMonthsSummary() {
        const row = document.getElementById('jumpsPastMonthsRow');
        const countEl = document.getElementById('jumpsPastMonthsCount');
        const input = document.getElementById('statsPastMonthsWindow');
        if (!row || !countEl || !input) return;
        if (this.jumps.length === 0) {
            row.hidden = true;
            return;
        }
        row.hidden = false;
        let n = parseInt(input.value, 10);
        if (!Number.isFinite(n) || n < 1) n = this.settings.statsPastMonthsWindow ?? 3;
        n = Math.min(240, Math.max(1, n));
        const cutoff = new Date();
        cutoff.setHours(0, 0, 0, 0);
        cutoff.setMonth(cutoff.getMonth() - n);
        let count = 0;
        for (const jump of this.jumps) {
            const jd = this._jumpDateAtLocalMidnight(jump);
            if (jd && jd >= cutoff) count++;
        }
        countEl.textContent = String(count);
    }

    _updateJumpsYearSummary() {
        const el = document.getElementById('jumpsYearSummary');
        if (!el) return;
        if (this.jumps.length === 0) {
            el.textContent = '';
            el.hidden = true;
            el.classList.remove('jumps-year-summary--clickable');
            el.removeAttribute('data-stats-year');
            el.removeAttribute('role');
            el.removeAttribute('tabindex');
            el.removeAttribute('aria-label');
            this._updateJumpsPastMonthsSummary();
            return;
        }
        const y = new Date().getFullYear();
        let thisYear = 0;
        let lastYear = 0;
        for (const jump of this.jumps) {
            const jy = this._jumpCalendarYear(jump);
            if (jy === y) thisYear++;
            else if (jy === y - 1) lastYear++;
        }
        if (thisYear > 0) {
            el.textContent = `Number of jumps this year: ${thisYear}`;
            el.hidden = false;
            el.classList.add('jumps-year-summary--clickable');
            el.setAttribute('data-stats-year', String(y));
            el.setAttribute('role', 'button');
            el.setAttribute('tabindex', '0');
            el.setAttribute('aria-label', `Open jump statistics for ${y} (${thisYear} jumps)`);
        } else if (lastYear > 0) {
            const prevY = y - 1;
            el.textContent = `Number of jumps last year: ${lastYear}`;
            el.hidden = false;
            el.classList.add('jumps-year-summary--clickable');
            el.setAttribute('data-stats-year', String(prevY));
            el.setAttribute('role', 'button');
            el.setAttribute('tabindex', '0');
            el.setAttribute('aria-label', `Open jump statistics for ${prevY} (${lastYear} jumps)`);
        } else {
            el.textContent = '';
            el.hidden = true;
            el.classList.remove('jumps-year-summary--clickable');
            el.removeAttribute('data-stats-year');
            el.removeAttribute('role');
            el.removeAttribute('tabindex');
            el.removeAttribute('aria-label');
        }
        this._updateJumpsPastMonthsSummary();
    }

    /** Calendar years strictly before the current year that have at least one jump. Newest first. */
    _getPreviousYearsWithJumps() {
        const cy = new Date().getFullYear();
        const years = new Set();
        for (const jump of this.jumps) {
            const jy = this._jumpCalendarYear(jump);
            if (jy != null && jy < cy) years.add(jy);
        }
        return [...years].sort((a, b) => b - a);
    }

    _jumpsInCalendarYear(year) {
        return this.jumps.filter(j => this._jumpCalendarYear(j) === year);
    }

    /** Same bucket key as year location stats (trimmed name or "No location"). */
    _jumpLocationStatsKey(jump) {
        const raw = (jump.location != null ? String(jump.location) : '').trim();
        return raw || 'No location';
    }

    _aggregateJumpsByLocationForYear(jumps) {
        const map = new Map();
        for (const j of jumps) {
            const key = this._jumpLocationStatsKey(j);
            map.set(key, (map.get(key) || 0) + 1);
        }
        return [...map.entries()]
            .map(([label, count]) => ({ label, count }))
            .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label, undefined, { sensitivity: 'base' }));
    }

    /** Count jumps per canopy (equipment id), merging all linesets for the same canopy. */
    _aggregateJumpsByCanopy(jumps) {
        const map = new Map();
        for (const j of (jumps || [])) {
            const id = j.equipment || '';
            map.set(id, (map.get(id) || 0) + 1);
        }
        const rows = [...map.entries()].map(([equipmentId, count]) => {
            const canopy = this.canopies.find(c => c.id === equipmentId);
            const label = canopy ? canopy.name : 'Unknown canopy';
            return { label, count, equipmentId };
        });
        rows.sort((a, b) => b.count - a.count || a.label.localeCompare(b.label, undefined, { sensitivity: 'base' }));
        return rows;
    }

    _aggregateJumpsByCanopyForYear(jumps) {
        return this._aggregateJumpsByCanopy(jumps);
    }

    /**
     * Logged jumps that snapshot this harness, grouped by canopy (jump.equipment).
     * Pre-app harness counts are not tied to a canopy and are excluded.
     */
    _aggregateLoggedJumpsByCanopyForHarness(harnessId) {
        const hid = harnessId;
        return this._aggregateJumpsByCanopy(
            this.jumps.filter(j => this._normalizeHarnessId(j.harnessId) === hid)
        );
    }

    _yearStatsPieColors() {
        return ['#1976D2', '#388E3C', '#F57C00', '#7B1FA2', '#C2185B', '#0097A7', '#5D4037', '#455A64', '#AFB42B', '#E91E63', '#3F51B5', '#689F38'];
    }

    /**
     * Stable canopy → color for year statistics pies, so the same equipment id keeps
     * the same color when switching years (order no longer follows that year's slice index).
     * Order: canopies by sortOrder, then any equipment ids seen on jumps but not in the list.
     */
    _yearStatsCanopyColorMap() {
        const palette = this._yearStatsPieColors();
        const bySort = [...this.canopies].sort((a, b) => (a.sortOrder ?? Infinity) - (b.sortOrder ?? Infinity));
        const orderedIds = bySort.map(c => c.id);
        const fromJumps = new Set();
        for (const j of this.jumps) {
            fromJumps.add(j.equipment || '');
        }
        const extras = [...fromJumps].filter(id => !orderedIds.includes(id)).sort();
        const allIds = [...orderedIds, ...extras];
        const map = new Map();
        allIds.forEach((id, i) => {
            map.set(id, palette[i % palette.length]);
        });
        return map;
    }

    _piePolar(cx, cy, r, rad) {
        return [cx + r * Math.cos(rad), cy + r * Math.sin(rad)];
    }

    /**
     * @param {{ label: string, count: number, equipmentId?: string }[]} entries
     * @param {{ getColor?: (entry: { label: string, count: number }, index: number) => string, interactiveLocationLegend?: { year: number } }} [options]
     * @returns {string} HTML (pie SVG + legend with counts)
     */
    _renderPieChartBlock(entries, options = {}) {
        const { getColor, interactiveLocationLegend } = options;
        const filtered = (entries || []).filter(e => e.count > 0);
        const total = filtered.reduce((s, e) => s + e.count, 0);
        if (total === 0) {
            return '<p class="no-items year-stats-pie-empty">No data for this chart.</p>';
        }
        const colors = this._yearStatsPieColors();
        const colorAt = (e, i) => (getColor ? getColor(e, i) : colors[i % colors.length]);
        const cx = 100;
        const cy = 100;
        const r = 90;
        const labelRadius = r * 0.7;
        let svgInner;
        if (filtered.length === 1) {
            const fill = colorAt(filtered[0], 0);
            const c0 = filtered[0].count;
            svgInner = `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${fill}" stroke="#fff" stroke-width="2"/>
                <text class="year-stats-pie-slice-value" x="${cx}" y="${cy}" text-anchor="middle" dominant-baseline="middle">${c0}</text>`;
        } else {
            let angle = -Math.PI / 2;
            const paths = [];
            const labels = [];
            filtered.forEach((e, i) => {
                const sweep = (e.count / total) * Math.PI * 2;
                const a0 = angle;
                const a1 = angle + sweep;
                const [sx, sy] = this._piePolar(cx, cy, r, a0);
                const [ex, ey] = this._piePolar(cx, cy, r, a1);
                const largeArc = sweep > Math.PI ? 1 : 0;
                const fill = colorAt(e, i);
                paths.push(
                    `<path d="M ${cx} ${cy} L ${sx.toFixed(3)} ${sy.toFixed(3)} A ${r} ${r} 0 ${largeArc} 1 ${ex.toFixed(3)} ${ey.toFixed(3)} Z" fill="${fill}" stroke="#fff" stroke-width="2"/>`
                );
                const mid = angle + sweep / 2;
                const [tx, ty] = this._piePolar(cx, cy, labelRadius, mid);
                labels.push(
                    `<text class="year-stats-pie-slice-value" x="${tx.toFixed(2)}" y="${ty.toFixed(2)}" text-anchor="middle" dominant-baseline="middle">${e.count}</text>`
                );
                angle = a1;
            });
            svgInner = paths.join('') + labels.join('');
        }
        const legendItems = filtered.map((e, i) => {
            const c = colorAt(e, i);
            if (interactiveLocationLegend && Number.isFinite(interactiveLocationLegend.year)) {
                const yr = interactiveLocationLegend.year;
                const enc = encodeURIComponent(e.label);
                const tip = `Show jumps per canopy at ${e.label} in ${yr}`;
                const tipEsc = this.escapeHtml(tip);
                return `<li>
                    <button type="button" class="year-stats-loc-legend-btn" data-year="${yr}" data-loc-key="${enc}" title="${tipEsc}" aria-label="${tipEsc}">
                        <span class="year-stats-swatch" style="background:${c}"></span>
                        <span class="year-stats-legend-label">${this.escapeHtml(e.label)}</span>
                        <span class="year-stats-legend-count">${e.count}</span>
                    </button>
                </li>`;
            }
            return `<li>
                <span class="year-stats-swatch" style="background:${c}"></span>
                <span class="year-stats-legend-label">${this.escapeHtml(e.label)}</span>
                <span class="year-stats-legend-count">${e.count}</span>
            </li>`;
        }).join('');
        return `
            <div class="year-stats-pie-visual">
                <svg viewBox="0 0 200 200" class="year-stats-svg" aria-hidden="true">${svgInner}</svg>
            </div>
            <ul class="year-stats-legend">${legendItems}</ul>
        `;
    }

    openYearStatisticsModal(year) {
        const modal = document.getElementById('yearStatsModal');
        if (!modal) return;
        this._renderYearStatisticsModalContent(year);
        modal.style.display = 'block';
    }

    closeYearStatisticsModal() {
        const modal = document.getElementById('yearStatsModal');
        if (modal) modal.style.display = 'none';
    }

    /**
     * @param {number} year Calendar year
     * @param {string} locationKeyEncoded `encodeURIComponent` of the location stats key (legend label)
     */
    openYearLocationCanopyPieModal(year, locationKeyEncoded) {
        let locationKey;
        try {
            locationKey = decodeURIComponent(locationKeyEncoded);
        } catch {
            return;
        }
        const yearInt = typeof year === 'number' ? year : parseInt(year, 10);
        if (!Number.isFinite(yearInt)) return;
        const modal = document.getElementById('harnessCanopyPieModal');
        if (!modal) return;
        const filtered = this._jumpsInCalendarYear(yearInt).filter(j =>
            this._jumpLocationStatsKey(j) === locationKey
        );
        if (filtered.length === 0) {
            this.showMessage('No jumps found for this location in that year.', 'error');
            return;
        }
        const byCan = this._aggregateJumpsByCanopy(filtered);
        const jumpCount = filtered.length;
        const locationText = locationKey;
        this._renderCanopyPieModal({
            headingText: `Jump statistics — ${yearInt}`,
            subText: `${locationText} · ${jumpCount} jump${jumpCount === 1 ? '' : 's'}`,
            entries: byCan
        });
        modal.style.display = 'block';
    }

    openHarnessCanopyPieModal(harnessId) {
        const modal = document.getElementById('harnessCanopyPieModal');
        if (!modal) return;
        this._renderHarnessCanopyPieModalContent(harnessId);
        modal.style.display = 'block';
    }

    closeHarnessCanopyPieModal() {
        const modal = document.getElementById('harnessCanopyPieModal');
        if (modal) modal.style.display = 'none';
    }

    _renderCanopyPieModal({ headingText, subText, entries, noteText = '' }) {
        const heading = document.getElementById('harnessCanopyPieHeading');
        const sub = document.getElementById('harnessCanopyPieSub');
        const root = document.getElementById('harnessCanopyPieRoot');
        const note = document.getElementById('harnessCanopyPiePreAppNote');
        if (!heading || !sub || !root || !note) return;

        heading.textContent = headingText;
        sub.textContent = subText;

        const canopyColorMap = this._yearStatsCanopyColorMap();
        const palette = this._yearStatsPieColors();
        root.innerHTML = this._renderPieChartBlock(entries, {
            getColor: (e) => canopyColorMap.get(e.equipmentId) ?? palette[0]
        });

        if (noteText) {
            note.hidden = false;
            note.textContent = noteText;
        } else {
            note.hidden = true;
            note.textContent = '';
        }
    }

    _renderHarnessCanopyPieModalContent(harnessId) {
        const harness = this.harnesses.find(h => h.id === harnessId);
        const name = harness?.name || 'Harness';
        const preApp = harness?.previousJumps ?? 0;

        const byCan = this._aggregateLoggedJumpsByCanopyForHarness(harnessId);
        const logged = byCan.reduce((s, e) => s + e.count, 0);
        const subText = logged === 0
            ? (preApp > 0
                ? 'No logged jumps on this harness yet (pre-app jumps are not split by canopy).'
                : 'No logged jumps on this harness yet.')
            : (logged === 1 ? '1 logged jump' : `${logged} logged jumps`);
        const noteText = preApp > 0
            ? `This harness also has ${preApp} pre-app jump${preApp === 1 ? '' : 's'} (not tied to a canopy); they are not included in the chart.`
            : '';
        this._renderCanopyPieModal({
            headingText: name,
            subText,
            entries: byCan.map(({ label, count, equipmentId }) => ({ label, count, equipmentId })),
            noteText
        });
    }

    openMonthLocationCanopyPieModal(groupId) {
        const modal = document.getElementById('harnessCanopyPieModal');
        if (!modal) return;
        this._renderMonthLocationCanopyPieModalContent(groupId);
        modal.style.display = 'block';
    }

    _renderMonthLocationCanopyPieModalContent(groupId) {
        const group = this._monthLocationPieGroups.get(groupId);
        if (!group) return;
        const byCan = this._aggregateJumpsByCanopy(group.jumps);
        const jumpCount = group.jumps.length;
        const locationText = group.location || 'No location';
        this._renderCanopyPieModal({
            headingText: group.monthLabel,
            subText: `${locationText} · ${jumpCount} jump${jumpCount === 1 ? '' : 's'}`,
            entries: byCan
        });
    }

    openDayLocationCanopyPieModal(groupId) {
        const modal = document.getElementById('harnessCanopyPieModal');
        if (!modal) return;
        this._renderDayLocationCanopyPieModalContent(groupId);
        modal.style.display = 'block';
    }

    _renderDayLocationCanopyPieModalContent(groupId) {
        const group = this._dayLocationPieGroups.get(groupId);
        if (!group) return;
        const byCan = this._aggregateJumpsByCanopy(group.jumps);
        const jumpCount = group.jumps.length;
        const locationText = group.location || 'No location';
        this._renderCanopyPieModal({
            headingText: group.dateLabel,
            subText: `${locationText} · ${jumpCount} jump${jumpCount === 1 ? '' : 's'}`,
            entries: byCan
        });
    }

    _bindHarnessStatsPieClicks(container) {
        const harnessList = container.querySelector('#harnessStatsList');
        if (!harnessList) return;
        const openFromRow = (row) => {
            const id = row.getAttribute('data-harness-id');
            if (id) this.openHarnessCanopyPieModal(id);
        };
        harnessList.addEventListener('click', (e) => {
            const row = e.target.closest('.stat-item-harness[data-harness-id]');
            if (!row || !harnessList.contains(row)) return;
            openFromRow(row);
        });
        harnessList.addEventListener('keydown', (e) => {
            if (e.key !== 'Enter' && e.key !== ' ') return;
            const row = e.target.closest('.stat-item-harness[data-harness-id]');
            if (!row || !harnessList.contains(row)) return;
            e.preventDefault();
            openFromRow(row);
        });
    }

    _renderYearStatisticsModalContent(year) {
        const heading = document.getElementById('yearStatsHeading');
        const totalLine = document.getElementById('yearStatsTotalLine');
        const locRoot = document.getElementById('yearStatsLocationPie');
        const canRoot = document.getElementById('yearStatsCanopyPie');
        const prevWrap = document.getElementById('yearStatsPrevYearsWrap');
        const prevBar = document.getElementById('yearStatsPrevYearsBar');
        if (!heading || !totalLine || !locRoot || !canRoot || !prevWrap || !prevBar) return;

        const jumps = this._jumpsInCalendarYear(year);
        const total = jumps.length;
        heading.textContent = `Jump statistics — ${year}`;
        totalLine.textContent = total === 1 ? '1 jump' : `${total} jumps`;

        const byLoc = this._aggregateJumpsByLocationForYear(jumps);
        const byCan = this._aggregateJumpsByCanopyForYear(jumps);
        locRoot.innerHTML = this._renderPieChartBlock(byLoc, { interactiveLocationLegend: { year } });
        const canopyColorMap = this._yearStatsCanopyColorMap();
        const palette = this._yearStatsPieColors();
        canRoot.innerHTML = this._renderPieChartBlock(
            byCan.map(({ label, count, equipmentId }) => ({ label, count, equipmentId })),
            {
                getColor: (e) => canopyColorMap.get(e.equipmentId) ?? palette[0]
            }
        );

        const prevYears = this._getPreviousYearsWithJumps();
        const cy = new Date().getFullYear();
        const yearsForNav = [];
        if (this._jumpsInCalendarYear(cy).length > 0) yearsForNav.push(cy);
        for (const py of prevYears) {
            if (!yearsForNav.includes(py)) yearsForNav.push(py);
        }
        yearsForNav.sort((a, b) => b - a);

        if (yearsForNav.length <= 1) {
            prevWrap.hidden = true;
            prevBar.innerHTML = '';
        } else {
            prevWrap.hidden = false;
            prevBar.innerHTML = yearsForNav.map(y => {
                const active = y === year ? ' is-active' : '';
                return `<button type="button" class="btn-secondary year-stats-year-btn${active}" onclick="logbook.openYearStatisticsModal(${y})">${y}</button>`;
            }).join('');
        }
    }

    renderJumpsList() {
        const jumpsList = document.getElementById('jumpsList');
        const totalEl = document.getElementById('recentJumpsTotal');
        const allJumpsByMonth = this.settings.recentJumpsDays === 0;
        this._monthLocationPieGroups = new Map();
        this._dayLocationPieGroups = new Map();

        const updateRecentTotal = (count) => {
            if (!totalEl) return;
            const hideTotal = allJumpsByMonth || this.settings.recentJumpsGroupByMonth;
            if (hideTotal) {
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
            this._updateJumpsYearSummary();
            this._updateRecentJumpsGroupByMonthUi();
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

        const byMonth =
            !allJumpsByMonth && this.settings.recentJumpsGroupByMonth;
        const useMergedMonthList = byMonth && recentJumps.length > 0;

        const PAGE_SIZE = 100;
        let html = '';
        let remaining = 0;

        if (useMergedMonthList) {
            // One month+location block per month/location across recent and older jumps
            this._useMergedListCache = true;
            this._mergedJumpsCache = [...recentJumps, ...olderJumps];
            this._pastYearBuckets = null;
            const targetFirst = Math.max(PAGE_SIZE, recentJumps.length);
            const endIndex = this._findMonthCompleteIndex(this._mergedJumpsCache, targetFirst);
            this._renderedMergedCount = endIndex;
            this._olderJumpsCache = olderJumps;
            this._renderedOlderCount = 0;

            if (endIndex > 0) {
                html += this._renderOlderMonthGroups(
                    this._mergedJumpsCache.slice(0, endIndex)
                );
            }
            remaining = this._mergedJumpsCache.length - this._renderedMergedCount;
        } else {
            this._useMergedListCache = false;
            this._mergedJumpsCache = [];
            this._renderedMergedCount = 0;

            const { sameYearOlder, pastYearBuckets } = this._splitOlderJumpsSameYearAndPastYears(olderJumps);
            this._olderJumpsCache = sameYearOlder;
            this._pastYearBuckets = pastYearBuckets;

            const endIndex = this._findMonthCompleteIndex(sameYearOlder, PAGE_SIZE);
            const initialOlder = sameYearOlder.slice(0, endIndex);
            this._renderedOlderCount = initialOlder.length;

            if (recentJumps.length > 0) {
                html += this.renderDayLocationGroups(recentJumps, { expandFirst: true });
            }
            if (sameYearOlder.length > 0) {
                html += `<div id="olderSameYearMonthsWrap">${
                    initialOlder.length ? this._renderOlderMonthGroups(initialOlder) : ''
                }</div>`;
            }
            remaining = sameYearOlder.length - this._renderedOlderCount;
        }

        if (remaining > 0) {
            html += `<button class="btn-secondary load-more-btn" id="loadMoreJumpsBtn" onclick="logbook.loadMoreJumps()">Load more (${remaining} remaining)</button>`;
        }

        if (!this._useMergedListCache && this._pastYearBuckets && this._pastYearBuckets.length > 0) {
            html += `<div id="olderPastYearsWrap">${this._renderPastYearCollapseGroups(this._pastYearBuckets)}</div>`;
        }

        updateRecentTotal(recentJumps.length);
        jumpsList.innerHTML = html;
        this._updateJumpsYearSummary();
        this._updateRecentJumpsGroupByMonthUi();
    }

    /**
     * Split non-recent jumps into (a) current calendar year — still paginated by month —
     * and (b) prior calendar years — shown as collapsed year rows so users open one year at a time.
     */
    _splitOlderJumpsSameYearAndPastYears(olderJumps) {
        const currentYear = new Date().getFullYear();
        const sameYearOlder = [];
        const byPastYear = new Map();

        for (const jump of olderJumps) {
            const y = this._jumpCalendarYear(jump);
            if (y === null) {
                if (!byPastYear.has('unknown')) byPastYear.set('unknown', []);
                byPastYear.get('unknown').push(jump);
            } else if (y >= currentYear) {
                sameYearOlder.push(jump);
            } else {
                if (!byPastYear.has(y)) byPastYear.set(y, []);
                byPastYear.get(y).push(jump);
            }
        }

        const numericYears = [...byPastYear.keys()].filter(k => k !== 'unknown').sort((a, b) => b - a);
        const pastYearBuckets = numericYears.map(year => ({ year, jumps: byPastYear.get(year) }));
        if (byPastYear.has('unknown')) {
            pastYearBuckets.push({ year: 'unknown', jumps: byPastYear.get('unknown') });
        }
        return { sameYearOlder, pastYearBuckets };
    }

    /** Collapsible headers for each past calendar year; body uses month+location groups. */
    _renderPastYearCollapseGroups(pastYearBuckets) {
        if (!pastYearBuckets.length) return '';
        return pastYearBuckets.map(({ year, jumps }) => {
            const slug = year === 'unknown' ? 'unknown' : String(year);
            const label = year === 'unknown' ? 'Unknown date' : String(year);
            const n = jumps.length;
            const countStr = n === 1 ? '1 jump' : `${n} jumps`;
            const statsButtonHtml = year === 'unknown'
                ? ''
                : `<button type="button" class="month-group-pie-btn" title="Open jump statistics" aria-label="Open jump statistics for ${this.escapeHtml(String(year))}" onclick="event.stopPropagation(); logbook.openYearStatisticsModal(${year})">◔</button>`;
            return `
                <div class="year-group" data-year="${this.escapeHtml(slug)}">
                    <div class="year-group-header" onclick="logbook.toggleYearGroup('${slug}')">
                        <span class="year-group-arrow" id="arrow-year-${slug}">&#9654;</span>
                        <span class="year-group-label">${this.escapeHtml(label)}</span>
                        <span class="year-group-count">${countStr}</span>
                        ${statsButtonHtml}
                    </div>
                    <div class="year-group-body" id="year-group-body-${slug}" style="display:none;">
                        ${this._renderOlderMonthGroups(jumps)}
                    </div>
                </div>`;
        }).join('');
    }

    /** Render jumps as collapsed month + location groups (day groups inside are collapsed too). */
    _renderOlderMonthGroups(jumps) {
        const pairKey = (monthKey, location) => `${monthKey}\x00${location.toLowerCase()}`;
        const monthLocationGroups = new Map();
        // `jumps` is sorted by jumpNumber descending; first encounter of a month+location
        // is that group's newest jump — use its index so within a month, latest-logged
        // location appears first (not alphabetical).
        jumps.forEach((jump, idx) => {
            const d = new Date(jump.date);
            const monthKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
            const location = (jump.location || '').trim();
            const pk = pairKey(monthKey, location);
            if (!monthLocationGroups.has(pk)) {
                monthLocationGroups.set(pk, {
                    monthKey,
                    monthLabel: d.toLocaleDateString(undefined, { year: 'numeric', month: 'long' }),
                    location,
                    newestJumpListIndex: idx,
                    jumps: []
                });
            }
            monthLocationGroups.get(pk).jumps.push(jump);
        });

        const entries = [...monthLocationGroups.values()].sort((a, b) => {
            if (a.monthKey !== b.monthKey) return b.monthKey.localeCompare(a.monthKey);
            if (a.newestJumpListIndex !== b.newestJumpListIndex) {
                return a.newestJumpListIndex - b.newestJumpListIndex;
            }
            return (a.location || '').localeCompare(b.location || '', undefined, { sensitivity: 'base' });
        });

        const usedDomIds = new Set();
        let html = '';
        for (const group of entries) {
            const jumpCount = group.jumps.length;
            const locSlug = (group.location || 'noloc').replace(/[^a-zA-Z0-9]+/g, '_').replace(/^_|_$/g, '') || 'noloc';
            let domId = `month_${group.monthKey}_${locSlug}`;
            let n = 1;
            while (usedDomIds.has(domId)) {
                domId = `month_${group.monthKey}_${locSlug}_${n++}`;
            }
            usedDomIds.add(domId);

            const locationHtml = group.location
                ? `<span class="month-group-location">📍 ${this.escapeHtml(group.location)}</span>`
                : '<span class="month-group-location month-group-location-empty">No location</span>';
            this._monthLocationPieGroups.set(domId, {
                monthLabel: group.monthLabel,
                location: group.location,
                jumps: [...group.jumps]
            });

            html += `
                <div class="month-group" data-month="${group.monthKey}" data-location="${this.escapeHtml(group.location)}">
                    <div class="month-group-header" onclick="logbook.toggleMonthGroup('${domId}')">
                        <span class="month-group-arrow" id="arrow-${domId}">&#9654;</span>
                        <span class="month-group-label">${group.monthLabel}</span>
                        ${locationHtml}
                        <span class="month-group-count">${jumpCount} jump${jumpCount !== 1 ? 's' : ''}</span>
                        <button type="button" class="month-group-pie-btn" title="Show jumps per canopy" aria-label="Show jumps per canopy" onclick="event.stopPropagation(); logbook.openMonthLocationCanopyPieModal('${domId}')">◔</button>
                    </div>
                    <div class="month-group-body" id="month-${domId}" style="display:none;">
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

    /** Append the next page of jumps to the list. */
    loadMoreJumps() {
        const PAGE_SIZE = 100;
        const btn = document.getElementById('loadMoreJumpsBtn');
        if (btn) btn.remove();

        const jumpsList = document.getElementById('jumpsList');
        if (!jumpsList) return;

        let remaining = 0;

        if (this._useMergedListCache) {
            const endIndex = this._findMonthCompleteIndex(
                this._mergedJumpsCache,
                this._renderedMergedCount + PAGE_SIZE
            );
            if (endIndex <= this._renderedMergedCount) return;
            this._renderedMergedCount = endIndex;

            let html = this._renderOlderMonthGroups(
                this._mergedJumpsCache.slice(0, this._renderedMergedCount)
            );
            remaining = this._mergedJumpsCache.length - this._renderedMergedCount;
            if (remaining > 0) {
                html += `<button class="btn-secondary load-more-btn" id="loadMoreJumpsBtn" onclick="logbook.loadMoreJumps()">Load more (${remaining} remaining)</button>`;
            }
            jumpsList.innerHTML = html;
            return;
        }

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

        const sameYearWrap = document.getElementById('olderSameYearMonthsWrap');
        const pastWrap = document.getElementById('olderPastYearsWrap');
        const fragment = document.createElement('div');
        fragment.innerHTML = this._renderOlderMonthGroups(nextBatch);
        const appendParent = sameYearWrap || jumpsList;
        while (fragment.firstChild) appendParent.appendChild(fragment.firstChild);

        remaining = this._olderJumpsCache.length - this._renderedOlderCount;
        if (remaining > 0) {
            const newBtn = document.createElement('button');
            newBtn.className = 'btn-secondary load-more-btn';
            newBtn.id = 'loadMoreJumpsBtn';
            newBtn.textContent = `Load more (${remaining} remaining)`;
            newBtn.onclick = () => this.loadMoreJumps();
            if (pastWrap) pastWrap.insertAdjacentElement('beforebegin', newBtn);
            else if (sameYearWrap) sameYearWrap.insertAdjacentElement('afterend', newBtn);
            else jumpsList.appendChild(newBtn);
        }
    }

    toggleYearGroup(yearSlug) {
        const body = document.getElementById(`year-group-body-${yearSlug}`);
        const arrow = document.getElementById(`arrow-year-${yearSlug}`);
        if (!body) return;
        const open = body.style.display !== 'none';
        body.style.display = open ? 'none' : 'block';
        if (arrow) arrow.innerHTML = open ? '&#9654;' : '&#9660;';
    }

    toggleMonthGroup(domId) {
        const body = document.getElementById('month-' + domId);
        const arrow = document.getElementById('arrow-' + domId);
        if (!body) return;
        const open = body.style.display !== 'none';
        body.style.display = open ? 'none' : 'block';
        if (arrow) arrow.innerHTML = open ? '&#9654;' : '&#9660;';
    }

    _updateRecentJumpsGroupByMonthUi() {
        const wrap = document.getElementById('recentJumpsGroupByMonthWrap');
        const toggle = document.getElementById('recentJumpsGroupByMonthToggle');
        const settingsChk = document.getElementById('recentJumpsGroupByMonthSettings');
        const hasRecentSection = (this.settings.recentJumpsDays || 0) > 0;
        if (wrap) wrap.hidden = !hasRecentSection;
        const val = !!this.settings.recentJumpsGroupByMonth;
        if (toggle) toggle.checked = val;
        if (settingsChk) settingsChk.checked = val;
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
            this._dayLocationPieGroups.set(dayId, {
                dateLabel: group.dateLabel,
                location: group.location,
                jumps: [...group.jumps]
            });
            return `
            <div class="day-location-group">
                <div class="day-location-header" onclick="logbook.toggleDayGroup('${dayId}')">
                    <span class="day-group-arrow" id="arrow-${dayId}">${expanded ? '&#9660;' : '&#9654;'}</span>
                    <span class="day-date">${group.dateLabel}</span>
                    ${group.location ? `<span class="day-location">📍 ${group.location}</span>` : ''}
                    <span class="day-jump-range">${this.getJumpCountLabel(group.jumps)}</span>
                    <button type="button" class="month-group-pie-btn" title="Show jumps per canopy" aria-label="Show jumps per canopy" onclick="event.stopPropagation(); logbook.openDayLocationCanopyPieModal('${dayId}')">◔</button>
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
        const hid = this._normalizeHarnessId(jump.harnessId);
        let harnessFrag = '';
        if (hid) {
            const h = this.harnesses.find(x => x.id === hid);
            const hname = h ? h.name : hid;
            harnessFrag = ` <span class="jump-harness" title="Harness at log time">\u00B7 ${this.escapeHtml(hname)}</span>`;
        }

        return `
            <div class="jump-row">
                <button type="button" class="jump-number jump-number-btn" onclick="logbook.openEditJumpModal('${encodedJumpId}')" title="Edit date, location, canopy, or note">#${jump.jumpNumber}</button>
                <span class="jump-canopy">🪂 ${canopyNameHtml}${harnessFrag}</span>
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

    _normalizeDateForInput(dateVal) {
        if (dateVal == null || dateVal === '') return '';
        const s = String(dateVal);
        if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
        const d = new Date(s);
        if (isNaN(d.getTime())) return '';
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${y}-${m}-${day}`;
    }

    /** Same chronological order as `renumberJumps()` (date asc, then timestamp). */
    _sortJumpsChronologically(jumps) {
        return [...jumps].sort((a, b) => {
            const da = Date.parse(a.date);
            const db = Date.parse(b.date);
            if (isNaN(da) && isNaN(db)) return 0;
            if (isNaN(da)) return 1;
            if (isNaN(db)) return -1;
            if (da !== db) return da - db;
            return Date.parse(a.timestamp) - Date.parse(b.timestamp);
        });
    }

    /** Calendar-day difference `to - from` for `YYYY-MM-DD` strings (UTC date math). */
    _isoDateDeltaDays(fromYyyyMmDd, toYyyyMmDd) {
        const [y1, m1, d1] = fromYyyyMmDd.split('-').map(Number);
        const [y2, m2, d2] = toYyyyMmDd.split('-').map(Number);
        return (Date.UTC(y2, m2 - 1, d2) - Date.UTC(y1, m1 - 1, d1)) / 86400000;
    }

    /** Add signed day offset to a calendar date; accepts any jump date string via normalization. */
    _addCalendarDaysToIsoDate(dateVal, dayDelta) {
        if (!dayDelta) return this._normalizeDateForInput(dateVal);
        const norm = /^\d{4}-\d{2}-\d{2}$/.test(String(dateVal).slice(0, 10))
            ? String(dateVal).slice(0, 10)
            : this._normalizeDateForInput(dateVal);
        if (!norm) return norm;
        const [y, m, d] = norm.split('-').map(Number);
        const ms = Date.UTC(y, m - 1, d + dayDelta);
        const dt = new Date(ms);
        const mm = String(dt.getUTCMonth() + 1).padStart(2, '0');
        const dd = String(dt.getUTCDate()).padStart(2, '0');
        return `${dt.getUTCFullYear()}-${mm}-${dd}`;
    }

    /**
     * Populate edit-jump canopy select. Option values are `canopyId:linesetNumber`
     * (lineset is always last segment after the final ':').
     */
    fillEditJumpEquipmentSelect(canopyId, linesetNumber) {
        const sel = document.getElementById('editJumpEquipment');
        if (!sel) return;

        sel.innerHTML = '';
        const ph = document.createElement('option');
        ph.value = '';
        ph.textContent = 'Select canopy';
        sel.appendChild(ph);

        const wantLs = Number(linesetNumber) || 1;
        let matchedValue = null;

        const addOpt = (id, num, label) => {
            const opt = document.createElement('option');
            opt.value = `${id}:${num}`;
            opt.textContent = label;
            sel.appendChild(opt);
            if (id === canopyId && Number(num) === wantLs) matchedValue = opt.value;
        };

        for (const canopy of this.canopies) {
            if (canopy.archived) continue;
            const lss = (canopy.linesets || []).filter(ls => !ls.archived).sort((a, b) => a.number - b.number);
            for (const ls of lss) {
                const hybridTag = ls.hybrid ? ' (Hybrid)' : '';
                addOpt(canopy.id, ls.number, `${canopy.name} — Lineset #${ls.number}${hybridTag}`);
            }
        }

        if (!matchedValue && canopyId) {
            const c = this.canopies.find(x => x.id === canopyId);
            const num = wantLs;
            const labelBase = c ? c.name : 'Unknown canopy';
            const extra = (!c || c.archived) ? ' (archived / inactive)' : '';
            addOpt(canopyId, num, `${labelBase} — Lineset #${num}${extra}`);
            matchedValue = `${canopyId}:${num}`;
        }

        if (sel.options.length === 1) {
            ph.textContent = 'No canopies available';
        }

        sel.value = matchedValue != null && matchedValue !== '' ? matchedValue : '';
    }

    openEditJumpModal(encodedJumpId) {
        const modal = document.getElementById('editJumpModal');
        const dateInput = document.getElementById('editJumpDate');
        const locInput = document.getElementById('editJumpLocation');
        const notesInput = document.getElementById('editJumpNotes');
        if (!modal || !dateInput || !locInput || !notesInput) return;

        const id = decodeURIComponent(encodedJumpId || '');
        const jump = this.jumps.find(j => j.id.toString() === id.toString());
        if (!jump) {
            this.showMessage('Jump not found', 'error');
            return;
        }

        this.closeJumpNotePopup();
        this.activeEditJumpId = id;

        const label = document.getElementById('editJumpNumberLabel');
        if (label) label.textContent = `#${jump.jumpNumber}`;

        dateInput.value = this._normalizeDateForInput(jump.date);
        locInput.value = jump.location || '';
        notesInput.value = typeof jump.notes === 'string' ? jump.notes : '';
        this.editJumpLocationAtOpen = (jump.location || '').trim();
        this.editJumpDateAtOpen = this._normalizeDateForInput(jump.date);
        this.fillEditJumpEquipmentSelect(jump.equipment, jump.linesetNumber);

        modal.style.display = 'block';
        this.syncEditJumpModalBulkOptions();
        notesInput.focus();
    }

    closeEditJumpModal() {
        const modal = document.getElementById('editJumpModal');
        const editDd = document.getElementById('editJumpLocationDropdown');
        if (editDd) editDd.classList.remove('open');
        if (modal) modal.style.display = 'none';
        this.activeEditJumpId = null;
        this.editJumpLocationAtOpen = null;
        this.editJumpDateAtOpen = null;
        const applyWrap = document.getElementById('editJumpApplyLocationToDayWrap');
        const applyChk = document.getElementById('editJumpApplyLocationToSameDay');
        if (applyWrap) applyWrap.hidden = true;
        if (applyChk) applyChk.checked = false;
        const shiftWrap = document.getElementById('editJumpShiftFollowingWrap');
        const shiftChk = document.getElementById('editJumpShiftFollowingChk');
        const shiftCnt = document.getElementById('editJumpShiftFollowingCount');
        if (shiftWrap) shiftWrap.hidden = true;
        if (shiftChk) shiftChk.checked = false;
        if (shiftCnt) shiftCnt.value = '1';
    }

    syncEditJumpModalBulkOptions() {
        this.syncEditJumpApplyLocationToDayOption();
        this.syncEditJumpShiftFollowingDatesOption();
    }

    /**
     * Show "apply location to all other jumps this day" when the location text changed
     * and there is at least one other jump on the date currently selected in the modal.
     */
    syncEditJumpApplyLocationToDayOption() {
        const wrap = document.getElementById('editJumpApplyLocationToDayWrap');
        const chk = document.getElementById('editJumpApplyLocationToSameDay');
        const dateInput = document.getElementById('editJumpDate');
        const locInput = document.getElementById('editJumpLocation');
        if (!wrap || !chk || !dateInput || !locInput || !this.activeEditJumpId) {
            if (wrap) wrap.hidden = true;
            if (chk) chk.checked = false;
            return;
        }

        const day = dateInput.value;
        const locNow = (locInput.value || '').trim();
        const opened = this.editJumpLocationAtOpen != null ? this.editJumpLocationAtOpen : '';
        const locChanged = locNow !== opened;

        const othersOnDay = this.jumps.filter(j => {
            if (j.id.toString() === this.activeEditJumpId.toString()) return false;
            return this._normalizeDateForInput(j.date) === day;
        });

        const show = !!day && locChanged && othersOnDay.length > 0;
        wrap.hidden = !show;
        if (!show) {
            chk.checked = false;
            const cap = document.getElementById('editJumpApplyLocationToSameDayCaption');
            if (cap) cap.textContent = 'Also apply this location to other jumps on this day';
            return;
        }

        const cap = document.getElementById('editJumpApplyLocationToSameDayCaption');
        if (cap && othersOnDay.length > 0) {
            const n = othersOnDay.length;
            cap.textContent =
                n === 1
                    ? 'Also apply this location to the other jump on this day'
                    : `Also apply this location to all ${n} other jumps on this day`;
        }
    }

    /**
     * When the jump date is changed from when the modal opened, offer to shift the next N
     * chronologically following jumps by the same calendar-day delta (same sort as renumber).
     */
    syncEditJumpShiftFollowingDatesOption() {
        const wrap = document.getElementById('editJumpShiftFollowingWrap');
        const chk = document.getElementById('editJumpShiftFollowingChk');
        const cntEl = document.getElementById('editJumpShiftFollowingCount');
        const hint = document.getElementById('editJumpShiftFollowingHint');
        const dateInput = document.getElementById('editJumpDate');
        if (!wrap || !chk || !cntEl || !dateInput || !this.activeEditJumpId || !this.editJumpDateAtOpen) {
            if (wrap) wrap.hidden = true;
            if (chk) chk.checked = false;
            if (hint) hint.textContent = '';
            return;
        }

        const dayNew = dateInput.value;
        const dayOld = this.editJumpDateAtOpen;
        const dateChanged = !!dayNew && !!dayOld && dayNew !== dayOld;

        const sorted = this._sortJumpsChronologically(this.jumps);
        const idx = sorted.findIndex(j => j.id.toString() === this.activeEditJumpId.toString());
        const followingCount = idx >= 0 ? sorted.length - idx - 1 : 0;

        const show = dateChanged && followingCount > 0;
        wrap.hidden = !show;
        if (!show) {
            chk.checked = false;
            if (hint) hint.textContent = '';
            return;
        }

        cntEl.max = String(followingCount);
        cntEl.min = '1';
        let n = parseInt(cntEl.value, 10);
        if (!Number.isFinite(n) || n < 1) n = 1;
        if (n > followingCount) {
            n = followingCount;
            cntEl.value = String(n);
        }

        if (hint) {
            hint.textContent =
                followingCount === 1
                    ? 'There is 1 later jump after this one in date order (ties broken by log order).'
                    : `There are ${followingCount} later jumps after this one in date order (ties broken by log order).`;
        }
    }

    saveEditedJump() {
        if (!this.activeEditJumpId) {
            this.showMessage('Jump not found', 'error');
            return;
        }

        const jump = this.jumps.find(j => j.id.toString() === this.activeEditJumpId.toString());
        if (!jump) {
            this.showMessage('Jump not found', 'error');
            this.closeEditJumpModal();
            return;
        }

        const dateInput = document.getElementById('editJumpDate');
        const locInput = document.getElementById('editJumpLocation');
        const eqSel = document.getElementById('editJumpEquipment');
        const notesInput = document.getElementById('editJumpNotes');
        if (!dateInput || !locInput || !eqSel || !notesInput) return;

        const date = dateInput.value;
        if (!date) {
            this.showMessage('Please select a date', 'error');
            return;
        }

        const eqVal = eqSel.value;
        if (!eqVal) {
            this.showMessage('Please select a canopy', 'error');
            return;
        }

        const li = eqVal.lastIndexOf(':');
        const equipment = li >= 0 ? eqVal.slice(0, li) : eqVal;
        const linesetNumber = li >= 0 ? (parseInt(eqVal.slice(li + 1), 10) || 1) : 1;

        const location = (locInput.value || '').trim();
        if (!location) {
            this.showMessage('Please enter a location.', 'error');
            return;
        }
        const applyLocToSameDay = !!document.getElementById('editJumpApplyLocationToSameDay')?.checked;

        const oldDateNorm = this._normalizeDateForInput(jump.date);
        const shiftChkEl = document.getElementById('editJumpShiftFollowingChk');
        const shiftCntEl = document.getElementById('editJumpShiftFollowingCount');
        const shiftFollowing =
            !!shiftChkEl?.checked && !!oldDateNorm && date !== oldDateNorm;
        let followerTargets = [];
        let deltaDays = 0;
        if (shiftFollowing) {
            deltaDays = this._isoDateDeltaDays(oldDateNorm, date);
            if (deltaDays !== 0) {
                const sorted = this._sortJumpsChronologically(this.jumps);
                const idx = sorted.findIndex(j => j.id.toString() === this.activeEditJumpId.toString());
                const maxFollow = idx >= 0 ? sorted.length - idx - 1 : 0;
                let n = parseInt(shiftCntEl?.value, 10) || 1;
                n = Math.min(Math.max(1, n), Math.max(0, maxFollow));
                for (let i = 1; i <= n && idx + i < sorted.length; i++) {
                    followerTargets.push(sorted[idx + i]);
                }
            }
        }

        jump.date = date;
        jump.location = location;
        jump.notes = notesInput.value || '';
        jump.equipment = equipment;
        jump.linesetNumber = linesetNumber;
        const harnessSnap = this._harnessIdSnapshotForJump(equipment);
        if (harnessSnap) jump.harnessId = harnessSnap;
        else delete jump.harnessId;

        if (followerTargets.length && deltaDays !== 0) {
            for (const j of followerTargets) {
                const cur = this._normalizeDateForInput(j.date);
                j.date = this._addCalendarDaysToIsoDate(cur, deltaDays);
            }
        }

        if (applyLocToSameDay) {
            for (const j of this.jumps) {
                if (j.id.toString() === this.activeEditJumpId.toString()) continue;
                if (this._normalizeDateForInput(j.date) !== date) continue;
                j.location = location;
            }
        }

        if (location) {
            const locationExists = this.locations.some(
                loc => loc.name.toLowerCase() === location.toLowerCase()
            );
            if (!locationExists) {
                const newId = 'loc_' + Date.now();
                const newLoc = { id: newId, name: location, lat: null, lng: null };
                this.locations.push(newLoc);
                DB.putAll('locations', this.locations).catch(err => console.error('[DB] Failed to save locations:', err));
                this.updateLocationDatalist();
                this.geocodeLocation(newLoc);
            }
        }

        this.renumberJumps();
        this.initializeCanopyLinesetJumpCounts();
        this.saveToLocalStorage();
        this.updateStats();
        this.renderJumpsList();
        if (this.currentView === 'equipment') {
            this.renderEquipmentView();
        }

        this.closeEditJumpModal();
        const shiftN = followerTargets.length;
        this.showMessage(
            shiftN > 0
                ? `Jump updated — ${shiftN} following jump${shiftN === 1 ? '' : 's'} shifted by the same calendar-day change.`
                : 'Jump updated',
            'success'
        );

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

    normalizeNoteSearchText(text) {
        return String(text)
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '');
    }

    executeNoteSearch(options = {}) {
        const input = document.getElementById('searchNotesInput');
        const resultsContainer = document.getElementById('searchNotesResults');
        if (!input || !resultsContainer) return;

        const presetTerms = Array.isArray(options.terms)
            ? options.terms.filter(t => typeof t === 'string' && t.trim())
            : [];
        const query = input.value.trim();
        const terms = presetTerms.length > 0 ? presetTerms : (query ? [query] : []);

        if (terms.length === 0) {
            resultsContainer.innerHTML =
                '<p class="search-notes-placeholder">Enter a search term to find jumps by note text.</p>';
            return;
        }

        const normalizedTerms = terms.map(t => this.normalizeNoteSearchText(t)).filter(Boolean);
        const matches = this.jumps
            .filter(j => typeof j.notes === 'string' && normalizedTerms.some(term =>
                this.normalizeNoteSearchText(j.notes).includes(term)))
            .sort((a, b) => b.jumpNumber - a.jumpNumber);

        if (matches.length === 0) {
            resultsContainer.innerHTML = '<p class="search-notes-placeholder">No jumps found matching that text.</p>';
            return;
        }

        const highlightPattern = [...terms]
            .sort((a, b) => b.length - a.length)
            .map(term => term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/[eéèêë]/gi, '[eéèêë]'))
            .join('|');
        const queryRegex = new RegExp(highlightPattern, 'gi');

        const countSuffix = options.label ? ` (${this.escapeHtml(options.label)})` : '';
        let html = `<p class="search-notes-count">${matches.length} result${matches.length !== 1 ? 's' : ''} found${countSuffix}</p>`;
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

            html += `
                <div class="search-result-item" onclick="logbook.closeSearchNotesModal(); logbook.openEditJumpModal('${encodedJumpId}')">
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

        if (this.settings.resequenceJumpsFromStartingNumber === false) {
            return;
        }
        // Renumber jumps starting from the configured starting number
        this.jumps.forEach((jump, index) => {
            jump.jumpNumber = this.settings.startingJumpNumber + index;
        });
    }

    /** After sync conflict resolution, always renumber from startingJumpNumber in date/time order. */
    forceRenumberJumpsAfterSync() {
        const prev = this.settings.resequenceJumpsFromStartingNumber;
        this.settings.resequenceJumpsFromStartingNumber = true;
        this.renumberJumps();
        this.settings.resequenceJumpsFromStartingNumber = prev;
    }

    openSettingsModal() {
        document.getElementById('startingJumpNumber').value = this.settings.startingJumpNumber;
        const reseqChk = document.getElementById('settingsResequenceJumpsCheckbox');
        if (reseqChk) {
            reseqChk.checked = this.settings.resequenceJumpsFromStartingNumber !== false;
            this._updateStartingJumpUiState();
        }
        const prev = this.settings.previousStartingJump;
        const current = this.settings.startingJumpNumber;
        const labelEl = document.getElementById('startingJumpNumberLabel');
        const showPrevious = prev != null && prev !== 1 && prev !== current;
        labelEl.textContent = showPrevious ? `Starting Jump Number (previous=${prev})` : 'Starting Jump Number';
        document.getElementById('recentJumpsDays').value = this.settings.recentJumpsDays ?? 3;
        const recentGrp = document.getElementById('recentJumpsGroupByMonthSettings');
        if (recentGrp) recentGrp.checked = !!this.settings.recentJumpsGroupByMonth;
        document.getElementById('standardRedThreshold').value = this.settings.standardRedThreshold ?? 160;
        document.getElementById('standardOrangeThreshold').value = this.settings.standardOrangeThreshold ?? 140;
        document.getElementById('hybridRedThreshold').value = this.settings.hybridRedThreshold ?? 80;
        document.getElementById('hybridOrangeThreshold').value = this.settings.hybridOrangeThreshold ?? 60;

        this._populateNavSettingsControls();

        const cacheVerEl = document.getElementById('settingsCacheVersion');
        if (cacheVerEl) {
            cacheVerEl.textContent =
                typeof CACHE_VERSION !== 'undefined'
                    ? `Cache version ${CACHE_VERSION}`
                    : '';
        }

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
            this.showMessage('Connected to Google Sheets!', 'success', 1000);
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

            this.showMessage('Connected to Google Sheets!', 'success', 1000);
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

    /** Disconnect OAuth and stop background sheet polling (no UI prompt). */
    async _disconnectGoogleSheetsSync() {
        await window.AuthManager.signOut();
        window.SheetsAPI.initialized = false;
        window.SheetsAPI.spreadsheetId = '';
        window.SheetsAPI._cancelPoll();
        window.SheetsAPI.updateSyncStatus('Not signed in');
    }

    async handleGoogleSignOut() {
        if (!confirm('Sign out and disconnect Google Sheets sync?')) return;

        await this._disconnectGoogleSheetsSync();

        this.showMessage('Signed out from Google Sheets', 'success');
        this.openSheetsModal(); // refresh modal state
    }

    openResetDbConfirmModal() {
        const modal = document.getElementById('resetDbConfirmModal');
        if (modal) modal.style.display = 'block';
    }

    closeResetDbConfirmModal() {
        const modal = document.getElementById('resetDbConfirmModal');
        if (modal) modal.style.display = 'none';
    }

    async confirmResetLocalDb() {
        await this._disconnectGoogleSheetsSync();
        localStorage.clear();
        try {
            await DB.open();
            await DB.clearAll();
        } catch (_) { /* IDB may not be available */ }
        this.closeResetDbConfirmModal();
        this.closeModal();
        this.showMessage('Local data removed. Reloading…', 'success');
        setTimeout(() => window.location.reload(), 300);
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
        const reseqEl = document.getElementById('settingsResequenceJumpsCheckbox');
        const nowResequence = reseqEl ? !!reseqEl.checked : true;
        const wasResequence = this.settings.resequenceJumpsFromStartingNumber !== false;

        const previousStartingJumpNumber = this.settings.startingJumpNumber;
        let startingJumpNumber = previousStartingJumpNumber;

        if (nowResequence) {
            startingJumpNumber = parseInt(document.getElementById('startingJumpNumber').value, 10);
            if (!startingJumpNumber || startingJumpNumber < 1) {
                this.showMessage('Please enter a valid starting jump number (1 or higher)', 'error');
                return;
            }
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

        this.settings.resequenceJumpsFromStartingNumber = nowResequence;
        if (nowResequence) {
            this.settings.startingJumpNumber = startingJumpNumber;
            // Persist the value we're leaving, so we can show (previous=XX) when opening settings
            this.settings.previousStartingJump = previousStartingJumpNumber;
        }
        this.settings.recentJumpsDays = recentJumpsDays;
        const recentGrpSettings = document.getElementById('recentJumpsGroupByMonthSettings');
        if (recentGrpSettings) {
            this.settings.recentJumpsGroupByMonth = recentGrpSettings.checked;
        }
        this.settings.standardRedThreshold = standardRedThreshold;
        this.settings.standardOrangeThreshold = standardOrangeThreshold;
        this.settings.hybridRedThreshold = hybridRedThreshold;
        this.settings.hybridOrangeThreshold = hybridOrangeThreshold;

        const visibleNavViews = this._readNavVisibilityFromForm();
        if (visibleNavViews.length === 0) {
            this.showMessage('Keep at least one navigation tab turned on', 'error');
            return;
        }
        this.settings.visibleNavViews = visibleNavViews;
        const startViewEl = document.getElementById('settingsStartView');
        const startView = startViewEl?.value;
        this.settings.startView = visibleNavViews.includes(startView) ? startView : visibleNavViews[0];
        this.normalizeNavSettings();
        this.applyNavVisibility();
        if (!this.settings.visibleNavViews.includes(this.currentView)) {
            this.showView(this.settings.startView);
        }

        localStorage.setItem('skydiving-settings', JSON.stringify(this.settings));
        this.markEquipmentModified();

        // Mark data as locally modified so the background poller detects pending changes.
        localStorage.setItem('skydiving-data-modified', new Date().toISOString());

        const needsJumpsRenumber = nowResequence
            && (!wasResequence || previousStartingJumpNumber !== this.settings.startingJumpNumber);
        if (needsJumpsRenumber) {
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
        if (needsJumpsRenumber) {
            this.showMessage('Settings saved. Jump numbers updated. Reloading...', 'success');
            setTimeout(() => window.location.reload(), 300);
        } else {
            this.showMessage('Settings saved successfully!', 'success');
            this.renderJumpsList();
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

    /**
     * Ensure every jump has a unique local `id` (IndexedDB keyPath) and a stable `jumpId` (Sheets sync / backups).
     * Duplicate or missing `id` causes silent data loss on save; missing `jumpId` breaks sync and merge.
     * @returns {{ repairedLocalIds: number, addedJumpIds: number }} counts of rows actually changed
     */
    ensureJumpIds() {
        const genJumpId = () => (typeof crypto !== 'undefined' && crypto.randomUUID)
            ? crypto.randomUUID()
            : 'jump-' + Date.now() + '-' + Math.random().toString(36).slice(2);
        let seq = 0;
        const seenLocalIds = new Set();
        /** New local id must not collide with any jump already processed (fixes reassignment clashing with existing ids). */
        const allocUniqueLocalId = (idx) => {
            if (typeof crypto !== 'undefined' && crypto.randomUUID) {
                const s = 'lb-' + crypto.randomUUID();
                if (!seenLocalIds.has(s)) {
                    return s;
                }
            }
            let candidate;
            let guard = 0;
            do {
                candidate = Date.now() + Math.random() + (++seq) * 1e-9 + idx * 1e-15 + (++guard) * 1e-18;
            } while (seenLocalIds.has(String(candidate)));
            return candidate;
        };

        let needsSave = false;
        let repairedLocalIds = 0;
        let addedJumpIds = 0;
        this.jumps.forEach((jump, idx) => {
            const idStr = jump.id != null && jump.id !== '' ? String(jump.id) : '';
            if (!idStr || seenLocalIds.has(idStr)) {
                jump.id = allocUniqueLocalId(idx);
                needsSave = true;
                repairedLocalIds++;
            }
            seenLocalIds.add(String(jump.id));

            if (!jump.jumpId) {
                jump.jumpId = genJumpId();
                needsSave = true;
                addedJumpIds++;
            }
        });
        if (needsSave) {
            DB.replaceAllJumps(this.jumps).catch(err => console.error('[DB] Failed to save jumps after jump id / jumpId fix:', err));
        }
        return { repairedLocalIds, addedJumpIds };
    }

    /**
     * One-click repair for duplicate or missing per-device jump `id` values (IndexedDB key) and missing `jumpId`.
     * Does not change jump numbers, dates, locations, or sheet sync identity (`jumpId` is only added when absent).
     */
    repairJumpLocalIdsFromSettings() {
        const { repairedLocalIds, addedJumpIds } = this.ensureJumpIds();
        if (repairedLocalIds === 0 && addedJumpIds === 0) {
            this.showMessage('No duplicate or missing jump IDs found. Your log is already consistent.', 'info');
            return;
        }
        this.markJumpsModified();
        if (this.currentView === 'jumps') {
            this.renderJumpsList();
        }
        if (this.currentView === 'equipment') {
            this.renderEquipmentView();
        }
        if (this.currentView === 'stats') {
            this.renderStats();
        }
        this.updateStats();
        const parts = [];
        if (repairedLocalIds > 0) {
            parts.push(`fixed ${repairedLocalIds} local storage id${repairedLocalIds === 1 ? '' : 's'} (missing or duplicate)`);
        }
        if (addedJumpIds > 0) {
            parts.push(`added ${addedJumpIds} missing sync id${addedJumpIds === 1 ? '' : 's'} (jumpId)`);
        }
        this.showMessage(`Repair complete: ${parts.join('; ')}. Your data was saved.`, 'success');
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

    loadTodos() {
        try {
            const parsed = JSON.parse(localStorage.getItem('skydiving-todos') || '[]');
            if (!Array.isArray(parsed)) return [];
            return parsed
                .filter(t => t && typeof t.text === 'string')
                .map(t => this._normalizeTodo(t));
        } catch (_) {
            return [];
        }
    }

    _normalizeTodo(t) {
        const createdAt = Number(t.createdAt) || Date.now();
        const done = Boolean(t.done);
        const doneAt = done ? (Number(t.doneAt) || createdAt) : null;
        const updatedAt = Number(t.updatedAt) || (doneAt || createdAt);
        return {
            id: String(t.id || this._newTodoId()),
            text: t.text,
            done,
            createdAt,
            doneAt,
            updatedAt
        };
    }

    loadDeletedTodos() {
        try {
            const parsed = JSON.parse(localStorage.getItem('skydiving-deleted-todos') || '[]');
            if (!Array.isArray(parsed)) return [];
            const byId = new Map();
            for (const d of parsed) {
                const id = d && String(d.id || '').trim();
                if (!id || byId.has(id)) continue;
                byId.set(id, {
                    id,
                    deletedAt: d.deletedAt || new Date().toISOString()
                });
            }
            return Array.from(byId.values());
        } catch (_) {
            return [];
        }
    }

    saveDeletedTodos() {
        try {
            localStorage.setItem('skydiving-deleted-todos', JSON.stringify(this.deletedTodos || []));
        } catch (err) {
            console.error('[TODOs] Failed to save deletions:', err);
        }
    }

    recordTodoDeletions(ids) {
        const now = new Date().toISOString();
        const have = new Set((this.deletedTodos || []).map(d => d.id));
        if (!Array.isArray(this.deletedTodos)) this.deletedTodos = [];
        for (const rawId of ids || []) {
            const id = String(rawId || '').trim();
            if (!id || have.has(id)) continue;
            this.deletedTodos.push({ id, deletedAt: now });
            have.add(id);
        }
        this.saveDeletedTodos();
    }

    saveTodos() {
        try {
            localStorage.setItem('skydiving-todos', JSON.stringify(this.todos));
        } catch (err) {
            console.error('[TODOs] Failed to save:', err);
        }
        if (this._applyingTodoSync) return;
        localStorage.setItem('skydiving-data-modified', new Date().toISOString());
        if (navigator.onLine && window.SheetsAPI?.initialized) {
            window.SheetsAPI.pushAllWithGuard();
        }
    }

    applyTodosFromSync(todos, deletedRecords) {
        this._applyingTodoSync = true;
        try {
            this.todos = Array.isArray(todos) ? todos.map(t => this._normalizeTodo(t)) : [];
            this.deletedTodos = Array.isArray(deletedRecords) ? deletedRecords : [];
            this.saveTodos();
            this.saveDeletedTodos();
            if (this.currentView === 'todos') this.renderTodosList();
        } finally {
            this._applyingTodoSync = false;
        }
    }

    _newTodoId() {
        return 'todo-' + Date.now() + '-' + Math.random().toString(36).slice(2, 8);
    }

    _todoCheckIconSvg() {
        return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path fill="currentColor" d="M9 16.17 4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg>';
    }

    /** Element.closest that works on text nodes and SVG (older Android WebView). */
    _todoEventEl(e) {
        let el = e.target;
        if (!el) return null;
        if (el.nodeType !== 1) el = el.parentElement;
        while (el && typeof el.closest !== 'function') el = el.parentElement;
        return el;
    }

    _bindTodoEvents() {
        document.getElementById('addTodoBtn')?.addEventListener('click', () => {
            this.showTodoAddForm();
        });

        document.getElementById('todoAddForm')?.addEventListener('submit', (e) => {
            e.preventDefault();
            this.addTodoFromInput();
        });

        document.getElementById('todoEditForm')?.addEventListener('submit', (e) => {
            e.preventDefault();
            this.saveTodoEdit();
        });

        document.getElementById('todoRemoveBtn')?.addEventListener('click', () => {
            this.removeEditingTodo();
        });

        document.getElementById('todoItemModalClose')?.addEventListener('click', () => {
            this.closeTodoItemModal();
        });

        const list = document.getElementById('todosList');
        if (!list) return;

        list.addEventListener('click', (e) => {
            if (this._todoSuppressClick) {
                e.preventDefault();
                e.stopPropagation();
                this._todoSuppressClick = false;
                return;
            }
            const el = this._todoEventEl(e);
            if (!el) return;
            const clearBtn = el.closest('#clearDoneTodosBtn');
            if (clearBtn) {
                this.clearDoneTodos();
                return;
            }
            const checkBtn = el.closest('.todo-check-btn');
            if (checkBtn) {
                this.toggleTodoDone(checkBtn.dataset.todoId);
            }
        });

        const startLongPress = (x, y, id) => {
            this._cancelTodoLongPress();
            this._todoLongPressId = id;
            this._todoPointerStartX = x;
            this._todoPointerStartY = y;
            this._todoLongPressTimer = setTimeout(() => {
                this._todoLongPressTimer = null;
                this._todoSuppressClick = true;
                if (typeof navigator.vibrate === 'function') {
                    try { navigator.vibrate(12); } catch (_) { /* iOS */ }
                }
                this.openTodoItemModal(id);
            }, 500);
        };

        const moveLongPress = (x, y) => {
            if (!this._todoLongPressTimer) return;
            if (Math.hypot(x - this._todoPointerStartX, y - this._todoPointerStartY) > 10) {
                this._cancelTodoLongPress();
            }
        };

        const itemIdFromEvent = (e) => {
            const el = this._todoEventEl(e);
            if (!el) return null;
            if (el.closest('.todo-check-btn') || el.closest('#clearDoneTodosBtn')) return null;
            const item = el.closest('.todo-item');
            return item ? item.dataset.todoId : null;
        };

        if (window.PointerEvent) {
            list.addEventListener('pointerdown', (e) => {
                if (e.pointerType === 'mouse' && e.button !== 0) return;
                const id = itemIdFromEvent(e);
                if (!id) return;
                startLongPress(e.clientX, e.clientY, id);
            });
            list.addEventListener('pointermove', (e) => moveLongPress(e.clientX, e.clientY));
            list.addEventListener('pointerup', () => this._cancelTodoLongPress());
            list.addEventListener('pointercancel', () => this._cancelTodoLongPress());
        } else {
            list.addEventListener('touchstart', (e) => {
                const t = e.changedTouches[0];
                if (!t) return;
                const id = itemIdFromEvent(e);
                if (!id) return;
                startLongPress(t.clientX, t.clientY, id);
            }, { passive: true });
            list.addEventListener('touchmove', (e) => {
                const t = e.changedTouches[0];
                if (t) moveLongPress(t.clientX, t.clientY);
            }, { passive: true });
            list.addEventListener('touchend', () => this._cancelTodoLongPress(), { passive: true });
            list.addEventListener('touchcancel', () => this._cancelTodoLongPress(), { passive: true });
        }

        list.addEventListener('contextmenu', (e) => {
            const el = this._todoEventEl(e);
            if (!el) return;
            const item = el.closest('.todo-item');
            if (!item || el.closest('.todo-check-btn')) return;
            e.preventDefault();
            this._cancelTodoLongPress();
            this.openTodoItemModal(item.dataset.todoId);
        });
    }

    _cancelTodoLongPress() {
        if (this._todoLongPressTimer) {
            clearTimeout(this._todoLongPressTimer);
            this._todoLongPressTimer = null;
        }
        this._todoLongPressId = null;
    }

    showTodoAddForm() {
        const form = document.getElementById('todoAddForm');
        const input = document.getElementById('todoAddInput');
        if (!form || !input) return;
        form.hidden = false;
        form.removeAttribute('hidden');
        requestAnimationFrame(() => {
            input.focus();
            try { input.scrollIntoView({ block: 'center', behavior: 'smooth' }); } catch (_) { /* older WebView */ }
        });
    }

    addTodoFromInput() {
        const input = document.getElementById('todoAddInput');
        if (!input) return;
        const text = input.value.trim();
        if (!text) return;
        const now = Date.now();
        this.todos.push({
            id: this._newTodoId(),
            text,
            done: false,
            createdAt: now,
            doneAt: null,
            updatedAt: now
        });
        this.saveTodos();
        input.value = '';
        this.renderTodosList();
        input.focus();
    }

    renderTodosList() {
        const list = document.getElementById('todosList');
        if (!list) return;

        const active = this.todos.filter(t => !t.done);
        const done = this.todos.filter(t => t.done);

        if (active.length === 0 && done.length === 0) {
            list.innerHTML = '<p class="todo-empty">No items yet. Tap + to add one.</p>';
            return;
        }

        const rows = items => items.map(item => this._todoItemHtml(item)).join('');
        let html = rows(active);
        if (done.length > 0) {
            html += `
                <div class="todo-done-separator">
                    <span class="todo-done-separator-label">Done</span>
                    <button type="button" id="clearDoneTodosBtn" class="todo-clear-done-btn">Remove all</button>
                </div>
                ${rows(done)}
            `;
        }
        list.innerHTML = html;
    }

    _todoItemHtml(item) {
        const doneClass = item.done ? ' todo-item-done' : '';
        const label = item.done ? 'Mark as not done' : 'Mark as done';
        const id = this.escapeHtml(item.id);
        return `
            <div class="todo-item${doneClass}" data-todo-id="${id}">
                <span class="todo-item-text">${this.escapeHtml(item.text)}</span>
                <button type="button" class="todo-check-btn" data-todo-id="${id}" title="${label}" aria-label="${label}">
                    ${this._todoCheckIconSvg()}
                </button>
            </div>
        `;
    }

    toggleTodoDone(id) {
        const idx = this.todos.findIndex(t => t.id === id);
        if (idx < 0) return;
        const [item] = this.todos.splice(idx, 1);
        const now = Date.now();
        if (item.done) {
            item.done = false;
            item.doneAt = null;
            item.updatedAt = now;
            const firstDone = this.todos.findIndex(t => t.done);
            if (firstDone === -1) this.todos.push(item);
            else this.todos.splice(firstDone, 0, item);
        } else {
            item.done = true;
            item.doneAt = now;
            item.updatedAt = now;
            this.todos.push(item);
        }
        this.saveTodos();
        this.renderTodosList();
    }

    clearDoneTodos() {
        if (!this.todos.some(t => t.done)) return;
        this.recordTodoDeletions(this.todos.filter(t => t.done).map(t => t.id));
        this.todos = this.todos.filter(t => !t.done);
        this.saveTodos();
        this.renderTodosList();
    }

    openTodoItemModal(id) {
        const item = this.todos.find(t => t.id === id);
        const modal = document.getElementById('todoItemModal');
        const input = document.getElementById('todoEditInput');
        if (!item || !modal || !input) return;
        this._editingTodoId = id;
        input.value = item.text;
        modal.style.display = 'block';
        setTimeout(() => {
            input.focus();
            try {
                input.setSelectionRange(input.value.length, input.value.length);
            } catch (_) { /* some mobile WebViews */ }
        }, 300);
    }

    closeTodoItemModal() {
        const modal = document.getElementById('todoItemModal');
        if (modal) modal.style.display = 'none';
        this._editingTodoId = null;
    }

    saveTodoEdit() {
        const id = this._editingTodoId;
        const input = document.getElementById('todoEditInput');
        if (!id || !input) return;
        const text = input.value.trim();
        if (!text) return;
        const item = this.todos.find(t => t.id === id);
        if (!item) return;
        item.text = text;
        item.updatedAt = Date.now();
        this.saveTodos();
        this.closeTodoItemModal();
        this.renderTodosList();
    }

    removeEditingTodo() {
        const id = this._editingTodoId;
        if (!id) return;
        this.recordTodoDeletions([id]);
        this.todos = this.todos.filter(t => t.id !== id);
        this.saveTodos();
        this.closeTodoItemModal();
        this.renderTodosList();
    }

    normalizeNavSettings() {
        const validIds = MAIN_NAV_VIEWS.map(v => v.id);
        let visible = this.settings.visibleNavViews;
        if (!Array.isArray(visible)) visible = [...validIds];
        visible = validIds.filter(id => visible.includes(id));
        if (visible.length === 0) visible = [...validIds];
        this.settings.visibleNavViews = visible;
        if (!visible.includes(this.settings.startView)) {
            this.settings.startView = visible[0];
        }
    }

    applyNavVisibility() {
        this.normalizeNavSettings();
        const visible = new Set(this.settings.visibleNavViews);
        document.querySelectorAll('.nav-btn').forEach(btn => {
            const viewId = btn.id.replace(/ViewBtn$/, '');
            const show = visible.has(viewId);
            btn.hidden = !show;
            btn.style.display = show ? '' : 'none';
        });
        const nav = document.querySelector('.main-nav');
        if (nav) {
            nav.style.display = this.settings.visibleNavViews.length <= 1 ? 'none' : '';
        }
    }

    _readNavVisibilityFromForm() {
        const validIds = MAIN_NAV_VIEWS.map(v => v.id);
        const checked = new Set(
            [...document.querySelectorAll('.nav-visibility-chk:checked')].map(el => el.value)
        );
        return validIds.filter(id => checked.has(id));
    }

    _populateNavSettingsControls() {
        this.normalizeNavSettings();
        const visible = new Set(this.settings.visibleNavViews);
        document.querySelectorAll('.nav-visibility-chk').forEach(chk => {
            chk.checked = visible.has(chk.value);
            chk.disabled = false;
        });
        this._syncStartViewSelect(this.settings.startView);
        this._lockLastNavVisibilityCheckbox();
    }

    _onNavVisibilityCheckboxChange(changed) {
        const checked = document.querySelectorAll('.nav-visibility-chk:checked');
        if (checked.length === 0) {
            changed.checked = true;
            this.showMessage('Keep at least one navigation tab turned on', 'error');
        }
        this._syncStartViewSelect();
        this._lockLastNavVisibilityCheckbox();
    }

    _lockLastNavVisibilityCheckbox() {
        const boxes = [...document.querySelectorAll('.nav-visibility-chk')];
        const checkedCount = boxes.filter(b => b.checked).length;
        boxes.forEach(b => {
            b.disabled = checkedCount === 1 && b.checked;
        });
    }

    _syncStartViewSelect(preferred) {
        const select = document.getElementById('settingsStartView');
        if (!select) return;
        const visible = this._readNavVisibilityFromForm();
        const keep = preferred || select.value;
        select.innerHTML = visible.map(id => {
            const meta = MAIN_NAV_VIEWS.find(v => v.id === id);
            return `<option value="${this.escapeHtml(id)}">${this.escapeHtml(meta ? meta.label : id)}</option>`;
        }).join('');
        select.value = visible.includes(keep) ? keep : (visible[0] || '');
    }

    showView(viewName) {
        const validIds = MAIN_NAV_VIEWS.map(v => v.id);
        if (!validIds.includes(viewName)) {
            viewName = this.settings.startView || validIds[0];
        }
        this.currentView = viewName;

        const viewEl = document.getElementById(`${viewName}View`);
        if (!viewEl) return;
        
        // Hide all views
        document.querySelectorAll('.view').forEach(view => {
            view.style.display = 'none';
        });
        
        // Show selected view
        viewEl.style.display = 'block';
        
        // Update navigation buttons
        document.querySelectorAll('.nav-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        document.getElementById(`${viewName}ViewBtn`)?.classList.add('active');
        
        // Load view-specific content
        if (viewName === 'jumps') {
            this.renderJumpsList();
        } else if (viewName === 'equipment') {
            this.renderEquipmentView();
        } else if (viewName === 'stats') {
            this.renderStats();
        } else if (viewName === 'flysight') {
            this.renderFlysightView();
        } else if (viewName === 'todos') {
            this.renderTodosList();
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

        this.rebuildCanopyPickerOptions();
    }

    setupCanopyPicker() {
        const wrap = document.getElementById('canopyPicker');
        const toggle = document.getElementById('canopyPickerToggle');
        const list = document.getElementById('canopyPickerList');
        const select = document.getElementById('equipment');
        if (!wrap || !toggle || !list || !select) return;
        if (this._canopyPickerBound) return;
        this._canopyPickerBound = true;

        toggle.addEventListener('click', (e) => {
            e.stopPropagation();
            if (list.classList.contains('open')) {
                this._closeCanopyPicker();
            } else {
                this._scrollCanopyFieldToTopIfNeeded();
                list.classList.add('open');
                list.setAttribute('aria-hidden', 'false');
                toggle.setAttribute('aria-expanded', 'true');
                const finishOpen = () => {
                    this._layoutCanopyPickerList();
                    this._highlightCanopyPickerSelection();
                    const sel = list.querySelector('.canopy-picker-option[aria-selected="true"]');
                    if (sel) sel.scrollIntoView({ block: 'nearest' });
                };
                requestAnimationFrame(() => {
                    requestAnimationFrame(finishOpen);
                });
            }
        });

        list.addEventListener('click', (e) => {
            const opt = e.target.closest('.canopy-picker-option');
            if (!opt) return;
            const value = opt.dataset.value;
            if (value === undefined || value === '') return;
            select.value = value;
            select.dispatchEvent(new Event('change', { bubbles: true }));
            this.syncCanopyPickerDisplay();
            this._closeCanopyPicker();
            toggle.focus();
        });

        document.addEventListener('click', (e) => {
            if (!list.classList.contains('open')) return;
            if (!wrap.contains(e.target)) this._closeCanopyPicker();
        });

        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape' && list.classList.contains('open')) {
                this._closeCanopyPicker();
                toggle.focus();
            }
        });

        window.addEventListener('resize', () => {
            if (list.classList.contains('open')) this._layoutCanopyPickerList();
        });
        window.addEventListener('orientationchange', () => {
            if (list.classList.contains('open')) this._layoutCanopyPickerList();
        });
        if (window.visualViewport) {
            window.visualViewport.addEventListener('resize', () => {
                if (list.classList.contains('open')) this._layoutCanopyPickerList();
            });
        }
    }

    /**
     * If the list would not fit in the current space under the field, scroll the
     * canopy form block toward the top so more viewport height is available below.
     */
    _canopyFieldNeedsTopScroll() {
        const select = document.getElementById('equipment');
        const toggle = document.getElementById('canopyPickerToggle');
        if (!select || !toggle) return false;
        const n = select.options.length - 1; // active canopies, excluding first placeholder
        if (n <= 0) return false;

        const rowEst = 50; // ~padding + line-height; conservative for multi-line labels
        const estContent = n * rowEst;
        const rect = toggle.getBoundingClientRect();
        const margin = 12;
        const vv = window.visualViewport;
        const preSpace = vv
            ? (vv.offsetTop + vv.height - margin - rect.bottom)
            : ((window.innerHeight || document.documentElement.clientHeight) - rect.bottom - margin);

        return estContent > preSpace;
    }

    _scrollCanopyFieldToTopIfNeeded() {
        if (!this._canopyFieldNeedsTopScroll()) return;
        const anchor = document.getElementById('canopyPickerScrollAnchor');
        if (anchor) {
            anchor.scrollIntoView({ block: 'start', behavior: 'auto' });
        }
    }

    /**
     * Set max-height from remaining space below the field (list always opens downward),
     * then shrink when all options fit to avoid a tall empty box.
     */
    _layoutCanopyPickerList() {
        const list = document.getElementById('canopyPickerList');
        const toggle = document.getElementById('canopyPickerToggle');
        if (!list || !toggle) return;

        const rect = toggle.getBoundingClientRect();
        const vv = window.visualViewport;
        const margin = 12;
        let spaceBelow;
        let vhCap;
        if (vv) {
            const visBottom = vv.offsetTop + vv.height - margin;
            spaceBelow = visBottom - rect.bottom;
            vhCap = vv.height;
        } else {
            const vh = window.innerHeight || document.documentElement.clientHeight;
            spaceBelow = vh - rect.bottom - margin;
            vhCap = vh;
        }

        const maxH = Math.max(140, Math.min(spaceBelow, vhCap * 0.92));
        const floorMax = Math.floor(maxH);
        list.style.maxHeight = `${floorMax}px`;

        requestAnimationFrame(() => {
            if (!list.classList.contains('open')) return;
            const natural = list.scrollHeight;
            if (natural > 0 && natural < floorMax - 1) {
                list.style.maxHeight = `${natural}px`;
            }
        });
    }

    _closeCanopyPicker() {
        const list = document.getElementById('canopyPickerList');
        const toggle = document.getElementById('canopyPickerToggle');
        if (list) {
            list.classList.remove('open');
            list.setAttribute('aria-hidden', 'true');
            list.style.maxHeight = '';
        }
        if (toggle) {
            toggle.setAttribute('aria-expanded', 'false');
        }
    }

    _highlightCanopyPickerSelection() {
        const list = document.getElementById('canopyPickerList');
        const select = document.getElementById('equipment');
        if (!list || !select) return;
        const val = select.value;
        list.querySelectorAll('.canopy-picker-option').forEach(el => {
            el.setAttribute('aria-selected', el.dataset.value === val ? 'true' : 'false');
        });
    }

    syncCanopyPickerDisplay() {
        const select = document.getElementById('equipment');
        const display = document.getElementById('canopyPickerDisplay');
        const toggle = document.getElementById('canopyPickerToggle');
        if (!select || !display) return;
        const opt = select.selectedOptions[0];
        display.textContent = opt && opt.value ? opt.textContent : 'Select Canopy';
        if (toggle) {
            const hasChoices = select.options.length > 1;
            toggle.disabled = !hasChoices;
        }
    }

    rebuildCanopyPickerOptions() {
        const select = document.getElementById('equipment');
        const list = document.getElementById('canopyPickerList');
        if (!select || !list) return;

        const wasOpen = list.classList.contains('open');
        const keptValue = select.value;

        list.innerHTML = '';
        for (let i = 0; i < select.options.length; i++) {
            const opt = select.options[i];
            if (!opt.value) continue;
            const div = document.createElement('div');
            div.className = 'canopy-picker-option';
            div.setAttribute('role', 'option');
            div.dataset.value = opt.value;
            div.textContent = opt.textContent;
            list.appendChild(div);
        }

        const stillValid = Array.from(select.options).some(o => o.value === keptValue);
        if (keptValue && !stillValid) {
            select.value = '';
            select.dispatchEvent(new Event('change', { bubbles: true }));
        }

        this.syncCanopyPickerDisplay();
        this._highlightCanopyPickerSelection();

        if (wasOpen) this._layoutCanopyPickerList();
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
        const total = (ls?.jumpCount || 0) + (ls?.previousJumps ?? 0);
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
            const { latest, older } = this._getLatestAndOlderLinesets(canopy);
            const showOlder = !!this.showOlderLinesetsByCanopyId?.[canopy.id];
            const olderId = 'older-linesets-' + String(canopy.id).replace(/[^a-zA-Z0-9_-]/g, '_');

            const renderLinesetRow = (ls) => {
                const logged = ls.jumpCount || 0;
                const preApp = ls.previousJumps ?? 0;
                const total = logged + preApp;
                const hybridBadge = ls.hybrid ? '<span class="hybrid-badge">Hybrid</span>' : '';
                return `
                    <div class="lineset-row ${ls.archived ? 'archived' : ''}">
                        <span class="lineset-info">
                            Lineset #${ls.number} ${hybridBadge}
                            <span class="lineset-jumps">${total} jumps${preApp !== 0 ? ` (${logged} logged + ${preApp} pre-app)` : ''}</span>
                        </span>
                        <span class="lineset-actions">
                            <button onclick="window.logbook.editLineset('${canopy.id}', ${ls.number})" class="btn-edit btn-sm">Edit</button>
                            ${logged === 0 ? `<button type="button" onclick="window.logbook.deleteLineset('${canopy.id}', ${ls.number})" class="btn-delete btn-sm" title="Remove this lineset (no jumps logged in this app)">Delete</button>` : ''}
                        </span>
                    </div>
                `;
            };

            const latestHtml = latest ? renderLinesetRow(latest) : '<p class="no-items" style="margin:4px 0;">No linesets</p>';
            const olderHtml = older.map(renderLinesetRow).join('');

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
                            ${latestHtml}
                            ${older.length ? `
                                <label class="show-older-linesets-label">
                                    <input type="checkbox" ${showOlder ? 'checked' : ''} onchange="window.logbook.toggleOlderLinesets('${canopy.id}', this.checked)">
                                    Show older linesets (${older.length})
                                </label>
                                <div id="${olderId}" class="older-linesets-group" style="display:${showOlder ? 'block' : 'none'};">
                                    ${olderHtml}
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

    /**
     * Latest/active lineset shown by default on a canopy card; remaining linesets are "older".
     * Latest is the highest-numbered non-archived lineset (same as jump logging).
     * If every lineset is archived, fall back to the highest-numbered lineset so the card is not empty.
     */
    _getLatestAndOlderLinesets(canopy) {
        const all = [...(canopy?.linesets || [])].sort((a, b) => b.number - a.number);
        if (all.length === 0) return { latest: null, older: [] };
        const active = all.filter(ls => !ls.archived);
        const latest = active.length > 0
            ? active.reduce((a, b) => (a.number >= b.number ? a : b))
            : all[0];
        const older = all.filter(ls => ls !== latest);
        return { latest, older };
    }

    toggleOlderLinesets(canopyId, show) {
        if (!this.showOlderLinesetsByCanopyId) this.showOlderLinesetsByCanopyId = Object.create(null);
        this.showOlderLinesetsByCanopyId[canopyId] = !!show;
        const id = 'older-linesets-' + String(canopyId).replace(/[^a-zA-Z0-9_-]/g, '_');
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

    /**
     * @returns {number|null} Parsed threshold if valid for stats override, else null.
     */
    _validLinesetStatThreshold(value) {
        const n = typeof value === 'number' ? value : parseInt(value, 10);
        return Number.isFinite(n) && n >= 1 ? n : null;
    }

    _prefillLinesetStatThresholdForm(primaryLs, fallbackLs, inputIds) {
        for (let i = 0; i < LINESET_STAT_THRESHOLD_PROPS.length; i++) {
            const prop = LINESET_STAT_THRESHOLD_PROPS[i];
            const el = document.getElementById(inputIds[i]);
            if (!el) continue;
            let v = this._validLinesetStatThreshold(primaryLs?.[prop]);
            if (v == null) v = this._validLinesetStatThreshold(fallbackLs?.[prop]);
            el.value = v != null ? String(v) : '';
        }
    }

    _applyLinesetStatThresholdFormToLineset(lineset, inputIds) {
        for (let i = 0; i < LINESET_STAT_THRESHOLD_PROPS.length; i++) {
            const prop = LINESET_STAT_THRESHOLD_PROPS[i];
            const el = document.getElementById(inputIds[i]);
            if (!el) continue;
            const raw = String(el.value).trim();
            if (raw === '') {
                delete lineset[prop];
            } else {
                const n = parseInt(raw, 10);
                if (Number.isFinite(n) && n >= 1) lineset[prop] = n;
                else delete lineset[prop];
            }
        }
        delete lineset.hybridOrangeThreshold;
        delete lineset.hybridRedThreshold;
    }

    /** Remove all per-lineset statistics threshold overrides (e.g. hybrid linesets use settings only). */
    _purgeLinesetStatThresholdOverrides(lineset) {
        delete lineset.standardOrangeThreshold;
        delete lineset.standardRedThreshold;
        delete lineset.hybridOrangeThreshold;
        delete lineset.hybridRedThreshold;
    }

    _syncLinesetModalStatThresholdSectionVisibility() {
        const hybrid = !!document.getElementById('linesetHybridCheck')?.checked;
        const sec = document.getElementById('linesetStatThresholdSection');
        if (sec) sec.style.display = hybrid ? 'none' : 'block';
    }

    _syncNewCanopyStatThresholdSectionVisibility() {
        const hybrid = !!document.getElementById('newCanopyHybridCheck')?.checked;
        const sec = document.getElementById('newCanopyStatThresholdSection');
        if (sec) sec.style.display = hybrid ? 'none' : 'block';
    }

    _setLinesetStatThresholdInputPlaceholders(inputIds) {
        const defaults = [
            this.settings.standardOrangeThreshold,
            this.settings.standardRedThreshold
        ];
        for (let i = 0; i < inputIds.length; i++) {
            const el = document.getElementById(inputIds[i]);
            if (el) el.placeholder = `default (${defaults[i]})`;
        }
    }

    /** Active non-archived lineset with the highest number (the one a new lineset typically replaces). */
    _getActiveReferenceLinesetForNewLineset(canopy) {
        const active = (canopy.linesets || []).filter(ls => !ls.archived);
        if (active.length === 0) return null;
        return active.reduce((a, b) => (a.number >= b.number ? a : b));
    }

    /** Lineset on the same canopy with the greatest number strictly less than `linesetNumber`. */
    _getPriorLinesetByNumber(canopy, linesetNumber) {
        const n = parseInt(linesetNumber, 10);
        const candidates = (canopy.linesets || []).filter(ls => ls.number < n);
        if (candidates.length === 0) return null;
        return candidates.reduce((a, b) => (a.number >= b.number ? a : b));
    }

    _effectiveLinesetStatOrangeRed(ls) {
        const hybrid = ls.hybrid || false;
        if (hybrid) {
            return {
                orangeThreshold: this.settings.hybridOrangeThreshold,
                redThreshold: this.settings.hybridRedThreshold
            };
        }
        return {
            orangeThreshold: this._validLinesetStatThreshold(ls.standardOrangeThreshold) ?? this.settings.standardOrangeThreshold,
            redThreshold: this._validLinesetStatThreshold(ls.standardRedThreshold) ?? this.settings.standardRedThreshold
        };
    }

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

        const refLs = this._getActiveReferenceLinesetForNewLineset(canopy);
        this._setLinesetStatThresholdInputPlaceholders(LINESET_MODAL_STAT_THRESHOLD_INPUT_IDS);
        this._prefillLinesetStatThresholdForm(null, refLs, LINESET_MODAL_STAT_THRESHOLD_INPUT_IDS);
        this._syncLinesetModalStatThresholdSectionVisibility();

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
        document.getElementById('linesetPreviousJumps').value = lineset.previousJumps ?? 0;

        const priorLs = this._getPriorLinesetByNumber(canopy, linesetNumber);
        this._setLinesetStatThresholdInputPlaceholders(LINESET_MODAL_STAT_THRESHOLD_INPUT_IDS);
        this._prefillLinesetStatThresholdForm(lineset, priorLs, LINESET_MODAL_STAT_THRESHOLD_INPUT_IDS);
        this._syncLinesetModalStatThresholdSectionVisibility();

        document.getElementById('linesetModal').style.display = 'block';
    }

    saveLineset() {
        const canopyId = document.getElementById('linesetCanopyId').value;
        const editNumber = document.getElementById('linesetEditNumber').value;
        const linesetNumber = parseInt(document.getElementById('linesetNumber').value) || 1;
        const hybrid = document.getElementById('linesetHybridCheck').checked;
        const prevRaw = String(document.getElementById('linesetPreviousJumps').value).trim();
        const prevParsed = parseInt(prevRaw, 10);
        const previousJumps = Number.isFinite(prevParsed) ? prevParsed : 0;
        
        const canopy = this.canopies.find(c => c.id === canopyId);
        if (!canopy) return;
        if (!Array.isArray(canopy.linesets)) canopy.linesets = [];
        
        if (editNumber) {
            // Edit existing lineset
            const lineset = canopy.linesets.find(ls => ls.number === parseInt(editNumber));
            if (lineset) {
                lineset.hybrid = hybrid;
                lineset.previousJumps = previousJumps;
                if (hybrid) {
                    this._purgeLinesetStatThresholdOverrides(lineset);
                } else {
                    this._applyLinesetStatThresholdFormToLineset(lineset, LINESET_MODAL_STAT_THRESHOLD_INPUT_IDS);
                }
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
            const newLs = {
                number: linesetNumber,
                hybrid: hybrid,
                previousJumps: previousJumps,
                jumpCount: 0,
                archived: false
            };
            if (hybrid) {
                this._purgeLinesetStatThresholdOverrides(newLs);
            } else {
                this._applyLinesetStatThresholdFormToLineset(newLs, LINESET_MODAL_STAT_THRESHOLD_INPUT_IDS);
            }
            canopy.linesets.push(newLs);
            canopy.linesets.sort((a, b) => a.number - b.number);
        }
        
        this.saveComponentsToLocalStorage();
        this.updateEquipmentOptions();
        this.renderEquipmentView();
        this.closeLinesetModal();
        this.showMessage('Lineset saved successfully!', 'success');
    }

    /**
     * If no non-archived lineset remains, mark the highest-numbered one as active
     * so jump logging still has a current lineset.
     */
    _activateHighestLinesetIfNoneActive(canopy) {
        if (!canopy || !Array.isArray(canopy.linesets) || canopy.linesets.length === 0) return;
        if (canopy.linesets.some(ls => !ls.archived)) return;
        const highest = canopy.linesets.reduce((a, b) => (a.number >= b.number ? a : b));
        highest.archived = false;
    }

    /**
     * Remove a lineset that has no jumps recorded in this logbook for that canopy/lineset.
     * Ensures at least one lineset remains (default #1).
     */
    deleteLineset(canopyId, linesetNumber) {
        const canopy = this.canopies.find(c => c.id === canopyId);
        if (!canopy || !Array.isArray(canopy.linesets)) return;
        const lineset = canopy.linesets.find(ls => ls.number === linesetNumber);
        if (!lineset) return;

        const logged = this.jumps.filter(j =>
            j.equipment === canopyId && j.linesetNumber === linesetNumber
        ).length;
        if (logged > 0) {
            this.showMessage('Cannot delete a lineset that has jumps logged in this app. Keep it for history, or delete those jumps first.', 'error');
            return;
        }

        const preApp = lineset.previousJumps ?? 0;
        let msg = `Delete lineset #${linesetNumber} for ${canopy.name}? This cannot be undone.`;
        if (preApp > 0) {
            msg = `Delete lineset #${linesetNumber}? It has ${preApp} pre-app jump(s) recorded (none in this logbook). This cannot be undone.`;
        } else if (preApp < 0) {
            msg = `Delete lineset #${linesetNumber}? It has a negative pre-app adjustment (${preApp}). This cannot be undone.`;
        }
        if (!confirm(msg)) return;

        const idx = canopy.linesets.findIndex(ls => ls.number === linesetNumber);
        if (idx === -1) return;
        canopy.linesets.splice(idx, 1);

        if (canopy.linesets.length === 0) {
            canopy.linesets.push({
                number: 1,
                hybrid: false,
                previousJumps: 0,
                jumpCount: 0,
                archived: false
            });
        } else {
            this._activateHighestLinesetIfNoneActive(canopy);
        }

        this.saveComponentsToLocalStorage();
        this.updateEquipmentOptions();
        this.renderEquipmentView();
        this.updateLinesetHint();
        if (navigator.onLine && window.SheetsAPI) window.SheetsAPI.syncEquipmentToSheet();
        this.showMessage(`Lineset #${linesetNumber} deleted.`, 'success');
    }

    closeLinesetModal() {
        document.getElementById('linesetModal').style.display = 'none';
    }

    _singularize(plural) {
        const map = { harnesses: 'harness', canopies: 'canopy', locations: 'location' };
        return map[plural] || plural.slice(0, -1);
    }

    /** Harness id stored on canopy/jump, or '' if none. */
    _normalizeHarnessId(v) {
        const s = (v == null ? '' : String(v)).trim();
        return s;
    }

    /** Harness id for a canopy at save/jump time, or ''. */
    _harnessIdForCanopyId(canopyId) {
        if (!canopyId) return '';
        const c = this.canopies.find(x => x.id === canopyId);
        return this._normalizeHarnessId(c?.harnessId);
    }

    /**
     * Snapshot harness id to store on a jump from the canopy's current assignment.
     * Returns undefined if no harness (omit property for cleaner legacy rows).
     */
    _harnessIdSnapshotForJump(canopyId) {
        const h = this._harnessIdForCanopyId(canopyId);
        return h || undefined;
    }

    /** Populate #canopyHarnessSelect; selectedId is current canopy.harnessId. */
    _fillCanopyHarnessSelect(selectedId) {
        const sel = document.getElementById('canopyHarnessSelect');
        if (!sel) return;
        const want = this._normalizeHarnessId(selectedId);
        sel.innerHTML = '<option value="">— None —</option>';
        const seen = new Set(['']);
        for (const h of this.harnesses) {
            if (!h?.id) continue;
            if (h.archived && h.id !== want) continue;
            const opt = document.createElement('option');
            opt.value = h.id;
            opt.textContent = h.name + (h.archived ? ' (Archived)' : '');
            sel.appendChild(opt);
            seen.add(h.id);
        }
        if (want && !seen.has(want)) {
            const opt = document.createElement('option');
            opt.value = want;
            opt.textContent = want + ' (missing)';
            sel.appendChild(opt);
        }
        sel.value = want && seen.has(want) ? want : (want && !seen.has(want) ? want : '');
    }

    addComponent(type) {
        document.getElementById('componentForm').reset();
        document.getElementById('componentId').value = '';
        document.getElementById('componentType').value = type;
        document.getElementById('componentNotes').value = '';
        document.getElementById('componentModalTitle').textContent = `Add ${type.charAt(0).toUpperCase() + type.slice(1)}`;
        const harnessPre = document.getElementById('harnessPreAppSection');
        const canopyPre = document.getElementById('canopyPreAppSection');
        const canopyHarness = document.getElementById('canopyHarnessSection');
        const canopyBackfill = document.getElementById('canopyHarnessBackfillWrap');
        if (harnessPre) harnessPre.style.display = type === 'harness' ? 'block' : 'none';
        if (canopyPre) canopyPre.style.display = type === 'canopy' ? 'block' : 'none';
        if (canopyHarness) canopyHarness.style.display = type === 'canopy' ? 'block' : 'none';
        if (canopyBackfill) canopyBackfill.style.display = 'none';
        if (type === 'harness') {
            const inp = document.getElementById('harnessPreviousJumps');
            if (inp) inp.value = '0';
        }
        if (type === 'canopy') {
            const inp = document.getElementById('canopyPreviousJumps');
            if (inp) inp.value = '0';
            this._fillCanopyHarnessSelect('');
        }
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
            this._setLinesetStatThresholdInputPlaceholders(NEW_CANOPY_STAT_THRESHOLD_INPUT_IDS);
            this._prefillLinesetStatThresholdForm(null, null, NEW_CANOPY_STAT_THRESHOLD_INPUT_IDS);
            this._syncNewCanopyStatThresholdSectionVisibility();
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
        let jumpsUpdatedForLocationRename = 0;
        let jumpsUpdatedForHarnessBackfill = 0;

        if (id) {
            // Edit existing
            const component = collection.find(c => c.id === id);
            if (component) {
                const prevLocationName = type === 'location' ? component.name : null;
                const nameChanged = type === 'location' && (component.name || '').trim() !== name;
                component.name = name;
                component.notes = notes;
                if (type === 'harness') {
                    const preRaw = String(document.getElementById('harnessPreviousJumps')?.value ?? '').trim();
                    const preParsed = Number(preRaw);
                    component.previousJumps = Number.isFinite(preParsed) ? preParsed : 0;
                }
                if (type === 'canopy') {
                    const preRaw = String(document.getElementById('canopyPreviousJumps')?.value ?? '').trim();
                    const preParsed = Number(preRaw);
                    component.previousJumps = Number.isFinite(preParsed) ? preParsed : 0;
                    const prevH = this._normalizeHarnessId(component.harnessId);
                    const newH = this._normalizeHarnessId(document.getElementById('canopyHarnessSelect')?.value);
                    if (newH) component.harnessId = newH;
                    else delete component.harnessId;
                    const backfill = !!(document.getElementById('canopyHarnessBackfillCheck')?.checked);
                    if (backfill && newH && !prevH) {
                        for (const j of this.jumps) {
                            if (j.equipment === component.id) {
                                j.harnessId = newH;
                                jumpsUpdatedForHarnessBackfill++;
                            }
                        }
                    }
                }
                if (type === 'location') {
                    if (manualLat !== null) {
                        // Manual coords override everything
                        component.lat = manualLat;
                        component.lng = manualLng;
                    } else {
                        if (nameChanged) { component.lat = null; component.lng = null; }
                        if (component.lat == null) this.geocodeLocation(component);
                    }
                    if (nameChanged && prevLocationName != null) {
                        const oldT = (prevLocationName || '').trim();
                        if (oldT && oldT !== name) {
                            for (const j of this.jumps) {
                                if ((j.location || '').trim() === oldT) {
                                    j.location = name;
                                    jumpsUpdatedForLocationRename++;
                                }
                            }
                        }
                    }
                }
            }
        } else {
            // Add new
            const newId = type + '_' + Date.now();
            const newComponent = { id: newId, name: name, notes: notes };
            if (type === 'harness') {
                const preRaw = String(document.getElementById('harnessPreviousJumps')?.value ?? '').trim();
                const preParsed = Number(preRaw);
                newComponent.previousJumps = Number.isFinite(preParsed) ? preParsed : 0;
            }
            if (type === 'canopy') {
                const preRaw = String(document.getElementById('canopyPreviousJumps')?.value ?? '').trim();
                const preParsed = Number(preRaw);
                newComponent.previousJumps = Number.isFinite(preParsed) ? preParsed : 0;
            }
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
                const ls1 = { number: 1, hybrid: hybrid, previousJumps: previousJumps, jumpCount: 0, archived: false };
                if (hybrid) {
                    this._purgeLinesetStatThresholdOverrides(ls1);
                } else {
                    this._applyLinesetStatThresholdFormToLineset(ls1, NEW_CANOPY_STAT_THRESHOLD_INPUT_IDS);
                }
                canopy.linesets = [ls1];
            }
            const hsel = this._normalizeHarnessId(document.getElementById('canopyHarnessSelect')?.value);
            if (hsel) canopy.harnessId = hsel;
            else delete canopy.harnessId;
        }
        
        this.saveComponentsToLocalStorage();
        this.renderEquipmentView();
        if (type === 'canopy') {
            this.updateEquipmentOptions();
            this.updateLinesetHint();
        }
        this.closeComponentModal();
        // Refresh autocomplete if a location was saved
        if (type === 'location') this.updateLocationDatalist();
        if (navigator.onLine && window.SheetsAPI) window.SheetsAPI.syncEquipmentToSheet();
        if (jumpsUpdatedForLocationRename > 0 || jumpsUpdatedForHarnessBackfill > 0) {
            this.saveToLocalStorage();
            this.updateStats();
            if (this.currentView === 'jumps') this.renderJumpsList();
            if (navigator.onLine && window.SheetsAPI?.initialized) {
                window.SheetsAPI.pushAllWithGuard();
            }
        }
        let savedMsg = `${type.charAt(0).toUpperCase() + type.slice(1)} saved successfully!`;
        if (jumpsUpdatedForLocationRename > 0) {
            savedMsg += ` ${jumpsUpdatedForLocationRename} jump${jumpsUpdatedForLocationRename === 1 ? '' : 's'} updated to the new location name.`;
        }
        if (jumpsUpdatedForHarnessBackfill > 0) {
            savedMsg += ` Harness applied to ${jumpsUpdatedForHarnessBackfill} existing jump${jumpsUpdatedForHarnessBackfill === 1 ? '' : 's'}.`;
        }
        this.showMessage(savedMsg, 'success');
    }
    
    updateLocationDatalist() {
        // No-op: replaced by custom autocomplete dropdown
    }

    setupLocationAutocomplete() {
        const pairs = [
            { input: document.getElementById('location'), dropdown: document.getElementById('locationDropdown') },
            { input: document.getElementById('editJumpLocation'), dropdown: document.getElementById('editJumpLocationDropdown') }
        ];
        pairs.forEach(({ input, dropdown }) => {
            if (input && dropdown) this._bindOneLocationAutocomplete(input, dropdown);
        });
    }

    /**
     * Location field + dropdown: same behavior for main jump form and edit-jump modal.
     */
    _bindOneLocationAutocomplete(input, dropdown) {
        if (input.dataset.locationAutocompleteBound === '1') return;
        input.dataset.locationAutocompleteBound = '1';

        const wrap = input.closest('.location-autocomplete');
        if (!wrap) {
            console.warn('[logbook] location input missing .location-autocomplete wrapper', input.id);
        }

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
            if (input.id === 'editJumpLocation') this.syncEditJumpModalBulkOptions();
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

        document.addEventListener('click', (e) => {
            if (wrap && wrap.contains(e.target)) return;
            dropdown.classList.remove('open');
        });
    }

    /**
     * Collapse whitespace and newlines for matching / single-line previews of jump notes.
     */
    _notesCollapseForMatch(note) {
        return String(note || '').replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim();
    }

    /**
     * Distinct non-empty notes from logged jumps, most recently logged first.
     */
    _distinctJumpNotesRecentFirst() {
        const seen = new Set();
        const out = [];
        const jumps = Array.isArray(this.jumps) ? this.jumps : [];
        for (let i = jumps.length - 1; i >= 0; i--) {
            const raw = jumps[i]?.notes;
            const n = typeof raw === 'string' ? raw.trim() : '';
            if (!n || seen.has(n)) continue;
            seen.add(n);
            out.push(n);
        }
        return out;
    }

    setupNotesAutocomplete() {
        const pairs = [
            ['notes', 'notesDropdown'],
            ['editJumpNotes', 'editJumpNotesDropdown'],
            ['jumpNoteContent', 'jumpNoteContentDropdown']
        ];
        for (const [inputId, dropdownId] of pairs) {
            const input = document.getElementById(inputId);
            const dropdown = document.getElementById(dropdownId);
            if (input && dropdown) this._bindJumpNotesAutocomplete(input, dropdown);
        }
    }

    /**
     * Textarea + dropdown: suggest full note text from previous jump notes (log form, edit jump, note modal).
     */
    _bindJumpNotesAutocomplete(input, dropdown) {
        if (input.dataset.notesAutocompleteBound === '1') return;
        input.dataset.notesAutocompleteBound = '1';

        const wrap = input.closest('.notes-autocomplete');
        if (!wrap) {
            console.warn('[logbook] notes field missing .notes-autocomplete wrapper', input.id);
        }

        let activeIndex = -1;
        let lastMatches = [];

        const previewHtmlForNote = (fullNote, queryTrimmed) => {
            const oneLine = this._notesCollapseForMatch(fullNote);
            const maxLen = 100;
            const qLower = queryTrimmed.toLowerCase();
            let start = 0;
            let end = Math.min(oneLine.length, maxLen);
            if (oneLine.length > maxLen) {
                if (queryTrimmed) {
                    const idx = oneLine.toLowerCase().indexOf(qLower);
                    if (idx >= 0) {
                        const half = Math.floor(maxLen / 2);
                        start = Math.max(0, idx - half);
                        end = Math.min(oneLine.length, start + maxLen);
                        if (end - start < maxLen) start = Math.max(0, end - maxLen);
                    }
                } else {
                    end = maxLen;
                }
            }
            const slice = oneLine.slice(start, end);
            const prefix = start > 0 ? '\u2026' : '';
            const suffix = end < oneLine.length ? '\u2026' : '';
            const relIdx = queryTrimmed ? slice.toLowerCase().indexOf(qLower) : -1;
            const qLen = queryTrimmed.length;
            let inner;
            if (relIdx >= 0 && qLen > 0) {
                inner = this.escapeHtml(slice.slice(0, relIdx))
                    + `<span class="match">${this.escapeHtml(slice.slice(relIdx, relIdx + qLen))}</span>`
                    + this.escapeHtml(slice.slice(relIdx + qLen));
            } else {
                inner = this.escapeHtml(slice);
            }
            return this.escapeHtml(prefix) + inner + this.escapeHtml(suffix);
        };

        const showDropdown = () => {
            const q = input.value.trim();
            const qLower = q.toLowerCase();
            const distinct = this._distinctJumpNotesRecentFirst();
            const matches = !q
                ? distinct.slice(0, 25)
                : distinct.filter(n => n.toLowerCase().includes(qLower)).slice(0, 25);

            if (matches.length === 0) {
                dropdown.classList.remove('open');
                activeIndex = -1;
                lastMatches = [];
                return;
            }

            lastMatches = matches;
            dropdown.innerHTML = matches.map((fullNote, i) =>
                `<div class="autocomplete-option notes-autocomplete-option" role="option" data-index="${i}">${previewHtmlForNote(fullNote, q)}</div>`
            ).join('');

            activeIndex = -1;
            dropdown.classList.add('open');
        };

        const selectOption = (index) => {
            const text = lastMatches[index];
            if (text === undefined) return;
            input.value = text;
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
                selectOption(activeIndex);
            } else if (e.key === 'Escape') {
                dropdown.classList.remove('open');
                activeIndex = -1;
            }
        });

        dropdown.addEventListener('mousedown', (e) => {
            const option = e.target.closest('.autocomplete-option');
            if (!option) return;
            e.preventDefault();
            const idx = parseInt(option.dataset.index, 10);
            if (!Number.isNaN(idx)) selectOption(idx);
        });

        document.addEventListener('click', (e) => {
            if (wrap && wrap.contains(e.target)) return;
            dropdown.classList.remove('open');
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
            if (type === 'canopies') {
                this.updateEquipmentOptions();
                this.updateLinesetHint();
            }
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
            const harnessPre = document.getElementById('harnessPreAppSection');
            const canopyPre = document.getElementById('canopyPreAppSection');
            const canopyHarness = document.getElementById('canopyHarnessSection');
            const canopyBackfill = document.getElementById('canopyHarnessBackfillWrap');
            const isHarness = singular === 'harness';
            const isCanopy = singular === 'canopy';
            if (harnessPre) harnessPre.style.display = isHarness ? 'block' : 'none';
            if (canopyPre) canopyPre.style.display = isCanopy ? 'block' : 'none';
            if (canopyHarness) canopyHarness.style.display = isCanopy ? 'block' : 'none';
            if (canopyBackfill) {
                canopyBackfill.style.display = (isCanopy && !this._normalizeHarnessId(component.harnessId)) ? 'block' : 'none';
                const bf = document.getElementById('canopyHarnessBackfillCheck');
                if (bf) bf.checked = false;
            }
            if (isHarness) {
                const inp = document.getElementById('harnessPreviousJumps');
                if (inp) inp.value = String(component.previousJumps ?? 0);
            }
            if (isCanopy) {
                const inp = document.getElementById('canopyPreviousJumps');
                if (inp) inp.value = String(component.previousJumps ?? 0);
                this._fillCanopyHarnessSelect(component.harnessId);
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
            if (type === 'harnesses') {
                const usedByCanopy = this.canopies.some(c => this._normalizeHarnessId(c.harnessId) === id);
                if (usedByCanopy) {
                    this.showMessage(`Cannot delete ${typeSingular} that is assigned to a canopy. Clear the harness on the canopy first.`, 'error');
                    return;
                }
            }
            
            const collection = this[type];
            const index = collection.findIndex(c => c.id === id);
            if (index !== -1) {
                collection.splice(index, 1);
                this.saveComponentsToLocalStorage();
                this.renderEquipmentView();
                if (type === 'canopies') {
                    this.updateEquipmentOptions();
                    this.updateLinesetHint();
                }
                if (type === 'locations') this.updateLocationDatalist();
                if (navigator.onLine && window.SheetsAPI) window.SheetsAPI.syncEquipmentToSheet();
                this.showMessage(`${typeSingular.charAt(0).toUpperCase() + typeSingular.slice(1)} deleted successfully!`, 'success');
            }
        }
    }
    
    closeComponentModal() {
        document.getElementById('componentModal').style.display = 'none';
    }

    _updateFlysightAvgLabel() {
        const avgValue = document.getElementById('flysightAvgPointsValue');
        const avgDuration = document.getElementById('flysightAvgDuration');
        if (!avgValue) return;

        const n = this.flysightAvgPoints;
        avgValue.textContent = String(n);

        if (!avgDuration) return;

        if (n <= 1) {
            avgDuration.textContent = '';
            return;
        }

        let intervalSec = typeof Flysight !== 'undefined'
            ? Flysight.DEFAULT_SAMPLE_INTERVAL_SEC
            : 0.1;

        if (typeof Flysight !== 'undefined' && this.flysightFiles.length) {
            const allPoints = [];
            for (const file of this.flysightFiles) {
                const parsed = Flysight.parseFlysightCsv(file.text);
                if (parsed.points?.length) allPoints.push(...parsed.points);
            }
            if (allPoints.length >= 2) {
                intervalSec = Flysight.medianSampleIntervalSec(allPoints);
            }
        }

        const durationSec = typeof Flysight !== 'undefined'
            ? Flysight.averagingWindowDurationSec(n, intervalSec)
            : (n - 1) * intervalSec;
        avgDuration.textContent = durationSec != null
            ? ` (${typeof Flysight !== 'undefined' ? Flysight.formatDurationSec(durationSec) : `${durationSec.toFixed(2)}s`})`
            : '';
    }

    _flysightHeightLimits() {
        const minM = (typeof Flysight !== 'undefined' && Number.isFinite(Flysight.MIN_MAX_HEIGHT_M))
            ? Flysight.MIN_MAX_HEIGHT_M
            : 1;
        const defaultM = (typeof Flysight !== 'undefined' && Number.isFinite(Flysight.DEFAULT_MAX_HEIGHT_M))
            ? Flysight.DEFAULT_MAX_HEIGHT_M
            : 500;
        const absMaxM = (typeof Flysight !== 'undefined' && Number.isFinite(Flysight.MAX_MAX_HEIGHT_M))
            ? Flysight.MAX_MAX_HEIGHT_M
            : 10000;
        return { minM, defaultM, absMaxM };
    }

    _parseFlysightMaxHeightSliderMax(value) {
        const { minM, defaultM, absMaxM } = this._flysightHeightLimits();
        const n = parseInt(value, 10);
        if (!Number.isFinite(n)) return defaultM;
        return Math.min(absMaxM, Math.max(minM, n));
    }

    _parseFlysightMaxHeight(value) {
        const { minM } = this._flysightHeightLimits();
        const n = parseInt(value, 10);
        const fallback = this.flysightMaxHeightSliderMaxM;
        if (!Number.isFinite(n)) return fallback;
        return Math.min(this.flysightMaxHeightSliderMaxM, Math.max(minM, n));
    }

    _updateFlysightMaxHeightLabel() {
        const maxHeightValue = document.getElementById('flysightMaxHeightValue');
        if (maxHeightValue) maxHeightValue.textContent = String(this.flysightMaxHeightM);
    }

    _syncFlysightMaxHeightSlider() {
        const slider = document.getElementById('flysightMaxHeight');
        if (slider) {
            slider.max = String(this.flysightMaxHeightSliderMaxM);
            slider.value = String(this.flysightMaxHeightM);
        }
        this._updateFlysightMaxHeightLabel();
    }

    _applyFlysightMaxHeightSliderMax(value) {
        this.flysightMaxHeightSliderMaxM = this._parseFlysightMaxHeightSliderMax(value);
        localStorage.setItem('flysight-max-height-slider-max', String(this.flysightMaxHeightSliderMaxM));
        const clamped = this._parseFlysightMaxHeight(this.flysightMaxHeightM);
        if (clamped !== this.flysightMaxHeightM) {
            this.flysightMaxHeightM = clamped;
            localStorage.setItem('flysight-max-height', String(this.flysightMaxHeightM));
        }
        this._syncFlysightMaxHeightSlider();
    }

    _updateFlysightSpeedModeButtons() {
        const verticalBtn = document.getElementById('flysightSpeedVertical');
        const totalBtn = document.getElementById('flysightSpeedTotal');
        const bothBtn = document.getElementById('flysightSpeedBoth');
        if (!verticalBtn || !totalBtn || !bothBtn) return;

        const metric = this.flysightSpeedMetric;
        verticalBtn.classList.toggle('is-active', metric === 'vertical');
        verticalBtn.setAttribute('aria-pressed', String(metric === 'vertical'));
        totalBtn.classList.toggle('is-active', metric === 'total');
        totalBtn.setAttribute('aria-pressed', String(metric === 'total'));
        bothBtn.classList.toggle('is-active', metric === 'both');
        bothBtn.setAttribute('aria-pressed', String(metric === 'both'));
    }

    _setFlysightSpeedMetric(metric) {
        this.flysightSpeedMetric = ['vertical', 'total', 'both'].includes(metric) ? metric : 'vertical';
        localStorage.setItem('flysight-speed-metric', this.flysightSpeedMetric);
        this._updateFlysightSpeedModeButtons();
        if (this.flysightFiles.length) this.renderFlysightView();
    }

    _parseFlysightCursorBDiveAngle(value) {
        const fallback = (typeof Flysight !== 'undefined' && Number.isFinite(Flysight.CURSOR_B_DIVE_ANGLE_DEG))
            ? Flysight.CURSOR_B_DIVE_ANGLE_DEG
            : 5.5;
        const n = parseFloat(value);
        if (!Number.isFinite(n)) return fallback;
        return Math.min(90, Math.max(0, n));
    }

    _parseFlysightCursorBAltTicks(value) {
        const fallback = (typeof Flysight !== 'undefined' && Number.isFinite(Flysight.CURSOR_B_ALT_TICKS))
            ? Flysight.CURSOR_B_ALT_TICKS
            : 2;
        const n = parseInt(value, 10);
        if (!Number.isFinite(n)) return fallback;
        return Math.min(100, Math.max(0, n));
    }

    _syncFlysightSettingsForm() {
        const avgSlider = document.getElementById('flysightAvgPoints');
        const sliderMaxInput = document.getElementById('flysightMaxHeightSliderMax');
        const angleInput = document.getElementById('flysightCursorBDiveAngle');
        const ticksInput = document.getElementById('flysightCursorBAltTicks');
        if (avgSlider) avgSlider.value = String(this.flysightAvgPoints);
        if (sliderMaxInput) {
            const { minM, absMaxM } = this._flysightHeightLimits();
            sliderMaxInput.min = String(minM);
            sliderMaxInput.max = String(absMaxM);
            sliderMaxInput.value = String(this.flysightMaxHeightSliderMaxM);
        }
        if (angleInput) angleInput.value = String(this.flysightCursorBDiveAngleDeg);
        if (ticksInput) ticksInput.value = String(this.flysightCursorBAltTicks);
        this._updateFlysightAvgLabel();
        this._syncFlysightMaxHeightSlider();
    }

    openFlysightSettingsModal() {
        this._syncFlysightSettingsForm();
        const modal = document.getElementById('flysightSettingsModal');
        if (modal) modal.style.display = 'block';
    }

    closeFlysightSettingsModal() {
        const modal = document.getElementById('flysightSettingsModal');
        if (modal) modal.style.display = 'none';
    }

    _restoreFlysightSettingsDefaults() {
        const defaultAvg = 3;
        const defaultAngle = (typeof Flysight !== 'undefined' && Number.isFinite(Flysight.CURSOR_B_DIVE_ANGLE_DEG))
            ? Flysight.CURSOR_B_DIVE_ANGLE_DEG
            : 5.5;
        const defaultTicks = (typeof Flysight !== 'undefined' && Number.isFinite(Flysight.CURSOR_B_ALT_TICKS))
            ? Flysight.CURSOR_B_ALT_TICKS
            : 2;
        this.flysightAvgPoints = defaultAvg;
        this.flysightCursorBDiveAngleDeg = defaultAngle;
        this.flysightCursorBAltTicks = defaultTicks;
        this._applyFlysightMaxHeightSliderMax(this._flysightHeightLimits().defaultM);
        localStorage.setItem('flysight-avg-points', String(this.flysightAvgPoints));
        localStorage.setItem('flysight-cursor-b-dive-angle', String(this.flysightCursorBDiveAngleDeg));
        localStorage.setItem('flysight-cursor-b-alt-ticks', String(this.flysightCursorBAltTicks));
        this._syncFlysightSettingsForm();
        if (this.flysightFiles.length) this.renderFlysightView();
        this._reapplyFlysightGraphDefaultCursors();
        this._rebuildFlysightGraphIfOpen();
    }

    _bindFlysightSettingsInputs() {
        const angleInput = document.getElementById('flysightCursorBDiveAngle');
        if (angleInput && angleInput.dataset.bound !== '1') {
            angleInput.dataset.bound = '1';
            const applyAngle = () => {
                this.flysightCursorBDiveAngleDeg = this._parseFlysightCursorBDiveAngle(angleInput.value);
                localStorage.setItem('flysight-cursor-b-dive-angle', String(this.flysightCursorBDiveAngleDeg));
                if (this.flysightFiles.length) this.renderFlysightView();
                this._reapplyFlysightGraphDefaultCursors();
            };
            angleInput.addEventListener('input', applyAngle);
            angleInput.addEventListener('change', applyAngle);
        }
        const sliderMaxInput = document.getElementById('flysightMaxHeightSliderMax');
        if (sliderMaxInput && sliderMaxInput.dataset.bound !== '1') {
            sliderMaxInput.dataset.bound = '1';
            const applySliderMax = () => {
                this._applyFlysightMaxHeightSliderMax(sliderMaxInput.value);
                if (this.flysightFiles.length) this.renderFlysightView();
            };
            sliderMaxInput.addEventListener('input', applySliderMax);
            sliderMaxInput.addEventListener('change', applySliderMax);
        }
        const ticksInput = document.getElementById('flysightCursorBAltTicks');
        if (ticksInput && ticksInput.dataset.bound !== '1') {
            ticksInput.dataset.bound = '1';
            const applyTicks = () => {
                this.flysightCursorBAltTicks = this._parseFlysightCursorBAltTicks(ticksInput.value);
                localStorage.setItem('flysight-cursor-b-alt-ticks', String(this.flysightCursorBAltTicks));
                if (this.flysightFiles.length) this.renderFlysightView();
                this._reapplyFlysightGraphDefaultCursors();
            };
            ticksInput.addEventListener('input', applyTicks);
            ticksInput.addEventListener('change', applyTicks);
        }
        document.getElementById('flysightSettingsBtn')?.addEventListener('click', () => {
            this.openFlysightSettingsModal();
        });
        document.getElementById('flysightSettingsModalClose')?.addEventListener('click', () => {
            this.closeFlysightSettingsModal();
        });
        document.getElementById('flysightSettingsRestoreBtn')?.addEventListener('click', () => {
            this._restoreFlysightSettingsDefaults();
        });
    }

    _reapplyFlysightGraphDefaultCursors() {
        const g = this._flysightGraph;
        const modal = document.getElementById('flysightGraphModal');
        if (!g?.samples?.length || typeof Flysight === 'undefined') return;
        if (!modal || modal.style.display !== 'flex') return;
        const cursors = Flysight.defaultSwoopCursorIndices(
            g.samples,
            this.flysightCursorBDiveAngleDeg,
            this.flysightCursorBAltTicks
        );
        g.idxA = cursors.idxA;
        g.idxB = cursors.idxB;
        this._drawFlysightGraph();
    }

    _rebuildFlysightGraphIfOpen() {
        const g = this._flysightGraph;
        const modal = document.getElementById('flysightGraphModal');
        if (!g?.fileId || !modal || modal.style.display !== 'flex') return;
        this.openFlysightGraphModal(g.fileId, g.mode);
    }

    _bindFlysightEvents() {
        const dropZone = document.getElementById('flysightDropZone');
        const dirInput = document.getElementById('flysightDirInput');
        const avgSlider = document.getElementById('flysightAvgPoints');
        const maxHeightSlider = document.getElementById('flysightMaxHeight');
        const speedVerticalBtn = document.getElementById('flysightSpeedVertical');
        const speedTotalBtn = document.getElementById('flysightSpeedTotal');
        const speedBothBtn = document.getElementById('flysightSpeedBoth');
        if (!dropZone || !dirInput || !avgSlider || !maxHeightSlider || !speedVerticalBtn || !speedTotalBtn || !speedBothBtn) return;

        avgSlider.value = String(this.flysightAvgPoints);
        this._syncFlysightMaxHeightSlider();
        this._updateFlysightAvgLabel();
        this._updateFlysightSpeedModeButtons();
        this._syncFlysightSettingsForm();
        this._bindFlysightSettingsInputs();

        const openPicker = () => this._openFlysightFolderPicker();
        dropZone.addEventListener('click', (e) => {
            if (e.target === dirInput) return;
            openPicker();
        });
        dropZone.addEventListener('keydown', (e) => {
            if (e.target === dirInput) return;
            if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                openPicker();
            }
        });

        dirInput.addEventListener('change', () => {
            if (!dirInput.files?.length) return;
            const files = Array.from(dirInput.files);
            dirInput.value = '';
            this._onFlysightDirectoryInputFiles(files);
        });

        document.getElementById('flysightBrowserModalClose')?.addEventListener('click', () => {
            this.closeFlysightBrowserModal();
        });
        document.getElementById('flysightBrowserCancelBtn')?.addEventListener('click', () => {
            this.closeFlysightBrowserModal();
        });
        document.getElementById('flysightBrowserChangeFolderBtn')?.addEventListener('click', (e) => {
            e.preventDefault();
            this._openFlysightFolderPicker();
        });
        document.getElementById('flysightBrowserImportBtn')?.addEventListener('click', (e) => {
            e.preventDefault();
            this._importFlysightBrowserSelection();
        });
        document.getElementById('flysightBrowserCrumbs')?.addEventListener('click', (e) => {
            const btn = e.target instanceof Element ? e.target.closest('[data-crumb-index]') : null;
            if (!btn) return;
            const idx = parseInt(btn.dataset.crumbIndex, 10);
            if (Number.isFinite(idx)) this._goToFlysightBrowserCrumb(idx);
        });
        document.getElementById('flysightBrowserList')?.addEventListener('click', (e) => {
            const dirBtn = e.target instanceof Element ? e.target.closest('.flysight-browser-dir-enter') : null;
            if (!dirBtn) return;
            e.preventDefault();
            this._enterFlysightBrowserDir(dirBtn.dataset.dirIndex);
        });
        document.getElementById('flysightBrowserList')?.addEventListener('change', (e) => {
            const input = e.target instanceof HTMLInputElement ? e.target : null;
            if (!input || !this._flysightBrowser) return;
            if (input.dataset.dirKey) {
                this._toggleFlysightBrowserSelection(
                    input.dataset.dirKey,
                    input.checked,
                    'dir',
                    Number(input.dataset.dirIndex)
                );
                return;
            }
            if (input.dataset.fileKey) {
                this._toggleFlysightBrowserSelection(
                    input.dataset.fileKey,
                    input.checked,
                    'file',
                    Number(input.dataset.fileIndex)
                );
            }
        });

        ['dragenter', 'dragover'].forEach(evt => {
            dropZone.addEventListener(evt, (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropZone.classList.add('is-dragover');
            });
        });
        ['dragleave', 'drop'].forEach(evt => {
            dropZone.addEventListener(evt, (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropZone.classList.remove('is-dragover');
            });
        });
        dropZone.addEventListener('drop', (e) => {
            this._addFlysightFilesFromDrop(e.dataTransfer);
        });

        avgSlider.addEventListener('input', () => {
            this.flysightAvgPoints = Math.min(20, Math.max(1, parseInt(avgSlider.value, 10) || 3));
            this._updateFlysightAvgLabel();
            localStorage.setItem('flysight-avg-points', String(this.flysightAvgPoints));
            if (this.flysightFiles.length) this.renderFlysightView();
            this._rebuildFlysightGraphIfOpen();
        });

        maxHeightSlider.addEventListener('input', () => {
            this.flysightMaxHeightM = this._parseFlysightMaxHeight(maxHeightSlider.value);
            this._updateFlysightMaxHeightLabel();
            localStorage.setItem('flysight-max-height', String(this.flysightMaxHeightM));
            if (this.flysightFiles.length) this.renderFlysightView();
        });

        speedVerticalBtn.addEventListener('click', () => this._setFlysightSpeedMetric('vertical'));
        speedTotalBtn.addEventListener('click', () => this._setFlysightSpeedMetric('total'));
        speedBothBtn.addEventListener('click', () => this._setFlysightSpeedMetric('both'));

        const results = document.getElementById('flysightResults');
        results?.addEventListener('click', (e) => {
            const el = e.target instanceof Element ? e.target : e.target?.parentElement;
            const graphBtn = el?.closest?.('.flysight-result-graph');
            if (graphBtn) {
                e.preventDefault();
                this.openFlysightGraphModal(graphBtn.dataset.flysightId, graphBtn.dataset.flysightGraph);
            }
        });

        document.getElementById('flysightGraphModalClose')?.addEventListener('click', () => {
            this.closeFlysightGraphModal();
        });

        const graphRoot = document.getElementById('flysightGraphRoot');
        if (graphRoot) {
            graphRoot.addEventListener('pointerdown', (e) => this._onFlysightGraphPointerDown(e));
            graphRoot.addEventListener('pointermove', (e) => this._onFlysightGraphPointerMove(e));
            graphRoot.addEventListener('pointerup', (e) => this._onFlysightGraphPointerUp(e));
            graphRoot.addEventListener('pointercancel', (e) => this._onFlysightGraphPointerUp(e));
            graphRoot.addEventListener('wheel', (e) => this._onFlysightGraphWheel(e), { passive: false });
            graphRoot.addEventListener('dblclick', (e) => {
                e.preventDefault();
                this._resetFlysightGraphZoom();
            });
        }
        document.getElementById('flysightGraphResetZoom')?.addEventListener('click', () => {
            this._resetFlysightGraphZoom();
        });
        document.getElementById('flysightGraphPrev')?.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            this.navigateFlysightGraph(-1);
        });
        document.getElementById('flysightGraphNext')?.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            this.navigateFlysightGraph(1);
        });

        const graphModal = document.getElementById('flysightGraphModal');
        graphModal?.addEventListener('touchmove', (e) => e.preventDefault(), { passive: false });

        const onViewport = () => this._onFlysightGraphViewportChange();
        window.addEventListener('resize', onViewport);
        window.addEventListener('orientationchange', onViewport);
        if (window.visualViewport) {
            window.visualViewport.addEventListener('resize', onViewport);
            window.visualViewport.addEventListener('scroll', onViewport);
        }
        if (graphRoot && typeof ResizeObserver !== 'undefined') {
            this._flysightGraphRo = new ResizeObserver(() => {
                const g = this._flysightGraph;
                const root = document.getElementById('flysightGraphRoot');
                if (!g || !root) return;
                const w = Math.round(root.clientWidth);
                const h = Math.round(root.clientHeight);
                if (w === g._layoutW && h === g._layoutH) return;
                this._drawFlysightGraph();
            });
            this._flysightGraphRo.observe(graphRoot);
        }
    }

    _flysightEmptyResultsHtml(message = 'No files analyzed yet.') {
        return `<p class="no-items flysight-empty">${message}</p>`;
    }

    _setFlysightEmptyResults(message) {
        if (this.flysightFiles.length) return;
        const container = document.getElementById('flysightResults');
        if (container) container.innerHTML = this._flysightEmptyResultsHtml(message);
    }

    _waitForFlysightUiPaint() {
        return new Promise(resolve => {
            requestAnimationFrame(() => requestAnimationFrame(resolve));
        });
    }

    _clickFlysightDirInput() {
        document.getElementById('flysightDirInput')?.click();
    }

    async _openFlysightFolderPicker() {
        if (typeof window.showDirectoryPicker === 'function') {
            try {
                const dirHandle = await window.showDirectoryPicker({
                    id: 'flysight-csv',
                    mode: 'read'
                });
                await this._openFlysightHandleBrowser(dirHandle);
                return;
            } catch (err) {
                if (err && (err.name === 'AbortError' || err.name === 'NotAllowedError')) return;
                console.error('[Flysight] Directory picker failed:', err);
                if (err && err.name === 'SecurityError') {
                    this._clickFlysightDirInput();
                    return;
                }
                this.showMessage('Could not read that folder.', 'error');
                return;
            }
        }
        this._clickFlysightDirInput();
    }

    _onFlysightDirectoryInputFiles(files) {
        if (typeof Flysight !== 'undefined' && Flysight.fileListHasRelativePaths(files)) {
            this._openFlysightVirtualBrowser(files);
            return;
        }
        this._addFlysightFiles(files);
    }

    _showFlysightBrowserModal() {
        const modal = document.getElementById('flysightBrowserModal');
        if (modal) modal.style.display = 'block';
    }

    closeFlysightBrowserModal() {
        const modal = document.getElementById('flysightBrowserModal');
        if (modal) modal.style.display = 'none';
        this._flysightBrowser = null;
        const list = document.getElementById('flysightBrowserList');
        if (list) list.innerHTML = '';
        const crumbs = document.getElementById('flysightBrowserCrumbs');
        if (crumbs) crumbs.innerHTML = '';
        const empty = document.getElementById('flysightBrowserEmpty');
        if (empty) empty.hidden = true;
        const hint = document.getElementById('flysightBrowserHint');
        if (hint) hint.hidden = true;
    }

    async _openFlysightHandleBrowser(dirHandle) {
        if (!dirHandle) return;
        this._flysightBrowser = {
            kind: 'handle',
            stack: [{ name: dirHandle.name || 'Folder', handle: dirHandle }],
            selected: new Map(),
            listing: { dirs: [], files: [] },
            csvCount: null,
            listGen: 0
        };
        this._showFlysightBrowserModal();
        await this._renderFlysightBrowser();
    }

    _openFlysightVirtualBrowser(fileList) {
        if (typeof Flysight === 'undefined') {
            this._addFlysightFiles(fileList);
            return;
        }
        const root = Flysight.buildVirtualTreeFromFileList(fileList);
        this._flysightBrowser = {
            kind: 'tree',
            stack: [{ name: root.name || 'Folder', node: root }],
            selected: new Map(),
            listing: { dirs: [], files: [] },
            csvCount: Flysight.countCsvFilesFromVirtualNode(root),
            listGen: 0
        };
        this._showFlysightBrowserModal();
        this._renderFlysightBrowser();
    }

    _flysightBrowserCurrent() {
        const stack = this._flysightBrowser?.stack;
        return stack?.length ? stack[stack.length - 1] : null;
    }

    _flysightBrowserRelDir() {
        return (this._flysightBrowser?.stack || []).map(s => s.name).filter(Boolean).join('/');
    }

    _flysightBrowserFileKey(name) {
        const rel = this._flysightBrowserRelDir();
        return rel ? `${rel}/${name}` : String(name || '');
    }

    async _renderFlysightBrowser() {
        const state = this._flysightBrowser;
        if (!state || typeof Flysight === 'undefined') return;
        const gen = ++state.listGen;
        const current = this._flysightBrowserCurrent();
        const hint = document.getElementById('flysightBrowserHint');
        if (hint) hint.hidden = state.kind !== 'tree';

        let dirs = [];
        let files = [];
        if (state.kind === 'handle') {
            const listed = await Flysight.listDirectoryHandleBrowserEntries(current?.handle);
            if (gen !== state.listGen || this._flysightBrowser !== state) return;
            dirs = listed.dirs.map((handle, i) => ({
                index: i,
                name: handle.name,
                handle,
                key: this._flysightBrowserFileKey(handle.name)
            }));
            files = listed.files.map((handle, i) => ({
                index: i,
                name: handle.name,
                handle,
                key: this._flysightBrowserFileKey(handle.name)
            }));
            state.csvCount = null;
            Flysight.countCsvFilesFromDirectoryHandle(current?.handle).then((n) => {
                if (this._flysightBrowser !== state || state.listGen !== gen) return;
                state.csvCount = n;
                this._updateFlysightBrowserImportButton();
            }).catch((err) => {
                console.error('[Flysight] Failed to count folder CSVs:', err);
                if (this._flysightBrowser !== state || state.listGen !== gen) return;
                state.csvCount = files.length;
                this._updateFlysightBrowserImportButton();
            });
        } else {
            const listed = Flysight.listVirtualTreeBrowserEntries(current?.node);
            dirs = listed.dirs.map((node, i) => ({
                index: i,
                name: node.name,
                node,
                key: this._flysightBrowserFileKey(node.name)
            }));
            files = listed.files.map((file, i) => ({
                index: i,
                name: file.name,
                file,
                key: this._flysightBrowserFileKey(file.name)
            }));
            state.csvCount = Flysight.countCsvFilesFromVirtualNode(current?.node);
        }
        if (gen !== state.listGen || this._flysightBrowser !== state) return;
        state.listing = { dirs, files };
        this._renderFlysightBrowserCrumbs();
        this._renderFlysightBrowserList();
        this._updateFlysightBrowserImportButton();
    }

    _renderFlysightBrowserCrumbs() {
        const nav = document.getElementById('flysightBrowserCrumbs');
        const state = this._flysightBrowser;
        if (!nav || !state) return;
        const last = state.stack.length - 1;
        nav.innerHTML = state.stack.map((crumb, i) => {
            const name = this.escapeHtml(crumb.name || 'Folder');
            const sep = i ? '<span class="flysight-browser-crumb-sep" aria-hidden="true">/</span>' : '';
            if (i === last) {
                return `${sep}<span class="flysight-browser-crumb is-current" aria-current="page">${name}</span>`;
            }
            return `${sep}<button type="button" class="flysight-browser-crumb" data-crumb-index="${i}">${name}</button>`;
        }).join('');
    }

    _renderFlysightBrowserList() {
        const list = document.getElementById('flysightBrowserList');
        const empty = document.getElementById('flysightBrowserEmpty');
        const state = this._flysightBrowser;
        if (!list || !state) return;
        const { dirs, files } = state.listing;
        if (!dirs.length && !files.length) {
            list.innerHTML = '';
            if (empty) empty.hidden = false;
            return;
        }
        if (empty) empty.hidden = true;
        const dirHtml = dirs.map((dir, i) => {
            const checked = state.selected?.get(dir.key)?.type === 'dir' ? ' checked' : '';
            const name = this.escapeHtml(dir.name);
            const key = this.escapeHtml(dir.key);
            return `
            <div class="flysight-browser-dir" role="listitem">
                <label class="flysight-browser-check">
                    <input type="checkbox" data-dir-key="${key}" data-dir-index="${i}"${checked} aria-label="Select folder ${name}">
                </label>
                <button type="button" class="flysight-browser-dir-enter" data-dir-index="${i}" aria-label="Open folder ${name}">
                    <span class="flysight-browser-dir-name">${name}</span>
                    <span class="flysight-browser-dir-chevron" aria-hidden="true">›</span>
                </button>
            </div>`;
        }).join('');
        const fileHtml = files.map((file, i) => {
            const checked = state.selected?.get(file.key)?.type === 'file' ? ' checked' : '';
            return `
            <label class="flysight-browser-file" role="listitem">
                <input type="checkbox" data-file-key="${this.escapeHtml(file.key)}" data-file-index="${i}"${checked}>
                <span class="flysight-browser-file-name">${this.escapeHtml(file.name)}</span>
            </label>`;
        }).join('');
        list.innerHTML = dirHtml + fileHtml;
    }

    _toggleFlysightBrowserSelection(key, checked, type, index) {
        const state = this._flysightBrowser;
        if (!state?.selected || !key) return;
        if (!checked) {
            state.selected.delete(key);
            this._updateFlysightBrowserImportButton();
            return;
        }
        if (type === 'dir') {
            const dir = state.listing?.dirs?.[index];
            if (!dir) return;
            state.selected.set(key, {
                type: 'dir',
                name: dir.name,
                handle: dir.handle,
                node: dir.node
            });
        } else {
            const file = state.listing?.files?.[index];
            if (!file) return;
            state.selected.set(key, {
                type: 'file',
                name: file.name,
                key,
                handle: file.handle,
                file: file.file
            });
        }
        this._updateFlysightBrowserImportButton();
    }

    _flysightBrowserSelectionCounts() {
        let files = 0;
        let dirs = 0;
        const selected = this._flysightBrowser?.selected;
        if (selected) {
            for (const item of selected.values()) {
                if (item.type === 'dir') dirs += 1;
                else files += 1;
            }
        }
        return { files, dirs };
    }

    _uniqueFlysightFiles(files) {
        const seen = new Set();
        const out = [];
        for (const file of files || []) {
            const key = String(file?.webkitRelativePath || file?.name || '');
            if (!key || seen.has(key)) continue;
            seen.add(key);
            out.push(file);
        }
        return out;
    }

    _updateFlysightBrowserImportButton() {
        const btn = document.getElementById('flysightBrowserImportBtn');
        const state = this._flysightBrowser;
        if (!btn) return;
        if (!state) {
            btn.disabled = true;
            btn.textContent = 'Import folder';
            return;
        }
        const { files, dirs } = this._flysightBrowserSelectionCounts();
        if (dirs || files) {
            btn.disabled = false;
            const parts = [];
            if (dirs) parts.push(dirs === 1 ? '1 folder' : `${dirs} folders`);
            if (files) parts.push(files === 1 ? '1 file' : `${files} files`);
            btn.textContent = `Import ${parts.join(', ')}`;
            return;
        }
        const n = state.csvCount;
        const maybeNested = (state.listing?.dirs?.length || 0) + (state.listing?.files?.length || 0) > 0;
        if (n == null) {
            btn.disabled = !maybeNested;
            btn.textContent = 'Import folder';
            return;
        }
        btn.textContent = n ? `Import folder (${n} CSV)` : 'Import folder';
        btn.disabled = n === 0;
    }

    async _enterFlysightBrowserDir(index) {
        const state = this._flysightBrowser;
        const dir = state?.listing?.dirs?.[Number(index)];
        if (!dir) return;
        if (state.kind === 'handle') {
            state.stack.push({ name: dir.name, handle: dir.handle });
        } else {
            state.stack.push({ name: dir.name, node: dir.node });
        }
        await this._renderFlysightBrowser();
    }

    async _goToFlysightBrowserCrumb(index) {
        const state = this._flysightBrowser;
        if (!state) return;
        const i = Number(index);
        if (!Number.isFinite(i) || i < 0 || i >= state.stack.length - 1) return;
        state.stack = state.stack.slice(0, i + 1);
        await this._renderFlysightBrowser();
    }

    _flysightBrowserRelDirOfKey(key) {
        const path = String(key || '');
        const i = path.lastIndexOf('/');
        return i < 0 ? '' : path.slice(0, i);
    }

    async _importFlysightBrowserSelection() {
        const state = this._flysightBrowser;
        if (!state || typeof Flysight === 'undefined') return;
        const btn = document.getElementById('flysightBrowserImportBtn');
        if (btn) btn.disabled = true;
        let files = [];
        try {
            const current = this._flysightBrowserCurrent();
            if (state.selected?.size) {
                const collected = [];
                for (const [key, item] of state.selected) {
                    if (item.type === 'dir') {
                        if (state.kind === 'handle') {
                            collected.push(...await Flysight.collectCsvFilesFromDirectoryHandle(item.handle));
                        } else {
                            collected.push(...Flysight.collectCsvFilesFromVirtualNode(item.node));
                        }
                    } else if (state.kind === 'handle') {
                        collected.push(...await Flysight.readCsvFilesFromFileHandles(
                            [item.handle],
                            this._flysightBrowserRelDirOfKey(item.key || key)
                        ));
                    } else if (item.file) {
                        collected.push(item.file);
                    }
                }
                files = this._uniqueFlysightFiles(collected);
            } else if (state.kind === 'handle') {
                files = await Flysight.collectCsvFilesFromDirectoryHandle(current?.handle);
            } else {
                files = Flysight.collectCsvFilesFromVirtualNode(current?.node);
            }
        } catch (err) {
            console.error('[Flysight] Failed to collect folder files:', err);
            this.showMessage('Could not read that folder.', 'error');
            this._updateFlysightBrowserImportButton();
            return;
        }
        this.closeFlysightBrowserModal();
        if (!files.length) {
            this.showMessage('No CSV files found in that folder.', 'error');
            return;
        }
        this._setFlysightEmptyResults('Analysing files');
        await this._waitForFlysightUiPaint();
        await this._addFlysightFiles(files);
    }

    async _collectFlysightFilesFromDataTransfer(dataTransfer) {
        if (!dataTransfer) return [];
        const items = dataTransfer.items;
        const entries = [];
        if (items?.length && typeof items[0].webkitGetAsEntry === 'function') {
            for (let i = 0; i < items.length; i++) {
                const entry = items[i].webkitGetAsEntry?.();
                if (entry) entries.push(entry);
            }
        }
        if (entries.length && typeof Flysight !== 'undefined') {
            return Flysight.collectCsvFilesFromEntries(entries);
        }
        if (typeof Flysight !== 'undefined') {
            return Flysight.collectCsvFilesFromFileList(dataTransfer.files);
        }
        return [...(dataTransfer.files || [])].filter(f => /\.csv$/i.test(f.name) || f.type === 'text/csv');
    }

    async _addFlysightFilesFromDrop(dataTransfer) {
        const files = await this._collectFlysightFilesFromDataTransfer(dataTransfer);
        await this._addFlysightFiles(files);
    }

    async _addFlysightFiles(fileList) {
        const files = typeof Flysight !== 'undefined'
            ? Flysight.collectCsvFilesFromFileList(fileList)
            : [...(fileList || [])].filter(f => /\.csv$/i.test(f.name) || f.type === 'text/csv');
        if (!files.length) {
            this.showMessage('Please select CSV files or a folder that contains them.', 'error');
            return;
        }

        for (const file of files) {
            try {
                const text = await file.text();
                this.flysightFiles.push({
                    id: `flysight_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`,
                    name: file.name,
                    text
                });
            } catch (err) {
                console.error('[Flysight] Failed to read file:', file.name, err);
                this.showMessage(`Could not read ${file.name}`, 'error');
            }
        }

        if (this.currentView === 'flysight') {
            this.renderFlysightView();
        } else {
            this.showView('flysight');
        }
    }

    removeFlysightFile(id) {
        this.flysightFiles = this.flysightFiles.filter(f => f.id !== id);
        this.renderFlysightView();
    }

    renderFlysightView() {
        const container = document.getElementById('flysightResults');
        const avgSlider = document.getElementById('flysightAvgPoints');
        if (!container) return;

        if (avgSlider) avgSlider.value = String(this.flysightAvgPoints);
        this._syncFlysightSettingsForm();
        this._updateFlysightSpeedModeButtons();

        if (!this.flysightFiles.length) {
            container.innerHTML = this._flysightEmptyResultsHtml();
            return;
        }

        if (typeof Flysight === 'undefined') {
            container.innerHTML = '<p class="flysight-result-error">Flysight module failed to load.</p>';
            return;
        }

        container.innerHTML = this.flysightFiles.map(file => {
            const result = Flysight.analyzeFlysightCsv(
                file.text,
                this.flysightAvgPoints,
                this.flysightMaxHeightM,
                this.flysightSpeedMetric
            );
            const title = Flysight.formatTrackStartTitle(result.points?.[0]?.time) || file.name;
            if (result.error) {
                return `
                    <div class="flysight-result-card is-error">
                        <div class="flysight-result-name">${this.escapeHtml(title)}</div>
                        <p class="flysight-result-error">${this.escapeHtml(result.error)}</p>
                        <button type="button" class="flysight-result-remove" onclick="logbook.removeFlysightFile('${file.id}')">Remove</button>
                    </div>`;
            }

            const altitude = Math.round(result.altitudeM);
            const recoverySec = Flysight.recoveryArcSec(
                result.points,
                this.flysightAvgPoints,
                this.flysightCursorBDiveAngleDeg,
                this.flysightCursorBAltTicks
            );
            const recoveryText = Number.isFinite(recoverySec)
                ? (Flysight.formatDurationSec(recoverySec) || `${recoverySec.toFixed(1)}s`)
                : '—';

            let speedMetricsHtml;
            let metaHtml = '';
            let metricsClass = 'flysight-result-metrics';
            let altLabel = 'Altitude';
            let recoveryLabel = 'Recovery arc';

            if (result.speedMetric === 'both') {
                const verticalSpeed = result.maxVerticalSpeedKmh.toFixed(1);
                const totalSpeed = result.maxTotalSpeedKmh.toFixed(1);
                const totalAltitude = Number.isFinite(result.totalPeakAltitudeM)
                    ? `${Math.round(result.totalPeakAltitudeM)} m`
                    : '—';
                speedMetricsHtml = `
                        <div>
                            <span class="flysight-metric-label">Max total</span>
                            <span class="flysight-metric-value">${totalSpeed} km/h</span>
                        </div>
                        <div class="flysight-metric-emphasis">
                            <span class="flysight-metric-label">Max vertical</span>
                            <span class="flysight-metric-value">${verticalSpeed} km/h</span>
                        </div>`;
                metaHtml = `${result.pointCount} points · vertical peak at ${altitude} m · total peak at ${totalAltitude}`;
            } else {
                const speed = result.maxVerticalSpeedKmh.toFixed(1);
                const speedLabel = result.speedMetric === 'total' ? 'Max total' : 'Max vertical';
                metricsClass += ' is-single';
                altLabel = 'Altitude';
                recoveryLabel = 'Recovery arc';
                speedMetricsHtml = `
                        <div>
                            <span class="flysight-metric-label">${speedLabel}</span>
                            <span class="flysight-metric-value">${speed} km/h</span>
                        </div>`;
            }

            return `
                <div class="flysight-result-card">
                    <div class="flysight-result-name">${this.escapeHtml(title)}</div>
                    <div class="${metricsClass}">
                        ${speedMetricsHtml}
                        <div>
                            <span class="flysight-metric-label">${altLabel}</span>
                            <span class="flysight-metric-value">${altitude} m</span>
                        </div>
                        <div>
                            <span class="flysight-metric-label">${recoveryLabel}</span>
                            <span class="flysight-metric-value">${recoveryText}</span>
                        </div>
                    </div>
                    ${metaHtml ? `<p class="flysight-result-meta">${metaHtml}</p>` : ''}
                    <div class="flysight-result-actions">
                        <button type="button" class="flysight-result-graph" data-flysight-id="${file.id}" data-flysight-graph="diveAngle">
                            <svg class="flysight-result-graph-icon" viewBox="0 0 24 24" aria-hidden="true" focusable="false">
                                <path fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" d="M3 3v18h18"/>
                                <path fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" d="M6 16c2.2-1.2 3.1-7 5.6-7s2.4 5.2 4.9 5.2 2.1-4.2 3.8-4.2"/>
                            </svg>
                            Dive Angle
                        </button>
                        <button type="button" class="flysight-result-graph" data-flysight-id="${file.id}" data-flysight-graph="verticalSpeed">
                            <svg class="flysight-result-graph-icon" viewBox="0 0 24 24" aria-hidden="true" focusable="false">
                                <path fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" d="M3 3v18h18"/>
                                <path fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" d="M6 16c2.2-1.2 3.1-7 5.6-7s2.4 5.2 4.9 5.2 2.1-4.2 3.8-4.2"/>
                            </svg>
                            Vertical Speed
                        </button>
                        <button type="button" class="flysight-result-remove" onclick="logbook.removeFlysightFile('${file.id}')">Remove</button>
                    </div>
                </div>`;
        }).join('');
    }

    openFlysightGraphModal(fileId, mode) {
        const file = this.flysightFiles.find(f => f.id === fileId);
        const modal = document.getElementById('flysightGraphModal');
        if (!file || !modal) return;
        if (typeof Flysight === 'undefined') {
            this.showMessage('Flysight module failed to load.', 'error');
            return;
        }

        const parsed = Flysight.parseFlysightCsv(file.text);
        if (parsed.error) {
            this.showMessage(parsed.error, 'error');
            return;
        }
        const series = Flysight.buildSwoopCursorSeries(
            parsed.points,
            this.flysightAvgPoints
        );
        if (series.error || !series.samples.length) {
            this.showMessage(series.error || 'Could not build graph.', 'error');
            return;
        }
        const cursors = Flysight.defaultSwoopCursorIndices(
            series.samples,
            this.flysightCursorBDiveAngleDeg,
            this.flysightCursorBAltTicks
        );
        const graphMode = mode === 'verticalSpeed' ? 'verticalSpeed' : 'diveAngle';
        this._flysightGraph = {
            fileId,
            mode: graphMode,
            samples: series.samples,
            idxA: cursors.idxA,
            idxB: cursors.idxB,
            drag: null,
            pan: null,
            pinch: null,
            pointers: new Map(),
            _layoutW: 0,
            _layoutH: 0
        };
        const title = Flysight.formatTrackStartTitle(parsed.points?.[0]?.time) || file.name;
        const titleEl = document.getElementById('flysightGraphTitle');
        if (titleEl) titleEl.textContent = title;
        const kind = graphMode === 'verticalSpeed' ? 'vertical speed' : 'dive angle';
        modal.setAttribute('aria-label', title ? `${title} ${kind} graph` : `${kind.charAt(0).toUpperCase()}${kind.slice(1)} graph`);
        this._initFlysightGraphView(this._flysightGraph);
        modal.style.display = 'flex';
        this._setFlysightGraphScrollLock(true);
        this._syncFlysightGraphViewport();
        this._syncFlysightGraphNav();
        this._drawFlysightGraph();
        requestAnimationFrame(() => this._drawFlysightGraph());
    }

    _isFlysightGraphable(file) {
        if (!file || typeof Flysight === 'undefined') return false;
        const result = Flysight.analyzeFlysightCsv(
            file.text,
            this.flysightAvgPoints,
            this.flysightMaxHeightM,
            this.flysightSpeedMetric
        );
        return !result.error;
    }

    _flysightGraphNeighbor(delta) {
        const fileId = this._flysightGraph?.fileId;
        if (!fileId || !delta) return null;
        const idx = this.flysightFiles.findIndex(f => f.id === fileId);
        if (idx < 0) return null;
        let i = idx + delta;
        while (i >= 0 && i < this.flysightFiles.length) {
            const file = this.flysightFiles[i];
            if (this._isFlysightGraphable(file)) return file;
            i += delta;
        }
        return null;
    }

    _syncFlysightGraphNav() {
        const prevBtn = document.getElementById('flysightGraphPrev');
        const nextBtn = document.getElementById('flysightGraphNext');
        if (!prevBtn || !nextBtn) return;
        prevBtn.hidden = !this._flysightGraphNeighbor(-1);
        nextBtn.hidden = !this._flysightGraphNeighbor(1);
    }

    navigateFlysightGraph(delta) {
        const g = this._flysightGraph;
        const neighbor = this._flysightGraphNeighbor(delta);
        if (!g || !neighbor) return;
        this.openFlysightGraphModal(neighbor.id, g.mode);
    }

    closeFlysightGraphModal() {
        const modal = document.getElementById('flysightGraphModal');
        if (modal) {
            modal.style.display = 'none';
            modal.style.top = '';
            modal.style.left = '';
            modal.style.right = '';
            modal.style.bottom = '';
            modal.style.width = '';
            modal.style.height = '';
        }
        this._setFlysightGraphScrollLock(false);
        this._clearFlysightGraphInteraction();
        this._syncFlysightGraphResetZoomBtn();
        const prevBtn = document.getElementById('flysightGraphPrev');
        const nextBtn = document.getElementById('flysightGraphNext');
        if (prevBtn) prevBtn.hidden = true;
        if (nextBtn) nextBtn.hidden = true;
    }

    _setFlysightGraphScrollLock(lock) {
        const html = document.documentElement;
        const body = document.body;
        if (lock) {
            this._flysightScrollY = window.scrollY || window.pageYOffset || 0;
            html.classList.add('flysight-graph-open');
            body.classList.add('flysight-graph-open');
            body.style.top = `-${this._flysightScrollY}px`;
        } else {
            html.classList.remove('flysight-graph-open');
            body.classList.remove('flysight-graph-open');
            body.style.top = '';
            window.scrollTo(0, this._flysightScrollY || 0);
        }
    }

    _syncFlysightGraphViewport() {
        const modal = document.getElementById('flysightGraphModal');
        if (!modal || modal.style.display === 'none' || !modal.style.display) return;
        const vv = window.visualViewport;
        if (vv) {
            modal.style.top = `${vv.offsetTop}px`;
            modal.style.left = `${vv.offsetLeft}px`;
            modal.style.right = 'auto';
            modal.style.bottom = 'auto';
            modal.style.width = `${vv.width}px`;
            modal.style.height = `${vv.height}px`;
        } else {
            modal.style.top = '0';
            modal.style.left = '0';
            modal.style.right = '0';
            modal.style.bottom = '0';
            modal.style.width = 'auto';
            modal.style.height = 'auto';
        }
    }

    _onFlysightGraphViewportChange() {
        const modal = document.getElementById('flysightGraphModal');
        if (!modal || modal.style.display !== 'flex') return;
        this._syncFlysightGraphViewport();
        this._drawFlysightGraph();
    }

    _flysightGraphLayout() {
        const root = document.getElementById('flysightGraphRoot');
        const width = Math.max(240, Math.round(root?.clientWidth || 720));
        const height = Math.max(140, Math.round(root?.clientHeight || 340));
        const compact = width < 520 || height < 300;
        const speedMode = this._flysightGraphMode() === 'verticalSpeed';
        return {
            width,
            height,
            compact,
            l: compact ? (speedMode ? 54 : 46) : (speedMode ? 72 : 64),
            r: compact ? 36 : 48,
            t: compact ? 14 : 22,
            b: compact ? 36 : 44
        };
    }

    _flysightGraphGroundHmsl(samples) {
        return samples.reduce((min, s) => {
            const h = s.hMSL;
            return Number.isFinite(h) && h < min ? h : min;
        }, Infinity);
    }

    _flysightGraphAglM(sample, ground) {
        if (!Number.isFinite(sample?.hMSL) || !Number.isFinite(ground)) return NaN;
        return Math.max(0, sample.hMSL - ground);
    }

    _flysightVelKmh(velDMs) {
        return velDMs * 3.6;
    }

    _flysightTotalSpeedKmh(sample) {
        if (!sample) return NaN;
        if (typeof Flysight !== 'undefined') {
            const ms = Flysight.trajectorySpeedMs(sample);
            return Number.isFinite(ms) ? ms * 3.6 : NaN;
        }
        const n = Number(sample.velN);
        const e = Number(sample.velE);
        const d = Number(sample.velD);
        if (![n, e, d].every(Number.isFinite)) {
            const v = this._flysightVelKmh(d);
            return Number.isFinite(v) ? v : NaN;
        }
        return Math.hypot(n, e, d) * 3.6;
    }

    _flysightGraphMode() {
        return this._flysightGraph?.mode === 'verticalSpeed' ? 'verticalSpeed' : 'diveAngle';
    }

    _flysightGraphYValue(sample) {
        if (this._flysightGraphMode() === 'verticalSpeed') {
            const v = this._flysightVelKmh(sample?.velD);
            return Number.isFinite(v) ? v : 0;
        }
        return this._flysightDiveAngleDeg(sample);
    }

    _flysightDiveAngleDeg(sample) {
        if (typeof Flysight !== 'undefined') return Flysight.diveAngleDeg(sample);
        const v = sample?.diveAngleDeg;
        return Number.isFinite(v) ? v : 0;
    }

    _flysightDegTicks(yMin, yMax) {
        const span = yMax - yMin;
        const step = span > 60 ? 15 : span > 30 ? 10 : span > 15 ? 5 : span > 6 ? 2 : span > 3 ? 1 : span > 1.5 ? 0.5 : 0.2;
        return this._flysightTicksInRange(yMin, yMax, step);
    }

    _flysightSpeedTicks(yMin, yMax) {
        const span = yMax - yMin;
        const step = span > 250 ? 50 : span > 120 ? 20 : span > 60 ? 10 : span > 30 ? 5 : span > 15 ? 2 : span > 6 ? 1 : span > 3 ? 0.5 : 0.2;
        return this._flysightTicksInRange(yMin, yMax, step);
    }

    _flysightGraphYTicks(yMin, yMax) {
        return this._flysightGraphMode() === 'verticalSpeed'
            ? this._flysightSpeedTicks(yMin, yMax)
            : this._flysightDegTicks(yMin, yMax);
    }

    _flysightTimeTicks(tMin, tMax) {
        const span = tMax - tMin;
        const step = span > 20 ? 5 : span > 10 ? 2 : span > 5 ? 1 : span > 2 ? 0.5 : span > 1 ? 0.2 : 0.1;
        return this._flysightTicksInRange(tMin, tMax, step);
    }

    _flysightTicksInRange(min, max, step) {
        const ticks = [];
        const start = Math.ceil((min - 1e-9) / step) * step;
        for (let i = 0; i < 48; i++) {
            const v = Number((start + i * step).toFixed(6));
            if (v > max + step * 0.01) break;
            ticks.push(v);
        }
        return { ticks, step };
    }

    _flysightTickLabel(v, step) {
        if (step >= 1) return String(Math.round(v));
        return v.toFixed(step >= 0.1 ? 1 : 2);
    }

    _flysightGraphFullExtents(samples) {
        let tMax = 0.001;
        for (let i = 0; i < samples.length; i++) {
            const t = samples[i]?.tRev;
            if (Number.isFinite(t) && t > tMax) tMax = t;
        }
        const values = samples.map(s => this._flysightGraphYValue(s)).filter(Number.isFinite);
        const floorMax = this._flysightGraphMode() === 'verticalSpeed' ? 40 : 15;
        const yMin = values.length ? Math.min(0, ...values) : 0;
        const yMax = (values.length ? Math.max(floorMax, ...values) : floorMax) * 1.08;
        return { tMin: 0, tMax, yMin, yMax };
    }

    _initFlysightGraphView(g) {
        const ext = this._flysightGraphFullExtents(g.samples);
        g.tFullMin = ext.tMin;
        g.tFullMax = ext.tMax;
        g.yFullMin = ext.yMin;
        g.yFullMax = ext.yMax;
        g.viewTMin = ext.tMin;
        g.viewTMax = ext.tMax;
        g.viewYMin = ext.yMin;
        g.viewYMax = ext.yMax;
    }

    _isFlysightGraphZoomed() {
        const g = this._flysightGraph;
        if (!g || !Number.isFinite(g.viewTMin) || !Number.isFinite(g.tFullMax)) return false;
        const tFull = g.tFullMax - g.tFullMin;
        const yFull = g.yFullMax - g.yFullMin;
        if (tFull <= 0 || yFull <= 0) return false;
        return ((g.viewTMax - g.viewTMin) / tFull) < 0.999
            || ((g.viewYMax - g.viewYMin) / yFull) < 0.999;
    }

    _syncFlysightGraphResetZoomBtn() {
        const btn = document.getElementById('flysightGraphResetZoom');
        if (!btn) return;
        const modal = document.getElementById('flysightGraphModal');
        const open = modal && modal.style.display === 'flex';
        btn.hidden = !open || !this._isFlysightGraphZoomed();
    }

    _resetFlysightGraphZoom() {
        const g = this._flysightGraph;
        if (!g || !Number.isFinite(g.tFullMax)) return;
        g.viewTMin = g.tFullMin;
        g.viewTMax = g.tFullMax;
        g.viewYMin = g.yFullMin;
        g.viewYMax = g.yFullMax;
        this._drawFlysightGraph();
    }

    _clampFlysightGraphView() {
        const g = this._flysightGraph;
        if (!g) return;
        const tFull = g.tFullMax - g.tFullMin;
        const yFull = g.yFullMax - g.yFullMin;
        const minTSpan = Math.max(0.5, tFull / 50);
        const minYSpan = Math.max(2, yFull / 50);
        let tSpan = Math.min(tFull, Math.max(minTSpan, g.viewTMax - g.viewTMin));
        let ySpan = Math.min(yFull, Math.max(minYSpan, g.viewYMax - g.viewYMin));

        if (tSpan >= tFull - 1e-9) {
            g.viewTMin = g.tFullMin;
            g.viewTMax = g.tFullMax;
        } else {
            if (g.viewTMin < g.tFullMin) {
                g.viewTMin = g.tFullMin;
                g.viewTMax = g.tFullMin + tSpan;
            }
            if (g.viewTMax > g.tFullMax) {
                g.viewTMax = g.tFullMax;
                g.viewTMin = g.tFullMax - tSpan;
            }
            g.viewTMin = Math.max(g.tFullMin, g.viewTMin);
            g.viewTMax = Math.min(g.tFullMax, g.viewTMax);
        }

        if (ySpan >= yFull - 1e-9) {
            g.viewYMin = g.yFullMin;
            g.viewYMax = g.yFullMax;
        } else {
            if (g.viewYMin < g.yFullMin) {
                g.viewYMin = g.yFullMin;
                g.viewYMax = g.yFullMin + ySpan;
            }
            if (g.viewYMax > g.yFullMax) {
                g.viewYMax = g.yFullMax;
                g.viewYMin = g.yFullMax - ySpan;
            }
            g.viewYMin = Math.max(g.yFullMin, g.viewYMin);
            g.viewYMax = Math.min(g.yFullMax, g.viewYMax);
        }
    }

    _flysightGraphClientToSvg(clientX, clientY) {
        const svg = document.querySelector('#flysightGraphRoot svg');
        const layout = this._flysightGraphLayout();
        if (!svg) return { x: layout.l, y: layout.t, layout, rect: null };
        const rect = svg.getBoundingClientRect();
        const x = rect.width ? ((clientX - rect.left) / rect.width) * layout.width : layout.l;
        const y = rect.height ? ((clientY - rect.top) / rect.height) * layout.height : layout.t;
        return { x, y, layout, rect };
    }

    _zoomFlysightGraphAtClient(clientX, clientY, spanFactor) {
        const g = this._flysightGraph;
        if (!g) return;
        const { x, y, layout } = this._flysightGraphClientToSvg(clientX, clientY);
        const sc = this._flysightGraphScales(g.samples, layout);
        const tRev = sc.viewTMax - ((x - layout.l) / sc.innerW) * sc.tSpan;
        const deg = sc.viewYMax - ((y - layout.t) / sc.innerH) * sc.ySpan;
        const leftFrac = sc.tSpan ? (sc.viewTMax - tRev) / sc.tSpan : 0.5;
        const topFrac = sc.ySpan ? (sc.viewYMax - deg) / sc.ySpan : 0.5;
        const tFull = g.tFullMax - g.tFullMin;
        const yFull = g.yFullMax - g.yFullMin;
        const minTSpan = Math.max(0.5, tFull / 50);
        const minYSpan = Math.max(2, yFull / 50);
        const tSpan = Math.min(tFull, Math.max(minTSpan, sc.tSpan * spanFactor));
        const ySpan = Math.min(yFull, Math.max(minYSpan, sc.ySpan * spanFactor));
        g.viewTMax = tRev + leftFrac * tSpan;
        g.viewTMin = g.viewTMax - tSpan;
        g.viewYMax = deg + topFrac * ySpan;
        g.viewYMin = g.viewYMax - ySpan;
        this._clampFlysightGraphView();
    }

    _panFlysightGraphByClientDelta(dx, dy) {
        const g = this._flysightGraph;
        const svg = document.querySelector('#flysightGraphRoot svg');
        if (!g || !svg) return;
        const layout = this._flysightGraphLayout();
        const sc = this._flysightGraphScales(g.samples, layout);
        const rect = svg.getBoundingClientRect();
        if (!rect.width || !rect.height) return;
        const dxSvg = (dx / rect.width) * layout.width;
        const dySvg = (dy / rect.height) * layout.height;
        g.viewTMin += (dxSvg / sc.innerW) * sc.tSpan;
        g.viewTMax += (dxSvg / sc.innerW) * sc.tSpan;
        g.viewYMin += (dySvg / sc.innerH) * sc.ySpan;
        g.viewYMax += (dySvg / sc.innerH) * sc.ySpan;
        this._clampFlysightGraphView();
    }

    _flysightGraphPinchSnapshot() {
        const g = this._flysightGraph;
        const pts = [...(g.pointers?.values() || [])];
        if (pts.length < 2) return null;
        const dx = pts[1].x - pts[0].x;
        const dy = pts[1].y - pts[0].y;
        return {
            dist: Math.hypot(dx, dy) || 1,
            midX: (pts[0].x + pts[1].x) / 2,
            midY: (pts[0].y + pts[1].y) / 2,
            viewTMin: g.viewTMin,
            viewTMax: g.viewTMax,
            viewYMin: g.viewYMin,
            viewYMax: g.viewYMax
        };
    }

    _applyFlysightGraphPinch() {
        const g = this._flysightGraph;
        const pts = [...(g?.pointers?.values() || [])];
        if (!g?.pinch || pts.length < 2) return;
        const dist = Math.hypot(pts[1].x - pts[0].x, pts[1].y - pts[0].y) || 1;
        const midX = (pts[0].x + pts[1].x) / 2;
        const midY = (pts[0].y + pts[1].y) / 2;
        g.viewTMin = g.pinch.viewTMin;
        g.viewTMax = g.pinch.viewTMax;
        g.viewYMin = g.pinch.viewYMin;
        g.viewYMax = g.pinch.viewYMax;
        this._zoomFlysightGraphAtClient(g.pinch.midX, g.pinch.midY, g.pinch.dist / dist);
        this._panFlysightGraphByClientDelta(midX - g.pinch.midX, midY - g.pinch.midY);
        this._drawFlysightGraph();
    }

    _clearFlysightGraphInteraction() {
        const g = this._flysightGraph;
        if (!g) return;
        g.drag = null;
        g.pan = null;
        g.pinch = null;
        g.pointers?.clear();
        document.getElementById('flysightGraphRoot')?.classList.remove('is-panning');
    }

    _flysightGraphScales(samples, layout) {
        const ext = this._flysightGraphFullExtents(samples);
        const g = this._flysightGraph;
        const viewTMin = g && Number.isFinite(g.viewTMin) ? g.viewTMin : ext.tMin;
        const viewTMax = g && Number.isFinite(g.viewTMax) ? g.viewTMax : ext.tMax;
        const viewYMin = g && Number.isFinite(g.viewYMin) ? g.viewYMin : ext.yMin;
        const viewYMax = g && Number.isFinite(g.viewYMax) ? g.viewYMax : ext.yMax;
        const tSpan = Math.max(0.001, viewTMax - viewTMin);
        const ySpan = Math.max(0.001, viewYMax - viewYMin);
        const innerW = layout.width - layout.l - layout.r;
        const innerH = layout.height - layout.t - layout.b;
        return {
            tMax: ext.tMax,
            viewTMin,
            viewTMax,
            viewYMin,
            viewYMax,
            tSpan,
            ySpan,
            innerW,
            innerH,
            xOf: (tRev) => layout.l + ((viewTMax - tRev) / tSpan) * innerW,
            yOf: (deg) => layout.t + (1 - (deg - viewYMin) / ySpan) * innerH
        };
    }

    _updateFlysightGraphStats() {
        const g = this._flysightGraph;
        if (!g) return;
        const a = g.samples[g.idxA];
        const b = g.samples[g.idxB];
        if (!a || !b) return;
        const dt = Math.abs(b.tRev - a.tRev);
        const aloft = Flysight.timeAloftSec(b);
        const dtEl = document.getElementById('flysightGraphDt');
        const aloftEl = document.getElementById('flysightGraphTimeAloft');
        const velAEl = document.getElementById('flysightGraphVelA');
        const velLabelEl = document.getElementById('flysightGraphVelLabel');
        if (dtEl) dtEl.textContent = Flysight.formatDurationSec(dt) || `${dt.toFixed(1)}s`;
        if (aloftEl) aloftEl.textContent = Flysight.formatDurationSec(aloft) || `${aloft.toFixed(1)}s`;
        const speedMode = this._flysightGraphMode() === 'verticalSpeed';
        if (velLabelEl) velLabelEl.textContent = speedMode ? 'Total Speed @ B' : 'Vertical@A';
        if (velAEl) {
            const vel = speedMode ? this._flysightTotalSpeedKmh(b) : this._flysightVelKmh(a.velD);
            velAEl.textContent = Number.isFinite(vel) ? `${vel.toFixed(1)} km/h` : '—';
        }
    }

    _drawFlysightGraph() {
        const root = document.getElementById('flysightGraphRoot');
        const g = this._flysightGraph;
        if (!root || !g) return;
        if (root.clientWidth < 40 || root.clientHeight < 40) return;
        const samples = g.samples;
        const layout = this._flysightGraphLayout();
        g._layoutW = layout.width;
        g._layoutH = layout.height;
        const sc = this._flysightGraphScales(samples, layout);
        const a = samples[g.idxA];
        const b = samples[g.idxB];
        const speedMode = this._flysightGraphMode() === 'verticalSpeed';
        const yValue = (s) => this._flysightGraphYValue(s);
        const poly = samples.map(s => `${sc.xOf(s.tRev).toFixed(1)},${sc.yOf(yValue(s)).toFixed(1)}`).join(' ');
        const yTick = this._flysightGraphYTicks(sc.viewYMin, sc.viewYMax);
        const xTick = this._flysightTimeTicks(sc.viewTMin, sc.viewTMax);
        const yBase = layout.height - layout.b;
        const plotRight = layout.width - layout.r;
        const bAngleThresh = Number.isFinite(this.flysightCursorBDiveAngleDeg)
            ? this.flysightCursorBDiveAngleDeg
            : Flysight.CURSOR_B_DIVE_ANGLE_DEG;
        const showBThresh = !speedMode && bAngleThresh > 0 && bAngleThresh >= sc.viewYMin && bAngleThresh <= sc.viewYMax;
        const fs = layout.compact ? 10 : 11;
        const pointFs = layout.compact ? 17 : 19;
        const handleR = layout.compact ? 8 : 10;
        const handleY = yBase + (layout.compact ? 10 : 12);
        const xLabelY = layout.compact ? layout.height - 4 : layout.height - 6;

        const ground = this._flysightGraphGroundHmsl(samples);
        const peakIdx = Flysight.maxVelDIdx(samples);
        const peak = samples[peakIdx];
        const maxVerticalMark = (() => {
            if (speedMode || !peak) return '';
            const x = sc.xOf(peak.tRev);
            const y = sc.yOf(yValue(peak));
            const vel = this._flysightVelKmh(peak.velD);
            const velText = Number.isFinite(vel) ? `${vel.toFixed(1)} km/h` : '—';
            const estimateW = layout.compact ? 264 : 312;
            const putRight = (x - estimateW) < 8;
            const dx = layout.compact ? 10 : 12;
            const labelX = putRight ? x + dx : x - dx;
            const putAbove = y > layout.t + (layout.compact ? 18 : 22);
            const labelY = putAbove ? y - 10 : y + 20;
            return `
                <g class="flysight-graph-max-vertical" pointer-events="none">
                    <circle cx="${x}" cy="${y}" r="4.5" fill="#C62828" stroke="#fff" stroke-width="1.5"/>
                    <text x="${labelX}" y="${labelY}" text-anchor="${putRight ? 'start' : 'end'}" fill="#C62828" font-size="${pointFs}" font-weight="600" stroke="#fff" stroke-width="4" paint-order="stroke">max vertical = ${velText}</text>
                </g>`;
        })();
        const cursor = (which, sample, color) => {
            if (!sample) return '';
            if (sample.tRev < sc.viewTMin - 1e-6 || sample.tRev > sc.viewTMax + 1e-6) return '';
            const x = sc.xOf(sample.tRev);
            const y = sc.yOf(yValue(sample));
            const alt = this._flysightGraphAglM(sample, ground);
            let altText = '';
            if (Number.isFinite(alt)) {
                const labelX = x + (layout.compact ? 10 : 12);
                altText = `<text x="${labelX}" y="${y + 5}" text-anchor="start" fill="${color}" font-size="${pointFs}" font-weight="600" stroke="#fff" stroke-width="4" paint-order="stroke" pointer-events="none">${Math.round(alt)} m</text>`;
            }
            const yOnPlot = y >= layout.t - 2 && y <= yBase + 2;
            const aYFs = layout.compact ? 14 : 16;
            const angleText = !yOnPlot ? '' : which === 'a'
                ? `<text x="${layout.l - 6}" y="${y + 5}" text-anchor="end" fill="#0D47A1" font-size="${aYFs}" font-weight="700" stroke="#fff" stroke-width="4" paint-order="stroke">${yValue(sample).toFixed(1)}</text>`
                : `<text x="${layout.l - 6}" y="${y + 4}" text-anchor="end" fill="${color}" font-size="${fs}">${yValue(sample).toFixed(1)}</text>`;
            const aToAxis = which === 'a' && yOnPlot
                ? `<line x1="${layout.l}" y1="${y}" x2="${x}" y2="${y}" stroke="${color}" stroke-width="1.5" stroke-dasharray="3 3" pointer-events="none"/>`
                : '';
            return `
                <g class="flysight-graph-cursor" data-cursor="${which}" style="cursor:ew-resize;touch-action:none">
                    <g clip-path="url(#flysightPlotClip)">
                        ${aToAxis}
                        <line x1="${x}" y1="${y}" x2="${x}" y2="${yBase}" stroke="${color}" stroke-width="1.5"/>
                        <circle cx="${x}" cy="${y}" r="4" fill="${color}"/>
                        ${altText}
                    </g>
                    <rect x="${x - 16}" y="${yBase - 6}" width="32" height="40" fill="transparent"/>
                    <circle cx="${x}" cy="${handleY}" r="${handleR}" fill="${color}"/>
                    <text x="${x}" y="${handleY + 4}" text-anchor="middle" fill="#fff" font-size="${fs}" font-weight="600" pointer-events="none">${which.toUpperCase()}</text>
                    ${angleText}
                </g>`;
        };

        const grid = yTick.ticks.map(v => {
            const y = sc.yOf(v);
            if (y < layout.t - 1 || y > yBase + 1) return '';
            return `
            <line x1="${layout.l}" y1="${y}" x2="${plotRight}" y2="${y}" stroke="#eee"/>
            <text x="${layout.l - 6}" y="${y + 4}" text-anchor="end" fill="#888" font-size="${fs}">${this._flysightTickLabel(v, yTick.step)}</text>`;
        }).join('');
        const xLabels = xTick.ticks.map(t => {
            const x = sc.xOf(t);
            if (x < layout.l - 1 || x > plotRight + 1) return '';
            return `<text x="${x}" y="${xLabelY}" text-anchor="middle" fill="#888" font-size="${fs}">${this._flysightTickLabel(t, xTick.step)}s</text>`;
        }).join('');
        const bThreshLine = showBThresh
            ? `<line x1="${layout.l}" y1="${sc.yOf(bAngleThresh)}" x2="${plotRight}" y2="${sc.yOf(bAngleThresh)}" stroke="#ddd" stroke-dasharray="3 3"/>`
            : '';
        const zoomedClass = this._isFlysightGraphZoomed() ? 'is-zoomed' : '';

        const yAxisTitle = speedMode ? 'km/h' : 'dive angle';
        const svgAria = speedMode
            ? 'Vertical speed in kilometres per hour versus seconds before landing. Pinch or scroll to zoom, drag to pan.'
            : 'Dive angle in degrees versus seconds before landing. Pinch or scroll to zoom, drag to pan.';
        root.innerHTML = `
            <svg viewBox="0 0 ${layout.width} ${layout.height}" width="${layout.width}" height="${layout.height}"
                 preserveAspectRatio="none" class="${zoomedClass}"
                 role="img" aria-label="${svgAria}">
                <defs>
                    <clipPath id="flysightPlotClip">
                        <rect x="${layout.l}" y="${layout.t}" width="${sc.innerW}" height="${sc.innerH}"/>
                    </clipPath>
                </defs>
                ${grid}
                ${xLabels}
                <line x1="${layout.l}" y1="${layout.t}" x2="${layout.l}" y2="${yBase}" stroke="#ccc"/>
                <line x1="${layout.l}" y1="${yBase}" x2="${plotRight}" y2="${yBase}" stroke="#ccc"/>
                ${bThreshLine}
                <g clip-path="url(#flysightPlotClip)">
                    <polyline fill="none" stroke="#1976D2" stroke-width="2" points="${poly}"/>
                </g>
                ${maxVerticalMark}
                ${cursor('a', a, '#1976D2')}
                ${cursor('b', b, '#555')}
                <text x="${layout.l}" y="${Math.max(10, layout.t - 4)}" text-anchor="middle" fill="#888" font-size="${fs}">${yAxisTitle}</text>
            </svg>`;
        this._updateFlysightGraphStats();
        this._syncFlysightGraphResetZoomBtn();
    }

    _flysightGraphIndexFromClientX(clientX) {
        const g = this._flysightGraph;
        const svg = document.querySelector('#flysightGraphRoot svg');
        if (!g || !svg) return 0;
        const layout = this._flysightGraphLayout();
        const sc = this._flysightGraphScales(g.samples, layout);
        const rect = svg.getBoundingClientRect();
        const x = ((clientX - rect.left) / rect.width) * layout.width;
        const tPlot = ((x - layout.l) / sc.innerW) * sc.tSpan;
        const tRev = Math.max(sc.viewTMin, Math.min(sc.viewTMax, sc.viewTMax - tPlot));
        let best = 0;
        let bestD = Infinity;
        for (let i = 0; i < g.samples.length; i++) {
            const d = Math.abs(g.samples[i].tRev - tRev);
            if (d < bestD) {
                bestD = d;
                best = i;
            }
        }
        return best;
    }

    _onFlysightGraphPointerDown(e) {
        const g = this._flysightGraph;
        if (!g) return;
        e.preventDefault();
        try { e.currentTarget.setPointerCapture(e.pointerId); } catch (_) { /* older WebView */ }
        if (!g.pointers) g.pointers = new Map();
        g.pointers.set(e.pointerId, { x: e.clientX, y: e.clientY });
        if (g.pointers.size === 1) {
            g.didPan = false;
            g.didDrag = false;
            g.didPinch = false;
        }
        if (g.pointers.size >= 2) {
            g.drag = null;
            g.pan = null;
            g.pinch = this._flysightGraphPinchSnapshot();
            g.didPinch = true;
            document.getElementById('flysightGraphRoot')?.classList.remove('is-panning');
            return;
        }
        const el = e.target instanceof Element ? e.target : e.target?.parentElement;
        const cursorEl = el?.closest?.('[data-cursor]');
        if (cursorEl) {
            g.drag = cursorEl.getAttribute('data-cursor');
            g.didDrag = true;
            return;
        }
        g.pan = { x: e.clientX, y: e.clientY };
    }

    _onFlysightGraphPointerMove(e) {
        const g = this._flysightGraph;
        if (!g) return;
        if (g.pointers?.has(e.pointerId)) {
            g.pointers.set(e.pointerId, { x: e.clientX, y: e.clientY });
        }
        if (g.pointers?.size >= 2) {
            if (!g.pinch) g.pinch = this._flysightGraphPinchSnapshot();
            this._applyFlysightGraphPinch();
            return;
        }
        if (g.drag) {
            const idx = this._flysightGraphIndexFromClientX(e.clientX);
            if (g.drag === 'a') g.idxA = Math.max(0, Math.min(idx, g.idxB - 1));
            else g.idxB = Math.min(g.samples.length - 1, Math.max(idx, g.idxA + 1));
            this._drawFlysightGraph();
            return;
        }
        if (g.pan && this._isFlysightGraphZoomed()) {
            const dx = e.clientX - g.pan.x;
            const dy = e.clientY - g.pan.y;
            if (!dx && !dy) return;
            g.pan.x = e.clientX;
            g.pan.y = e.clientY;
            if (Math.hypot(dx, dy) > 1) {
                g.didPan = true;
                document.getElementById('flysightGraphRoot')?.classList.add('is-panning');
            }
            this._panFlysightGraphByClientDelta(dx, dy);
            this._drawFlysightGraph();
        }
    }

    _onFlysightGraphPointerUp(e) {
        const g = this._flysightGraph;
        if (!g) return;
        if (e) g.pointers?.delete(e.pointerId);
        if ((g.pointers?.size || 0) < 2) g.pinch = null;
        if (g.pointers?.size) {
            g.drag = null;
            g.pan = null;
            document.getElementById('flysightGraphRoot')?.classList.remove('is-panning');
            return;
        }
        const canResetTap = e && e.pointerType !== 'mouse' && !g.didPan && !g.didDrag && !g.didPinch;
        this._clearFlysightGraphInteraction();
        if (!canResetTap) return;
        const now = Date.now();
        const dx = e.clientX - (g.lastTapX || 0);
        const dy = e.clientY - (g.lastTapY || 0);
        if (now - (g.lastTapAt || 0) < 320 && Math.hypot(dx, dy) < 28) {
            this._resetFlysightGraphZoom();
            g.lastTapAt = 0;
        } else {
            g.lastTapAt = now;
            g.lastTapX = e.clientX;
            g.lastTapY = e.clientY;
        }
    }

    _onFlysightGraphWheel(e) {
        const g = this._flysightGraph;
        if (!g) return;
        e.preventDefault();
        let dy = e.deltaY;
        if (e.deltaMode === 1) dy *= 16;
        else if (e.deltaMode === 2) dy *= 400;
        const spanFactor = Math.exp(Math.max(-0.35, Math.min(0.35, dy * 0.0015)));
        this._zoomFlysightGraphAtClient(e.clientX, e.clientY, spanFactor);
        this._drawFlysightGraph();
    }

    /** Canopy-level previous jumps plus all lineset previous jumps (not including logged jumps). */
    _canopyPreAppTotal(canopy) {
        const canopyPre = Number(canopy?.previousJumps);
        const canopyPart = Number.isFinite(canopyPre) ? canopyPre : 0;
        const linesetPart = (canopy?.linesets || []).reduce((sum, ls) => {
            const n = Number(ls?.previousJumps);
            return sum + (Number.isFinite(n) ? n : 0);
        }, 0);
        return canopyPart + linesetPart;
    }

    /** True when any canopy/lineset/harness has a non-zero previous-jumps value. */
    _hasEquipmentPreviousJumps() {
        if (this.canopies.some(c => this._canopyPreAppTotal(c) !== 0)) return true;
        return this.harnesses.some(h => {
            const n = Number(h?.previousJumps);
            return Number.isFinite(n) && n !== 0;
        });
    }

    /**
     * Canopy totals for the statistics page (equipment order: active then archived).
     * Includes canopies with only previous jumps and no logged jumps.
     */
    _buildCanopyTotalsStats() {
        const canopiesForTotals = [...this.canopies].sort((a, b) => !!a.archived - !!b.archived);
        return canopiesForTotals.map(canopy => {
            const logged = this.jumps.filter(j => j.equipment === canopy.id).length;
            const preApp = this._canopyPreAppTotal(canopy);
            return {
                name: canopy.name,
                count: logged + preApp,
                logged,
                preApp,
                archived: !!canopy.archived
            };
        }).filter(s => s.logged > 0 || s.preApp !== 0);
    }

    /** Harness rows for the statistics page (logged snapshots + harness.previousJumps). */
    _buildHarnessStats() {
        const harnessStats = [];
        this.harnesses.forEach(h => {
            if (!h?.id) return;
            const hid = h.id;
            const logged = this.jumps.filter(j => this._normalizeHarnessId(j.harnessId) === hid).length;
            const preApp = h.previousJumps ?? 0;
            harnessStats.push({
                id: hid,
                name: h.name,
                count: logged + preApp,
                logged,
                preApp,
                archived: !!h.archived
            });
        });
        return harnessStats;
    }

    renderStats() {
        const container = document.getElementById('statsContent');
        
        if (this.jumps.length === 0 && !this._hasEquipmentPreviousJumps()) {
            container.innerHTML = '<p class="no-items">No jumps logged yet.</p>';
            this._updateJumpsYearSummary();
            return;
        }
        
        // Build canopy/lineset stats (replaces old rig stats)
        const linesetStats = [];
        this.canopies.forEach(canopy => {
            (canopy.linesets || []).forEach(ls => {
                const logged = this.jumps.filter(j => j.equipment === canopy.id && j.linesetNumber === ls.number).length;
                const preApp = ls.previousJumps ?? 0;
                const total = logged + preApp;
                const hybridSuffix = ls.hybrid ? ' (Hybrid)' : '';
                const { orangeThreshold, redThreshold } = this._effectiveLinesetStatOrangeRed(ls);
                linesetStats.push({
                    name: `${canopy.name} — Lineset #${ls.number}${hybridSuffix}`,
                    count: total,
                    logged,
                    preApp,
                    archived: canopy.archived || ls.archived,
                    hybrid: ls.hybrid || false,
                    orangeThreshold,
                    redThreshold
                });
            });
        });
        
        const activeStats = linesetStats.filter(s => !s.archived && (s.logged > 0 || s.preApp !== 0));
        const archivedStats = linesetStats.filter(s => s.archived);
        const sortedStats = this.showArchivedStats ? [...activeStats, ...archivedStats] : activeStats;
        
        const hasArchivedLinesets = archivedStats.length > 0;
        const archivedTotal = archivedStats.length;
        const archivedBtnLabel = this.showArchivedStats ? 'Hide Archived' : `Show Archived (${archivedTotal})`;
        const archivedToggleBtn = hasArchivedLinesets
            ? `<button type="button" class="btn-secondary btn-sm" onclick="window.logbook.toggleArchivedStats()">${archivedBtnLabel}</button>`
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
                const redThreshold = Math.max(stat.redThreshold, 1);
                const orangeThreshold = stat.orangeThreshold;
                const percentage = Math.min((stat.count / redThreshold) * 100, 100);
                let barColorClass = '';
                if (stat.count >= redThreshold) barColorClass = 'stat-fill-red';
                else if (stat.count >= orangeThreshold) barColorClass = 'stat-fill-orange';
                const breakdown = stat.preApp !== 0
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

        // Add canopy aggregate statistics: same order as equipment (non-archived first, then archived;
        // within each group, order matches the canopies list / sortOrder — see renderCanopiesWithLinesets).
        const canopyTotalsArrayAll = this._buildCanopyTotalsStats();
        const hasArchivedCanopyTotals = canopyTotalsArrayAll.some(s => s.archived);
        const canopyTotalsArray = this.showArchivedCanopyTotals
            ? canopyTotalsArrayAll
            : canopyTotalsArrayAll.filter(s => !s.archived);
        const archivedCanopyTotalsCount = canopyTotalsArrayAll.filter(s => s.archived).length;
        const canopyTotalsArchivedBtnLabel = this.showArchivedCanopyTotals
            ? 'Hide Archived'
            : `Show Archived (${archivedCanopyTotalsCount})`;
        const canopyTotalsHeaderExtra = hasArchivedCanopyTotals
            ? `<button type="button" class="btn-secondary btn-sm" onclick="window.logbook.toggleArchivedCanopyTotals()">${canopyTotalsArchivedBtnLabel}</button>`
            : '';
        html += this.renderOrderedComponentStats('Canopy Totals', canopyTotalsArray, canopyTotalsHeaderExtra);

        // Harness stats (from jump.harnessId snapshots + harness.previousJumps).
        // Bar uses default fill only; width scales to the busiest harness (like Canopy Totals), not lineset orange/red thresholds.
        const harnessStats = this._buildHarnessStats();
        const activeHarnessStats = harnessStats.filter(s => !s.archived && (s.logged > 0 || s.preApp !== 0));
        const archivedHarnessStats = harnessStats.filter(s => s.archived);
        const sortedHarnessStats = this.showArchivedHarnessStats
            ? [...activeHarnessStats, ...archivedHarnessStats]
            : activeHarnessStats;

        const hasArchivedHarnesses = this.harnesses.some(h => h?.archived);
        const archivedHarnessCount = archivedHarnessStats.length;
        const harnessArchivedBtnLabel = this.showArchivedHarnessStats
            ? 'Hide Archived'
            : `Show Archived (${archivedHarnessCount})`;
        const harnessHeaderExtra = hasArchivedHarnesses
            ? `<button type="button" class="btn-secondary btn-sm" onclick="window.logbook.toggleArchivedHarnessStats()">${harnessArchivedBtnLabel}</button>`
            : '';

        html += `
            <div class="stats-section">
                <div class="stats-section-header">
                    <h3>Harness</h3>
                    ${harnessHeaderExtra}
                </div>
                <p class="stats-harness-hint" style="color:#888;font-size:12px;margin:0 0 8px 0;">Counts use harness saved on each jump (from the canopy's harness assignment when logged). Tap a harness row for a pie chart of logged jumps per canopy.</p>
                <div class="stats-list" id="harnessStatsList">
        `;
        if (sortedHarnessStats.length > 0) {
            const maxHarnessCount = Math.max(...sortedHarnessStats.map(s => s.count), 1);
            sortedHarnessStats.forEach(stat => {
                const percentage = stat.count > 0 ? Math.min((stat.count / maxHarnessCount) * 100, 100) : 0;
                // Default blue/green bar only — never orange/red from lineset thresholds; never red below 5000 jumps.
                const barColorClass = '';
                const breakdown = stat.preApp !== 0
                    ? `${stat.count} total (${stat.logged} logged + ${stat.preApp} pre-app)`
                    : `${stat.count} jumps`;
                html += `
                    <div class="stat-item stat-item-harness${stat.archived ? ' archived' : ''}"
                        role="button" tabindex="0" data-harness-id="${this.escapeHtml(stat.id)}"
                        title="Show jumps per canopy">
                        <div class="stat-info stat-info-stacked">
                            <span class="stat-name">${this.escapeHtml(stat.name)} ${stat.archived ? '(Archived)' : ''}</span>
                            <span class="stat-count">${breakdown}</span>
                        </div>
                        <div class="stat-bar">
                            <div class="stat-fill ${barColorClass}" style="width: ${percentage}%"></div>
                        </div>
                    </div>
                `;
            });
        } else {
            html += '<p class="no-items">No harness statistics yet. Assign a harness to a canopy (Equipment) and log jumps, or add pre-app jumps on the harness.</p>';
        }
        html += '</div></div>';

        container.innerHTML = html;
        this._bindHarnessStatsPieClicks(container);
        this._updateJumpsYearSummary();
    }
    
    toggleArchivedStats() {
        this.showArchivedStats = !this.showArchivedStats;
        this.renderStats();
    }

    toggleArchivedCanopyTotals() {
        this.showArchivedCanopyTotals = !this.showArchivedCanopyTotals;
        this.renderStats();
    }

    toggleArchivedHarnessStats() {
        this.showArchivedHarnessStats = !this.showArchivedHarnessStats;
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

    // Like renderComponentStats but accepts a pre-ordered array { name, count, archived? }
    // so the display order is controlled by the caller (not sorted by count).
    // Optional headerExtra: HTML for the right side of stats-section-header (e.g. show-archived button).
    renderOrderedComponentStats(title, statsArray, headerExtra = '') {
        let html = `
            <div class="stats-section">
                <div class="stats-section-header">
                    <h3>${title}</h3>
                    ${headerExtra}
                </div>
                <div class="stats-list">
        `;

        if (statsArray.length > 0) {
            const maxCount = Math.max(...statsArray.map(s => s.count), 1);
            statsArray.forEach(stat => {
                const percentage = stat.count > 0 ? Math.min((stat.count / maxCount) * 100, 100) : 0;
                const preApp = stat.preApp ?? 0;
                const hasBreakdown = preApp !== 0;
                const breakdown = hasBreakdown
                    ? `${stat.count} total (${stat.logged} logged + ${preApp} pre-app)`
                    : `${stat.count} jumps`;
                html += `
                    <div class="stat-item${stat.archived ? ' archived' : ''}">
                        <div class="stat-info${hasBreakdown ? ' stat-info-stacked' : ''}">
                            <span class="stat-name">${stat.name}${stat.archived ? ' (Archived)' : ''}</span>
                            <span class="stat-count">${breakdown}</span>
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

    canShareBackupFile() {
        try {
            if (!navigator.share) return false;
            if (!navigator.canShare) return true;
            return navigator.canShare({
                files: [new File(['{}'], 'share-test.json', { type: 'application/json' })]
            });
        } catch {
            return false;
        }
    }

    handleExportClick() {
        if (!this.hasExportableData()) {
            this.showMessage('No data to export', 'error');
            return;
        }
        if (this.canShareBackupFile()) {
            this.showExportChoiceModal();
            return;
        }
        this.exportData();
    }

    showExportChoiceModal() {
        const modal = document.getElementById('exportChoiceModal');
        if (modal) modal.style.display = 'block';
    }

    closeExportChoiceModal() {
        const modal = document.getElementById('exportChoiceModal');
        if (modal) modal.style.display = 'none';
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
            if (typeof ExternalCsvImport !== 'undefined'
                && ExternalCsvImport.isExternalSkydivingLogbookCsv(text)) {
                await ExternalCsvImport.importExternalLogbookCsv(this, text);
                return;
            }
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

    /**
     * Merge import: keep local jumps and equipment; add/update from import. No deletions.
     * Rows are matched by `jumpId` when set, otherwise by legacy numeric `id` (same key twice drops a duplicate row).
     * After import, `ensureJumpIds()` assigns missing `jumpId` and fixes duplicate/missing local `id` for IndexedDB.
     */
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
        if (this.settings.recentJumpsGroupByMonth === undefined) this.settings.recentJumpsGroupByMonth = false;
        if (this.settings.autoDetectDropZone === undefined) this.settings.autoDetectDropZone = true;
        if (this.settings.resequenceJumpsFromStartingNumber === undefined) {
            this.settings.resequenceJumpsFromStartingNumber = true;
        }
        if (importJumps.length && SkydivingLogbook.importJumpsHaveExplicitNumbers(importJumps)) {
            this.settings.resequenceJumpsFromStartingNumber = false;
        }
        this.normalizeNavSettings();
        this.applyNavVisibility();
        if (!this.settings.visibleNavViews.includes(this.currentView)) {
            this.showView(this.settings.startView);
        }

        this.canopies.forEach(canopy => {
            if (!Number.isFinite(Number(canopy.previousJumps))) canopy.previousJumps = 0;
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
        this.applyAutoDetectDropZoneUi(false);
        this.showMessage('Data merged successfully!', 'success');
    }

    /** Replace all: replace jumps and equipment with import file; merge settings. Local-only data is removed. */
    applyImportReplace(payload) {
        const importJumps = Array.isArray(payload.jumps) ? payload.jumps : [];
        this.jumps = importJumps;
        this.harnesses = Array.isArray(payload.harnesses) ? payload.harnesses : [];
        this.canopies = Array.isArray(payload.canopies) ? payload.canopies : [];
        this.locations = Array.isArray(payload.locations) ? payload.locations : [];

        if (payload.settings && typeof payload.settings === 'object') {
            this.settings = { ...this.settings, ...payload.settings };
        }
        if (this.settings.recentJumpsDays === undefined) this.settings.recentJumpsDays = 16;
        if (this.settings.recentJumpsGroupByMonth === undefined) this.settings.recentJumpsGroupByMonth = false;
        if (this.settings.autoDetectDropZone === undefined) this.settings.autoDetectDropZone = true;
        if (this.settings.resequenceJumpsFromStartingNumber === undefined) {
            this.settings.resequenceJumpsFromStartingNumber = true;
        }
        if (importJumps.length && SkydivingLogbook.importJumpsHaveExplicitNumbers(importJumps)) {
            this.settings.resequenceJumpsFromStartingNumber = false;
        }
        this.normalizeNavSettings();
        this.applyNavVisibility();
        if (!this.settings.visibleNavViews.includes(this.currentView)) {
            this.showView(this.settings.startView);
        }

        this.canopies.forEach(canopy => {
            if (!Number.isFinite(Number(canopy.previousJumps))) canopy.previousJumps = 0;
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
        this.applyAutoDetectDropZoneUi(false);
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

    showMessage(message, type = 'info', durationMs = 3000) {
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
        }, durationMs);
    }

    showSyncConflictModal(conflictData) {
        this._pendingSyncConflict = conflictData;
        const modal = document.getElementById('conflictModal');
        const listEl = document.getElementById('conflictList');
        const summaryEl = document.getElementById('syncConflictSummary');
        const introEl = document.getElementById('syncConflictIntro');
        if (!modal || !listEl) return;

        const jumpItems = conflictData.jumpItems || conflictData.items || [];
        const equipmentItems = conflictData.equipmentItems || [];
        const totalDiffs = jumpItems.length + equipmentItems.length;
        const localCount = conflictData.localJumps?.length ?? 0;
        const sheetCount = conflictData.sheetJumps?.length ?? 0;

        if (summaryEl) {
            summaryEl.textContent =
                `${localCount} jump(s) on this device, ${sheetCount} on the sheet. `
                + `${totalDiffs} difference(s) to review when merging `
                + `(${jumpItems.length} jump, ${equipmentItems.length} equipment). `
                + `Jump numbers will be renumbered from #${this.settings.startingJumpNumber} after sync.`;
        }
        if (introEl) {
            introEl.textContent = totalDiffs
                ? 'This device and Google Sheets both have changes since the last sync. Use local data to discard sheet changes, or review each jump and equipment difference below and merge.'
                : 'This device and Google Sheets both have changes since the last sync. Use local data to overwrite the sheet, or merge to combine both logbooks.';
        }

        listEl.innerHTML = '';
        this._appendSyncConflictSection(listEl, 'Jumps', jumpItems,
            'No individual jump differences detected — merge will combine both jump lists using the default rules.');
        this._appendSyncConflictSection(listEl, 'Equipment', equipmentItems,
            'No equipment differences detected — merge will combine harnesses, canopies, locations, and settings using the default rules.');

        modal.style.display = 'block';
    }

    _appendSyncConflictSection(listEl, title, items, emptyMessage) {
        const heading = document.createElement('h3');
        heading.className = 'sync-conflict-section-title';
        heading.textContent = title;
        listEl.appendChild(heading);

        if (!items.length) {
            const empty = document.createElement('p');
            empty.className = 'sync-conflict-empty';
            empty.style.cssText = 'color:#666;font-size:13px;margin:0 0 0.5rem;';
            empty.textContent = emptyMessage;
            listEl.appendChild(empty);
            return;
        }

        for (const item of items) {
            listEl.appendChild(this._renderSyncConflictItem(item));
        }
    }

    closeSyncConflictModal() {
        const modal = document.getElementById('conflictModal');
        if (modal) modal.style.display = 'none';
        this._pendingSyncConflict = null;
    }

    _renderSyncConflictItem(item) {
        const wrap = document.createElement('div');
        wrap.className = `conflict-item conflict-item-${item.type.replace(/_/g, '-')}`;
        wrap.dataset.conflictId = item.id;

        const title = document.createElement('h4');
        title.textContent = item.title;
        wrap.appendChild(title);

        if (item.type === 'modified') {
            const options = document.createElement('div');
            options.className = 'conflict-options';
            options.innerHTML = `
                <div class="conflict-option selected" data-choice="sheet">
                    <label>Sheet</label>
                    ${this._formatSyncConflictDetails(item.sheet, item)}
                </div>
                <div class="conflict-option" data-choice="local">
                    <label>This device</label>
                    ${this._formatSyncConflictDetails(item.local, item)}
                </div>
            `;
            options.querySelectorAll('.conflict-option').forEach(opt => {
                opt.addEventListener('click', () => {
                    options.querySelectorAll('.conflict-option').forEach(o => o.classList.remove('selected'));
                    opt.classList.add('selected');
                });
            });
            wrap.appendChild(options);
        } else {
            const options = document.createElement('div');
            options.className = 'conflict-options';
            const entity = item.local || item.sheet;
            const defaultChecked = item.type !== 'deleted_on_sheet';
            const labelText = item.type === 'deleted_on_sheet'
                ? 'Keep on this device (ignore sheet deletion)'
                : 'Include in merged logbook';
            options.innerHTML = `
                <div class="conflict-option conflict-option-keep">
                    <label class="conflict-keep-toggle">
                        <input type="checkbox" data-conflict-keep ${defaultChecked ? 'checked' : ''}>
                        <span>${labelText}</span>
                    </label>
                    ${this._formatSyncConflictDetails(entity, item)}
                </div>
            `;
            const keepRow = options.querySelector('.conflict-option-keep');
            const keepChk = options.querySelector('[data-conflict-keep]');
            const syncKeepVisual = () => {
                keepRow.classList.toggle('conflict-option-excluded', !keepChk.checked);
            };
            keepChk.addEventListener('change', syncKeepVisual);
            syncKeepVisual();
            wrap.appendChild(options);
        }

        return wrap;
    }

    _formatSyncConflictDetails(entity, item) {
        if (!entity) return '';
        if (item.equipmentKind) {
            return this._formatEquipmentConflictDetails(entity, item.equipmentKind);
        }
        return this._formatJumpConflictDetails(entity);
    }

    _formatEquipmentConflictDetails(entity, kind) {
        const esc = (s) => String(s ?? '').replace(/</g, '&lt;');
        switch (kind) {
            case 'harness':
                return `
                    <div class="conflict-detail">${esc(entity.name)}</div>
                    ${entity.notes ? `<div class="conflict-detail">${esc(entity.notes)}</div>` : ''}
                `;
            case 'canopy': {
                const linesets = Array.isArray(entity.linesets) ? entity.linesets.length : 0;
                const archived = entity.archived ? ' · archived' : '';
                return `
                    <div class="conflict-detail">${esc(entity.name)}${archived}</div>
                    <div class="conflict-detail">${linesets} lineset(s)</div>
                `;
            }
            case 'location':
                return `
                    <div class="conflict-detail">${esc(entity.name)}</div>
                    <div class="conflict-detail">${entity.lat != null && entity.lng != null ? `${entity.lat}, ${entity.lng}` : 'No coordinates'}</div>
                `;
            case 'settings':
                return `
                    <div class="conflict-detail">Starting jump #${entity.startingJumpNumber ?? 1}</div>
                    <div class="conflict-detail">Resequence jumps: ${entity.resequenceJumpsFromStartingNumber !== false ? 'yes' : 'no'}</div>
                    <div class="conflict-detail">Recent jumps window: ${entity.recentJumpsDays ?? 7} day(s)</div>
                `;
            default:
                return `<div class="conflict-detail">${esc(entity.name || entity.id || '')}</div>`;
        }
    }

    _formatJumpConflictDetails(jump) {
        if (!jump) return '';
        const date = jump.date || '—';
        const location = (jump.location || '—').replace(/</g, '&lt;');
        const notes = (jump.notes || '').trim();
        const notesHtml = notes
            ? `<div class="conflict-detail">Note: ${notes.replace(/</g, '&lt;')}</div>`
            : '';
        const ts = jump.timestamp
            ? `<div class="conflict-ts">${new Date(jump.timestamp).toLocaleString()}</div>`
            : '';
        return `
            <div class="conflict-detail">#${jump.jumpNumber} · ${date}</div>
            <div class="conflict-detail">${location}</div>
            ${notesHtml}
            ${ts}
        `;
    }

    _collectSyncConflictSelections() {
        const selections = {};
        const listEl = document.getElementById('conflictList');
        if (!listEl) return selections;

        listEl.querySelectorAll('.conflict-item').forEach(itemEl => {
            const id = itemEl.dataset.conflictId;
            const modified = itemEl.querySelector('.conflict-option[data-choice].selected');
            if (modified) {
                selections[id] = modified.dataset.choice;
                return;
            }
            const keepChk = itemEl.querySelector('[data-conflict-keep]');
            if (keepChk) {
                if (id.startsWith('deleted:')) {
                    selections[id] = keepChk.checked ? 'keep' : 'discard';
                } else {
                    selections[id] = keepChk.checked;
                }
            }
        });
        return selections;
    }

    async applySyncConflictMerge() {
        const conflict = this._pendingSyncConflict || window.SheetsAPI?._pendingConflict;
        if (!conflict || !window.SheetsAPI) return;

        const btn = document.getElementById('resolveConflictsBtn');
        if (btn) btn.disabled = true;

        try {
            const selections = this._collectSyncConflictSelections();
            const mergedJumps = window.SheetsAPI.buildMergedJumpsFromSelections(conflict, selections);
            const mergedEquipment = window.SheetsAPI.buildMergedEquipmentFromSelections(conflict, selections);
            await window.SheetsAPI.completeConflictResolution(mergedJumps, mergedEquipment);
            this.closeSyncConflictModal();
            this.showMessage('Changes merged and synced to Google Sheets.', 'success');
        } catch (err) {
            console.error('[Sync] Merge resolution failed:', err);
            this.showMessage('Failed to merge sync changes. Try again.', 'error');
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    async applySyncConflictOverwrite() {
        const conflict = this._pendingSyncConflict || window.SheetsAPI?._pendingConflict;
        if (!conflict || !window.SheetsAPI) return;

        const btn = document.getElementById('syncConflictOverwriteBtn');
        if (btn) btn.disabled = true;

        try {
            const localJumps = [...(conflict.localJumps || this.jumps)];
            this.jumps = localJumps;
            this.ensureJumpIds();
            this.forceRenumberJumpsAfterSync();
            const localEquipment = conflict.localEquipment || {
                harnesses: [...this.harnesses],
                canopies: JSON.parse(JSON.stringify(this.canopies)),
                locations: [...this.locations],
                settings: { ...this.settings }
            };
            await window.SheetsAPI.completeConflictResolution(this.jumps, localEquipment);
            this.closeSyncConflictModal();
            this.showMessage('Local logbook uploaded — sheet changes were overwritten.', 'success');
        } catch (err) {
            console.error('[Sync] Overwrite resolution failed:', err);
            this.showMessage('Failed to upload local data. Try again.', 'error');
        } finally {
            if (btn) btn.disabled = false;
        }
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
