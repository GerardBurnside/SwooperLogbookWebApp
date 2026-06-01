/**
 * One-off import for CSV exports from another logbook app (French columns).
 * Intentionally isolated from app.js — safe to delete after a successful import.
 */
(function (global) {
    'use strict';

    function parseCsvLine(line) {
        const out = [];
        let cur = '';
        let inQuotes = false;
        for (let i = 0; i < line.length; i++) {
            const c = line[i];
            if (inQuotes) {
                if (c === '"') {
                    if (line[i + 1] === '"') {
                        cur += '"';
                        i++;
                    } else {
                        inQuotes = false;
                    }
                } else {
                    cur += c;
                }
            } else if (c === '"') {
                inQuotes = true;
            } else if (c === ',') {
                out.push(cur);
                cur = '';
            } else {
                cur += c;
            }
        }
        out.push(cur);
        return out;
    }

    function isExternalSkydivingLogbookCsv(text) {
        if (!text || typeof text !== 'string') return false;
        const firstLine = text.replace(/^\uFEFF/, '').split(/\r?\n/).find(l => l.trim().length > 0);
        if (!firstLine) return false;
        const cells = parseCsvLine(firstLine).map(c => c.trim().replace(/^\uFEFF/, ''));
        return cells[0] === 'Saut #'
            && cells[1] === 'Date'
            && cells[2] === 'Zone de saut'
            && cells.includes('Équipement')
            && cells.includes('Libération')
            && cells.includes('Notes');
    }

    /** Trailing two-digit canopy size (e.g. "Petra65" → "65", "PI-71" → "71"). */
    function extractTrailingTwoDigitSize(str) {
        const s = (str || '').trim();
        const m = s.match(/(\d{2})\s*$/);
        return m ? m[1] : '';
    }

    /** Lowercase, strip spaces / hyphens / underscores / + for loose canopy matching. */
    function normalizeForCanopyMatch(str) {
        return String(str || '').toLowerCase().replace(/[\s_\-+]/g, '');
    }

    /**
     * Canopies that may match CSV equipment: trailing two-digit size on name or id,
     * then normalized segment / suggested name against canopy name and id.
     */
    function collectExternalCsvCanopyCandidates(logbook, parsed) {
        const active = (logbook.canopies || []).filter(c => !c.archived);
        const ordered = [];
        const pushUnique = (c) => {
            if (!c || ordered.some(x => x.id === c.id)) return;
            ordered.push(c);
        };

        const sizeDigits = (parsed.sizeDigits || '').trim();
        if (sizeDigits) {
            for (const c of active) {
                if (extractTrailingTwoDigitSize(c.name) === sizeDigits) pushUnique(c);
            }
            for (const c of active) {
                if (extractTrailingTwoDigitSize(c.id) === sizeDigits) pushUnique(c);
            }
        }

        const seg = (parsed.canopySegment || '').trim();
        const suggested = (parsed.suggestedCanopyName || '').trim();
        const segN = normalizeForCanopyMatch(seg);
        const sugN = normalizeForCanopyMatch(suggested);
        const MIN_LEN = 4;

        const textMatchesCanopy = (canopyName, canopyId) => {
            const nameN = normalizeForCanopyMatch(canopyName);
            const idN = normalizeForCanopyMatch(canopyId);
            if (segN.length >= MIN_LEN) {
                if (nameN === segN || idN === segN) return true;
                if (nameN.includes(segN) || segN.includes(nameN)) return true;
                if (idN.length >= MIN_LEN && (idN.includes(segN) || segN.includes(idN))) return true;
            }
            if (sugN.length >= MIN_LEN) {
                if (nameN === sugN || idN === sugN) return true;
                if (nameN.includes(sugN) || sugN.includes(nameN)) return true;
            }
            return false;
        };

        for (const c of active) {
            if (textMatchesCanopy(c.name, c.id)) pushUnique(c);
        }

        return ordered;
    }

    function parseExternalEquipmentString(raw) {
        const s = (raw || '').trim();
        const out = {
            raw: s,
            harness: '',
            canopySegment: '',
            sizeDigits: '',
            lineset: 1,
            suggestedCanopyName: ''
        };
        if (!s) return out;

        let rest = s;
        const linesetRe = /^(.*?)[-\s]+lineset\s*(\d+)\s*$/i;
        /** Same intent as "lineset1" but some exports use "-line1" (not the word "lineset"). */
        const lineNumSuffixRe = /^(.*?)[-\s]+line\s*(\d+)\s*$/i;
        const linesetL = /^(.*?)-\s*l\s*(\d+)\s*$/i;
        let m = rest.match(linesetRe);
        if (m) {
            rest = m[1].trim();
            out.lineset = parseInt(m[2], 10) || 1;
        } else {
            m = rest.match(lineNumSuffixRe);
            if (m) {
                rest = m[1].trim();
                out.lineset = parseInt(m[2], 10) || 1;
            } else {
                m = rest.match(linesetL);
                if (m) {
                    rest = m[1].trim();
                    out.lineset = parseInt(m[2], 10) || 1;
                }
            }
        }

        const plus = rest.indexOf('+');
        if (plus !== -1) {
            out.harness = rest.slice(0, plus).trim();
            rest = rest.slice(plus + 1).trim();
        }
        out.canopySegment = rest;

        const sz = rest.match(/^(.*?)(\d{2})\s*$/);
        if (sz) {
            const base = sz[1].replace(/[-_\s]+$/g, '').trim();
            out.sizeDigits = sz[2];
            out.suggestedCanopyName = base ? `${base} ${sz[2]}` : sz[2];
        } else {
            out.suggestedCanopyName = rest;
        }
        return out;
    }

    function parseExternalSkydivingCsv(text) {
        const lines = text.replace(/^\uFEFF/, '').split(/\r?\n/).filter(l => l.length > 0);
        if (lines.length < 2) {
            return { ok: false, rows: [], error: 'CSV has no data rows' };
        }
        const header = parseCsvLine(lines[0]).map(c => c.trim().replace(/^\uFEFF/, ''));
        const idx = (name) => header.indexOf(name);
        const iSaut = idx('Saut #');
        const iDate = idx('Date');
        const iZone = idx('Zone de saut');
        const iEquip = idx('Équipement');
        const iLib = idx('Libération');
        const iNotes = idx('Notes');
        if (iSaut < 0 || iDate < 0 || iEquip < 0 || iLib < 0 || iNotes < 0) {
            return { ok: false, rows: [], error: 'CSV is missing required columns' };
        }

        const rows = [];
        for (let li = 1; li < lines.length; li++) {
            const cells = parseCsvLine(lines[li]);
            const jumpNumber = parseInt((cells[iSaut] || '').trim(), 10);
            if (!Number.isFinite(jumpNumber)) continue;

            const dateRaw = (cells[iDate] || '').trim();
            let date = dateRaw;
            if (dateRaw && !/^\d{4}-\d{2}-\d{2}$/.test(dateRaw)) {
                const d = new Date(dateRaw);
                date = isNaN(d.getTime()) ? dateRaw : d.toISOString().slice(0, 10);
            }

            const location = iZone >= 0 ? (cells[iZone] || '').trim() : '';
            const equipmentRaw = (cells[iEquip] || '').trim();
            const lib = (cells[iLib] || '').trim().toLowerCase();
            const noteText = (cells[iNotes] || '').trim();
            const noteParts = [];
            if (lib === 'oui') noteParts.push('cutaway');
            if (noteText) noteParts.push(noteText);
            const notesComposed = noteParts.join('\n');

            rows.push({
                jumpNumber,
                date,
                location,
                equipmentRaw,
                notesComposed
            });
        }
        return { ok: true, rows, error: null };
    }

    function closeImportExternalCsvModal(logbook) {
        const modal = document.getElementById('importExternalCsvModal');
        if (modal) modal.style.display = 'none';
        logbook._rejectExternalCsvEquipmentStep = null;
        logbook._resolveExternalCsvEquipmentStep = null;
    }

    function createCanopyForExternalImport(logbook, name, linesetNumber) {
        const ls = Math.max(1, parseInt(linesetNumber, 10) || 1);
        const id = 'canopy_' + Date.now() + '_' + Math.random().toString(36).slice(2, 9);
        const linesets = [];
        for (let n = 1; n <= ls; n++) {
            linesets.push({ number: n, hybrid: false, previousJumps: 0, jumpCount: 0, archived: false });
        }
        const canopy = { id, name: name.trim(), notes: '', linesets };
        logbook.canopies.push(canopy);
        return id;
    }

    function ensureLinesetExistsOnCanopy(logbook, canopyId, linesetNumber) {
        const c = logbook.canopies.find(x => x.id === canopyId);
        if (!c) return;
        if (!Array.isArray(c.linesets)) c.linesets = [];
        const n = Math.max(1, parseInt(linesetNumber, 10) || 1);
        if (!c.linesets.some(ls => ls.number === n)) {
            c.linesets.push({
                number: n, hybrid: false, previousJumps: 0, jumpCount: 0, archived: false
            });
        }
        c.linesets.sort((a, b) => a.number - b.number);
    }

    function ensureLocationExistsForImport(logbook, name) {
        const trimmed = (name || '').trim();
        if (!trimmed) return;
        const exists = logbook.locations.some(
            loc => loc.name.toLowerCase() === trimmed.toLowerCase()
        );
        if (exists) return;
        const newId = 'loc_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
        const newLoc = { id: newId, name: trimmed, lat: null, lng: null };
        logbook.locations.push(newLoc);
        logbook.geocodeLocation(newLoc);
    }

    function escapeHtml(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function mergeExternalCsvRowIntoJump(logbook, jump, row, equipResolution) {
        const addNotes = (row.notesComposed || '').trim();
        if (addNotes) {
            const cur = (jump.notes || '').trim();
            if (!cur) {
                jump.notes = addNotes;
            } else if (!cur.includes(addNotes)) {
                jump.notes = `${cur}\n${addNotes}`;
            }
        }
        if (!String(jump.date || '').trim() && row.date) {
            jump.date = row.date;
        }
        if (!String(jump.location || '').trim() && row.location) {
            jump.location = row.location;
            ensureLocationExistsForImport(logbook, row.location);
        }
        const eq = (row.equipmentRaw || '').trim();
        const res = equipResolution.get(eq) || { type: 'none' };
        if (!String(jump.equipment || '').trim() && res.type === 'existing' && res.canopyId) {
            jump.equipment = res.canopyId;
            jump.linesetNumber = res.lineset || 1;
        }
    }

    function promptExternalCsvEquipmentMapping(logbook, ctx) {
        const {
            equipRaw,
            parsed,
            index,
            total,
            candidates,
            toImportCount,
            toMergeCount,
            skippedFileDup,
            sessionCanopyIds
        } = ctx;
        const sessionIds = sessionCanopyIds instanceof Set ? sessionCanopyIds : new Set();

        return new Promise((resolve, reject) => {
            const modal = document.getElementById('importExternalCsvModal');
            const intro = document.getElementById('importExternalCsvIntro');
            const prog = document.getElementById('importExternalCsvEquipProgress');
            const rawEl = document.getElementById('importExternalCsvEquipRaw');
            const parsedEl = document.getElementById('importExternalCsvEquipParsed');
            const radioList = document.getElementById('importExternalCsvRadioList');
            const newWrap = document.getElementById('importExternalCsvNewWrap');
            const nameInput = document.getElementById('importExternalCsvNewCanopyName');

            if (!modal || !radioList || !nameInput) {
                reject(new Error('Import UI missing'));
                return;
            }

            const finishReject = () => {
                closeImportExternalCsvModal(logbook);
                reject(Object.assign(new Error('cancel'), { code: 'EXT_CSV_CANCEL' }));
            };

            logbook._rejectExternalCsvEquipmentStep = finishReject;

            const PICK_FROM_LIST_VALUE = '__pick__';
            const MANUAL_CANOPY_SELECT_ID = 'importExternalCsvManualCanopySelect';

            logbook._resolveExternalCsvEquipmentStep = () => {
                const picked = modal.querySelector('input[name="extCsvEquipPick"]:checked');
                if (!picked) {
                    logbook.showMessage('Choose an existing canopy or add new.', 'error');
                    return;
                }
                const val = picked.value;
                if (val === '__new__') {
                    const name = nameInput.value.trim();
                    if (!name) {
                        logbook.showMessage('Enter a name for the new canopy.', 'error');
                        return;
                    }
                    closeImportExternalCsvModal(logbook);
                    resolve({ type: 'new', name, lineset: parsed.lineset });
                    return;
                }
                if (val === PICK_FROM_LIST_VALUE) {
                    const dd = modal.querySelector('#' + MANUAL_CANOPY_SELECT_ID);
                    const canopyId = (dd && dd.value) || '';
                    if (!canopyId) {
                        logbook.showMessage('Select a canopy from the list.', 'error');
                        return;
                    }
                    closeImportExternalCsvModal(logbook);
                    resolve({ type: 'existing', canopyId, lineset: parsed.lineset });
                    return;
                }
                if (val.startsWith('existing:')) {
                    const canopyId = val.slice('existing:'.length);
                    closeImportExternalCsvModal(logbook);
                    resolve({ type: 'existing', canopyId, lineset: parsed.lineset });
                }
            };

            if (index === 1 && intro) {
                const parts = [];
                if (toImportCount) parts.push(`${toImportCount} new jump(s) will be added.`);
                if (toMergeCount) {
                    parts.push(`${toMergeCount} row(s) match existing jump # — they will be merged (notes appended; empty fields filled only).`);
                }
                if (skippedFileDup) parts.push(`${skippedFileDup} duplicate row(s) in the file were skipped.`);
                intro.textContent = parts.join(' ');
            }

            prog.textContent = `Equipment mapping: step ${index} of ${total}`;
            rawEl.textContent = equipRaw || '(empty)';
            const pBits = [];
            if (parsed.harness) pBits.push(`Harness prefix: ${parsed.harness}`);
            if (parsed.sizeDigits) pBits.push(`Detected size: ${parsed.sizeDigits}`);
            pBits.push(`Lineset from file: ${parsed.lineset}`);
            parsedEl.textContent = pBits.join(' · ');

            nameInput.value = parsed.suggestedCanopyName || parsed.canopySegment || equipRaw;

            const allCanopies = (logbook.canopies || [])
                .filter(c => !c.archived)
                .slice()
                .sort((a, b) => (a.name || '').localeCompare(b.name || '', undefined, { sensitivity: 'base' }));

            const linesetHint = `(lineset #${parsed.lineset} from file name: lineset / line / -l)`;
            const pickOpts = ['<option value="">Select canopy…</option>']
                .concat(allCanopies.map(c =>
                    `<option value="${escapeHtml(c.id)}">${escapeHtml(c.name)}</option>`
                ));

            const radios = [];
            radios.push(
                `<label class="import-external-csv-manual-pick">`
                + `<input type="radio" name="extCsvEquipPick" value="${PICK_FROM_LIST_VALUE}"> `
                + `<span class="import-external-csv-manual-pick-label">Choose canopy from list</span> `
                + `<select id="${MANUAL_CANOPY_SELECT_ID}" class="import-external-csv-manual-select" disabled>`
                + pickOpts.join('')
                + `</select> <span class="import-external-csv-sub">${escapeHtml(linesetHint)}</span>`
                + `</label>`
            );
            if (candidates.length) {
                radios.push('<p class="import-external-csv-suggested-hint">Suggested matches</p>');
            }
            for (const c of candidates) {
                const fromSession = sessionIds.has(c.id);
                const sub = fromSession
                    ? `(lineset #${parsed.lineset} from file — same canopy you added in this import)`
                    : `(lineset #${parsed.lineset})`;
                radios.push(
                    `<label><input type="radio" name="extCsvEquipPick" value="existing:${escapeHtml(c.id)}"> `
                    + `${escapeHtml(c.name)} <span class="import-external-csv-sub">${escapeHtml(sub)}</span></label>`
                );
            }
            radios.push(
                '<label><input type="radio" name="extCsvEquipPick" value="__new__"> '
                + 'Add new canopy… <span class="import-external-csv-sub">(edit name below)</span></label>'
            );
            radioList.innerHTML = radios.join('');

            const manualSelect = radioList.querySelector('#' + MANUAL_CANOPY_SELECT_ID);

            const syncNewState = () => {
                const sel = modal.querySelector('input[name="extCsvEquipPick"]:checked');
                const isNew = sel && sel.value === '__new__';
                const isPickList = sel && sel.value === PICK_FROM_LIST_VALUE;
                nameInput.disabled = !isNew;
                if (newWrap) newWrap.style.opacity = isNew ? '1' : '0.65';
                if (manualSelect) {
                    manualSelect.disabled = !isPickList;
                }
            };

            radioList.querySelectorAll('input[name="extCsvEquipPick"]').forEach(r => {
                r.addEventListener('change', syncNewState);
            });

            if (manualSelect) {
                const checkPickListRadio = () => {
                    const inp = radioList.querySelector(
                        `input[name="extCsvEquipPick"][value="${PICK_FROM_LIST_VALUE}"]`
                    );
                    if (inp) inp.checked = true;
                    syncNewState();
                };
                manualSelect.addEventListener('mousedown', checkPickListRadio);
                manualSelect.addEventListener('focus', checkPickListRadio);
            }

            const pickExistingValue = (wantVal) => {
                const inputs = radioList.querySelectorAll('input[name="extCsvEquipPick"]');
                let found = false;
                inputs.forEach(inp => {
                    const on = inp.value === wantVal;
                    inp.checked = on;
                    if (on) found = true;
                });
                return found;
            };

            if (candidates.length > 0) {
                const preferSession = candidates.find(c => sessionIds.has(c.id));
                const defaultCandidate = preferSession || candidates[0];
                pickExistingValue(`existing:${defaultCandidate.id}`);
                if (manualSelect) {
                    manualSelect.value = defaultCandidate.id;
                }
            } else {
                pickExistingValue(PICK_FROM_LIST_VALUE);
                if (manualSelect) {
                    manualSelect.value = '';
                }
            }
            if (candidates.length === 0 && allCanopies.length === 0) {
                pickExistingValue('__new__');
                syncNewState();
            } else {
                syncNewState();
            }

            modal.style.display = 'block';
        });
    }

    function normalizeDateYmd(dateVal) {
        if (dateVal == null || dateVal === '') return '';
        const s = String(dateVal);
        if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
        const d = new Date(s);
        return isNaN(d.getTime()) ? '' : d.toISOString().slice(0, 10);
    }

    /**
     * Missing jump # range strictly below the imported CSV block: (maxBelow+1) .. (csvMin-1).
     */
    function computeJumpNumberGapInfo(logbook, csvMin) {
        if (csvMin == null || !Number.isFinite(csvMin)) return null;
        const below = logbook.jumps.filter(j => j.jumpNumber < csvMin).map(j => j.jumpNumber);
        const maxBelow = below.length > 0 ? Math.max(...below) : null;
        const start = maxBelow != null ? maxBelow + 1 : (logbook.settings?.startingJumpNumber ?? 1);
        const end = csvMin - 1;
        if (end < start) return null;
        const maxBelowJump = maxBelow != null
            ? logbook.jumps.find(j => j.jumpNumber === maxBelow)
            : null;
        const suggestedDate = normalizeDateYmd(maxBelowJump?.date);
        return {
            gapKind: 'belowCsvBlock',
            gapStart: start,
            gapEnd: end,
            gapCount: end - start + 1,
            csvMin,
            maxBelow,
            suggestedDate
        };
    }

    /**
     * Missing jump # range strictly above the CSV file's highest # and strictly below
     * the smallest jump # that was in the log before import: (csvMax+1) .. (preImportDbMin-1).
     * Typical case: importing an older CSV block (e.g. #1–100) while the log already had #500+.
     */
    function computeJumpNumberGapBetweenCsvMaxAndPreDbMin(parsedCsvRows, preImportDbMinJump) {
        if (!parsedCsvRows || !parsedCsvRows.length) return null;
        if (preImportDbMinJump == null || !Number.isFinite(preImportDbMinJump)) return null;
        const csvMax = Math.max(...parsedCsvRows.map(r => r.jumpNumber));
        if (!Number.isFinite(csvMax)) return null;
        const start = csvMax + 1;
        const end = preImportDbMinJump - 1;
        if (end < start) return null;
        let suggestedDate = '';
        for (const r of parsedCsvRows) {
            if (r.jumpNumber === csvMax) {
                const d = normalizeDateYmd(r.date);
                if (d) suggestedDate = d;
            }
        }
        return {
            gapKind: 'betweenCsvMaxAndPreDbMin',
            gapStart: start,
            gapEnd: end,
            gapCount: end - start + 1,
            csvMax,
            preImportDbMinJump,
            suggestedDate
        };
    }

    function buildGapFillCanopySelect(logbook) {
        const sel = document.createElement('select');
        sel.className = 'import-gap-fill-canopy';
        const ph = document.createElement('option');
        ph.value = '';
        ph.textContent = 'Select canopy';
        sel.appendChild(ph);
        const canopies = (logbook.canopies || []).slice().sort((a, b) => {
            const aArch = a.archived ? 1 : 0;
            const bArch = b.archived ? 1 : 0;
            if (aArch !== bArch) return aArch - bArch;
            return (a.name || '').localeCompare(b.name || '', undefined, { sensitivity: 'base' });
        });
        for (const canopy of canopies) {
            const opt = document.createElement('option');
            opt.value = canopy.id;
            opt.textContent = canopy.archived ? `${canopy.name} (archived)` : canopy.name;
            sel.appendChild(opt);
        }
        return sel;
    }

    /** Repopulate lineset dropdown for a segment row; disables until a canopy is chosen. */
    function populateGapFillLinesetSelect(logbook, linesetSel, canopyId) {
        while (linesetSel.firstChild) {
            linesetSel.removeChild(linesetSel.firstChild);
        }
        const ph = document.createElement('option');
        ph.value = '';
        ph.textContent = 'Select lineset';
        linesetSel.appendChild(ph);
        if (!canopyId) {
            linesetSel.disabled = true;
            return;
        }
        const canopy = (logbook.canopies || []).find(c => c.id === canopyId);
        if (!canopy) {
            linesetSel.disabled = true;
            return;
        }
        linesetSel.disabled = false;
        /** Include archived linesets so gap-fill can match historical gear (still sorted by #). */
        const lss = (canopy.linesets || []).slice().sort((a, b) => a.number - b.number);
        for (const ls of lss) {
            const opt = document.createElement('option');
            opt.value = String(ls.number);
            const hybridTag = ls.hybrid ? ' (Hybrid)' : '';
            const archivedTag = ls.archived ? ' (archived)' : '';
            opt.textContent = `Lineset #${ls.number}${hybridTag}${archivedTag}`;
            linesetSel.appendChild(opt);
        }
        if (lss.length === 1) {
            linesetSel.value = String(lss[0].number);
        }
    }

    function appendGapFillSegmentRow(logbook, segmentsRoot, onChange) {
        const row = document.createElement('div');
        row.className = 'import-gap-fill-row';
        row.dataset.role = 'segment';

        const countLab = document.createElement('label');
        countLab.textContent = 'Jumps';
        const countIn = document.createElement('input');
        countIn.type = 'number';
        countIn.min = '1';
        countIn.step = '1';
        countIn.className = 'import-gap-fill-count';
        countIn.value = '';
        countLab.appendChild(countIn);

        const canopyWrap = document.createElement('div');
        canopyWrap.className = 'import-gap-fill-canopy-wrap';
        const canopyLab = document.createElement('label');
        canopyLab.textContent = 'Canopy';
        const canopySel = buildGapFillCanopySelect(logbook);
        canopyLab.appendChild(canopySel);
        canopyWrap.appendChild(canopyLab);

        const linesetWrap = document.createElement('div');
        linesetWrap.className = 'import-gap-fill-lineset-wrap';
        const linesetLab = document.createElement('label');
        linesetLab.textContent = 'Lineset';
        const linesetSel = document.createElement('select');
        linesetSel.className = 'import-gap-fill-lineset';
        linesetSel.disabled = true;
        populateGapFillLinesetSelect(logbook, linesetSel, '');

        linesetLab.appendChild(linesetSel);
        linesetWrap.appendChild(linesetLab);

        const dateLab = document.createElement('label');
        dateLab.textContent = 'Date (optional)';
        const dateIn = document.createElement('input');
        dateIn.type = 'date';
        dateIn.className = 'import-gap-fill-date';
        dateLab.appendChild(dateIn);

        const rm = document.createElement('button');
        rm.type = 'button';
        rm.className = 'btn-secondary import-gap-fill-remove';
        rm.textContent = 'Remove';
        rm.addEventListener('click', () => {
            if (segmentsRoot.querySelectorAll('.import-gap-fill-row').length <= 1) {
                return;
            }
            row.remove();
            onChange();
        });

        row.appendChild(countLab);
        row.appendChild(canopyWrap);
        row.appendChild(linesetWrap);
        row.appendChild(dateLab);
        row.appendChild(rm);

        countIn.addEventListener('input', onChange);
        countIn.addEventListener('change', onChange);
        canopySel.addEventListener('change', () => {
            populateGapFillLinesetSelect(logbook, linesetSel, canopySel.value);
            onChange();
        });
        linesetSel.addEventListener('change', onChange);
        dateIn.addEventListener('change', onChange);

        segmentsRoot.appendChild(row);
        onChange();
    }

    function readGapFillSegments(gap, defaultDateInput, segmentsRoot) {
        const defaultDate = normalizeDateYmd(defaultDateInput.value)
            || gap.suggestedDate
            || new Date().toISOString().slice(0, 10);
        const rows = segmentsRoot.querySelectorAll('.import-gap-fill-row');
        const segments = [];
        let sum = 0;
        for (const row of rows) {
            const c = parseInt(row.querySelector('.import-gap-fill-count')?.value, 10);
            const dateRaw = row.querySelector('.import-gap-fill-date')?.value || '';
            if (!Number.isFinite(c) || c <= 0) continue;
            const canopyId = row.querySelector('select.import-gap-fill-canopy')?.value || '';
            const linesetStr = row.querySelector('select.import-gap-fill-lineset')?.value || '';
            if (!canopyId) {
                return { error: 'Each segment with a jump count needs a canopy selected.' };
            }
            if (!linesetStr) {
                return { error: 'Each segment with a jump count needs a lineset selected.' };
            }
            const linesetNumber = parseInt(linesetStr, 10) || 1;
            const date = normalizeDateYmd(dateRaw) || defaultDate;
            segments.push({ count: c, equipment: canopyId, linesetNumber, date });
            sum += c;
        }
        if (sum !== gap.gapCount) {
            return { error: `Segment jump counts must add up to ${gap.gapCount} (currently ${sum}).` };
        }
        if (segments.length === 0) {
            return { error: 'Add at least one segment with jump counts.' };
        }
        return { segments, defaultDate };
    }

    function applyGapFillJumps(logbook, gap, segments, genJumpId) {
        let num = gap.gapStart;
        let idSeq = 0;
        for (const seg of segments) {
            for (let i = 0; i < seg.count; i++) {
                const jump = {
                    id: Date.now() + Math.random() + (idSeq++ * 1e-6),
                    jumpId: genJumpId(),
                    jumpNumber: num,
                    date: seg.date,
                    location: '',
                    equipment: seg.equipment,
                    linesetNumber: seg.linesetNumber,
                    notes: '',
                    timestamp: seg.date ? `${seg.date}T12:00:00.000Z` : new Date().toISOString()
                };
                logbook.jumps.push(jump);
                num += 1;
            }
        }
    }

    function promptJumpGapFillModal(logbook, gap) {
        return new Promise((resolve) => {
            const modal = document.getElementById('importCsvGapFillModal');
            const explain = document.getElementById('importCsvGapFillExplain');
            const segmentsRoot = document.getElementById('importCsvGapFillSegments');
            const totalLine = document.getElementById('importCsvGapFillTotalLine');
            const defaultDateInput = document.getElementById('importCsvGapFillDefaultDate');
            const addBtn = document.getElementById('importCsvGapFillAddRow');
            const skipBtn = document.getElementById('importCsvGapFillSkip');
            const applyBtn = document.getElementById('importCsvGapFillApply');
            const closeX = document.getElementById('importCsvGapFillClose');

            if (!modal || !segmentsRoot || !totalLine || !defaultDateInput || !addBtn || !skipBtn || !applyBtn) {
                resolve(null);
                return;
            }

            const refreshTotals = () => {
                let sum = 0;
                segmentsRoot.querySelectorAll('.import-gap-fill-count').forEach(inp => {
                    const v = parseInt(inp.value, 10);
                    if (Number.isFinite(v) && v > 0) sum += v;
                });
                totalLine.textContent = `Assigned: ${sum} / ${gap.gapCount} jump(s) for #${gap.gapStart}–#${gap.gapEnd}`;
                let ok = sum === gap.gapCount && sum > 0;
                if (ok) {
                    segmentsRoot.querySelectorAll('.import-gap-fill-row').forEach(row => {
                        const c = parseInt(row.querySelector('.import-gap-fill-count')?.value, 10);
                        if (!Number.isFinite(c) || c <= 0) return;
                        const canopyId = row.querySelector('select.import-gap-fill-canopy')?.value;
                        const ls = row.querySelector('select.import-gap-fill-lineset')?.value;
                        if (!canopyId || !ls) ok = false;
                    });
                }
                applyBtn.disabled = !ok;
            };

            const closeModal = () => {
                modal.style.display = 'none';
                segmentsRoot.innerHTML = '';
                defaultDateInput.value = '';
            };

            const finish = (segments) => {
                modal.removeEventListener('click', onBackdrop);
                closeModal();
                resolve(segments);
            };

            function onBackdrop(e) {
                if (e.target === modal) finish(null);
            }

            if (gap.gapKind === 'betweenCsvMaxAndPreDbMin') {
                explain.textContent = `Jump numbers #${gap.gapStart} through #${gap.gapEnd} (${gap.gapCount} jumps) are missing between the highest jump number in the CSV (#${gap.csvMax}) and your log before this import (which started at jump #${gap.preImportDbMinJump}). For each segment, set jump count, canopy, and lineset; optional per-segment date. Counts must add up exactly.`;
            } else {
                explain.textContent = `Jump numbers #${gap.gapStart} through #${gap.gapEnd} (${gap.gapCount} jumps) are missing between your existing log and the imported block (starts at #${gap.csvMin}). For each segment, set jump count, canopy, and lineset; optional per-segment date. Counts must add up exactly.`;
            }

            defaultDateInput.value = gap.suggestedDate || '';

            segmentsRoot.innerHTML = '';
            appendGapFillSegmentRow(logbook, segmentsRoot, refreshTotals);
            appendGapFillSegmentRow(logbook, segmentsRoot, refreshTotals);

            addBtn.onclick = () => appendGapFillSegmentRow(logbook, segmentsRoot, refreshTotals);
            skipBtn.onclick = () => finish(null);
            closeX.onclick = () => finish(null);
            applyBtn.onclick = () => {
                const parsed = readGapFillSegments(gap, defaultDateInput, segmentsRoot);
                if (parsed.error) {
                    logbook.showMessage(parsed.error, 'error');
                    return;
                }
                finish(parsed.segments);
            };

            modal.addEventListener('click', onBackdrop);
            modal.style.display = 'block';
            refreshTotals();
        });
    }

    function persistAfterGapFill(logbook) {
        logbook.ensureJumpIds();
        logbook.initializeCanopyLinesetJumpCounts();
        logbook.saveToLocalStorage();
        logbook.saveComponentsToLocalStorage();
        logbook.markEquipmentModified();
        logbook.updateEquipmentOptions();
        logbook.renderJumpsList();
        logbook.updateStats();
        if (typeof logbook.currentView === 'string' && logbook.currentView === 'equipment') {
            logbook.renderEquipmentView();
        }
        logbook.renderStats();
        logbook.applyAutoDetectDropZoneUi(false);

        if (typeof navigator !== 'undefined' && navigator.onLine && global.SheetsAPI?.initialized) {
            global.SheetsAPI.pushAllWithGuard();
        }
    }

    async function maybeOfferJumpGapFill(logbook, appliedJumpNumbers) {
        if (!appliedJumpNumbers.length) return;
        const csvMin = Math.min(...appliedJumpNumbers);
        const gap = computeJumpNumberGapInfo(logbook, csvMin);
        if (!gap || gap.gapCount <= 0) return;

        const segments = await promptJumpGapFillModal(logbook, gap);
        if (!segments || !segments.length) return;

        const genJumpId = () => ((typeof crypto !== 'undefined' && crypto.randomUUID)
            ? crypto.randomUUID()
            : 'jump-' + Date.now() + '-' + Math.random().toString(36).slice(2));

        applyGapFillJumps(logbook, gap, segments, genJumpId);
        persistAfterGapFill(logbook);
        logbook.showMessage(`Added ${gap.gapCount} jump(s) for numbers #${gap.gapStart}–#${gap.gapEnd}.`, 'success');
    }

    /** Gap between max jump # in the CSV file and smallest # in the log before import (editable canopy/lineset in modal). */
    async function maybeOfferJumpGapFillBetweenCsvMaxAndPreDbMin(logbook, parsedCsvRows, preImportDbMinJump) {
        const gap = computeJumpNumberGapBetweenCsvMaxAndPreDbMin(parsedCsvRows, preImportDbMinJump);
        if (!gap || gap.gapCount <= 0) return;

        const segments = await promptJumpGapFillModal(logbook, gap);
        if (!segments || !segments.length) return;

        const genJumpId = () => ((typeof crypto !== 'undefined' && crypto.randomUUID)
            ? crypto.randomUUID()
            : 'jump-' + Date.now() + '-' + Math.random().toString(36).slice(2));

        applyGapFillJumps(logbook, gap, segments, genJumpId);
        persistAfterGapFill(logbook);
        logbook.showMessage(`Added ${gap.gapCount} jump(s) for numbers #${gap.gapStart}–#${gap.gapEnd}.`, 'success');
    }

    async function importExternalLogbookCsv(logbook, text) {
        const parsedCsv = parseExternalSkydivingCsv(text);
        if (!parsedCsv.ok) {
            logbook.showMessage(parsedCsv.error || 'Could not read CSV', 'error');
            return;
        }

        const usedNumbers = new Set(logbook.jumps.map(j => j.jumpNumber));
        const dbJumpNums = new Set(usedNumbers);
        const toImport = [];
        const toMerge = [];
        let skippedFileDup = 0;
        const seenInFile = new Set();

        for (const row of parsedCsv.rows) {
            if (seenInFile.has(row.jumpNumber)) {
                skippedFileDup++;
                continue;
            }
            seenInFile.add(row.jumpNumber);
            if (dbJumpNums.has(row.jumpNumber)) {
                toMerge.push(row);
            } else {
                usedNumbers.add(row.jumpNumber);
                toImport.push(row);
            }
        }

        if (toImport.length === 0 && toMerge.length === 0) {
            logbook.showMessage(
                skippedFileDup
                    ? `Nothing to import (${skippedFileDup} duplicate row(s) in the file).`
                    : 'Nothing to import.',
                'info'
            );
            return;
        }

        const seenEq = new Set();
        const uniqueEquip = [];
        const pushEquip = (r) => {
            const k = (r.equipmentRaw || '').trim();
            if (!seenEq.has(k)) {
                seenEq.add(k);
                uniqueEquip.push(k);
            }
        };
        toImport.forEach(pushEquip);
        toMerge.forEach(pushEquip);

        const equipResolution = new Map();
        uniqueEquip.forEach(eq => {
            if (!eq) equipResolution.set(eq, { type: 'none' });
        });
        const equipNeedsPrompt = uniqueEquip.filter(eq => eq.length > 0);

        const sessionCanopyIds = new Set();

        try {
            const preImportDbMinJumpNumber = logbook.jumps.length
                ? Math.min(...logbook.jumps.map(j => j.jumpNumber))
                : null;

            for (let i = 0; i < equipNeedsPrompt.length; i++) {
                const eq = equipNeedsPrompt[i];
                const parsedEq = parseExternalEquipmentString(eq);
                let candidates = collectExternalCsvCanopyCandidates(logbook, parsedEq);
                candidates = candidates.slice().sort((a, b) => {
                    const as = sessionCanopyIds.has(a.id) ? 0 : 1;
                    const bs = sessionCanopyIds.has(b.id) ? 0 : 1;
                    if (as !== bs) return as - bs;
                    return (a.name || '').localeCompare(b.name || '');
                });

                let res = await promptExternalCsvEquipmentMapping(logbook, {
                    equipRaw: eq,
                    parsed: parsedEq,
                    index: i + 1,
                    total: equipNeedsPrompt.length,
                    candidates,
                    toImportCount: toImport.length,
                    toMergeCount: toMerge.length,
                    skippedFileDup,
                    sessionCanopyIds
                });
                if (res.type === 'new') {
                    const id = createCanopyForExternalImport(logbook, res.name, res.lineset);
                    sessionCanopyIds.add(id);
                    res = { type: 'existing', canopyId: id, lineset: res.lineset };
                }
                equipResolution.set(eq, res);
            }

            for (const eq of uniqueEquip) {
                const res = equipResolution.get(eq);
                if (res && res.type === 'existing' && res.canopyId) {
                    ensureLinesetExistsOnCanopy(logbook, res.canopyId, res.lineset);
                }
            }

            const genJumpId = () => ((typeof crypto !== 'undefined' && crypto.randomUUID)
                ? crypto.randomUUID()
                : 'jump-' + Date.now() + '-' + Math.random().toString(36).slice(2));

            let idSeq = 0;
            for (const row of toImport) {
                const eq = (row.equipmentRaw || '').trim();
                const res = equipResolution.get(eq) || { type: 'none' };
                let equipment = '';
                let linesetNumber = 1;
                if (res.type === 'existing' && res.canopyId) {
                    equipment = res.canopyId;
                    linesetNumber = res.lineset || 1;
                }
                if (row.location) ensureLocationExistsForImport(logbook, row.location);

                const jump = {
                    id: Date.now() + Math.random() + (idSeq++ * 1e-6),
                    jumpId: genJumpId(),
                    jumpNumber: row.jumpNumber,
                    date: row.date,
                    location: row.location || '',
                    equipment,
                    linesetNumber,
                    notes: row.notesComposed,
                    timestamp: row.date ? `${row.date}T12:00:00.000Z` : new Date().toISOString()
                };
                logbook.jumps.push(jump);
            }

            for (const row of toMerge) {
                const jump = logbook.jumps.find(j => j.jumpNumber === row.jumpNumber);
                if (jump) {
                    mergeExternalCsvRowIntoJump(logbook, jump, row, equipResolution);
                }
            }

            logbook.ensureJumpIds();
            logbook.initializeCanopyLinesetJumpCounts();
            logbook.saveToLocalStorage();
            logbook.saveComponentsToLocalStorage();
            logbook.markEquipmentModified();
            logbook.updateEquipmentOptions();
            logbook.renderJumpsList();
            logbook.updateStats();
            logbook.renderEquipmentView();
            logbook.renderStats();
            logbook.applyAutoDetectDropZoneUi(false);

            if (typeof navigator !== 'undefined' && navigator.onLine && global.SheetsAPI?.initialized) {
                global.SheetsAPI.pushAllWithGuard();
            }

            let msg = '';
            if (toImport.length) msg += `Imported ${toImport.length} new jump(s) from external CSV.`;
            if (toMerge.length) {
                msg += (msg ? ' ' : '')
                    + `Merged ${toMerge.length} row(s) into existing jump(s) with the same # (existing jumps were not renumbered).`;
            }
            if (skippedFileDup) msg += ` Skipped ${skippedFileDup} duplicate row(s) in file.`;
            logbook.showMessage(msg.trim(), 'success');

            const appliedNums = [...new Set([...toImport, ...toMerge].map(r => r.jumpNumber))];
            await maybeOfferJumpGapFill(logbook, appliedNums);
            await maybeOfferJumpGapFillBetweenCsvMaxAndPreDbMin(
                logbook,
                parsedCsv.rows,
                preImportDbMinJumpNumber
            );
        } catch (e) {
            if (e && e.code === 'EXT_CSV_CANCEL') {
                logbook.showMessage('Import cancelled.', 'info');
            } else {
                console.error('External CSV import failed:', e);
                logbook.showMessage('External CSV import failed.', 'error');
            }
        } finally {
            closeImportExternalCsvModal(logbook);
        }
    }

    global.ExternalCsvImport = {
        isExternalSkydivingLogbookCsv,
        parseExternalSkydivingCsv,
        parseExternalEquipmentString,
        extractTrailingTwoDigitSize,
        collectExternalCsvCanopyCandidates,
        mergeExternalCsvRowIntoJump,
        importExternalLogbookCsv,
        computeJumpNumberGapInfo,
        computeJumpNumberGapBetweenCsvMaxAndPreDbMin
    };
})(typeof window !== 'undefined' ? window : globalThis);
