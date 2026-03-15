// js/storage.js

/**
 * Hanterar all datauthållighet: att spara och ladda projektet,
 * hantera säkerhetskopior i IndexedDB och implementera ångra/gör om.
 */

// === HISTORIKHANTERING (Undo/Redo) ===
let history = [];
let historyIndex = -1;
// Håller nuvarande filhandtag (File System Access API). Kan vara null.
let currentFileHandle = null;
let currentFileName = '';
const DEFAULT_PROJECT_TITLE = 'Gantt-schema';

try {
    currentFileName = localStorage.getItem('gantt_fileName') || '';
} catch (_) {
    currentFileName = '';
}

function stripFileExtension(fileName) {
    return String(fileName || '').replace(/\.[^.]+$/, '');
}

function getProjectTitle(fallback = DEFAULT_PROJECT_TITLE) {
    const rawName = currentFileHandle?.name || currentFileName || '';
    const displayName = stripFileExtension(rawName).trim();
    return displayName || fallback;
}

function updateProjectTitle(fallback = DEFAULT_PROJECT_TITLE) {
    const title = getProjectTitle(fallback);
    document.title = title;
    return title;
}

function setCurrentProjectFile(handle = null, fileName = '') {
    currentFileHandle = handle || null;
    currentFileName = handle?.name || fileName || '';
    try {
        if (currentFileName) {
            localStorage.setItem('gantt_fileName', currentFileName);
        } else {
            localStorage.removeItem('gantt_fileName');
        }
    } catch (_) {}
    return updateProjectTitle();
}

updateProjectTitle();

/**
 * Tar en ögonblicksbild av det nuvarande tillståndet och lägger till i historiken.
 */
function pushHistory() {
    if (historyIndex < history.length - 1) {
        history = history.slice(0, historyIndex + 1);
    }
    // Använder en anpassad replacer för att hantera Date-objekt korrekt (lokal tid utan tidszonskift)
    const snapshot = JSON.stringify(activities, (key, value) => {
        if (value instanceof Date) {
            return isTimeSpecific(value) ? formatDateTimeLocal(value) : formatDateLocal(value);
        }
        return value;
    });
    if (history.length > 0 && history[historyIndex] === snapshot) {
        return;
    }
    history.push(snapshot);
    historyIndex++;
    saveToStorage();
}

/**
 * Återställer det föregående tillståndet från historiken.
 */
function undo() {
    if (historyIndex > 0) {
        historyIndex--;
        restoreStateFromHistory(historyIndex);
    }
}

/**
 * Återställer ett ångrat tillstånd från historiken.
 */
function redo() {
    if (historyIndex < history.length - 1) {
        historyIndex++;
        restoreStateFromHistory(historyIndex);
    }
}

/**
 * Hjälpfunktion för att återställa ett tillstånd och rendera om.
 */
function restoreStateFromHistory(index) {
    const state = JSON.parse(history[index]);
    activities = state.map(a => {
        const restoredAct = { ...a };
        if (restoredAct.segments) {
            restoredAct.segments = restoredAct.segments.map(s => ({
                start: parseImportedDate(s.start),
                end: parseImportedDate(s.end),
                info: s.info,
                name: s.name,
                color: s.color,
                responsibles: Array.isArray(s.responsibles) ? s.responsibles : [],
                hoursPerDay: typeof s.hoursPerDay === 'number' ? s.hoursPerDay : undefined
            }));
        }
        restoredAct.start = a.start ? parseImportedDate(a.start) : null;
        restoredAct.end = a.end ? parseImportedDate(a.end) : null;
        return restoredAct;
    });
    try { if (typeof rebuildActivityIndex === 'function') rebuildActivityIndex(); } catch(_) {}
    saveToStorage();
    if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render();
}

// Robust parser för importerade datumvärden (sträng eller Date).
function parseImportedDate(input) {
    if (!input) return null;
    if (input instanceof Date) return new Date(input); // Copy date
    if (typeof input === 'string') {
        // Om strängen innehåller tidszonsmarkör (Z eller ±HH:MM), använd inbyggda Date-parsning
        if (/Z$|[+-]\d{2}:\d{2}$/.test(input) || input.includes('T')) { 
            const parsed = new Date(input);
            return isNaN(parsed.getTime()) ? null : parsed;
        }
        // Annars använd vår lokala normalisering (stöd för 'YYYY-MM-DD' och 'YYYY-MM-DDTHH:MM')
        const nd = normalizeDate(input);
        if (nd) return nd;
        const parsed = new Date(input);
        return isNaN(parsed.getTime()) ? null : parsed;
    }
    return null;
}

/**
 * Läser in projektdata från ett JSON-objekt och återställer applikationens tillstånd.
 */
function loadProjectData(data) {
    if (!data) return;

    if (data.title) {
        updateProjectTitle(data.title);
    }

    if (data.settings) {
        if (data.settings.startDate) document.getElementById('startDate').value = data.settings.startDate;
        if (data.settings.endDate) document.getElementById('endDate').value = data.settings.endDate;
        if (typeof data.settings.showHolidays === 'boolean') document.getElementById('showHolidays').checked = data.settings.showHolidays;
        if (data.settings.country) document.getElementById('dateFormat').value = data.settings.country;
        
        // Spara till localStorage så att de minns vid nästa omladdning även om man inte laddar fil
        localStorage.setItem('startDate', document.getElementById('startDate').value);
        localStorage.setItem('endDate', document.getElementById('endDate').value);
    }

    if (data.groupOpenState) {
        groupOpenState = data.groupOpenState;
    } else {
        groupOpenState = {};
    }

    if (data.logo) {
        localStorage.setItem('logo', data.logo);
        // initializeAI(); // Inte relevant längre
    }

    // Normalisera och läs in aktiviteter
    window.activities = normalizeImportedActivities(data);
    
    // Rensa urval
    selectedActivityIds.clear();
    selectedSegmentKeys.clear();

    pushHistory();
    if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render();
}
function normalizeImportedActivities(data) {
    const rawActs = Array.isArray(data.activities) ? data.activities : [];
    // Klona och gör grundkonverteringar
    let acts = rawActs.map(a => ({ ...a }));

    // 1) Konvertera segment/start/end fält tidigt
    acts.forEach(act => {
        act.start = parseImportedDate(act.start);
        act.end = parseImportedDate(act.end);
        if (Array.isArray(act.segments) && act.segments.length) {
            act.segments = act.segments.map(s => ({
                ...s,
                start: parseImportedDate(s.start),
                end: parseImportedDate(s.end),
                responsibles: Array.isArray(s.responsibles) ? s.responsibles : [],
                hoursPerDay: typeof s.hoursPerDay === 'number' ? s.hoursPerDay : undefined
            }));
        }
        if (!Array.isArray(act.children)) act.children = Array.isArray(act.children) ? act.children : [];
    });

    // 2) parentId -> parent
    const mapById = new Map();
    acts.forEach(a => { if (a && a.id) mapById.set(a.id, a); });
    acts.forEach(a => { if (!a.parent && a.parentId) a.parent = a.parentId; });

    // 3) Populera parent.children från parent-pointers
    acts.forEach(a => {
        if (a.parent) {
            const p = mapById.get(a.parent);
            if (p) {
                if (!Array.isArray(p.children)) p.children = [];
                if (!p.children.includes(a.id)) p.children.push(a.id);
            } else {
                delete a.parent; // okänt parent
            }
        }
    });

    // 4) Sätt barns parent från children-arrayer
    acts.forEach(a => {
        if (Array.isArray(a.children)) {
            a.children.forEach(cid => {
                const child = mapById.get(cid);
                if (child) child.parent = a.id;
            });
        }
    });

    // 5) Deduplicera children-arrayer
    acts.forEach(a => { if (Array.isArray(a.children)) a.children = Array.from(new Set(a.children)); });

    // 6) Om inga segments men start+end finns -> skapa segment (icke-grupp)
    acts.forEach(a => {
        if (!Array.isArray(a.segments) || a.segments.length === 0) {
            if (a.start && a.end && !a.isGroup) {
                a.segments = [{ start: a.start, end: a.end, name: a.name, color: a.color }];
            } else {
                a.segments = a.segments || [];
            }
        }
    });

    // 7) Deduplikera aktiviteter med samma id (merge fält)
    (function dedupeActivities() {
        const uniq = new Map();
        acts.forEach(a => {
            if (!a || !a.id) return;
            if (!uniq.has(a.id)) { uniq.set(a.id, a); return; }
            const existing = uniq.get(a.id);
            existing.children = Array.from(new Set([...(existing.children||[]), ...(a.children||[])]));
            if (!existing.parent && a.parent) existing.parent = a.parent;
            if (Array.isArray(a.segments) && a.segments.length) { existing.segments = existing.segments || []; a.segments.forEach(s => existing.segments.push(s)); }
        });
        acts = Array.from(uniq.values());
    })();

    // 8) Final normalisering och tidsjustering (bump sena UTC-midnights)
    const adjust = d => {
        if (!d) return;
        const h = d.getHours();
        if (h >= 21) d.setDate(d.getDate() + 1);
        d.setHours(0,0,0,0);
    };
    acts.forEach(a => {
        if (a.start) adjust(a.start);
        if (a.end) adjust(a.end);
        if (Array.isArray(a.segments)) {
            a.segments.forEach(s => { adjust(s.start); adjust(s.end); });
            if (!a.isGroup && a.segments.length) {
                updateActivityStartEndFromSegments(a);
                a.totalHours = calculateTotalHours(a);
            }
        }
    });

    // 9) Justera vy-start om någon aktivitet börjar tidigare än nuvarande vy
    try {
        let earliest = null;
        acts.forEach(a => {
            if (a.segments?.length) {
                a.segments.forEach(s => { if (!earliest || s.start < earliest) earliest = s.start; });
            } else if (a.start && (!earliest || a.start < earliest)) earliest = a.start;
        });
        if (earliest) {
            const viewStartInput = document.getElementById('startDate');
            const currentStart = normalizeDate(viewStartInput.value);
            if (currentStart && earliest < currentStart) viewStartInput.value = formatDateLocal(earliest);
        }
    } catch (_) { /* UI may not be ready */ }

    return acts;
}


// === LOCALSTORAGE & INDEXEDDB HANTERING ===

/**
 * Sparar hela applikationens tillstånd till webbläsarens localStorage.
 */
function saveToStorage() {
    try {
        const settings = {
            startDate: document.getElementById('startDate').value,
            endDate: document.getElementById('endDate').value,
            showHolidays: document.getElementById('showHolidays').checked,
            country: document.getElementById('dateFormat').value,
        };
        localStorage.setItem('gantt_settings', JSON.stringify(settings));
        localStorage.removeItem('gantt_title');
        localStorage.setItem('gantt_fileName', currentFileHandle?.name || currentFileName || '');
        localStorage.setItem('gantt_groupOpenState', JSON.stringify(groupOpenState));
        
        const activitiesToSave = activities.map(act => {
            const actToSave = { ...act };
            const useDate = d => (d && isTimeSpecific(d) ? formatDateTimeLocal(d) : formatDateLocal(d));
            if (actToSave.segments) {
                actToSave.segments = actToSave.segments.map(s => ({
                    start: useDate(s.start),
                    end: useDate(s.end),
                    info: s.info,
                    name: s.name,
                    color: s.color,
                    responsibles: Array.isArray(s.responsibles) ? s.responsibles : [],
                    hoursPerDay: typeof s.hoursPerDay === 'number' ? s.hoursPerDay : undefined
                }));
            }
            actToSave.start = act.start ? useDate(act.start) : null;
            actToSave.end = act.end ? useDate(act.end) : null;
            return actToSave;
        });
        localStorage.setItem('gantt_activities', JSON.stringify(activitiesToSave));
        saveBackup();
    } catch (error) {
        console.error("Kunde inte spara till localStorage:", error);
        showCenteredModal(`<h2>Lagringsfel</h2><p>Kunde inte spara projektet. Detta kan bero på att du använder webbläsaren i privat/inkognito-läge.</p><button onclick="closeCenteredModal()">Stäng</button>`);
    }
}

/**
 * Laddar applikationens tillstånd från localStorage.
 */
function loadFromStorage() {
    const savedActivities = localStorage.getItem('gantt_activities');
    if (!savedActivities) {
        activities = [];
        pushHistory();
        return;
    }

    try {
        const settings = JSON.parse(localStorage.getItem('gantt_settings'));
        if (settings) {
            document.getElementById('startDate').value = settings.startDate;
            document.getElementById('endDate').value = settings.endDate;
            document.getElementById('showHolidays').checked = settings.showHolidays;
            document.getElementById('dateFormat').value = settings.country;
        }

        setCurrentProjectFile(null, localStorage.getItem('gantt_fileName') || '');
        
        const savedGroupState = JSON.parse(localStorage.getItem('gantt_groupOpenState'));
        if (savedGroupState) {
            Object.assign(groupOpenState, savedGroupState);
        }

        activities = JSON.parse(savedActivities).map(a => {
            const loadedAct = { ...a };
            if (loadedAct.segments) {
                loadedAct.segments = loadedAct.segments.map(s => ({
                    start: normalizeDate(s.start),
                    end: normalizeDate(s.end),
                    info: s.info,
                    name: s.name,
                    color: s.color,
                    responsibles: Array.isArray(s.responsibles) ? s.responsibles : [],
                    hoursPerDay: typeof s.hoursPerDay === 'number' ? s.hoursPerDay : undefined
                }));
            }
            loadedAct.start = a.start ? normalizeDate(a.start) : null;
            loadedAct.end = a.end ? normalizeDate(a.end) : null;
            return loadedAct;
        });

    try { if (typeof rebuildActivityIndex === 'function') rebuildActivityIndex(); } catch(_) {}
        history = [];
        historyIndex = -1;
        pushHistory();
    } catch (error) {
        console.error("Kunde inte ladda från localStorage, börjar med ett tomt projekt.", error);
        activities = [];
        pushHistory();
    }
}

/**
 * Laddar applikationens tillstånd från en sparad JSON-fil.
 */
function loadFromFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = JSON.parse(e.target.result);

            setCurrentProjectFile(null, file.name);
            if (data.settings) {
                document.getElementById('startDate').value = data.settings.startDate;
                document.getElementById('endDate').value = data.settings.endDate;
                document.getElementById('showHolidays').checked = data.settings.showHolidays;
                document.getElementById('dateFormat').value = data.settings.country;
            }
            if(data.logo) {
                localStorage.setItem('logo', data.logo);
            }

            Object.assign(groupOpenState, data.groupOpenState || {});
            // Centraliserad normalisering
            activities = normalizeImportedActivities(data);
            try { if (typeof rebuildActivityIndex === 'function') rebuildActivityIndex(); } catch(_) {}

            // --- Normalisera tider (ta bort 22:00-drift) och justera vy-start ---
            (function normalizeAfterLoad(){
                const adjust = d => {
                    if (!d) return;
                    // Om sena kvällen (21-23) -> bump till nästa dag (troligen UTC-midnight sparad)
                    const h = d.getHours();
                    if (h >= 21) d.setDate(d.getDate() + 1);
                    d.setHours(0,0,0,0);
                };
                activities.forEach(a => {
                    if (a.start) adjust(a.start);
                    if (a.end) adjust(a.end);
                    if (Array.isArray(a.segments)) {
                        a.segments.forEach(s => { adjust(s.start); adjust(s.end); });
                        if (!a.isGroup && a.segments.length) {
                            updateActivityStartEndFromSegments(a);
                            a.totalHours = calculateTotalHours(a);
                        }
                    }
                });
                // Justera startDate om någon aktivitet börjar tidigare än vald vy
                let earliest = null;
                activities.forEach(a => {
                    if (a.segments?.length) {
                        a.segments.forEach(s => { if (!earliest || s.start < earliest) earliest = s.start; });
                    } else if (a.start && (!earliest || a.start < earliest)) earliest = a.start;
                });
                if (earliest) {
                    const viewStartInput = document.getElementById('startDate');
                    const currentStart = normalizeDate(viewStartInput.value);
                    if (currentStart && earliest < currentStart) {
                        viewStartInput.value = formatDateLocal(earliest);
                    }
                }
            })();

            history = [];
            historyIndex = -1;
            pushHistory();
            if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render();
            alert(`Projektet "${data.title || file.name}" har laddats.`);
        } catch (error) {
            alert("Kunde inte läsa filen. Se till att det är en giltig JSON-fil från detta program.");
            console.error("Fel vid import av JSON:", error);
        }
    };
    reader.readAsText(file);
}

/**
 * Öppnar filväljaren för att ladda ett projekt.
 */
async function openFile() {
    // Försök använda File System Access API (Windows/Edge/Chrome) för bästa filhantering
    if (window.showOpenFilePicker) {
        try {
            let startInOption;
            try {
                startInOption = await getSavedFileHandle();
            } catch (_) {}

            const [handle] = await window.showOpenFilePicker({
                multiple: false,
                types: [{ description: 'Gantt-projekt (JSON)', accept: { 'application/json': ['.json'] } }],
                startIn: startInOption || undefined
            });
            const file = await handle.getFile();
            const text = await file.text();
            // Ladda in data (återanvänd logiken från loadFromFile)
            try {
                const data = JSON.parse(text);
                setCurrentProjectFile(handle, handle.name || file.name);
                if (data.settings) {
                    document.getElementById('startDate').value = data.settings.startDate;
                    document.getElementById('endDate').value = data.settings.endDate;
                    document.getElementById('showHolidays').checked = data.settings.showHolidays;
                    document.getElementById('dateFormat').value = data.settings.country;
                }
                if (data.logo) {
                    localStorage.setItem('logo', data.logo);
                }
                Object.assign(groupOpenState, data.groupOpenState || {});
                // Centraliserad normalisering
                activities = normalizeImportedActivities(data);
                try { if (typeof rebuildActivityIndex === 'function') rebuildActivityIndex(); } catch(_) {}
                history = [];
                historyIndex = -1;
                pushHistory();
                if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render();
                alert(`Projektet "${data.title || file.name}" har laddats.`);
                await setSavedFileHandle(handle);
                return;
            } catch (err) {
                alert('Kunde inte läsa filen. Se till att det är en giltig JSON-fil från detta program.');
                console.error('Fel vid import av JSON:', err);
            }
        } catch (e) {
            if (e && e.name === 'AbortError') return; // användaren avbröt
            if (window.developerMode) console.warn('showOpenFilePicker misslyckades, faller tillbaka till input[type=file]', e);
        }
    }
    // Fallback: traditionell filväljare
    document.getElementById('openFileInput').click();
}

/**
 * Samlar ihop all projektdata till ett enda objekt redo för att sparas.
 */
function prepareSaveData() {
    const settings = {
        startDate: document.getElementById('startDate').value,
        endDate: document.getElementById('endDate').value,
        showHolidays: document.getElementById('showHolidays').checked,
        country: document.getElementById('dateFormat').value,
    };
    const activitiesToSave = activities.map(act => {
        const actToSave = { ...act };
        // Use date-only when time is not set to avoid TZ-related shifts
        const useDate = d => (d && isTimeSpecific(d) ? formatDateTimeLocal(d) : formatDateLocal(d));
        if (actToSave.segments) {
            actToSave.segments = actToSave.segments.map(s => ({
                start: useDate(s.start),
                end: useDate(s.end),
                info: s.info,
                name: s.name,
                color: s.color,
                hoursPerDay: typeof s.hoursPerDay === 'number' ? s.hoursPerDay : undefined
            }));
        }
        actToSave.start = act.start ? useDate(act.start) : null;
        actToSave.end = act.end ? useDate(act.end) : null;
        return actToSave;
    });

    return {
        title: getProjectTitle(),
        settings: settings,
        groupOpenState: groupOpenState,
        activities: activitiesToSave,
        logo: localStorage.getItem('logo')
    };
}

/**
 * Skapar ett JSON-objekt av endast nu synliga aktiviteter i hierarkin.
 * Inkluderar aggregerad procent för containers (viktad efter timmar).
 */
function prepareVisibleData() {
    if (typeof getFilteredActivitiesWithHierarchy !== 'function') {
        return prepareSaveData(); // fallback
    }
    const viewStart = normalizeDate(document.getElementById('startDate').value);
    const viewEnd = normalizeDate(document.getElementById('endDate').value);
    const visible = getFilteredActivitiesWithHierarchy().map(v => findActivityById(v.id)).filter(Boolean);
    const idSet = new Set(visible.map(a => a.id));

        function cloneAct(act) {
        const c = { ...act };
        if (c.segments) {
            // Filtrera och klipp segment till synligt spann
            const clipped = [];
            c.segments.forEach(s => {
                const sStart = normalizeDate(s.start);
                const sEnd = normalizeDate(s.end);
                if (!sStart || !sEnd) return;
                if (sEnd < viewStart || sStart > viewEnd) return; // helt utanför
                const clipStart = new Date(Math.max(sStart.getTime(), viewStart.getTime()));
                const clipEnd = new Date(Math.min(sEnd.getTime(), viewEnd.getTime()));
                // Behåll ursprunglig tidskomponent endast om segmentet inte klippt i den änden
                const isTruncStart = clipStart.getTime() !== sStart.getTime();
                const isTruncEnd = clipEnd.getTime() !== sEnd.getTime();
                    const useDate = d => (d && isTimeSpecific(d) ? formatDateTimeLocal(d) : formatDateLocal(d));
                    const outSeg = {
                    start: useDate(clipStart),
                    end: useDate(clipEnd),
                    info: s.info,
                    name: s.name,
                    color: s.color,
                    responsibles: Array.isArray(s.responsibles) ? s.responsibles : [],
                        hoursPerDay: typeof s.hoursPerDay === 'number' ? s.hoursPerDay : undefined,
                };
                if (isTruncStart || isTruncEnd) outSeg.truncated = true;
                clipped.push(outSeg);
            });
            c.segments = clipped;
        }
            const useDateTop = d => (d && isTimeSpecific(d) ? formatDateTimeLocal(d) : formatDateLocal(d));
            c.start = c.start ? useDateTop(c.start) : null;
            c.end = c.end ? useDateTop(c.end) : null;
        // Filtrera children till synliga
        if (c.children) c.children = c.children.filter(id => idSet.has(id));
        return c;
    }

    // Förberäkna synliga timmar per aktivitet (efter klippning)
    const visibleHoursMap = new Map();
    visible.forEach(a => {
        if (!a.segments || !a.segments.length || a.isGroup) return;
        let h = 0;
        a.segments.forEach(s => {
            const sStart = normalizeDate(s.start);
            const sEnd = normalizeDate(s.end);
            if (sEnd < viewStart || sStart > viewEnd) return;
            const clipStart = new Date(Math.max(sStart.getTime(), viewStart.getTime()));
            const clipEnd = new Date(Math.min(sEnd.getTime(), viewEnd.getTime()));
            const segHpd = (typeof s.hoursPerDay === 'number') ? s.hoursPerDay : (a.hoursPerDay || 0);
            h += getWorkDays(clipStart, clipEnd) * segHpd;
        });
        visibleHoursMap.set(a.id, h);
    });

    // Rekursiv viktad completion och visible hours för containers
    function computeContainer(act) {
        if (!act.isGroup) return { hours: visibleHoursMap.get(act.id) || 0, comp: act.completed || 0 };
        const children = act.children.map(id => findActivityById(id)).filter(ch => ch && idSet.has(ch.id));
        let hoursSum = 0, weighted = 0;
        children.forEach(ch => {
            const { hours, comp } = computeContainer(ch);
            hoursSum += hours;
            weighted += comp * hours;
        });
        act._visibleHours = hoursSum;
        act._visibleCompleted = hoursSum > 0 ? Math.round(weighted / (hoursSum || 1)) : 0;
        return { hours: hoursSum, comp: act._visibleCompleted };
    }
    visible.filter(a => a.isGroup).forEach(computeContainer);

    const activitiesToSave = visible.map(a => {
        // Hoppa över aktiviteter utan segment i spann (om ej container)
        if (!a.isGroup && (!a.segments || !a.segments.some(s => {
            const sStart = normalizeDate(s.start); const sEnd = normalizeDate(s.end);
            return !(sEnd < viewStart || sStart > viewEnd);
        }))) return null;
        const c = cloneAct(a);
        if (a.isGroup) {
            if (typeof a._visibleCompleted === 'number') c.spanCompleted = a._visibleCompleted; // spann-procent
            c.totalCompleted = a.completed || 0;
            if (typeof a._visibleHours === 'number') c.visibleHours = a._visibleHours;
        } else {
            c.visibleHours = visibleHoursMap.get(a.id) || 0;
        }
        return c;
    });
    const filteredActs = activitiesToSave.filter(Boolean);
    return {
        title: `${getProjectTitle()} (synligt)`,
        settings: {
            startDate: document.getElementById('startDate').value,
            endDate: document.getElementById('endDate').value
        },
        groupOpenState: groupOpenState,
        activities: filteredActs
    };
}

async function exportVisibleJSON() {
    const data = prepareVisibleData();
    const json = JSON.stringify(data, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const suggestedName = `${sanitizeFileName(data.title)}.json`;

    if (window.showSaveFilePicker) {
        try {
            const handle = await window.showSaveFilePicker({
                suggestedName,
                types: [{ description: 'Synligt Gantt (JSON)', accept: { 'application/json': ['.json'] } }]
            });
            if (!(await verifyPermission(handle, 'readwrite'))) throw new Error('Permission denied');
            const writable = await handle.createWritable();
            await writable.write(blob);
            await writable.close();
            alert('Synligt utdrag exporterat.');
            return;
        } catch (e) {
            if (e && e.name === 'AbortError') return;
            if (window.developerMode) console.warn('showSaveFilePicker misslyckades, använder nedladdning', e);
        }
    }
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = suggestedName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
}

/**
 * Hjälpare: Sanera filnamn
 */
function sanitizeFileName(name) {
    return (name || 'gantt-projekt').replace(/[^a-z0-9._-]/gi, '_');
}

/**
 * Verifiera/efterfråga filbehörighet
 */
async function verifyPermission(fileHandle, mode = 'readwrite') {
    if (!fileHandle) return false;
    const opts = { mode };
    if (await fileHandle.queryPermission(opts) === 'granted') return true;
    if (await fileHandle.requestPermission(opts) === 'granted') return true;
    return false;
}

/**
 * Spara till befintlig fil om möjligt, annars öppna Spara som...
 */
async function saveFile() {
    const data = prepareSaveData();
    const json = JSON.stringify(data, null, 2);
    const blob = new Blob([json], { type: 'application/json' });

    // Om vi har ett aktivt handtag, skriv direkt
    if (currentFileHandle) {
        try {
            if (!(await verifyPermission(currentFileHandle, 'readwrite'))) throw new Error('Permission denied');
            const writable = await currentFileHandle.createWritable();
            await writable.write(blob);
            await writable.close();
            setCurrentProjectFile(currentFileHandle, currentFileHandle.name);
            alert('Projektet har sparats.');
            await setSavedFileHandle(currentFileHandle); // uppdatera cache
            return;
        } catch (e) {
            if (window.developerMode) console.warn('Kunde inte spara till nuvarande fil, försöker Spara som...', e);
        }
    }

    // Försök återanvända senast använda handtag från IndexedDB
    try {
        const savedHandle = await getSavedFileHandle();
        if (savedHandle && (await verifyPermission(savedHandle, 'readwrite'))) {
            const writable = await savedHandle.createWritable();
            await writable.write(blob);
            await writable.close();
            setCurrentProjectFile(savedHandle, savedHandle.name);
            alert('Projektet har sparats.');
            return;
        }
    } catch (e) {
    if (window.developerMode) console.warn('Återanvändning av sparat filhandtag misslyckades:', e);
    }

    // Falla tillbaka till "Spara som..."
    return saveAsFile();
}

/**
 * Spara som: öppna Windows filhanterare för att välja nytt filnamn.
 * Faller tillbaka till nedladdning om API inte stöds.
 */
async function saveAsFile() {
    const data = prepareSaveData();
    const json = JSON.stringify(data, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const suggestedName = `${sanitizeFileName(getProjectTitle('gantt-projekt'))}.json`;

    if (window.showSaveFilePicker) {
        try {
            let startInOption;
            try {
                startInOption = await getSavedFileHandle();
            } catch (_) {}

            const handle = await window.showSaveFilePicker({
                suggestedName,
                types: [{ description: 'Gantt-projekt (JSON)', accept: { 'application/json': ['.json'] } }],
                startIn: startInOption || undefined
            });
            if (!(await verifyPermission(handle, 'readwrite'))) throw new Error('Permission denied');
            const writable = await handle.createWritable();
            await writable.write(blob);
            await writable.close();
            setCurrentProjectFile(handle, handle.name);
            await setSavedFileHandle(handle);
            alert('Projektet har sparats.');
            return;
        } catch (e) {
            if (e && e.name === 'AbortError') return; // användaren avbröt
            if (window.developerMode) console.warn('showSaveFilePicker misslyckades, faller tillbaka till nedladdning', e);
        }
    }

    // Fallback: nedladdning
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = suggestedName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
    setCurrentProjectFile(null, suggestedName);
}

    // === Export: MS Project (XML) ===

    function xmlEsc(str) {
        return String(str ?? '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    function toIsoLocal(date, hh = 8, mm = 0, ss = 0) {
        if (!(date instanceof Date) || isNaN(date.getTime())) return '';
        const d = new Date(date.getFullYear(), date.getMonth(), date.getDate(), hh, mm, ss, 0);
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const da = String(d.getDate()).padStart(2, '0');
        const h = String(d.getHours()).padStart(2, '0');
        const mi = String(d.getMinutes()).padStart(2, '0');
        const s = String(d.getSeconds()).padStart(2, '0');
        return `${y}-${m}-${da}T${h}:${mi}:${s}`; // local, no Z, avoids timezone shifts in MSP
    }

    function durationWorkISO(hoursFloat) {
        const hrs = Math.max(0, Math.round((hoursFloat || 0) * 60) / 60);
        const h = Math.floor(hrs);
        const m = Math.round((hrs - h) * 60);
        return `PT${h}H${m}M0S`;
    }

    function durationFromDatesISO(start, end) {
        const s = normalizeDate(start);
        const e = normalizeDate(end);
        if (!s || !e || e < s) return 'PT0H0M0S';
        // Antag arbetsdag 8h: Duration = dagar * 8h
        const days = getWorkDays(s, e);
        const hours = days * getStandardHoursPerDay();
        return durationWorkISO(hours);
    }

    function collectTopLevelActivities() {
        // Keep order as in activities array
        return (activities || []).filter(a => !a.parent);
    }

    function getChildren(act) {
        if (!act || !Array.isArray(act.children)) return [];
        return act.children.map(id => (typeof findActivityById === 'function' ? findActivityById(id) : null)).filter(Boolean);
    }

    function computeSegmentHours(act, seg) {
        try {
            const hpd = (typeof seg.hoursPerDay === 'number') ? seg.hoursPerDay : (typeof act.hoursPerDay === 'number' ? act.hoursPerDay : getStandardHoursPerDay());
            const days = getWorkDays(seg.start, seg.end);
            return hpd * days;
        } catch(_) { return 0; }
    }

    function buildMSPTasksXML(flatten = false) {
        let uid = 0, id = 0;
        let flatSeq = 0;
        const lines = [];
        function addTask(task) {
            lines.push('      <Task>');
            Object.keys(task).forEach(k => {
                const v = task[k];
                if (v == null) return;
                lines.push(`        <${k}>${v}</${k}>`);
            });
            lines.push('      </Task>');
        }

        const startView = normalizeDate(document.getElementById('startDate').value);
        const endView = normalizeDate(document.getElementById('endDate').value);
        const defaultStart = toIsoLocal(startView || new Date(), 8, 0, 0);

        function walk(node, outlineLevel, outlineNumber) {
            if (!node) return;
            const thisUID = ++uid;
            const thisID = ++id;
            const isSummary = !!node.isGroup || (Array.isArray(node.segments) && node.segments.length > 1);
            const name = xmlEsc(node.name || (node.isProject ? 'Projekt' : (node.isGroup ? 'Grupp' : 'Aktivitet')));
            const start = node.start ? toIsoLocal(node.start, 8, 0, 0) : defaultStart;
            const finish = node.end ? toIsoLocal(node.end, 17, 0, 0) : start;

            if (!flatten || (!node.isGroup && (!node.segments || node.segments.length <= 1))) {
                const wbs = flatten ? String(++flatSeq) : outlineNumber;
                const outlineLvl = flatten ? 1 : outlineLevel;
                addTask({
                UID: thisUID,
                ID: thisID,
                Name: name,
                Type: 1,
                IsNull: 0,
                WBS: wbs,
                OutlineLevel: outlineLvl,
                OutlineNumber: wbs,
                Start: start,
                Finish: finish,
                Duration: durationFromDatesISO(node.start || start, node.end || finish),
                    Summary: (flatten ? 0 : (isSummary ? 1 : 0)),
                Active: 1,
                CalendarUID: 1,
                Manual: 0
                });
            }

            // Children: either real children (grupper) eller segment som deluppgifter
            if (node.isGroup) {
                const kids = getChildren(node);
                kids.forEach((child, idx) => walk(child, outlineLevel + 1, `${outlineNumber}.${idx + 1}`));
            } else if (Array.isArray(node.segments)) {
                if (node.segments.length > 1) {
                    node.segments.forEach((seg, idx) => {
                        const segUID = ++uid; const segID = ++id;
                        const sName = xmlEsc((seg.name && seg.name.trim()) ? seg.name : `${name} (del ${idx + 1})`);
                        const sStart = toIsoLocal(seg.start, 8, 0, 0);
                        const sFinish = toIsoLocal(seg.end, 17, 0, 0);
                        const workH = computeSegmentHours(node, seg);
                        const wbs = flatten ? String(++flatSeq) : `${outlineNumber}.${idx + 1}`;
                        const outlineLvl = flatten ? 1 : (outlineLevel + 1);
                        const task = {
                            UID: segUID,
                            ID: segID,
                            Name: sName,
                            Type: 1,
                            IsNull: 0,
                            WBS: wbs,
                            OutlineLevel: outlineLvl,
                            OutlineNumber: wbs,
                            Start: sStart,
                            Finish: sFinish,
                            Duration: durationFromDatesISO(seg.start, seg.end),
                            Summary: 0,
                            Active: 1,
                            CalendarUID: 1,
                            Manual: 0
                        };
                        if (!flatten) task.Work = durationWorkISO(workH);
                        addTask(task);
                    });
                } else if (node.segments.length === 1) {
                    // Already exported as non-summary above (using node.start/end). Add Work if available
                    const workH = computeSegmentHours(node, node.segments[0]);
                    // Inject a Work line for the previously added task
                    // Simpler approach: append a separate task with same dates but keep hierarchy simple is not ideal.
                    // Instead, we’ll add an extra Task with zero outline delta to set Work.
                    // For schema simplicity, we skip this and let MSP compute duration from dates.
                }
            }
        }

        const tops = collectTopLevelActivities();
        tops.forEach((t, i) => walk(t, 1, String(i + 1)));
        return lines.join('\n');
    }

    function buildMSProjectXML(options = {}) {
        const title = getProjectTitle('Gantt-projekt');
        const vs = normalizeDate(document.getElementById('startDate').value) || new Date();
        const ve = normalizeDate(document.getElementById('endDate').value) || vs;
        const start = toIsoLocal(vs, 8, 0, 0);
        const finish = toIsoLocal(ve, 17, 0, 0);
        const nowIso = toIsoLocal(new Date());
            const tasksXML = buildMSPTasksXML(!!options.flatten);
            // Minimal header to match importer-friendly sample
            // Standardkalender UID 1, mån-fre 08-12 & 13-17
            const calendars = `  <Calendars>
            <Calendar>
                <UID>1</UID>
                <Name>Standard</Name>
                <IsBaseCalendar>1</IsBaseCalendar>
                <BaseCalendarUID>0</BaseCalendarUID>
                <WeekDays>
                    <WeekDay><DayType>1</DayType><DayWorking>0</DayWorking></WeekDay>
                    <WeekDay>
                        <DayType>2</DayType><DayWorking>1</DayWorking>
                        <WorkingTimes>
                            <WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime>
                            <WorkingTime><FromTime>13:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime>
                        </WorkingTimes>
                    </WeekDay>
                    <WeekDay>
                        <DayType>3</DayType><DayWorking>1</DayWorking>
                        <WorkingTimes>
                            <WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime>
                            <WorkingTime><FromTime>13:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime>
                        </WorkingTimes>
                    </WeekDay>
                    <WeekDay>
                        <DayType>4</DayType><DayWorking>1</DayWorking>
                        <WorkingTimes>
                            <WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime>
                            <WorkingTime><FromTime>13:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime>
                        </WorkingTimes>
                    </WeekDay>
                    <WeekDay>
                        <DayType>5</DayType><DayWorking>1</DayWorking>
                        <WorkingTimes>
                            <WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime>
                            <WorkingTime><FromTime>13:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime>
                        </WorkingTimes>
                    </WeekDay>
                    <WeekDay>
                        <DayType>6</DayType><DayWorking>1</DayWorking>
                        <WorkingTimes>
                            <WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime>
                            <WorkingTime><FromTime>13:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime>
                        </WorkingTimes>
                    </WeekDay>
                    <WeekDay><DayType>7</DayType><DayWorking>0</DayWorking></WeekDay>
                </WeekDays>
            </Calendar>
        </Calendars>`;

            return `<?xml version="1.0" encoding="UTF-8"?>
    <Project xmlns="http://schemas.microsoft.com/project">
        <Name>${xmlEsc(title)}</Name>
        <Title>${xmlEsc(title)}</Title>
            <CreationDate>${nowIso}</CreationDate>
        <StartDate>${start}</StartDate>
        <FinishDate>${finish}</FinishDate>
        <ScheduleFromStart>1</ScheduleFromStart>
        <CalendarUID>1</CalendarUID>
        <CurrencySymbol>kr</CurrencySymbol>
        <DefaultStartTime>08:00:00</DefaultStartTime>
        <DefaultFinishTime>17:00:00</DefaultFinishTime>
            <MinutesPerDay>480</MinutesPerDay>
            <MinutesPerWeek>2400</MinutesPerWeek>
        <DaysPerMonth>20</DaysPerMonth>
    ${calendars}
        <Tasks>
    ${tasksXML}
        </Tasks>
            <Resources />
            <Assignments />
    </Project>`;
    }

    async function exportToMSProject() {
        try {
            let xml = buildMSProjectXML();
            xml = toTwoLineXML(xml);
            const blob = new Blob([xml], { type: 'application/xml' });
            const suggestedName = `${sanitizeFileName(getProjectTitle('gantt-projekt'))}.xml`;
            if (window.showSaveFilePicker) {
                try {
                    const handle = await window.showSaveFilePicker({
                        suggestedName,
                        types: [{ description: 'MS Project (XML)', accept: { 'application/xml': ['.xml'] } }]
                    });
                    if (!(await verifyPermission(handle, 'readwrite'))) throw new Error('Permission denied');
                    const writable = await handle.createWritable();
                    await writable.write(blob);
                    await writable.close();
                    alert('MS Project XML exporterat.');
                    return;
                } catch (e) {
                    if (e && e.name === 'AbortError') return;
                    if (window.developerMode) console.warn('showSaveFilePicker misslyckades, använder nedladdning', e);
                }
            }
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = suggestedName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(a.href);
        } catch (err) {
            console.error('MS Project-export misslyckades', err);
            alert('Kunde inte exportera till MS Project XML.');
        }
    }

// --- INDEXEDDB BACKUP ---
let db;

function openDB() {
    return new Promise((resolve, reject) => {
        if (db) return resolve(db);
        const request = indexedDB.open('GanttBackupDB', 2);
        request.onupgradeneeded = e => {
            const dbi = e.target.result;
            if (!dbi.objectStoreNames.contains('projects')) {
                dbi.createObjectStore('projects', { keyPath: 'id' });
            }
            if (!dbi.objectStoreNames.contains('handles')) {
                dbi.createObjectStore('handles', { keyPath: 'id' });
            }
        };
        request.onsuccess = e => { db = e.target.result; resolve(db); };
        request.onerror = e => { console.error('IndexedDB-fel:', e.target.error); reject(e.target.error); };
    });
}

function buildAstaProjectXML() {
    // Force flattened tasks to mirror Asta sample
    return buildMSProjectXML({ flatten: true });
}

async function exportToAstaXML() {
    try {
        let xml = buildAstaProjectXML();
        xml = toTwoLineXML(xml);
        const blob = new Blob([xml], { type: 'application/xml' });
        const suggestedName = `${sanitizeFileName(getProjectTitle('gantt-projekt'))}_asta.xml`;
        if (window.showSaveFilePicker) {
            try {
                const handle = await window.showSaveFilePicker({
                    suggestedName,
                    types: [{ description: 'Asta/Powerproject (XML)', accept: { 'application/xml': ['.xml'] } }]
                });
                if (!(await verifyPermission(handle, 'readwrite'))) throw new Error('Permission denied');
                const writable = await handle.createWritable();
                await writable.write(blob);
                await writable.close();
                alert('Asta-kompatibel XML exporterat.');
                return;
            } catch (e) {
                if (e && e.name === 'AbortError') return;
                if (window.developerMode) console.warn('showSaveFilePicker misslyckades, använder nedladdning', e);
            }
        }
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = suggestedName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(a.href);
    } catch (err) {
        console.error('Asta-export misslyckades', err);
        alert('Kunde inte exportera Asta-kompatibel XML.');
    }
}

try { window.exportToAstaXML = exportToAstaXML; } catch(_) {}

function toTwoLineXML(xml) {
    if (!xml) return xml;
    // Split declaration from body (handle CRLF, LF, or CR)
    const parts = xml.split(/(?:\r\n|\r|\n)/);
    let decl = parts[0] && parts[0].startsWith('<?xml') ? parts[0] : '<?xml version="1.0" encoding="UTF-8"?>';
    let body = (parts[0] && parts[0].startsWith('<?xml')) ? parts.slice(1).join('') : xml;
    // Collapse whitespace between tags
    const collapsed = body.replace(/>\s+</g, '><')
                          .replace(/\s{2,}/g, ' ')
                          .trim();
    // Validate collapse; if we lost structure, keep original
    if (!collapsed || collapsed.indexOf('<Project') === -1) {
        // Ensure two lines anyway
        return `${decl}\r\n${body.trim()}`;
    }
    // Ensure CRLF between the two lines for Windows tools
    return `${decl}\r\n${collapsed}`;
}

async function saveBackup() {
    try {
        const dbInstance = await openDB();
        const dataToBackup = {
            id: 'current_project',
            timestamp: new Date(),
            data: prepareSaveData()
        };
        const transaction = dbInstance.transaction(['projects'], 'readwrite');
        transaction.objectStore('projects').put(dataToBackup);
    } catch (error) {
        console.error("Kunde inte spara backup till IndexedDB:", error);
    }
}

// Exponera för global åtkomst från inline onclick i HTML
try { window.exportToMSProject = exportToMSProject; } catch(_) {}

function loadBackup() {
    if (confirm("Är du säker på att du vill återställa den senast sparade sessionen? Nuvarande ändringar kommer att förloras.")) {
        loadFromStorage();
        if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render();
    }
}

// Spara/återhämta senast använda filhandtag i IndexedDB (endast Chromium)
async function setSavedFileHandle(handle) {
    try {
        const dbInstance = await openDB();
        const tx = dbInstance.transaction(['handles'], 'readwrite');
        tx.objectStore('handles').put({ id: 'current', handle });
        return new Promise((res, rej) => {
            tx.oncomplete = () => res();
            tx.onerror = e => rej(e);
            tx.onabort = e => rej(e);
        });
    } catch (e) {
    if (window.developerMode) console.warn('Kunde inte spara filhandtag (ok, fortsätter ändå):', e);
    }
}

async function getSavedFileHandle() {
    try {
        const dbInstance = await openDB();
        return await new Promise((resolve, reject) => {
            const tx = dbInstance.transaction(['handles'], 'readonly');
            const req = tx.objectStore('handles').get('current');
            req.onsuccess = () => resolve(req.result ? req.result.handle : null);
            req.onerror = e => reject(e);
        });
    } catch (e) {
    if (window.developerMode) console.warn('Kunde inte läsa sparat filhandtag:', e);
        return null;
    }
}

// === EXCEL IMPORT ===

/**
 * Öppnar filväljaren för Excel-import
 */
function importExcelFile() {
    document.getElementById('excelFileInput').click();
}

/**
 * Parsear veckodata från header-raden för Excel-import
 */
function parseWeekHeaders(headerRow) {
    const weekInfo = [];
    
    for (let col = 4; col < headerRow.length; col++) {
        const cellValue = headerRow[col];
        if (typeof cellValue === 'number' && cellValue >= 15 && cellValue <= 52) {
            // Detta är ett veckoanummer
            weekInfo.push({
                column: col,
                week: cellValue,
                startDate: getWeekStartDate(cellValue, 2025)
            });
        }
    }
    
    return weekInfo;
}

/**
 * Konverterar veckoanummer och år till startdatum för veckan
 */
function getWeekStartDate(weekNumber, year = 2025) {
    // ISO 8601 veckoberäkning
    const jan4 = new Date(year, 0, 4);
    const jan4WeekDay = jan4.getDay() || 7; // Måndag = 1, Söndag = 7
    const jan4Week = new Date(jan4.getTime() - (jan4WeekDay - 1) * 24 * 60 * 60 * 1000);
    
    const weekStart = new Date(jan4Week.getTime() + (weekNumber - 1) * 7 * 24 * 60 * 60 * 1000);
    return weekStart;
}

/**
 * Skapar aktivitetssegment från Excel-arbetschema
 */
function createActivitySegments(personName, taskCells, weekInfo) {
    const segments = [];
    let currentSegment = null;
    
    for (let i = 0; i < weekInfo.length; i++) {
        const week = weekInfo[i];
        const cellValue = taskCells[week.column] || '';
        const cellStr = cellValue?.toString().trim() || '';
        const hasWork = cellStr !== '' && cellStr !== '.';
        
        if (hasWork) {
            if (!currentSegment) {
                // Starta nytt segment
                currentSegment = {
                    start: new Date(week.startDate),
                    taskName: cellStr
                };
            }
            // Uppdatera slutdatum (slutet av veckan)
            currentSegment.end = new Date(week.startDate);
            currentSegment.end.setDate(currentSegment.end.getDate() + 4); // Fredag
        } else {
            if (currentSegment) {
                // Avsluta pågående segment
                segments.push({
                    start: formatDateTimeLocal(currentSegment.start),
                    end: formatDateTimeLocal(currentSegment.end),
                    name: currentSegment.taskName,
                    info: `${personName} - ${currentSegment.taskName}`
                });
                currentSegment = null;
            }
        }
    }
    
    // Avsluta eventuellt pågående segment
    if (currentSegment) {
        segments.push({
            start: formatDateTimeLocal(currentSegment.start),
            end: formatDateTimeLocal(currentSegment.end),
            name: currentSegment.taskName,
            info: `${personName} - ${currentSegment.taskName}`
        });
    }
    
    return segments;
}

/**
 * Hanterar Excel-filimport
 */
function handleExcelImport(file) {
    if (!file.name.toLowerCase().endsWith('.xlsx') && !file.name.toLowerCase().endsWith('.xls')) {
        alert('Vänligen välj en Excel-fil (.xlsx eller .xls)');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Hitta Service worksheet
            const serviceSheet = workbook.Sheets['Service'];
            if (!serviceSheet) {
                alert('Worksheet "Service" hittades inte i Excel-filen. Kontrollera att filen har rätt format.');
                return;
            }
            
            // Konvertera till array
            const sheetData = XLSX.utils.sheet_to_json(serviceSheet, { 
                header: 1,
                defval: '',
                range: 0
            });
            
            if (sheetData.length < 3) {
                alert('För få rader i Service-bladet');
                return;
            }
            
            // Parsea header för att hitta veckor
            const headerRow = sheetData[0];
            const weekInfo = parseWeekHeaders(headerRow);
            
            if (weekInfo.length === 0) {
                alert('Kunde inte hitta veckodata i Excel-filen');
                return;
            }
            
            if (window.developerMode) console.log(`Hittade ${weekInfo.length} veckor från vecka ${weekInfo[0]?.week} till ${weekInfo[weekInfo.length-1]?.week}`);
            
            // Bekräfta import
            const confirmMsg = `Importera bemanningsplan från Excel?\n\nHittade ${weekInfo.length} veckor och ${sheetData.length - 2} personer.\nDetta kommer ersätta befintligt schema.`;
            if (!confirm(confirmMsg)) {
                return;
            }
            
            // Skapa ny Gantt struktur
            const ganttData = {
                title: 'Bemanningsplan från Excel',
                settings: {
                    startDate: formatDateLocal(weekInfo[0].startDate),
                    endDate: formatDateLocal(new Date(weekInfo[weekInfo.length-1].startDate.getTime() + 6*24*60*60*1000)),
                    showHolidays: true,
                    country: 'SE'
                },
                groupOpenState: {},
                activities: []
            };
            
            // Skapa huvudgrupp för servicemontörer
            const serviceGroup = {
                id: 'service-group',
                name: 'Servicemontörer',
                isGroup: true,
                isProject: false,
                children: [],
                parent: null,
                color: '#2196f3',
                colorName: 'Service',
                completed: 0,
                segments: [],
                _order: 1
            };
            
            ganttData.activities.push(serviceGroup);
            ganttData.groupOpenState['service-group'] = true;
            
            let activityOrder = 2;
            let importedCount = 0;
            
            // Procesessa varje person (från rad 3 och framåt)
            for (let row = 2; row < sheetData.length; row++) {
                const rowData = sheetData[row];
                const personName = rowData[0]?.toString().trim();
                
                if (!personName || personName === '') continue;
                
                // Skapa aktivitetssegment från schemat
                const segments = createActivitySegments(personName, rowData, weekInfo);
                
                if (segments.length > 0) {
                    // Skapa personaktivitet
                    const personActivity = {
                        id: generateId(),
                        name: personName,
                        segments: segments,
                        isGroup: false,
                        isProject: false,
                        children: [],
                        parent: 'service-group',
                        color: '#4caf50',
                        colorName: 'Person',
                        completed: 0,
                        hoursPerDay: 8,
                        totalHours: segments.length * 5 * 8, // Approximation
                        _order: activityOrder++
                    };
                    
                    // Beräkna övergripande start och slutdatum
                    const allDates = segments.flatMap(s => [new Date(s.start), new Date(s.end)]);
                    personActivity.start = formatDateTimeLocal(new Date(Math.min(...allDates)));
                    personActivity.end = formatDateTimeLocal(new Date(Math.max(...allDates)));
                    
                    ganttData.activities.push(personActivity);
                    serviceGroup.children.push(personActivity.id);
                    importedCount++;
                }
            }
            
            if (importedCount === 0) {
                alert('Inga personer med schemalagd tid hittades i Excel-filen.');
                return;
            }
            
            // Uppdatera applikationens tillstånd
            setCurrentProjectFile(null, file.name);
            document.getElementById('startDate').value = ganttData.settings.startDate;
            document.getElementById('endDate').value = ganttData.settings.endDate;
            document.getElementById('showHolidays').checked = ganttData.settings.showHolidays;
            document.getElementById('dateFormat').value = ganttData.settings.country;
            
            Object.assign(groupOpenState, ganttData.groupOpenState);
            activities = normalizeImportedActivities(ganttData);
            try { if (typeof rebuildActivityIndex === 'function') rebuildActivityIndex(); } catch(_) {}

            history = [];
            historyIndex = -1;
            pushHistory();
            if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render();
            alert('Bemanningsplan importerad från Excel.');
        } catch (err) {
            console.error('Fel vid Excel-import:', err);
            alert('Kunde inte importera Excel-filen. Kontrollera formatet.');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Koppla input-change händelser för filval (fallback-lägen)
setTimeout(() => {
    const excelEl = document.getElementById('excelFileInput');
    if (excelEl && !excelEl._ganttBound) {
        excelEl.addEventListener('change', (e) => {
            const f = e.target.files && e.target.files[0];
            if (f) handleExcelImport(f);
            e.target.value = '';
        });
        excelEl._ganttBound = true;
    }
    const openJsonEl = document.getElementById('openFileInput');
    if (openJsonEl && !openJsonEl._ganttBound) {
        openJsonEl.addEventListener('change', (e) => {
            const f = e.target.files && e.target.files[0];
            if (f) loadFromFile(f);
            e.target.value = '';
        });
        openJsonEl._ganttBound = true;
    }
}, 0);

// Exponera nyckelfunktioner globalt (för onclick i HTML)
try {
    window.openFile = openFile;
    window.saveFile = saveFile;
    window.saveAsFile = saveAsFile;
    window.exportVisibleJSON = exportVisibleJSON;
    window.importExcelFile = importExcelFile;
    window.pushHistory = pushHistory;
} catch(_) {}

// === Export to Excel (.xlsx) with preserved hierarchy ===

function computeGroupDates(group) {
    let minStart = null, maxEnd = null;
    const stack = Array.isArray(group.children) ? group.children.slice() : [];
    while (stack.length) {
        const cid = stack.shift();
        const child = typeof findActivityById === 'function' ? findActivityById(cid) : null;
        if (!child) continue;
        if (child.isGroup) {
            if (Array.isArray(child.children)) stack.push(...child.children);
        }
        if (Array.isArray(child.segments) && child.segments.length) {
            child.segments.forEach(s => {
                const sStart = normalizeDate(s.start);
                const sEnd = normalizeDate(s.end);
                if (sStart && (!minStart || sStart < minStart)) minStart = sStart;
                if (sEnd && (!maxEnd || sEnd > maxEnd)) maxEnd = sEnd;
            });
        } else {
            const sStart = normalizeDate(child.start);
            const sEnd = normalizeDate(child.end);
            if (sStart && (!minStart || sStart < minStart)) minStart = sStart;
            if (sEnd && (!maxEnd || sEnd > maxEnd)) maxEnd = sEnd;
        }
    }
    return { start: minStart, end: maxEnd };
}

function sumActivitySegmentDays(act, countWorkingDaysFn) {
    if (!act || !Array.isArray(act.segments) || act.segments.length === 0) return 0;
    return act.segments.reduce((days, s) => days + countWorkingDaysFn(s.start, s.end), 0);
}

function countCalendarDaysBetween(start, end) {
    const s = normalizeDate(start); const e = normalizeDate(end);
    if (!s || !e || e < s) return 0;
    const s0 = new Date(s.getFullYear(), s.getMonth(), s.getDate());
    const e0 = new Date(e.getFullYear(), e.getMonth(), e.getDate());
    return Math.round((e0 - s0) / (1000*60*60*24)) + 1; // inclusive
}

function sumActivityCalendarDays(act) {
    if (!act || !Array.isArray(act.segments) || act.segments.length === 0) return act && act.start && act.end ? countCalendarDaysBetween(act.start, act.end) : 0;
    return act.segments.reduce((days, s) => days + countCalendarDaysBetween(s.start, s.end), 0);
}

function computeGroupAggregates(group, countWorkingDaysFn, recursive = false) {
    let totalHours = 0;
    let totalDays = 0;
    let totalCalendarDays = 0;
    const children = Array.isArray(group.children) ? group.children : [];

    children.forEach(cid => {
        const child = (typeof findActivityById === 'function') ? findActivityById(cid) : null;
        if (!child) return;

        if (child.isGroup) {
            // UI-mode: recurse so group totals equal what the program shows.
            if (recursive) {
                const sub = computeGroupAggregates(child, countWorkingDaysFn, true);
                totalHours += sub.hours;
                totalDays += sub.days;
                totalCalendarDays += sub.calDays;
            }
            return; // detailed-mode: skip nested groups
        }

        // Leaf activity (with or without segments)
        if (Array.isArray(child.segments) && child.segments.length) {
            child.segments.forEach(s => {
                const hpd = (typeof s.hoursPerDay === 'number')
                    ? s.hoursPerDay
                    : (typeof child.hoursPerDay === 'number' ? child.hoursPerDay : getStandardHoursPerDay());
                const d = countWorkingDaysFn(s.start, s.end);
                totalHours += hpd * d;
                totalDays += d;
                totalCalendarDays += countCalendarDaysBetween(s.start, s.end);
            });
        } else if (child.start && child.end) {
            const d = countWorkingDaysFn(child.start, child.end);
            const hpd = (typeof child.hoursPerDay === 'number') ? child.hoursPerDay : getStandardHoursPerDay();
            totalHours += hpd * d;
            totalDays += d;
            totalCalendarDays += countCalendarDaysBetween(child.start, child.end);
        }
    });

    return { hours: totalHours, days: totalDays, calDays: totalCalendarDays };
}

async function buildExcelAOA(mode = 'ui') {
    // Use same day logic as UI, with safe fallbacks
    function countEffectiveDaysBetween(start, end) {
        const fn = (typeof getEffectiveDays === 'function') ? getEffectiveDays : getWorkDays;
        return fn(start, end);
    }

    const STD_HPD = (typeof getStandardHoursPerDay === 'function') ? getStandardHoursPerDay() : 8;

    // Only two numeric columns we care about
    const headers = [
        'WBS', 'Level', 'Type', 'Name', 'Start', 'End',
        'WorkDays', 'WorkHours',
        'Completed', 'ID', 'ParentWBS', 'Color'
    ];
    const aoa = [headers];
    const tops = (activities || []).filter(a => !a.parent);

    function fmtDate(d) { const nd = normalizeDate(d); return nd ? formatDateLocal(nd) : ''; }

    function computeGroupAggregatesHours(group, recursive) {
        let hours = 0;
        const children = Array.isArray(group.children) ? group.children : [];
        children.forEach(cid => {
            const child = (typeof findActivityById === 'function') ? findActivityById(cid) : null;
            if (!child) return;
            if (child.isGroup) {
                if (recursive) hours += computeGroupAggregatesHours(child, true);
                return;
            }
            if (Array.isArray(child.segments) && child.segments.length) {
                child.segments.forEach(s => {
                    const hpd = (typeof s.hoursPerDay === 'number')
                        ? s.hoursPerDay
                        : (typeof child.hoursPerDay === 'number' ? child.hoursPerDay : STD_HPD);
                    const d = countEffectiveDaysBetween(s.start, s.end);
                    hours += hpd * d;
                });
            } else if (child.start && child.end) {
                const d = countEffectiveDaysBetween(child.start, child.end);
                const hpd = (typeof child.hoursPerDay === 'number') ? child.hoursPerDay : STD_HPD;
                hours += hpd * d;
            }
        });
        return hours;
    }

    function computeActivityHours(act) {
        if (Array.isArray(act.segments) && act.segments.length) {
            return act.segments.reduce((sum, s) => {
                const hpd = (typeof s.hoursPerDay === 'number')
                    ? s.hoursPerDay
                    : (typeof act.hoursPerDay === 'number' ? act.hoursPerDay : STD_HPD);
                const d = countEffectiveDaysBetween(s.start, s.end);
                return sum + (hpd * d);
            }, 0);
        } else if (act.start && act.end) {
            const hpd = (typeof act.hoursPerDay === 'number') ? act.hoursPerDay : STD_HPD;
            const d = countEffectiveDaysBetween(act.start, act.end);
            return hpd * d;
        }
        return 0;
    }

    function walk(node, level, wbs, parentWbs) {
        if (!node) return;

        const mirrorUI = (mode === 'ui');
        const name = node.name || (node.isGroup ? 'Group' : 'Activity');
        const completed = Number.isFinite(node.completed) ? node.completed : 0;
        const color = node.color || '';

        // Datum på rader: grupper får sina aggregerade datum
        let start = node.start, end = node.end;
        if (node.isGroup) {
            const range = computeGroupDates(node);
            start = range.start; end = range.end;
        }

        // Hours på denna rad
        const segCount = Array.isArray(node.segments) ? node.segments.length : 0;
        let workHours = 0;
        let showOnThisRow = true;

        if (node.isGroup) {
            // ui = rekursivt (matchar programmet), detailed = tomt på grupper
            workHours = computeGroupAggregatesHours(node, mirrorUI);
            showOnThisRow = mirrorUI;
        } else {
            workHours = computeActivityHours(node);
            // ui: visa även på multi-segment; detailed: endast löv (<=1 segment)
            showOnThisRow = mirrorUI || segCount <= 1;
        }

        const workDays = showOnThisRow ? (workHours / STD_HPD) : '';

        aoa.push([
            wbs,
            level,
            node.isGroup ? 'Group' : 'Activity',
            (level > 1 ? '  '.repeat(level - 1) : '') + name,
            fmtDate(start),
            fmtDate(end),
            workDays,
            showOnThisRow ? workHours : '',
            completed,
            node.id || '',
            parentWbs || '',
            color
        ]);

        // Barn / segment
        if (node.isGroup) {
            const kids = getChildren(node);
            kids.forEach((child, idx) => walk(child, level + 1, `${wbs}.${idx + 1}`, wbs));
        } else if (segCount > 1) {
            // Segmentrader i "Summerbar" (summationsbar). I "Översikt" lämnas de tomma.
            if (!mirrorUI) {
                node.segments.forEach((seg, idx) => {
                    const sName = (seg.name && seg.name.trim()) ? seg.name : `${name} (del ${idx + 1})`;
                    const sStart = seg.start, sEnd = seg.end;
                    const hpd = (typeof seg.hoursPerDay === 'number')
                        ? seg.hoursPerDay
                        : (typeof node.hoursPerDay === 'number' ? node.hoursPerDay : STD_HPD);
                    const shours = (sStart && sEnd) ? hpd * countEffectiveDaysBetween(sStart, sEnd) : 0;
                    aoa.push([
                        `${wbs}.${idx + 1}`,
                        level + 1,
                        'Segment',
                        '  '.repeat(level) + sName,
                        fmtDate(sStart),
                        fmtDate(sEnd),
                        shours / STD_HPD,
                        shours,
                        completed,
                        `${node.id || ''}#${idx + 1}`,
                        wbs,
                        seg.color || ''
                    ]);
                });
            }
        }
    }

    // Top-nivåer och körning
    tops.forEach((t, i) => walk(t, 1, String(i + 1), ''));
    return aoa;
}

// Skapa XLSX med två blad: Översikt (som UI) och Summerbar (summationsbar)
async function exportToExcel() {
    try {
        const aoaOverview = await buildExcelAOA('ui');
        const aoaSummable = await buildExcelAOA('detailed');

        const wsOverview = XLSX.utils.aoa_to_sheet(aoaOverview);
        const wsSummable = XLSX.utils.aoa_to_sheet(aoaSummable);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, wsOverview, 'Översikt');
        XLSX.utils.book_append_sheet(wb, wsSummable, 'Summerbar');

        const data = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const suggestedName = `${sanitizeFileName(getProjectTitle('gantt-projekt'))}.xlsx`;

        if (window.showSaveFilePicker) {
            try {
                const handle = await window.showSaveFilePicker({
                    suggestedName,
                    types: [{ description: 'Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
                });
                if (!(await verifyPermission(handle, 'readwrite'))) throw new Error('Permission denied');
                const writable = await handle.createWritable();
                await writable.write(blob);
                await writable.close();
                alert('Excel exporterat.');
                return;
            } catch (e) {
                if (e && e.name === 'AbortError') return;
                if (window.developerMode) console.warn('showSaveFilePicker misslyckades, använder nedladdning', e);
            }
        }
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = suggestedName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(a.href);
    } catch (err) {
        console.error('Excel-export misslyckades', err);
        alert('Kunde inte exportera Excel.');
    }
}

try { window.exportToExcel = exportToExcel; } catch(_) {}
