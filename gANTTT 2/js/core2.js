// js/core2.js (consolidated)

/**
 * Applikationens affärslogik: aktiviteter, grupper/projekt, timmar och flytt/CRUD.
 * Denna fil är UI-agnostisk; rendering triggas via renderAndPreserveScroll() när möjligt.
 */

// === Hjälpfunktioner för aktiviteter ===

function updateActivityStartEndFromSegments(activity) {
    if (activity.isGroup || !activity.segments || activity.segments.length === 0) return;
    activity.start = activity.segments.reduce((min, s) => new Date(s.start) < min ? new Date(s.start) : min, new Date(activity.segments[0].start));
    activity.end = activity.segments.reduce((max, s) => new Date(s.end) > max ? new Date(s.end) : max, new Date(activity.segments[0].end));
}

function calculateTotalHours(activity) {
    if (!activity || activity.isGroup || !activity.segments || activity.segments.length === 0) return 0;
    return activity.segments.reduce((total, segment) => {
        const segmentStart = normalizeDate(segment.start);
        const segmentEnd = normalizeDate(segment.end);
        const hoursPerDay = (typeof segment.hoursPerDay === 'number') ? segment.hoursPerDay : (activity.hoursPerDay || 0);
        const workDays = getEffectiveDays(segmentStart, segmentEnd);
        return total + (workDays * hoursPerDay);
    }, 0);
}

function cleanupAndMergeSegments(act) {
    if (!act || !act.segments || act.segments.length === 0) return;
    act.segments = act.segments
        .map(s => ({ ...s, start: normalizeDate(s.start), end: normalizeDate(s.end) }))
        .sort((a, b) => a.start.getTime() - b.start.getTime());
}

// === CRUD: skapa, radera, projekt/grupp ===

function addActivity() {
    const sel = selectedActivityId ? findActivityById(selectedActivityId) : null;
    const parentId = (sel && sel.isGroup) ? sel.id : (sel && sel.parent ? sel.parent : null);
    const startDate = parentId ? findActivityById(parentId).start : normalizeDate(document.getElementById('startDate').value);
    const endDate = new Date(startDate); endDate.setDate(startDate.getDate() + 4);

    const newActivity = {
        id: generateId(),
        name: 'Ny aktivitet',
        segments: [{ start: startDate, end: endDate }],
        isGroup: false,
        children: [],
        parent: parentId,
        color: '#007bff',
        colorName: 'Standard',
        completed: 0,
        hoursPerDay: getStandardHoursPerDay(),
    };
    updateActivityStartEndFromSegments(newActivity);
    newActivity.totalHours = calculateTotalHours(newActivity);
    newActivity._order = Date.now();

    if (parentId) {
        activities.push(newActivity);
        const parent = findActivityById(parentId);
        if (parent) { parent.children.push(newActivity.id); updateGroupAndAncestors(parentId); }
    } else {
        if (sel && !sel.parent && !sel.isGroup && !sel.isProject) {
            const idx = activities.findIndex(a => a.id === sel.id);
            activities.splice(idx + 1, 0, newActivity);
        } else {
            activities.push(newActivity);
        }
    }
    if (typeof rowCapacity === 'number' && activities.filter(a => !a.parent || a.isGroup || a.parent).length > rowCapacity) {
        rowCapacity = activities.length;
    }
    pushHistory();
    if (typeof renderAndPreserveScroll === 'function') { renderAndPreserveScroll(); } else { render(); }
}

function addGroup() {
    const startDate = normalizeDate(document.getElementById('startDate').value);
    const newGroup = {
        id: generateId(),
        name: 'Ny grupp',
        start: startDate,
        end: new Date(startDate),
        isGroup: true,
        isProject: false,
        children: [],
        parent: null,
        color: '#6c757d',
        colorName: 'Grupp',
        completed: 0,
        maxTotalHours: null,
        totalHours: 0,
    };
    newGroup._order = Date.now();
    activities.push(newGroup);
    if (typeof rowCapacity === 'number' && activities.length > rowCapacity) rowCapacity = activities.length;
    groupOpenState[newGroup.id] = true;
    pushHistory();
    if (typeof renderAndPreserveScroll === 'function') { renderAndPreserveScroll(); } else { render(); }
}

function addProject() {
    const startDate = normalizeDate(document.getElementById('startDate').value);
    const newProject = {
        id: generateId(),
        name: 'Nytt prosjekt',
        start: startDate,
        end: new Date(startDate),
        isGroup: true,
        isProject: true,
        children: [],
        parent: null,
        color: '#343a40',
        colorName: 'Prosjekt',
        completed: 0,
        maxTotalHours: null,
        totalHours: 0,
    };
    newProject._order = Date.now();
    activities.push(newProject);
    if (typeof rowCapacity === 'number' && activities.length > rowCapacity) rowCapacity = activities.length;
    groupOpenState[newProject.id] = true;
    pushHistory();
    if (typeof renderAndPreserveScroll === 'function') { renderAndPreserveScroll(); } else { render(); }
}

// === Lookup och färgpalett ===

function findActivityById(id) {
    if (!id) return null;
    if (typeof activityIndex?.get === 'function') {
        const hit = activityIndex.get(id);
        if (hit) return hit;
    }
    const found = activities.find(a => a && a.id === id) || null;
    if (found && typeof activityIndex?.set === 'function') activityIndex.set(id, found);
    return found;
}

function ensureProjectHasColor(projectOrId, color, name) {
    const proj = typeof projectOrId === 'string' ? findActivityById(projectOrId) : projectOrId;
    if (!proj) return;
    proj.colorPalette = proj.colorPalette || [];
    const exists = proj.colorPalette.find(p => p.color && p.color.toLowerCase() === (color||'').toLowerCase());
    if (!exists) {
        proj.colorPalette.push({ color: color, name: name || `Färg (${color})` });
    } else if (name && (!exists.name || exists.name === `Färg (${color})`)) {
        exists.name = name;
    }
}

// === Radera aktivitet ===

function deleteActivity(actId) {
    const actToDelete = findActivityById(actId);
    if (!actToDelete) return;
    if (actToDelete.parent) removeChildFromParent(actToDelete.id, actToDelete.parent);

    const toDelete = new Set([actId]);
    (function collect(id){
        const a = findActivityById(id);
        if (a?.isGroup) (a.children||[]).forEach(cid => { toDelete.add(cid); collect(cid); });
    })(actId);

    activities = activities.filter(a => !toDelete.has(a.id));
    pushHistory();
    if (typeof renderAndPreserveScroll === 'function') { renderAndPreserveScroll(); } else { render(); }
}

// === Hierarki och flytt ===

function removeChildFromParent(childId, parentId) {
    const parent = findActivityById(parentId);
    if (parent?.isGroup) parent.children = parent.children.filter(id => id !== childId);
}

function isDescendant(childId, potentialParentId) {
    let current = findActivityById(childId);
    while (current?.parent) { if (current.parent === potentialParentId) return true; current = findActivityById(current.parent); }
    return false;
}

function moveActivity(sourceId, targetId, position) {
    const sourceAct = findActivityById(sourceId);
    const targetAct = findActivityById(targetId);
    if (!sourceAct || !targetAct) return;
    if (position === 'on' && isDescendant(targetId, sourceId)) { alert('En grupp kan inte flyttas till en av dess egna underaktiviteter.'); return; }

    const oldParentId = sourceAct.parent || null;
    if (oldParentId) removeChildFromParent(sourceId, oldParentId);
    sourceAct.parent = null;

    const sourceIndex = activities.findIndex(a => a.id === sourceId);
    activities.splice(sourceIndex, 1);
    const targetIndex = activities.findIndex(a => a.id === targetId);

    if (position === 'on') {
        if (!targetAct.isGroup) { targetAct.isGroup = true; targetAct.children = targetAct.children || []; }
        targetAct.children = (targetAct.children || []).filter(id => id !== sourceId);
        targetAct.children.push(sourceId);
        sourceAct.parent = targetId;
        activities.splice(targetIndex + 1, 0, sourceAct);
    } else {
        const newParentId = targetAct.parent || null;
        sourceAct.parent = newParentId;
        if (newParentId) {
            const parent = findActivityById(newParentId);
            if (parent && parent.isGroup) {
                parent.children = parent.children || [];
                const pos = parent.children.indexOf(targetId);
                const insertAt = (pos >= 0) ? (position === 'above' ? pos : pos + 1) : parent.children.length;
                parent.children.splice(insertAt, 0, sourceId);
            }
        }
        let newIndex = (position === 'below') ? targetIndex + 1 : targetIndex;
        if (sourceIndex < targetIndex && position === 'above') newIndex--;
        if (sourceIndex < targetIndex && position === 'below') newIndex--;
        activities.splice(newIndex, 0, sourceAct);
    }

    updateAllGroups();
    pushHistory();
    reindexTopLevelOrders();
    if (typeof renderAndPreserveScroll === 'function') { renderAndPreserveScroll(); } else { render(); }
}

function reindexTopLevelOrders() {
    let idx = 0;
    activities.forEach(a => { if (!a.parent) a._order = idx++; });
}

// === Grupp- och timberäkningar ===

function updateGroupFromChildren(group) {
    if (!group?.isGroup) return;
    if (group.children.length === 0) { group.totalHours = 0; group.completed = 0; group.end = new Date(group.start); return; }
    let minStart = null, maxEnd = null, totalHours = 0, weightedCompletedSum = 0;
    group.children.forEach(childId => {
        const child = findActivityById(childId);
        if (!child) return;
        if (child.isGroup) updateGroupFromChildren(child); else updateActivityStartEndFromSegments(child);
        if (child.start && (!minStart || child.start < minStart)) minStart = child.start;
        if (child.end && (!maxEnd || child.end > maxEnd)) maxEnd = child.end;
        const childHours = child.totalHours || 0;
        totalHours += childHours;
        weightedCompletedSum += ((child.completed || 0) / 100) * childHours;
    });
    if (minStart && maxEnd) { group.start = minStart; group.end = maxEnd; }
    group.totalHours = totalHours;
    group.completed = totalHours > 0 ? Math.round((weightedCompletedSum / totalHours) * 100) : 0;
}

function updateAllGroups() { activities.filter(a => a.isGroup).forEach(g => updateGroupFromChildren(g)); }

function updateGroupAndAncestors(groupId) {
    let current = findActivityById(groupId);
    while (current) { if (current.isGroup) updateGroupFromChildren(current); current = current.parent ? findActivityById(current.parent) : null; }
}

function getHoursForActivity(act, from = null, to = null) {
    if (act.isGroup) {
        return act.children.reduce((sum, childId) => {
            const child = findActivityById(childId);
            return sum + (child ? getHoursForActivity(child, from, to) : 0);
        }, 0);
    }
    if (!act.segments) return act.totalHours || 0;
    return act.segments.reduce((total, segment) => {
        const segmentStart = normalizeDate(segment.start);
        const segmentEnd = normalizeDate(segment.end);
        const segHpd = (typeof segment.hoursPerDay === 'number') ? segment.hoursPerDay : (act.hoursPerDay || 0);
        if (!from || !to) return total + (getEffectiveDays(segmentStart, segmentEnd) * segHpd);
        const overlapStart = new Date(Math.max(segmentStart.getTime(), normalizeDate(from).getTime()));
        const overlapEnd = new Date(Math.min(segmentEnd.getTime(), normalizeDate(to).getTime()));
        if (overlapStart > overlapEnd) return total;
        return total + (getEffectiveDays(overlapStart, overlapEnd) * segHpd);
    }, 0);
}

function getHoursForActivityFiltered(act, visibleSet, from = null, to = null) {
    if (!act) return 0;
    if (act.isGroup) {
        const kids = (act.children || []).map(id => findActivityById(id)).filter(Boolean);
        return kids.reduce((sum, child) => {
            if (child.isGroup) return sum + getHoursForActivityFiltered(child, visibleSet, from, to);
            return sum + (visibleSet?.has(child.id) ? getHoursForActivity(child, from, to) : 0);
        }, 0);
    }
    return (visibleSet?.has(act.id)) ? getHoursForActivity(act, from, to) : 0;
}

function getProjectPlannedHours(projectOrId) {
    const proj = typeof projectOrId === 'string' ? findActivityById(projectOrId) : projectOrId;
    if (!proj) return 0;
    return getHoursForActivity(proj);
}

function findProjectForActivity(actOrId) {
    let cur = typeof actOrId === 'string' ? findActivityById(actOrId) : actOrId;
    if (!cur) return null;
    while (cur && !cur.isProject) cur = findActivityById(cur.parent);
    return cur || null;
}

function canAllocateHours(projectId, deltaHours = 0) {
    const proj = typeof projectId === 'string' ? findActivityById(projectId) : projectId;
    if (!proj) return true;
    const max = proj.maxTotalHours || proj.projectInfo?.maxHours || null;
    if (!max) return true;
    const current = getProjectPlannedHours(proj);
    return (current + deltaHours) <= (max + 1e-6);
}

// === Avancerat: segmentdelning ===

function splitActivity(actId, splitDate, targetSegmentIndex) {
    const activity = findActivityById(actId);
    if (!activity || activity.isGroup) return;
    const splitDateNorm = normalizeDate(splitDate);
    let segmentIndex = -1;
    if (typeof targetSegmentIndex === 'number' && activity.segments[targetSegmentIndex]) {
        const s = activity.segments[targetSegmentIndex];
        if (splitDateNorm > normalizeDate(s.start) && splitDateNorm < normalizeDate(s.end)) segmentIndex = targetSegmentIndex;
    }
    if (segmentIndex === -1) {
        segmentIndex = activity.segments.findIndex(s => splitDateNorm > normalizeDate(s.start) && splitDateNorm < normalizeDate(s.end));
    }
    if (segmentIndex === -1) { alert('Delningsdatumet måste ligga inom ett befintligt segment av aktiviteten.'); return; }

    const originalSegment = activity.segments[segmentIndex];
    const originalSegmentEnd = new Date(originalSegment.end);
    const newSegment = {
        start: splitDate,
        end: originalSegmentEnd,
        name: originalSegment.name || undefined,
        info: originalSegment.info || undefined,
        color: originalSegment.color || undefined,
        responsibles: Array.isArray(originalSegment.responsibles) ? [...originalSegment.responsibles] : undefined,
    };
    originalSegment.end = new Date(splitDate.getTime() - 24 * 60 * 60 * 1000);
    activity.segments.splice(segmentIndex + 1, 0, newSegment);
    updateActivityStartEndFromSegments(activity);
    activity.totalHours = calculateTotalHours(activity);
    if (activity.parent) updateGroupAndAncestors(activity.parent);
    pushHistory();
    if (typeof renderAndPreserveScroll === 'function') { renderAndPreserveScroll(); } else { render(); }
}

function splitActivityAtWeekends(activityId) {
    const act = findActivityById(activityId);
    if (!act || act.isGroup) return;
    let needsAnotherPass = true;
    while (needsAnotherPass) {
        needsAnotherPass = false;
        let segmentToSplitIndex = -1;
        let fridayDate = null;
        for (let i = 0; i < act.segments.length; i++) {
            const segment = act.segments[i];
            let d = new Date(segment.start);
            while (d < segment.end) {
                if (d.getDay() === 5) {
                    const nextDay = new Date(d); nextDay.setDate(d.getDate() + 1);
                    if (nextDay <= segment.end) { segmentToSplitIndex = i; fridayDate = new Date(d); break; }
                }
                d.setDate(d.getDate() + 1);
            }
            if (segmentToSplitIndex !== -1) break;
        }
        if (segmentToSplitIndex !== -1) {
            const originalSegment = act.segments[segmentToSplitIndex];
            const originalEnd = new Date(originalSegment.end);
            originalSegment.end = fridayDate;
            const monday = new Date(fridayDate); monday.setDate(fridayDate.getDate() + 3);
            if (monday <= originalEnd) {
                const newSegment = {
                    start: monday, end: originalEnd,
                    name: originalSegment.name || undefined,
                    info: originalSegment.info || undefined,
                    color: originalSegment.color || undefined,
                    responsibles: Array.isArray(originalSegment.responsibles) ? [...originalSegment.responsibles] : undefined,
                };
                act.segments.splice(segmentToSplitIndex + 1, 0, newSegment);
            }
            needsAnotherPass = true;
        }
    }
    updateActivityStartEndFromSegments(act);
    act.totalHours = calculateTotalHours(act);
    if (act.parent) updateGroupAndAncestors(act.parent);
    pushHistory();
    if (typeof renderAndPreserveScroll === 'function') { renderAndPreserveScroll(); } else { render(); }
}

// === UI-relaterade hjälpfunktioner ===

function showAbout() {
    const content = `
        <h2>Om Gantt-planeraren</h2>
        <p>En modern och flexibel Gantt-planerare byggd med JavaScript.</p>
        <p><strong>Version:</strong> 3.0 (Sammanslagen)</p>
        <p>Kombinerar en robust modulär arkitektur med ett beprövat och användarvänligt gränssnitt.</p>
        <hr>
        <h4>Tangentbordsgenvägar:</h4>
        <ul>
            <li><strong>Ctrl + Z:</strong> Ångra</li>
            <li><strong>Ctrl + Y / Ctrl + Shift + Z:</strong> Gör om</li>
            <li><strong>Esc:</strong> Stäng popup/modal</li>
        </ul>
        <button class="tonal-button" onclick="closeCenteredModal()">Stäng</button>
    `;
    showCenteredModal(content);
}

function showCreatorForm() {
    const name = localStorage.getItem('creatorName') || '';
    const revision = localStorage.getItem('revisionNumber') || '1';
    const content = `
        <h2>Projektinformation</h2>
        <div class="popup-row"><label for="creatorNameInput">Skapare:</label><input type="text" id="creatorNameInput" value="${name}"></div>
        <div class="popup-row"><label for="revisionNumberInput">Revision:</label><input type="number" id="revisionNumberInput" value="${revision}" min="1"></div>
        <div class="popup-buttons"><button class="filled-button" onclick="saveCreator()">Spara</button><button class="text-button" onclick="closeCenteredModal()">Avbryt</button></div>
    `;
    showCenteredModal(content);
}

function saveCreator() {
    localStorage.setItem('creatorName', document.getElementById('creatorNameInput').value);
    localStorage.setItem('revisionNumber', document.getElementById('revisionNumberInput').value);
    closeCenteredModal();
    if (typeof renderAndPreserveScroll === 'function') { renderAndPreserveScroll(); } else { render(); }
}

function printView() { window.print(); }

function scrollToToday() {
    const todayLine = document.getElementById('todayLine');
    if (todayLine) todayLine.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
}
