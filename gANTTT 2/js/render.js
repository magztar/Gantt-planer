// js/render.js

/**
 * Ansvarar för att rita upp allt på skärmen baserat på det aktuella tillståndet.
 * Använder en CSS Grid-baserad metod för att säkerställa synkronisering.
 */
let _renderVersion = 0;
const monthYearFormatterSv = new Intl.DateTimeFormat('sv-SE', { month: 'long', year: 'numeric' });
const shortWeekdayFormatterSv = new Intl.DateTimeFormat('sv-SE', { weekday: 'short' });
let sidebarVirtualState = null;

function createHoursResolver(visibleSet = null) {
    const cache = new Map();

    return function resolveHours(act, from = null, to = null) {
        if (!act) return 0;

        const rangeKey = (from && to)
            ? `${formatDateLocal(normalizeDate(from))}:${formatDateLocal(normalizeDate(to))}`
            : 'all';
        const key = `${act.id}|${visibleSet ? 'filtered' : 'all'}|${rangeKey}`;
        if (cache.has(key)) return cache.get(key);

        let value = 0;
        if (act.isGroup) {
            value = (act.children || []).reduce((sum, childId) => {
                const child = findActivityById(childId);
                if (!child) return sum;
                if (!visibleSet) return sum + resolveHours(child, from, to);
                if (child.isGroup) return sum + resolveHours(child, from, to);
                return sum + (visibleSet.has(child.id) ? getHoursForActivity(child, from, to) : 0);
            }, 0);
        } else {
            value = visibleSet
                ? (visibleSet.has(act.id) ? getHoursForActivity(act, from, to) : 0)
                : getHoursForActivity(act, from, to);
        }

        cache.set(key, value);
        return value;
    };
}

function findRowIndexForOffset(rowOffsets, offset) {
    if (!Array.isArray(rowOffsets) || rowOffsets.length < 2) return 0;
    let low = 0;
    let high = rowOffsets.length - 2;
    const target = Math.max(0, offset);
    while (low < high) {
        const mid = Math.floor((low + high + 1) / 2);
        if (rowOffsets[mid] <= target) low = mid;
        else high = mid - 1;
    }
    return low;
}

function getActivitySidebarRowOffsetTop(activityId) {
    const state = sidebarVirtualState;
    if (!state || !state.indexById || !state.indexById.has(activityId)) return null;
    const rowIndex = state.indexById.get(activityId);
    return state.paddingTop + state.rowOffsets[rowIndex];
}

function requestActivitySidebarViewportUpdate(scrollTop = null, force = false) {
    const state = sidebarVirtualState;
    if (!state || !state.container) return;

    if (typeof scrollTop === 'number') state.pendingScrollTop = scrollTop;
    if (force) state.forceRerender = true;
    if (state.frameId) return;

    state.frameId = requestAnimationFrame(() => {
        state.frameId = 0;
        flushActivitySidebarViewport();
    });
}

function buildActivitySidebarRow(act, idx, state) {
    const item = document.createElement('div');
    item.className = 'activity-item';
    item.dataset.id = act.id;
    item.dataset.depth = String(act.depth || 0);
    item.style.height = `${state.rowHeights[idx] || state.baseRowHeight}px`;
    item.style.paddingLeft = `${(act.depth || 0) * 20}px`;
    item.draggable = false;
    if (act.isGroup) item.classList.add('group-row');
    if (act.isProject) item.classList.add('project-row');
    if (selectedActivityIds.has(act.id)) item.classList.add('selected');

    const leftCol = document.createElement('div');
    leftCol.className = 'activity-left';

    if (act.isGroup) {
        const toggle = document.createElement('span');
        toggle.className = 'group-toggle';
        toggle.title = groupOpenState[act.id] !== false ? 'Fäll ihop grupp' : 'Expandera grupp';
        toggle.innerHTML = `<span class="material-icons-outlined">${groupOpenState[act.id] !== false ? 'expand_more' : 'chevron_right'}</span>`;
        toggle.onclick = () => {
            try {
                window.__forceListAnchorId = act.id;
                const listEl = document.getElementById('activityList');
                const rowEl = toggle.closest('.activity-item');
                if (listEl && rowEl) {
                    window.__forceListAnchorOffset = listEl.scrollTop - rowEl.offsetTop;
                }
            } catch(_) {}
            groupOpenState[act.id] = !(groupOpenState[act.id] !== false);
            if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render();
        };
        leftCol.appendChild(toggle);
    } else {
        const spacer = document.createElement('span');
        spacer.className = 'group-toggle spacer';
        leftCol.appendChild(spacer);
    }

    const dragHandle = document.createElement('span');
    dragHandle.className = 'drag-handle';
    dragHandle.title = 'Flytta raden';
    dragHandle.draggable = true;
    dragHandle.innerHTML = '<span class="material-icons-outlined">drag_indicator</span>';
    dragHandle.addEventListener('dragstart', e => handleDragStart(e, act.id));
    dragHandle.addEventListener('dragend', () => hideInteractionTooltip());
    leftCol.appendChild(dragHandle);

    const nameInput = document.createElement('input');
    nameInput.type = 'text';
    nameInput.value = act.name;
    nameInput.oninput = e => {
        e.stopPropagation();
        const original = findActivityById(act.id);
        if (original) original.name = e.target.value;
    };
    nameInput.addEventListener('mousedown', e => { e.stopPropagation(); });
    nameInput.addEventListener('dragstart', e => e.preventDefault());
    nameInput.onblur = () => { pushHistory(); if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render(); };
    nameInput.onkeydown = e => {
        if (e.key === 'Enter') { e.preventDefault(); e.currentTarget.blur(); }
        if (e.key === 'Escape') { e.preventDefault(); e.currentTarget.blur(); }
    };
    leftCol.appendChild(nameInput);
    item.appendChild(leftCol);

    try {
        const total = state.resolveHours(act);
        const span = state.resolveHours(act, state.viewStart, state.viewEnd);
        const completionInfo = ` • ${(act.completed || 0)}%`;
        const meta = document.createElement('span');
        meta.className = 'meta';
        meta.textContent = `${span.toFixed(1)}h / ${total.toFixed(1)}h${completionInfo}`;
        if (act.maxTotalHours && total > act.maxTotalHours) meta.classList.add('warn');
        item.appendChild(meta);
    } catch (_) {}

    item.addEventListener('dragover', handleDragOver);
    item.addEventListener('dragenter', handleDragOver);
    item.addEventListener('dragleave', handleDragLeave);
    item.addEventListener('drop', e => handleDrop(e, act.id));
    item.addEventListener('click', (e) => {
        if (e.shiftKey) {
            const anchor = selectionAnchorActivityIndex != null ? selectionAnchorActivityIndex : idx;
            const [start, end] = anchor <= idx ? [anchor, idx] : [idx, anchor];
            if (!e.ctrlKey && !e.metaKey) { selectedActivityIds.clear(); selectedSegmentKeys.clear(); }
            for (let i = start; i <= end; i++) {
                const id = state.filteredActivities[i]?.id;
                if (id) selectedActivityIds.add(id);
            }
            selectedActivityId = act.id;
            highlightSelection();
        } else if (e.ctrlKey || e.metaKey) {
            toggleActivitySelection(act.id, true);
            selectionAnchorActivityIndex = idx;
        } else {
            selectedActivityId = act.id;
            selectedActivityIds.clear();
            selectedSegmentKeys.clear();
            selectedActivityIds.add(act.id);
            selectionAnchorActivityIndex = idx;
            highlightSelection();
        }
    });

    return item;
}

function flushActivitySidebarViewport() {
    const state = sidebarVirtualState;
    if (!state || !state.container) return;

    const scrollTop = typeof state.pendingScrollTop === 'number'
        ? state.pendingScrollTop
        : (state.container.scrollTop || 0);
    const viewportHeight = state.container.clientHeight || 0;
    const visibleTop = Math.max(0, scrollTop - state.paddingTop);
    const startIndex = state.totalRows
        ? findRowIndexForOffset(state.rowOffsets, Math.max(0, visibleTop - state.overscanPx))
        : 0;
    const endIndex = state.totalRows
        ? Math.min(state.totalRows, findRowIndexForOffset(state.rowOffsets, visibleTop + viewportHeight + state.overscanPx) + 2)
        : 0;

    if (!state.forceRerender && state.renderStart === startIndex && state.renderEnd === endIndex) return;

    state.forceRerender = false;
    state.renderStart = startIndex;
    state.renderEnd = endIndex;

    const fragment = document.createDocumentFragment();

    const topSpacer = document.createElement('div');
    topSpacer.className = 'activity-top-spacer activity-virtual-spacer';
    topSpacer.style.height = `${state.rowOffsets[startIndex] || 0}px`;
    fragment.appendChild(topSpacer);

    for (let idx = startIndex; idx < endIndex; idx++) {
        fragment.appendChild(buildActivitySidebarRow(state.filteredActivities[idx], idx, state));
    }

    const bottomSpacer = document.createElement('div');
    bottomSpacer.className = 'activity-bottom-spacer activity-virtual-spacer';
    bottomSpacer.style.height = `${Math.max(0, state.totalContentHeight - (state.rowOffsets[endIndex] || 0))}px`;
    fragment.appendChild(bottomSpacer);

    state.container.replaceChildren(fragment);
    state.container.scrollTop = scrollTop;
}

async function render() {
    const myVersion = ++_renderVersion;
    try { if (typeof rebuildActivityIndex === 'function') rebuildActivityIndex(); } catch(_) {}
    const ganttEl = document.getElementById('gantt');
    const activityListContainer = document.getElementById('activityList');
    if (!ganttEl || !activityListContainer) return;

    // Om en ny render startat innan vi hunnit rensa, avbryt denna direkt
    if (myVersion !== _renderVersion) return;

    // Rensa tidigare innehåll (endast om fortfarande aktuell)
    ganttEl.innerHTML = '';
    activityListContainer.innerHTML = '';
    sidebarVirtualState = null;
    try {
        window.__updateActivitySidebarViewport = null;
        window.__getActivityRowOffsetTop = null;
    } catch (_) {}
    const startDate = normalizeDate(document.getElementById('startDate').value);
    const endDate = normalizeDate(document.getElementById('endDate').value);
    if (!startDate || !endDate || endDate < startDate) {
        const showManpowerEl = document.getElementById('showManpower');
        const showManpower = !!(showManpowerEl && showManpowerEl.checked);
        ganttEl.innerHTML = '<p class="error-message">Ogiltigt datumintervall.</p>';
        return;
    }

    updateAllGroups();

    const showHolidays = document.getElementById('showHolidays').checked;
    const country = document.getElementById('dateFormat').value;
    const holidays = showHolidays ? await getHolidays(country, startDate.getFullYear(), endDate.getFullYear()) : new Set();
    
    // Om en ny render startat medan vi väntade på holidays, avbryt denna direkt
    if (myVersion !== _renderVersion) return;

    let filteredActivities = getFilteredActivitiesWithHierarchy();
    // Sätt global synlighetsmängd (för filter-justerad summering)
    try {
        const searchTermRaw = document.getElementById('searchInput')?.value || '';
        window.__filterActive = !!searchTermRaw && searchTermRaw.trim().length > 0;
        window.__visibleIdSet = new Set((filteredActivities || []).map(a => a.id));
    } catch(_) { window.__filterActive = false; window.__visibleIdSet = new Set(); }
    if (overviewMode) {
        // Visa endast projekt och grupper för översiktlig vy
        filteredActivities = filteredActivities.filter(a => a.isGroup || a.isProject);
    }
    
    // Huvudrendering av Gantt-schemat (header, grid, bars)
    renderGanttUI(ganttEl, startDate, endDate, holidays, filteredActivities);
    if (myVersion !== _renderVersion) return; // avbruten

    // Rendering av sidopanel och övrig information
    const activitySidebar = document.getElementById('activitySidebar');
    if (activitySidebar && activitySidebar.style.display === 'none') {
        activityListContainer.innerHTML = '';
        sidebarVirtualState = null;
        try {
            window.__updateActivitySidebarViewport = null;
            window.__getActivityRowOffsetTop = null;
        } catch (_) {}
    } else {
        renderActivitySidebar(activityListContainer, filteredActivities, startDate, endDate);
    }
    // Reapply selection highlight efter DOM rebuild
    if (typeof highlightSelection === 'function') highlightSelection();
    if (myVersion !== _renderVersion) return;
    renderColorPanel();
    renderFooterInfo();
    const projectPanel = document.getElementById('projectPanel');
    if (!projectPanel || projectPanel.style.display !== 'none') {
        renderProjectOverview();
    }

}

/**
 * Renderar hela Gantt-schemats gränssnitt, inklusive header, grid och staplar i en enda grid-container.
 */
function renderGanttUI(container, startDate, endDate, holidays, filteredActivities) {
    const totalDays = getWorkDays(startDate, endDate);
    const cssRowH = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--row-height') || '38', 10);
    const baseRowPx = cssRowH;            // tidigare hårdkodat 38
    const normalBarHeightPx = 28;
    const ultraZoom = (zoomLevel === 0);
    const headerRows = ultraZoom ? 1 : 3;
    const headerHeight = headerRows * 30;
    window.__headerHeight = headerHeight; // exporteras för sidolistan
    
    // Synka exakt med filtrerade aktiviteter för att låsa scrollhöjden
    rowCapacity = filteredActivities.length;
    const totalRows = filteredActivities.length;
    let dayWidth = zoomDayWidths?.[zoomLevel] ?? 35; // styrs av zoom
    if (overviewMode && ultraZoom) {
        // Gör månader smalare men tydliga i översikt
        dayWidth = Math.max(8, dayWidth);
    }
    interaction.dayWidth = dayWidth; // Tillgängliggör för interaktionslogiken

    // Beräkna dynamiska radhöjder baserat på max antal globala lanes per aktivitet (multi-dags-stöd)
    const basePaddingPx = Math.max(0, Math.round((baseRowPx - normalBarHeightPx) / 2));
    const rowHeights = [];
    const laneLayoutCache = new Map();
    for (let r = 0; r < totalRows; r++) {
        const act = filteredActivities[r];
        let height = baseRowPx;
        if (act && !act.isGroup && act.segments?.length) {
            let laneInfo = laneLayoutCache.get(act.id);
            if (!laneInfo) {
                laneInfo = assignGlobalLanes(act);
                laneLayoutCache.set(act.id, laneInfo);
            }
            if (laneInfo.maxLanes > 1) {
                const miniHeightPx = 12, miniGapPx = 3, miniStepPx = miniHeightPx + miniGapPx;
                const needed = (2 * basePaddingPx) + miniHeightPx + ((laneInfo.maxLanes - 1) * miniStepPx);
                height = Math.max(baseRowPx, needed);
            }
        }
        rowHeights.push(height);
    }

    // Expose for sidebar alignment
    try { window.__rowHeights = rowHeights.slice(); } catch(_) {}

    const ganttGrid = document.createElement('div');
    ganttGrid.className = 'gantt-grid';
    ganttGrid.style.display = 'grid';
    ganttGrid.style.gridTemplateColumns = `repeat(${totalDays}, ${dayWidth}px)`;
    // Expose header height as CSS var for masking overlay (både på grid och root för ev. annan styling)
    ganttGrid.style.setProperty('--gantt-header-height', `${headerHeight}px`);
    document.documentElement.style.setProperty('--gantt-header-height', `${headerHeight}px`);
    
    // Header rader + dynamiska radhöjder
    if (ultraZoom) {
        ganttGrid.style.gridTemplateRows = `30px ${rowHeights.map(h => `${h}px`).join(' ')}`;
    } else {
        ganttGrid.style.gridTemplateRows = `30px 30px 30px ${rowHeights.map(h => `${h}px`).join(' ')}`;
    }

    // --- Render Header ---
    renderTimelineHeader(ganttGrid, startDate, endDate, totalDays);

    // --- Render Grid Cells ---
    renderGridCells(ganttGrid, startDate, holidays, totalDays, totalRows, filteredActivities);
    
    // --- Render Activity Bars ---
    renderActivityBars(ganttGrid, filteredActivities, startDate, endDate, laneLayoutCache);

    // --- Render Today Line ---
    renderTodayLine(ganttGrid, startDate, endDate);

    container.appendChild(ganttGrid);
}

/**
 * Justerar sidolistans paddingTop dynamiskt så att första aktivitetsraden
 * linjerar exakt med första stapelraden i gantt-griden, även om
 * sidopanelens headerhöjd eller zoomläge varierar.
 */
// (Tidigare transform-baserad linjering borttagen – ersatt av deterministisk padding-formel)

// Tilldela globala lanes över en aktivitets alla segment så att överlapp inte hamnar i samma lane
function assignGlobalLanes(activity) {
    const segments = (activity.segments || []).map((seg, idx) => {
        if (!seg || !seg.start || !seg.end) return null;
        const s = new Date(seg.start);
        const e = new Date(seg.end);
        // Grouping key bias: if responsibles exist use them, else use segment name. If neither, don't group.
        const owners = Array.isArray(seg.responsibles) && seg.responsibles.length ? [...seg.responsibles].sort().join('|') : '';
        const key = owners || (seg.name ? String(seg.name) : null);
        return { idx, s, e, key };
    }).filter(Boolean);
    // Sort by start then end, tie-break by key to keep similar items adjacent
    segments.sort((a, b) => (a.s - b.s) || (a.e - b.e) || ((a.key || '').localeCompare(b.key || '')));

    // Global lanes to avoid collisions; each lane holds the last end Date
    const laneEnds = [];
    // Preferred lane indices per group key
    const reservedByKey = new Map(); // key -> number[] of lane indices (stable order)
    const laneIndexBySegment = new Map();

    for (const item of segments) {
        const start = item.s;
        const end = item.e;
        const k = item.key; // may be null
        let placed = false;

        // 1) Try group's reserved lanes first (if any)
        if (k && reservedByKey.has(k)) {
            const reserved = reservedByKey.get(k);
            for (let i = 0; i < reserved.length; i++) {
                const li = reserved[i];
                if (laneEnds[li] < start) {
                    laneEnds[li] = end;
                    laneIndexBySegment.set(item.idx, li);
                    placed = true;
                    break;
                }
            }
        }

        // 2) Try any free global lane
        if (!placed) {
            for (let li = 0; li < laneEnds.length; li++) {
                if (laneEnds[li] < start) {
                    laneEnds[li] = end;
                    laneIndexBySegment.set(item.idx, li);
                    // If grouped, remember this lane for the key (to keep future items close)
                    if (k) {
                        if (!reservedByKey.has(k)) reservedByKey.set(k, []);
                        const arr = reservedByKey.get(k);
                        if (!arr.includes(li)) arr.push(li);
                    }
                    placed = true;
                    break;
                }
            }
        }

        // 3) No free lane; open a new one and reserve it for the group (if any)
        if (!placed) {
            const newIndex = laneEnds.length;
            laneEnds.push(end);
            laneIndexBySegment.set(item.idx, newIndex);
            if (k) {
                if (!reservedByKey.has(k)) reservedByKey.set(k, []);
                reservedByKey.get(k).push(newIndex);
            }
        }
    }

    return { laneIndexBySegment, maxLanes: laneEnds.length };
}

function renderTimelineHeader(gridContainer, startDate, endDate, totalDays) {
    const monthGroups = new Map();
    const weekGroups = new Map();
    const quarterGroups = new Map();
    const headerFragment = document.createDocumentFragment();
    const ultraZoom = (zoomLevel === 0);

    for (let d = new Date(startDate), i = 0; i < totalDays; d.setDate(d.getDate() + 1), i++) {
        const date = new Date(d);
        const monthKey = `${date.getFullYear()}-${date.getMonth()}`;
        const weekKey = `${date.getFullYear()}-${getWeekNumber(date)}`;
        const quarter = Math.floor(date.getMonth() / 3) + 1;
        const quarterKey = `${date.getFullYear()}-Q${quarter}`;

        if (!monthGroups.has(monthKey)) monthGroups.set(monthKey, { label: monthYearFormatterSv.format(date), start: i + 1, count: 0 });
        monthGroups.get(monthKey).count++;
        if (!quarterGroups.has(quarterKey)) quarterGroups.set(quarterKey, { label: `Q${quarter} ${date.getFullYear()}`, start: i + 1, count: 0 });
        quarterGroups.get(quarterKey).count++;

        if (!ultraZoom) {
            if (!weekGroups.has(weekKey)) weekGroups.set(weekKey, { label: `v${getWeekNumber(date)}`, start: i + 1, count: 0 });
            weekGroups.get(weekKey).count++;

            const dayDiv = document.createElement('div');
            dayDiv.className = 'timeline-day';
            dayDiv.style.gridColumn = `${i + 1}`;
            dayDiv.style.gridRow = '3';
            dayDiv.innerHTML = `<span>${date.getDate()}</span><span class="weekday">${shortWeekdayFormatterSv.format(date)}</span>`;
            headerFragment.appendChild(dayDiv);
        }
    }

    if (overviewMode && ultraZoom) {
        quarterGroups.forEach(group => {
            const qDiv = document.createElement('div');
            qDiv.className = 'timeline-month'; // reuse month styles
            qDiv.textContent = group.label;
            qDiv.style.gridColumn = `${group.start} / span ${group.count}`;
            qDiv.style.gridRow = '1';
            headerFragment.appendChild(qDiv);
        });
    } else {
        monthGroups.forEach(group => {
            const monthDiv = document.createElement('div');
            monthDiv.className = 'timeline-month';
            monthDiv.textContent = group.label;
            monthDiv.style.gridColumn = `${group.start} / span ${group.count}`;
            monthDiv.style.gridRow = '1';
            headerFragment.appendChild(monthDiv);
        });
    }

    if (!ultraZoom) {
        weekGroups.forEach(group => {
            const weekDiv = document.createElement('div');
            weekDiv.className = 'timeline-week';
            weekDiv.textContent = group.label;
            weekDiv.style.gridColumn = `${group.start} / span ${group.count}`;
            weekDiv.style.gridRow = '2';
            headerFragment.appendChild(weekDiv);
        });
    }
    
    gridContainer.appendChild(headerFragment);
}

function renderGridCells(gridContainer, startDate, holidays, totalDays, totalRows, filteredActivities) {
    const today = normalizeDate(new Date());
    const ultraZoom = (zoomLevel === 0);
    const headerRows = ultraZoom ? 1 : 3;
    const startRowIndex = headerRows + 1; // first content row
    const rowHeights = Array.isArray(window.__rowHeights) ? window.__rowHeights : new Array(totalRows).fill(38);

    const frag = document.createDocumentFragment();

    // Create one column element per day spanning all content rows
    for (let col = 0; col < totalDays; col++) {
        const d = new Date(startDate);
        d.setDate(d.getDate() + col);
        const dateStr = formatDateLocal(d);
        const colDiv = document.createElement('div');
        let classes = 'grid-col';
        if (holidays && typeof holidays.has === 'function' && holidays.has(dateStr)) {
            classes += ' holiday';
            try {
                const name = holidays.get(dateStr);
                if (name) {
                    colDiv.title = name;
                    colDiv.setAttribute('aria-label', name);
                }
            } catch(_) {}
        }
        if (d.getDay() === 0 || d.getDay() === 6) classes += ' weekend';
        if (d.getTime() === today.getTime()) classes += ' today-bg';
        colDiv.className = classes;
        colDiv.style.gridColumn = String(col + 1);
        // span all content rows
        colDiv.style.gridRow = `${startRowIndex} / span ${totalRows}`;
        frag.appendChild(colDiv);
    }

    // Add horizontal row dividers (one per row) spanning all columns
    for (let row = 0; row < totalRows; row++) {
        const divider = document.createElement('div');
        divider.className = 'row-divider';
        divider.style.gridColumn = `1 / span ${totalDays}`;
        divider.style.gridRow = String(row + startRowIndex);
        frag.appendChild(divider);
    }

    gridContainer.appendChild(frag);

    // Container-level interactions for empty grid area
    const parentScrollEl = document.querySelector('.gantt-container');
    const dayWidth = interaction?.dayWidth || 35;

    function getGridIndicesFromEvent(e) {
        // Use grid's left for X, scroller's top for Y to avoid sticky/header offsets
        const gridRect = gridContainer.getBoundingClientRect();
        const scrollerRect = parentScrollEl ? parentScrollEl.getBoundingClientRect() : gridRect;
        const scrollLeft = parentScrollEl ? parentScrollEl.scrollLeft : 0;
        const scrollTop = parentScrollEl ? parentScrollEl.scrollTop : 0;
        const xInGrid = (e.clientX - gridRect.left);
        const yInGrid = (e.clientY - scrollerRect.top) + scrollTop;
        const col = Math.max(0, Math.floor(xInGrid / dayWidth));
        // Map y to row using variable row heights
        const headerHeight = headerRows * 30;
        const yContent = yInGrid - headerHeight;
        if (yContent < 0) return { row: -1, col };
        let acc = 0, row = -1;
        for (let i = 0; i < rowHeights.length; i++) {
            const h = rowHeights[i] || 38;
            if (yContent >= acc && yContent < acc + h) { row = i; break; }
            acc += h;
        }
        return { row, col };
    }

    // Right-click empty space: create segment
    gridContainer.addEventListener('contextmenu', async (e) => {
        // If right-clicked on a bar, let bar handler handle it
        if (e.target.closest('.bar')) return;
        e.preventDefault();
        const { row, col } = getGridIndicesFromEvent(e);
        if (row < 0 || row >= filteredActivities.length || col < 0 || col >= totalDays) return;
        await handleEmptyCellContextMenu(row, col, startDate, filteredActivities);
    });

    // Alt-click sets paste anchor date
    gridContainer.addEventListener('click', (e) => {
        if (!e.altKey) return;
        if (e.target.closest('.bar')) return;
        const { col } = getGridIndicesFromEvent(e);
        const d2 = new Date(startDate);
        d2.setDate(d2.getDate() + col);
        pasteAnchorDate = normalizeDate(d2);
        showInteractionTooltip(`Inklistringsdatum: ${formatDateLocal(pasteAnchorDate)}`, e.clientX, e.clientY);
        setTimeout(hideInteractionTooltip, 800);
    });
}

async function handleEmptyCellContextMenu(rowIndex, colIndex, viewStartDate, filteredActivities) {
    try {
        // Räkna ut datumet från kolumn-index
        const date = new Date(viewStartDate);
        date.setDate(date.getDate() + colIndex);

        const act = filteredActivities[rowIndex];
        // Om raden ligger utanför det synliga aktivitetsspannet: skapa inget
        if (rowIndex >= filteredActivities.length) {
            return; // tom buffert-rad, ignoreras
        }
        if (act && !act.isGroup) {
            // Anchor restoration around the acted row (existing in OLD DOM)
            try { window.__forceListAnchorId = act.id; } catch(_) {}
            // Skapa ett endagssegment för aktiviteten
            act.segments = act.segments || [];
            act.segments.push({ start: normalizeDate(date), end: normalizeDate(date) });
            cleanupAndMergeSegments(act);
            updateActivityStartEndFromSegments(act);
            act.totalHours = calculateTotalHours(act);
            if (act.parent) updateGroupAndAncestors(act.parent);
            pushHistory();
            if (typeof renderAndPreserveScroll === 'function') await renderAndPreserveScroll(); else render();
            return;
        }

        // Om grupprad eller tom rad: skapa en ny aktivitet (i gruppen om vald rad är en grupp)
        const parentId = act && act.isGroup ? act.id : null;
        const newActivity = {
            id: generateId(),
            name: 'Ny aktivitet',
            segments: [{ start: normalizeDate(date), end: normalizeDate(date) }],
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
        activities.push(newActivity);
        if (parentId) {
            const parent = findActivityById(parentId);
            if (parent) {
                parent.children.push(newActivity.id);
                updateGroupAndAncestors(parentId);
            }
        }
        // Anchor restoration around the parent/group row (existing in OLD DOM)
        try { window.__forceListAnchorId = parentId || null; } catch(_) {}
        pushHistory();
        if (typeof renderAndPreserveScroll === 'function') await renderAndPreserveScroll(); else render();
    } catch (err) {
        console.error('Kunde inte skapa segment:', err);
    }
}

function renderActivityBars(gridContainer, filteredActivities, viewStartDate, viewEndDate, laneLayoutCache = null) {
    // Lokala konstanter för mini-stapelhöjd och mellanrum (måste matcha CSS)
    const miniHeightPx = 12;
    const miniGapPx = 3;
    const miniStepPx = miniHeightPx + miniGapPx;
    // Samma topp/botten-luft som singelstapelrad (baserad på baseRowPx 38 och .bar 28)
    const normalBarHeightPx = 28;
    const baseRowPx = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--row-height') || '34', 10);
    const basePaddingPx = Math.max(0, Math.round((baseRowPx - normalBarHeightPx) / 2));
    const barsFragment = document.createDocumentFragment();

    // Manpower overlay on bars only (per activity/segment only; no cross-activity aggregation)
    const showManpower = !!document.getElementById('showManpower')?.checked;

    filteredActivities.forEach((act, index) => {
    if (act.isGroup) {
            // Rendera grupp/projekt som union av barns aktiva dagar inom vyn
            const vs = normalizeDate(viewStartDate);
            const ve = normalizeDate(viewEndDate);
            const activeDays = new Set(); // number (timestamp for normalized day)
            const dayHours = new Map();   // number -> total hours (för badge)

            function collectDescendants(id) {
                const node = findActivityById(id);
                if (!node) return;
                if (node.isGroup) {
                    (node.children || []).forEach(collectDescendants);
                } else if (Array.isArray(node.segments)) {
                    node.segments.forEach(seg => {
                        if (!seg || !seg.start || !seg.end) return;
                        const s = normalizeDate(seg.start);
                        const e = normalizeDate(seg.end);
                        if (e < vs || s > ve) return;
                        const hpd = (typeof seg.hoursPerDay === 'number' ? seg.hoursPerDay : (typeof node.hoursPerDay === 'number' ? node.hoursPerDay : null));
                        for (let d = new Date(Math.max(s, vs)); d <= Math.min(e, ve); d.setDate(d.getDate() + 1)) {
                            const dn = new Date(d.getFullYear(), d.getMonth(), d.getDate());
                            const key = dn.getTime();
                            activeDays.add(key);
                            // Lägg till timmar enligt inställning: vardagar eller även helg
                            if (typeof hpd === 'number' && (shouldCountWeekends() || (dn.getDay() !== 0 && dn.getDay() !== 6))) {
                                dayHours.set(key, (dayHours.get(key) || 0) + hpd);
                            }
                        }
                    });
                }
            }
            (act.children || []).forEach(collectDescendants);

            if (activeDays.size === 0) return; // Ingen stapel om inga aktiva dagar

            // Konvertera till sorterade datumobjekt
            const dates = Array.from(activeDays).sort((a,b) => a - b).map(ts => new Date(ts));
            // Komprimera till sammanhängande datumintervall
            const ranges = [];
            let rangeStart = dates[0];
            let prev = dates[0];
            for (let i = 1; i < dates.length; i++) {
                const cur = dates[i];
                const diff = Math.round((cur - prev) / (1000*60*60*24));
                if (diff === 1) {
                    prev = cur;
                } else {
                    ranges.push([rangeStart, prev]);
                    rangeStart = cur;
                    prev = cur;
                }
            }
            ranges.push([rangeStart, prev]);

            // Rendera en stapel per intervall
            ranges.forEach(([rs, re]) => {
                const startOffsetDays = getWorkDays(viewStartDate, rs);
                const durationDays = getWorkDays(rs, re);

                const barWrapper = document.createElement('div');
                barWrapper.className = 'bar-wrapper';
                barWrapper.style.gridRow = `${index + (zoomLevel === 0 ? 2 : 4)}`;
                barWrapper.style.gridColumn = `${startOffsetDays} / span ${durationDays}`;

                const bar = document.createElement('div');
                bar.className = 'bar ' + (act.isProject ? 'project' : 'group');
                bar.dataset.id = act.id;
                bar.style.backgroundColor = act.color;
                try {
                    const sStr = formatDateLocal(rs);
                    const eStr = formatDateLocal(re);
                    bar.title = `${act.name} • ${sStr} → ${eStr}`;
                } catch(_) {}

                if (!overviewMode) {
                    const label = document.createElement('span');
                    label.className = 'bar-label';
                    label.textContent = act.name;
                    bar.appendChild(label);
                }

                // Men-badge: summera endast gruppens egna barn per dag; visa topp för intervallet
                const showManpower = !!document.getElementById('showManpower')?.checked;
                if (showManpower) {
                    let peakHours = 0;
                    for (let d = new Date(rs); d <= re; d.setDate(d.getDate() + 1)) {
                        const dn = new Date(d.getFullYear(), d.getMonth(), d.getDate());
                        const key = dn.getTime();
                        const h = dayHours.get(key) || 0;
                        if (h > peakHours) peakHours = h;
                    }
                    const men = Math.ceil(peakHours / (getStandardHoursPerDay() || 8));
                    if (men > 1) {
                        const badge = document.createElement('span');
                        badge.className = 'bar-manpower-badge';
                        badge.textContent = String(men);
                        try {
                            const dw = interaction?.dayWidth || 35;
                            if (dw < 50) {
                                const scale = Math.max(0.8, dw / 50);
                                badge.style.transform = `translateY(-50%) scale(${scale})`;
                                badge.style.transformOrigin = 'right center';
                            }
                        } catch(_) {}
                        bar.appendChild(badge);
                    }
                }

                bar.addEventListener('contextmenu', e => { e.preventDefault(); openPopup(act.id, e); });
                bar.addEventListener('click', e => {
                    if (e.ctrlKey || e.metaKey) {
                        toggleActivitySelection(act.id, true);
                    } else if (e.shiftKey) {
                        toggleActivitySelection(act.id, false);
                    } else {
                        toggleActivitySelection(act.id, false);
                    }
                });

                barWrapper.appendChild(bar);
                barsFragment.appendChild(barWrapper);
            });

    } else if (!overviewMode && act.segments && act.segments.length > 0) {
            // Rendera segment-staplar för vanliga aktiviteter med globala lanes
            let laneInfo = laneLayoutCache && laneLayoutCache.get(act.id);
            if (!laneInfo) {
                laneInfo = assignGlobalLanes(act);
                if (laneLayoutCache) laneLayoutCache.set(act.id, laneInfo);
            }
            const { laneIndexBySegment, maxLanes } = laneInfo;
            act.segments.forEach((segment, segmentIndex) => {
                const segStart = normalizeDate(segment.start);
                const segEnd = normalizeDate(segment.end);
                
                if (!segStart || !segEnd || getWorkDays(segStart, segEnd) <= 0) return;

                // Clamp rendering to visible range to avoid expensive off-screen spans.
                const overlapStart = (segStart < viewStartDate) ? viewStartDate : segStart;
                const overlapEnd = (segEnd > viewEndDate) ? viewEndDate : segEnd;
                if (segEnd < viewStartDate || segStart > viewEndDate || overlapStart > overlapEnd) return;

                const startOffsetDays = getWorkDays(viewStartDate, overlapStart);
                const durationDays = getWorkDays(overlapStart, overlapEnd);
                if (durationDays <= 0) return;

                const barWrapper = document.createElement('div');
                barWrapper.className = 'bar-wrapper';
                barWrapper.style.gridRow = `${index + (zoomLevel === 0 ? 2 : 4)}`;
                barWrapper.style.gridColumn = `${startOffsetDays} / span ${durationDays}`;
                
                const bar = document.createElement('div');
                bar.className = 'bar';
                bar.dataset.id = act.id;
                bar.dataset.segmentIndex = segmentIndex;
                // Prefer segment-level color when set, otherwise fallback to activity color
                bar.style.backgroundColor = (segment.color || act.color);

                // Tooltip med datum/tid, ev. ansvariga och info
                try {
                    const sStr = isTimeSpecific(segment.start) ? formatDateTimeLocal(segment.start).replace('T', ' ') : formatDateLocal(segment.start);
                    const eStr = isTimeSpecific(segment.end) ? formatDateTimeLocal(segment.end).replace('T', ' ') : formatDateLocal(segment.end);
                    const resp = Array.isArray(segment.responsibles) && segment.responsibles.length ? ` • Ansvariga: ${segment.responsibles.join(', ')}` : '';
                    const infoStr = segment.info ? ` • ${segment.info}` : '';
                    const segTitle = segment.name ? `${segment.name} — ${act.name}` : act.name;
                    bar.title = `${segTitle} • ${sStr} → ${eStr}${resp}${infoStr}`;
                } catch(_) {}

                const isSingleDay = durationDays === 1;
                const laneIndex = laneIndexBySegment.get(segmentIndex) ?? 0;
                const totalLanes = Math.max(1, maxLanes);
                if (totalLanes > 1) {
                    // Krymp och stacka uppifrån och ner över hela längden
                    bar.classList.add('mini');
                    barWrapper.style.alignItems = 'flex-start';
                    // Första minis staplas med samma topp-luft som singelstapel
                    bar.style.marginTop = `${basePaddingPx + (laneIndex * miniStepPx)}px`;
                    // Om det är endagssegment och tid finns: visa endast delen av dagen
                    try {
                        if (isSingleDay && (isTimeSpecific(segment.start) || isTimeSpecific(segment.end))) {
                            const startMin = segment.start.getHours() * 60 + segment.start.getMinutes();
                            const endMin = segment.end.getHours() * 60 + segment.end.getMinutes();
                            const dayMin = 24 * 60;
                            const leftPct = Math.max(0, Math.min(100, (startMin / dayMin) * 100));
                            const widthPct = Math.max(0, Math.min(100, ((Math.max(endMin, startMin) - startMin) / dayMin) * 100));
                            bar.style.marginLeft = `${leftPct}%`;
                            bar.style.width = widthPct > 0 ? `${widthPct}%` : '4px';
                        }
                    } catch(_) {}
                } else {
                    // Fullhöjd om bara en lane totalt
                    if (segment.name || segmentIndex === 0) {
                        const label = document.createElement('span');
                        label.className = 'bar-label';
                        label.textContent = segment.name || act.name;
                        bar.appendChild(label);
                    }
                    if (isSingleDay && (isTimeSpecific(segment.start) || isTimeSpecific(segment.end))) {
                        try {
                            const startMin = segment.start.getHours() * 60 + segment.start.getMinutes();
                            const endMin = segment.end.getHours() * 60 + segment.end.getMinutes();
                            const dayMin = 24 * 60;
                            const leftPct = Math.max(0, Math.min(100, (startMin / dayMin) * 100));
                            const widthPct = Math.max(0, Math.min(100, ((Math.max(endMin, startMin) - startMin) / dayMin) * 100));
                            bar.style.marginLeft = `${leftPct}%`;
                            bar.style.width = widthPct > 0 ? `${widthPct}%` : '4px';
                        } catch(_) {}
                    }
                }

                    // Manpower badge on bar (per segment only)
                    if (showManpower) {
                        try {
                            const hpd = (typeof segment.hoursPerDay === 'number') ? segment.hoursPerDay : (typeof act.hoursPerDay === 'number' ? act.hoursPerDay : null);
                            if (typeof hpd === 'number') {
                                const men = Math.ceil(hpd / (getStandardHoursPerDay() || 8));
                                if (men > 1) {
                                    const badge = document.createElement('span');
                                    badge.className = 'bar-manpower-badge';
                                    badge.textContent = String(men);
                                    try {
                                        const dw = interaction?.dayWidth || 35;
                                        if (dw < 50) {
                                            const scale = Math.max(0.8, dw / 50);
                                            badge.style.transform = `translateY(-50%) scale(${scale})`;
                                            badge.style.transformOrigin = 'right center';
                                        }
                                    } catch(_) {}
                                    bar.appendChild(badge);
                                }
                            }
                        } catch(_) { /* no-op */ }
                    }

                    // Progress-logik borttagen härifrån

                // Alltid tillåt drag/resize, även för endagssegment (mini)
                const leftHandle = document.createElement('div');
                leftHandle.className = 'resize-handle left';
                const rightHandle = document.createElement('div');
                rightHandle.className = 'resize-handle right';
                bar.appendChild(leftHandle);
                bar.appendChild(rightHandle);

                bar.addEventListener('mousedown', e => {
                    if (e.target.classList.contains('resize-handle')) {
                        const side = e.target.classList.contains('left') ? 'left' : 'right';
                        startResize(e, bar, side);
                    } else {
                        // Tips: håll Ctrl/Shift för att flytta flera markerade segment samtidigt
                        startDrag(e, bar);
                    }
                });

                bar.addEventListener('contextmenu', e => { e.preventDefault(); openPopup(act.id, e); });
                bar.addEventListener('click', e => {
                    if (e.ctrlKey || e.metaKey || e.shiftKey) {
                        e.stopPropagation();
                        toggleSegmentSelection(act.id, segmentIndex, true);
                    } else {
                        openPopup(act.id, e);
                    }
                });
                
                barWrapper.appendChild(bar);
                barsFragment.appendChild(barWrapper);
            });

            // NY LOGIK: Rendera en övergripande progress-stapel för hela aktiviteten
            if (act.completed > 0) {
                const actStart = normalizeDate(act.start);
                const actEnd = normalizeDate(act.end);
                if (!actStart || !actEnd) return;
                if (actEnd < viewStartDate || actStart > viewEndDate) return;

                const spanStart = (actStart < viewStartDate) ? viewStartDate : actStart;
                const spanEnd = (actEnd > viewEndDate) ? viewEndDate : actEnd;
                const startOffsetDays = getWorkDays(viewStartDate, spanStart);
                const durationDays = getWorkDays(spanStart, spanEnd);
                if (durationDays <= 0) return;
                const progressDuration = Math.max(1, Math.round(durationDays * (act.completed / 100)));

                const progressBar = document.createElement('div');
                progressBar.className = 'bar-progress-overall';
                progressBar.style.gridRow = `${index + (zoomLevel === 0 ? 2 : 4)}`;
                progressBar.style.gridColumn = `${startOffsetDays} / span ${progressDuration}`;
                
                // Använd en mörkare nyans av aktivitetens färg för progress-stapeln
                try {
                    progressBar.style.backgroundColor = chroma(act.color).darken(0.8).alpha(0.7).hex();
                } catch (e) {
                    progressBar.style.backgroundColor = '#333'; // Fallback
                }

                const progressLabel = document.createElement('span');
                progressLabel.className = 'bar-label progress';
                progressLabel.textContent = `${act.completed}%`;
                progressBar.appendChild(progressLabel);

                barsFragment.appendChild(progressBar);
            }
        }
    });
    gridContainer.appendChild(barsFragment);
}

function renderTodayLine(gridContainer, startDate, endDate) {
    const today = normalizeDate(new Date());
    if (today < startDate || today > endDate) return;

    const todayIndex = getWorkDays(startDate, today);

    const ultraZoom = (zoomLevel === 0);
    const headerRows = ultraZoom ? 1 : 3;
    const contentStartRow = headerRows + 1;

    const todayHeaderLine = document.createElement('div');
    todayHeaderLine.id = 'todayLineHeader';
    todayHeaderLine.className = 'today-line today-line-header';
    todayHeaderLine.style.gridRow = `1 / span ${headerRows}`;
    todayHeaderLine.style.gridColumn = `${todayIndex}`;

    const todayBodyLine = document.createElement('div');
    todayBodyLine.id = 'todayLine';
    todayBodyLine.className = 'today-line today-line-body';
    todayBodyLine.style.gridRow = `${contentStartRow} / -1`;
    todayBodyLine.style.gridColumn = `${todayIndex}`;

    gridContainer.appendChild(todayHeaderLine);
    gridContainer.appendChild(todayBodyLine);
}

/**
 * KORRIGERAD: Enkel och robust funktion för att filtrera och bygga hierarkin.
 */
function getFilteredActivitiesWithHierarchy() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase();
    const visibleActivities = [];
    const viewStart = normalizeDate(document.getElementById('startDate').value);
    const viewEnd = normalizeDate(document.getElementById('endDate').value);

    function createVisibleRow(activity, depth) {
        const row = Object.create(activity);
        row.depth = depth;
        return row;
    }

    // Hjälp: kolla om aktivitet (eller container) överlappar visningsspann
    function intersectsView(act) {
        if (!act) return false;
        if (act.isGroup) {
            // grupper/projekt utan barn visas alltid om de ligger inom någon dag (start<=end)
            if (!act.start || !act.end) return true;
            return !(act.end < viewStart || act.start > viewEnd);
        }
        if (!act.segments || !act.segments.length) return true;
        return act.segments.some(seg => {
            const s = normalizeDate(seg.start);
            const e = normalizeDate(seg.end);
            return !(e < viewStart || s > viewEnd);
        });
    }

    // Filtrera bort toppnivåer (utan parent) som inte överlappar intervallet, men behåll deras barnlogik via rekursion.
    const topLevel = activities.filter(a => !a.parent);
    // Sortera projekt först men använd endast manuellt satta _order för ordning
    topLevel.sort((a,b)=>{
        const aProj = a.isProject ? 0 : 1;
        const bProj = b.isProject ? 0 : 1;
        if (aProj !== bProj) return aProj - bProj;
        // Endast använd _order så att automatisk omordning inte sker
        return (a._order||0) - (b._order||0);
    });

    // Om inget sökord finns, bygg hela den synliga hierarkin
    if (!searchTerm) {
        function addVisible(activity, depth) {
            if (!activity) return;
            if (!intersectsView(activity)) return; // helt utanför
            visibleActivities.push(createVisibleRow(activity, depth));
            if (activity.isGroup && groupOpenState[activity.id] !== false) {
                activity.children.forEach(childId => addVisible(findActivityById(childId), depth + 1));
            }
        }
        topLevel.forEach(act => addVisible(act, 0));
        return visibleActivities;
    }

    // Om sökord finns, hitta alla matchningar och deras föräldrar
    const matches = new Set();
    activities.forEach(act => {
        if (act.name.toLowerCase().includes(searchTerm)) {
            matches.add(act.id);
            let parent = findActivityById(act.parent);
            while (parent) {
                matches.add(parent.id);
                parent = findActivityById(parent.parent);
            }
        }
    });

    // Bygg hierarkin med endast de matchande aktiviteterna
    function addFiltered(activity, depth) {
        if (!activity || !matches.has(activity.id)) return;
        if (!intersectsView(activity)) return; // helt utanför
        
        visibleActivities.push(createVisibleRow(activity, depth));
        
        // Vid aktivt sökfilter: ignorera manuellt kollapsläge och gå igenom barn
        if (activity.isGroup && (groupOpenState[activity.id] !== false || searchTerm)) {
            activity.children.forEach(childId => addFiltered(findActivityById(childId), depth + 1));
        }
    }

    topLevel.forEach(act => addFiltered(act, 0));
    return visibleActivities;
}


function renderActivitySidebar(container, filteredActivities, viewStart, viewEnd) {
    if (overviewMode) {
        filteredActivities = filteredActivities.filter(a => a.isGroup || a.isProject);
    }

    const rowHeights = Array.isArray(window.__rowHeights) ? window.__rowHeights : [];
    const visibleSet = (window.__visibleIdSet instanceof Set) ? window.__visibleIdSet : new Set();
    const filterActive = !!window.__filterActive;
    const resolveHours = createHoursResolver(filterActive ? visibleSet : null);
    const baseRowHeight = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--row-height') || '34', 10);

    let paddingTop = 0;
    try {
        const sidebarHeader = document.querySelector('.sidebar-header');
        const sbHeaderH = sidebarHeader ? sidebarHeader.offsetHeight : 0;
        const ganttHeaderH = window.__headerHeight || 90;
        const ganttGrid = document.querySelector('.gantt-grid');

        container.style.boxSizing = 'border-box';
        container.style.paddingBottom = '0px';
        if (ganttGrid) ganttGrid.style.marginTop = '0px';

        const diff = ganttHeaderH - sbHeaderH;
        paddingTop = Math.max(0, diff);
        container.style.paddingTop = `${paddingTop}px`;
        if (diff < 0 && ganttGrid) ganttGrid.style.marginTop = `${Math.abs(diff)}px`;
    } catch (_) {
        container.style.paddingTop = '0px';
    }

    container.style.display = 'block';
    container.style.gridTemplateRows = '';
    container.style.gridAutoRows = '';
    container.style.alignContent = '';

    const rowOffsets = new Array(filteredActivities.length + 1).fill(0);
    const indexById = new Map();
    for (let idx = 0; idx < filteredActivities.length; idx++) {
        const rowHeight = rowHeights[idx] || baseRowHeight;
        rowOffsets[idx + 1] = rowOffsets[idx] + rowHeight;
        indexById.set(filteredActivities[idx].id, idx);
    }

    sidebarVirtualState = {
        container,
        filteredActivities,
        viewStart,
        viewEnd,
        resolveHours,
        rowHeights,
        rowOffsets,
        indexById,
        totalRows: filteredActivities.length,
        totalContentHeight: rowOffsets[rowOffsets.length - 1] || 0,
        baseRowHeight,
        paddingTop,
        overscanPx: Math.max(240, baseRowHeight * 8),
        renderStart: -1,
        renderEnd: -1,
        pendingScrollTop: document.querySelector('.gantt-container')?.scrollTop || container.scrollTop || 0,
        forceRerender: true,
        frameId: 0
    };

    try {
        window.__updateActivitySidebarViewport = requestActivitySidebarViewportUpdate;
        window.__getActivityRowOffsetTop = getActivitySidebarRowOffsetTop;
    } catch (_) {}

    flushActivitySidebarViewport();
}

function renderColorPanel() {
    const container = document.getElementById('colorPanelContainer');
    if (!container) return;
    container.innerHTML = '<h3>Färgpalett</h3>';

    // If a project is selected, show its palette. Otherwise show a global palette derived from activities.
    const selProj = typeof window.__selectedProjectId !== 'undefined' ? findActivityById(window.__selectedProjectId) : null;

    // Helper to find activities that belong to a given project
    const actsForProject = proj => activities.filter(a => findProjectForActivity(a) === proj);

    if (selProj) {
        // Ensure palette exists on project
        selProj.colorPalette = selProj.colorPalette || [];

        const list = document.createElement('div');
        list.className = 'project-palette-list';

        selProj.colorPalette.forEach((entry, idx) => {
            const item = document.createElement('div');
            item.className = 'color-item';
            const colorInput = document.createElement('input');
            colorInput.type = 'color';
            colorInput.value = entry.color || '#007bff';
            colorInput.onchange = e => {
                const old = entry.color;
                entry.color = e.target.value;
                // Update activities in this project that used the old color to use the new
                actsForProject(selProj).forEach(a => { if (a.color === old) { a.color = entry.color; a.colorName = entry.name; } });
                pushHistory(); if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render();
            };
            const nameInput = document.createElement('input');
            nameInput.type = 'text';
            nameInput.value = entry.name || `Färg (${entry.color||'#007bff'})`;
            nameInput.onchange = e => {
                entry.name = e.target.value;
                // Update activity colorName for matching colors in this project
                actsForProject(selProj).forEach(a => { if (a.color === entry.color) a.colorName = entry.name; });
                pushHistory(); if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render();
            };
            const deleteBtn = document.createElement('button');
            deleteBtn.textContent = '✕';
            deleteBtn.onclick = () => {
                // Reassign activities with this color to default and remove entry
                actsForProject(selProj).forEach(a => { if (a.color === entry.color) { a.color = '#007bff'; delete a.colorName; } });
                selProj.colorPalette.splice(idx, 1);
                pushHistory(); if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render();
            };
            item.appendChild(colorInput); item.appendChild(nameInput); item.appendChild(deleteBtn);
            list.appendChild(item);
        });

    container.appendChild(list);
    return;
        return;
    }

    // Fallback: global palette derived from activities
    const uniqueColors = new Map();
    activities.forEach(act => {
        if (!uniqueColors.has(act.color)) {
            uniqueColors.set(act.color, act.colorName || `Färg (${act.color})`);
        }
    });

    if (uniqueColors.size === 0) {
        container.innerHTML += '<p class="empty-state muted">Inga aktiviteter.</p>';
        return;
    }

    uniqueColors.forEach((name, color) => {
        const item = document.createElement('div');
        item.className = 'color-item';
        const colorInput = document.createElement('input');
        colorInput.type = 'color';
        colorInput.value = color;
        colorInput.onchange = e => {
            const newColor = e.target.value;
            activities.forEach(a => { if (a.color === color) a.color = newColor; });
            // If a project is selected, ensure it has the new color in its palette
            try {
                const selProj = typeof window.__selectedProjectId !== 'undefined' ? findActivityById(window.__selectedProjectId) : null;
                if (selProj) ensureProjectHasColor(selProj, newColor);
            } catch(_) {}
            pushHistory(); if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render();
        };
        const nameInput = document.createElement('input');
        nameInput.type = 'text';
        nameInput.value = name;
    nameInput.onchange = e => { activities.forEach(a => { if (a.color === color) a.colorName = e.target.value; }); pushHistory(); if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render(); };
        const deleteBtn = document.createElement('button');
        deleteBtn.textContent = '✕';
    deleteBtn.onclick = () => { activities.forEach(a => { if (a.color === color) { a.color = '#007bff'; delete a.colorName; } }); pushHistory(); if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render(); };
        item.appendChild(colorInput); item.appendChild(nameInput); item.appendChild(deleteBtn);
        container.appendChild(item);
    });
}

function renderFooterInfo() {
    const creatorInfoEl = document.getElementById('creatorInfo');
    const startDateVal = document.getElementById('startDate').value;
    const endDateVal = document.getElementById('endDate').value;
    // Creator info and global project summary are now handled in the project panel UI.
    // Clear legacy elements so they do not render duplicate info.
    if (creatorInfoEl) creatorInfoEl.innerHTML = '';
    const summaryEl = document.getElementById('projectSummary');
    if (summaryEl) summaryEl.innerHTML = '';
}

// === Projektöversiktsträd ===
function expandAllProjects(expand) {
    activities.filter(a => a.isGroup).forEach(g => { groupOpenState[g.id] = expand; });
    if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render();
}

// Expose for inline onclick in HTML
try { window.expandAllProjects = expandAllProjects; } catch(_) {}

// === Manpower helpers (8 timmar = 1 man) ===
function computeNameDayHours(viewStart, viewEnd) {
    const map = new Map(); // name -> Map(dateStr -> hours)
    const vs = normalizeDate(viewStart);
    const ve = normalizeDate(viewEnd);
    // Iterate activities with segments
    activities.forEach(act => {
        if (!act || act.isGroup || !Array.isArray(act.segments) || !act.name) return;
        const hpd = act.hoursPerDay || getStandardHoursPerDay();
        act.segments.forEach(seg => {
            if (!seg || !seg.start || !seg.end) return;
            // overlap with view
            const s = normalizeDate(seg.start);
            const e = normalizeDate(seg.end);
            if (e < vs || s > ve) return;
            // walk days
            for (let d = new Date(Math.max(s, vs)); d <= Math.min(e, ve); d.setDate(d.getDate() + 1)) {
                const key = formatDateLocal(d);
                if (!map.has(act.name)) map.set(act.name, new Map());
                const inner = map.get(act.name);
                inner.set(key, (inner.get(key) || 0) + hpd);
            }
        });
    });
    return map;
}

function computePeakMenForName(name, nameDayHours) {
    const inner = nameDayHours.get(name);
    if (!inner) return 0;
    let peakHours = 0;
    inner.forEach(h => { if (h > peakHours) peakHours = h; });
    // 8 timmar = 1 man; använd tak för att vara konservativ
    return Math.ceil(peakHours / 8);
}

function renderProjectOverview() {
    const container = document.getElementById('projectSummary');
    if (!container) return;
    container.innerHTML = '';

    const viewStart = normalizeDate(document.getElementById('startDate').value);
    const viewEnd = normalizeDate(document.getElementById('endDate').value);

    function hoursInSpan(act) { return getHoursForActivity(act, viewStart, viewEnd); }

    // Visa endast projekt i projektöversikten — inga aktiviteter under projekten
    const projects = activities.filter(a => a.isProject).sort((a,b)=> (a._order||0) - (b._order||0));
    if (!projects.length) {
        container.innerHTML = '<div class="empty-state muted">Inga projekt att visa</div>';
        return;
    }

    // Populate project selector at top
    try {
        const selector = document.getElementById('projectSelector');
        if (selector) {
            selector.innerHTML = '';
            projects.forEach(p => {
                const opt = document.createElement('option'); opt.value = p.id; opt.textContent = p.name; selector.appendChild(opt);
            });
            if (!window.__selectedProjectId || !projects.find(p=>p.id===window.__selectedProjectId)) window.__selectedProjectId = projects[0].id;
            selector.value = window.__selectedProjectId;
            selector.onchange = () => { window.__selectedProjectId = selector.value; if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render(); };
        }
    } catch(_) {}

    // Render details for selected project
    try {
        const detailsEl = document.getElementById('selectedProjectDetails');
        if (detailsEl) {
            const sel = projects.find(p => p.id === window.__selectedProjectId) || projects[0];
            const vis = (window.__visibleIdSet instanceof Set) ? window.__visibleIdSet : new Set();
            const filterActive = !!window.__filterActive;
            const viewStartN = normalizeDate(document.getElementById('startDate').value);
            const viewEndN = normalizeDate(document.getElementById('endDate').value);
            const tot = filterActive ? getHoursForActivityFiltered(sel, vis) : getHoursForActivity(sel);
            const span = filterActive ? getHoursForActivityFiltered(sel, vis, viewStartN, viewEndN) : getHoursForActivity(sel, viewStartN, viewEndN);
            const maxHours = sel.maxTotalHours || sel.projectInfo?.maxHours || null;
            const overrun = maxHours ? (tot - maxHours) : 0;

            // Compute weighted percent complete for the project within the current view span
            let pct = sel.completed || 0;
            try {
                const kids = (sel.children || []).map(id => findActivityById(id)).filter(Boolean);
                let hs = 0, wc = 0;
                kids.forEach(k => {
                    const spanH = filterActive ? getHoursForActivityFiltered(k, vis) : getHoursForActivity(k, viewStart, viewEnd);
                    if (spanH <= 0) return;
                    const comp = (k.isGroup ? (k.visibleCompleted ?? k.completed) : (k.completed || 0)) || 0;
                    hs += spanH; wc += comp * spanH;
                });
                if (hs > 0) pct = Math.round(wc / hs);
            } catch (_) { pct = sel.completed || 0; }

            // Render info summary and stat cards for the selected project
            detailsEl.innerHTML = `
                <div class="project-detail-shell">
                    <section class="project-hero-card">
                        <div id="projInfoSummary" class="project-info-card"></div>
                        <div class="project-stat-grid">
                            <div class="project-stat-card">
                                <span class="project-stat-label">Visningsspan</span>
                                <strong class="project-stat-value">${span.toFixed(1)}h</strong>
                            </div>
                            <div class="project-stat-card">
                                <span class="project-stat-label">Planerade timmar</span>
                                <strong class="project-stat-value">${tot.toFixed(1)}h</strong>
                            </div>
                            <div class="project-stat-card">
                                <span class="project-stat-label">Max timmar</span>
                                <strong class="project-stat-value">${maxHours ? maxHours.toFixed(1) + 'h' : '-'}</strong>
                            </div>
                            <div class="project-stat-card">
                                <span class="project-stat-label">Färdigt</span>
                                <strong class="project-stat-value">${pct}%</strong>
                            </div>
                            <div class="project-stat-card ${overrun > 0 ? 'is-warn' : (overrun < 0 ? 'is-good' : '')}">
                                <span class="project-stat-label">Under/överskott</span>
                                <strong class="project-stat-value">${overrun > 0 ? '+' : ''}${overrun.toFixed(1)}h</strong>
                            </div>
                        </div>
                    </section>
                </div>`;

            // Fill info summary
            const infoSummary = document.getElementById('projInfoSummary');
            if (infoSummary) {
                const info = sel.projectInfo || {};
                const infoRows = [
                    info.planner ? `<div class="project-info-row"><span>Planerare</span><strong>${info.planner}</strong></div>` : '',
                    info.company ? `<div class="project-info-row"><span>Företag</span><strong>${info.company}</strong></div>` : '',
                    info.contact ? `<div class="project-info-row"><span>Kontakt</span><strong>${info.contact}</strong></div>` : ''
                ].filter(Boolean).join('');
                infoSummary.innerHTML = `
                    <div class="project-info-eyebrow">Valt projekt</div>
                    <div class="project-info-title">${sel.name}</div>
                    <div class="project-info-grid">${infoRows || '<div class="project-info-empty">Ingen projektinfo är sparad ännu.</div>'}</div>
                    ${info.comment ? `<div class="project-info-note">${info.comment}</div>` : ''}
                `;

                // Always show an editable vertical palette under the other info.
                try {
                    // Build or derive the palette data on the project object so edits persist.
                    if (!sel.colorPalette || !Array.isArray(sel.colorPalette) || sel.colorPalette.length === 0) {
                        // Derive from activities in the project
                        const projActs = activities.filter(a => { const p = findProjectForActivity(a); return p && p.id === sel.id; });
                        const map = new Map();
                        projActs.forEach(a => { if (a.color) map.set(a.color, a.colorName || `Färg (${a.color})`); });
                        sel.colorPalette = Array.from(map.entries()).map(([color, name]) => ({ color, name }));
                    }

                    // Create the palette container
                    const palWrapper = document.createElement('section');
                    palWrapper.className = 'project-card project-palette-card';
                    const palTitle = document.createElement('h4');
                    palTitle.className = 'project-section-title';
                    palTitle.textContent = 'Palett';
                    palWrapper.appendChild(palTitle);

                    const list = document.createElement('div');
                    list.className = 'project-palette-list';

                    // Helper to get activities belonging to this project
                    const projActs = () => activities.filter(a => { const p = findProjectForActivity(a); return p && p.id === sel.id; });

                    sel.colorPalette.forEach((entry, idx) => {
                        const row = document.createElement('div');
                        row.className = 'project-palette-row';

                        const colorInput = document.createElement('input');
                        colorInput.type = 'color';
                        colorInput.className = 'project-palette-color';
                        colorInput.value = entry.color || '#007bff';
                        colorInput.title = 'Ändra färg';
                        colorInput.onchange = e => {
                            const old = entry.color;
                            entry.color = e.target.value;
                            // Update activities in this project that used the old color
                            projActs().forEach(a => { if (a.color === old) { a.color = entry.color; a.colorName = entry.name; } });
                            pushHistory(); if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render();
                        };

                        const nameInput = document.createElement('input');
                        nameInput.type = 'text';
                        nameInput.className = 'project-palette-name';
                        nameInput.value = entry.name || entry.color || '';
                        nameInput.title = 'Ändra namn på färgen';
                        nameInput.onchange = e => {
                            entry.name = e.target.value;
                            // Update activity colorName for matching colors in this project
                            projActs().forEach(a => { if (a.color === entry.color) a.colorName = entry.name; });
                            pushHistory(); if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render();
                        };

                        const delBtn = document.createElement('button');
                        delBtn.className = 'project-palette-delete';
                        delBtn.innerHTML = '<span class="material-icons-outlined">close</span>';
                        delBtn.title = 'Ta bort färg från paletten';
                        delBtn.onclick = () => {
                            // Reassign activities with this color to default and remove from palette
                            projActs().forEach(a => { if (a.color === entry.color) { a.color = '#007bff'; delete a.colorName; } });
                            sel.colorPalette.splice(idx, 1);
                            pushHistory(); if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render();
                        };

                        row.appendChild(colorInput);
                        row.appendChild(nameInput);
                        row.appendChild(delBtn);

                        list.appendChild(row);
                    });

                    palWrapper.appendChild(list);
                    detailsEl.appendChild(palWrapper);
                } catch (_) {}
            }

            // Summor: timmar per aktivitetsnamn (för valt projekt)
            try {
                const vs = viewStart;
                const ve = viewEnd;
                const acts = [];
                (function walk(node){
                    if (!node) return;
                    if (node.isGroup) {
                        (node.children||[]).forEach(id => walk(findActivityById(id)));
                    } else {
                        acts.push(node);
                    }
                })(sel);

                const byName = new Map();
                const filterActive = !!window.__filterActive;
                const vis = (window.__visibleIdSet instanceof Set) ? window.__visibleIdSet : new Set();
                acts.forEach(a => {
                    const key = (a.name || 'Namnlös').trim();
                    if (filterActive && !vis.has(a.id)) return; // hoppa över icke-synliga aktiviter i sökfilter
                    const spanH = getHoursForActivity(a, vs, ve);
                    const totH = getHoursForActivity(a);
                    const curr = byName.get(key) || { span: 0, total: 0 };
                    curr.span += spanH;
                    curr.total += totH;
                    byName.set(key, curr);
                });

                if (byName.size > 0) {
                    const rows = Array.from(byName.entries()).sort((a,b)=> b[1].span - a[1].span);
                    const box = document.createElement('section');
                    box.className = 'project-card project-hours-card';
                    const title = document.createElement('h4');
                    title.className = 'project-section-title';
                    title.textContent = 'Timmar per aktivitetsnamn';
                    box.appendChild(title);
                    const grid = document.createElement('div');
                    grid.className = 'project-hours-table';
                    const hN = document.createElement('div'); hN.className = 'project-hours-head'; hN.innerHTML = '<strong>Namn</strong>';
                    const hS = document.createElement('div'); hS.className = 'project-hours-head align-right'; hS.innerHTML = '<strong>Span</strong>';
                    const hT = document.createElement('div'); hT.className = 'project-hours-head align-right'; hT.innerHTML = '<strong>Totalt</strong>';
                    grid.appendChild(hN); grid.appendChild(hS); grid.appendChild(hT);
                    rows.forEach(([name, v]) => {
                        const n = document.createElement('div'); n.className = 'project-hours-name'; n.textContent = name;
                        const s = document.createElement('div'); s.className = 'project-hours-value align-right'; s.textContent = `${v.span.toFixed(1)} h`;
                        const t = document.createElement('div'); t.className = 'project-hours-value align-right'; t.textContent = `${v.total.toFixed(1)} h`;
                        grid.appendChild(n); grid.appendChild(s); grid.appendChild(t);
                    });
                    box.appendChild(grid);
                    detailsEl.appendChild(box);
                }
            } catch(_) {}
        }
    } catch(_) {}

        // Project list cards removed — selector + selected-project details are shown instead.
}
