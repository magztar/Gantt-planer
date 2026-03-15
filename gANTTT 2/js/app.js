// js/app.js

/**
 * Detta är applikationens "lim". Den binder samman alla delar genom att
 * sätta upp händelselyssnare och hantera användarinteraktioner.
 */

// === INITIALISERING ===

function isElementVisible(element) {
    return !!element && element.style.display !== 'none';
}

const SEARCH_RENDER_DEBOUNCE_MS = 120;
let searchRenderTimer = null;
let queuedRenderRequest = null;
let queuedRenderResolvers = [];
let renderFlushScheduled = false;
let activeQueuedRenderPromise = null;

function enqueueRenderRequest(options = {}) {
    const preserveScroll = options.preserveScroll !== false;
    if (!queuedRenderRequest) queuedRenderRequest = { preserveScroll: false };
    queuedRenderRequest.preserveScroll = queuedRenderRequest.preserveScroll || preserveScroll;

    return new Promise((resolve, reject) => {
        queuedRenderResolvers.push({ resolve, reject });
        if (!renderFlushScheduled) {
            renderFlushScheduled = true;
            requestAnimationFrame(flushRenderQueue);
        }
    });
}

async function flushRenderQueue() {
    if (activeQueuedRenderPromise) {
        renderFlushScheduled = false;
        return;
    }

    const request = queuedRenderRequest;
    const resolvers = queuedRenderResolvers;
    queuedRenderRequest = null;
    queuedRenderResolvers = [];
    renderFlushScheduled = false;

    if (!request) {
        resolvers.forEach(({ resolve }) => resolve());
        return;
    }

    activeQueuedRenderPromise = (async () => {
        if (request.preserveScroll) await performRenderAndPreserveScroll();
        else if (typeof render === 'function') await render();
    })();

    try {
        await activeQueuedRenderPromise;
        resolvers.forEach(({ resolve }) => resolve());
    } catch (error) {
        resolvers.forEach(({ reject }) => reject(error));
        console.error('Render queue failed:', error);
    } finally {
        activeQueuedRenderPromise = null;
        if (queuedRenderRequest && !renderFlushScheduled) {
            renderFlushScheduled = true;
            requestAnimationFrame(flushRenderQueue);
        }
    }
}

function scheduleSearchRender() {
    if (searchRenderTimer) clearTimeout(searchRenderTimer);
    searchRenderTimer = setTimeout(() => {
        searchRenderTimer = null;
        renderAndPreserveScroll();
    }, SEARCH_RENDER_DEBOUNCE_MS);
}

function setFabState(button, options) {
    if (!button || !options) return;
    const icon = button.querySelector('.material-icons-outlined');
    const label = button.querySelector('.fab-label');
    const visible = !!options.visible;
    if (icon && options.icon) icon.textContent = options.icon;
    if (label && options.label) label.textContent = options.label;
    button.classList.toggle('active', visible);
    button.setAttribute('aria-pressed', visible ? 'true' : 'false');
    if (options.title) button.title = options.title;
}

function updatePanelToggleButtons() {
    const activitySidebar = document.getElementById('activitySidebar');
    const projectPanel = document.getElementById('projectPanel');
    const activityFab = document.getElementById('fabActivityPanel');
    const projectFab = document.getElementById('fabProjectPanel');

    const activityVisible = isElementVisible(activitySidebar);
    const projectVisible = isElementVisible(projectPanel);

    setFabState(activityFab, {
        visible: activityVisible,
        icon: activityVisible ? 'menu_open' : 'menu',
        label: activityVisible ? 'Dölj aktiviteter' : 'Visa aktiviteter',
        title: activityVisible ? 'Dölj aktivitetspanelen' : 'Visa aktivitetspanelen'
    });
    setFabState(projectFab, {
        visible: projectVisible,
        icon: projectVisible ? 'dashboard_customize' : 'dashboard',
        label: projectVisible ? 'Dölj projekt' : 'Visa projekt',
        title: projectVisible ? 'Dölj projektöversikten' : 'Visa projektöversikten'
    });
}

function setActivityPanelVisibility(visible) {
    const sidebar = document.getElementById('activitySidebar');
    const resizer = document.getElementById('sidebarResizer');
    if (!sidebar) return;

    sidebar.style.display = visible ? 'flex' : 'none';
    if (resizer) resizer.style.display = visible ? 'block' : 'none';
    localStorage.setItem('activityPanelVisible', visible ? 'true' : 'false');
    updatePanelToggleButtons();
}

function setProjectPanelVisibility(visible) {
    const projectPanel = document.getElementById('projectPanel');
    if (!projectPanel) return;

    projectPanel.style.display = visible ? 'flex' : 'none';
    localStorage.setItem('projectPanelVisible', visible ? 'true' : 'false');
    updatePanelToggleButtons();
}

function setupUnifiedScrollController() {
    const ganttContainerEl = document.querySelector('.gantt-container');
    const activityListEl = document.querySelector('#activityList');
    if (!ganttContainerEl || !activityListEl) return;

    if (window.__ganttUnifiedScrollBound) return;
    window.__ganttUnifiedScrollBound = true;

    let syncLock = false;
    const updateVirtualViewport = (scrollTop) => {
        if (typeof window.__updateActivitySidebarViewport === 'function') {
            window.__updateActivitySidebarViewport(scrollTop);
        }
    };
    const syncVerticalScroll = (source, target) => {
        if (!source || !target || syncLock) return;
        const nextTop = source.scrollTop;
        if (Math.abs(target.scrollTop - nextTop) < 0.5) return;
        syncLock = true;
        target.scrollTop = nextTop;
        updateVirtualViewport(nextTop);
        requestAnimationFrame(() => { syncLock = false; });
    };

    ganttContainerEl.addEventListener('scroll', () => {
        if (window.__suppressSyncDuringRestore) return;
        updateVirtualViewport(ganttContainerEl.scrollTop);
        syncVerticalScroll(ganttContainerEl, activityListEl);
    }, { passive: true });

    activityListEl.addEventListener('scroll', () => {
        if (window.__suppressSyncDuringRestore) return;
        updateVirtualViewport(activityListEl.scrollTop);
        syncVerticalScroll(activityListEl, ganttContainerEl);
    }, { passive: true });

    activityListEl.addEventListener('wheel', (e) => {
        const deltaY = e.deltaY || 0;
        const deltaX = e.deltaX || 0;
        const horizontalFromShift = deltaX !== 0 ? deltaX : ((e.shiftKey && deltaY !== 0) ? deltaY : 0);
        if (deltaY !== 0) ganttContainerEl.scrollTop += deltaY;
        if (horizontalFromShift !== 0) ganttContainerEl.scrollLeft += horizontalFromShift;
        e.preventDefault();
    }, { passive: false });
}

window.addEventListener('load', () => {
    const today = new Date();
    const startInput = document.getElementById('startDate');
    const endInput = document.getElementById('endDate');
    // Återställ datumintervall från localStorage om det finns
    const savedStart = localStorage.getItem('startDate');
    const savedEnd = localStorage.getItem('endDate');
    if (savedStart) {
        startInput.value = savedStart;
    } else if (!startInput.value) {
        startInput.value = formatDateLocal(today);
    }
    if (savedEnd) {
        endInput.value = savedEnd;
    } else if (!endInput.value) {
        endInput.value = formatDateLocal(new Date(today.getFullYear(), today.getMonth() + 2, 0));
    }

    const savedTheme = localStorage.getItem('theme') || 'light';
    document.documentElement.setAttribute('data-theme', savedTheme);

    // Återställ panelsynlighet från localStorage
    const projectPanelVisible = localStorage.getItem('projectPanelVisible') !== 'false';
    const activityPanelVisible = localStorage.getItem('activityPanelVisible') !== 'false';
    setProjectPanelVisibility(projectPanelVisible);
    setActivityPanelVisibility(activityPanelVisible);

    setupEventListeners();
    setupUnifiedScrollController();
    setupPopupInteractions();
    closePopup();
    // initializeAI(); // Borttagen

    renderAndPreserveScroll();
});
// Panel visibility toggle for activity panel (FAB)
function toggleActivityPanel() {
    const sidebar = document.getElementById('activitySidebar');
    if (!sidebar) return;
    setActivityPanelVisibility(!isElementVisible(sidebar));
    if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll();
}

// === Interaktions-Tooltip (visar delta och längd vid drag/resize) ===
let interactionTooltipEl = null;
function ensureInteractionTooltip() {
    if (!interactionTooltipEl) {
        interactionTooltipEl = document.createElement('div');
        interactionTooltipEl.id = 'interactionTooltip';
        interactionTooltipEl.className = 'interaction-tooltip';
        interactionTooltipEl.style.position = 'fixed';
        interactionTooltipEl.style.zIndex = '2000';
        interactionTooltipEl.style.pointerEvents = 'none';
        interactionTooltipEl.style.display = 'none';
        document.body.appendChild(interactionTooltipEl);
    }
}
function showInteractionTooltip(text, x, y) {
    ensureInteractionTooltip();
    interactionTooltipEl.textContent = text;
    interactionTooltipEl.style.left = `${x + 12}px`;
    interactionTooltipEl.style.top = `${y + 12}px`;
    interactionTooltipEl.style.display = 'block';
}
function updateInteractionTooltip(text, x, y) {
    if (!interactionTooltipEl) return;
    interactionTooltipEl.textContent = text;
    interactionTooltipEl.style.left = `${x + 12}px`;
    interactionTooltipEl.style.top = `${y + 12}px`;
}
function hideInteractionTooltip() {
    if (interactionTooltipEl) interactionTooltipEl.style.display = 'none';
}

// === MARQUEE (rektangel) URVAL ===
let marqueeEl = null;
function ensureMarqueeEl() {
    if (!marqueeEl) {
        marqueeEl = document.createElement('div');
        marqueeEl.id = 'marqueeSelection';
        marqueeEl.style.position = 'fixed';
        marqueeEl.style.background = 'rgba(30,144,255,0.15)';
        marqueeEl.style.border = '1px solid var(--md-sys-color-primary)';
        marqueeEl.style.pointerEvents = 'none';
        marqueeEl.style.zIndex = '1500';
        document.body.appendChild(marqueeEl);
    }
}
function updateMarqueeEl() {
    if (!marquee.active || !marqueeEl) return;
    const x1 = Math.min(marquee.startX, marquee.currentX);
    const y1 = Math.min(marquee.startY, marquee.currentY);
    const x2 = Math.max(marquee.startX, marquee.currentX);
    const y2 = Math.max(marquee.startY, marquee.currentY);
    marqueeEl.style.left = x1 + 'px';
    marqueeEl.style.top = y1 + 'px';
    marqueeEl.style.width = (x2 - x1) + 'px';
    marqueeEl.style.height = (y2 - y1) + 'px';
}
function removeMarqueeEl() {
    if (marqueeEl) { marqueeEl.remove(); marqueeEl = null; }
}
function applyMarqueeSelection() {
    const x1 = Math.min(marquee.startX, marquee.currentX);
    const y1 = Math.min(marquee.startY, marquee.currentY);
    const x2 = Math.max(marquee.startX, marquee.currentX);
    const y2 = Math.max(marquee.startY, marquee.currentY);
    const newlySelectedActs = new Set();
    const newlySelectedSegs = new Set();
    // Staplar
    document.querySelectorAll('.bar').forEach(bar => {
        const r = bar.getBoundingClientRect();
        if (r.right < x1 || r.left > x2 || r.bottom < y1 || r.top > y2) return;
        const aId = bar.dataset.id;
        const segIdx = bar.dataset.segmentIndex;
        if (segIdx != null) {
            newlySelectedSegs.add(`${aId}:${segIdx}`);
        } else if (aId) {
            newlySelectedActs.add(aId);
        }
    });
    // Rader i sidopanel
    document.querySelectorAll('.activity-item').forEach(row => {
        const r = row.getBoundingClientRect();
        if (r.right < x1 || r.left > x2 || r.bottom < y1 || r.top > y2) return;
        const id = row.dataset.id;
        if (id) newlySelectedActs.add(id);
    });
    // Återställ baserat på additive eller snapshot
    if (!marquee.additive) { selectedActivityIds.clear(); selectedSegmentKeys.clear(); }
    else {
        // Börja från snapshot för additive
        selectedActivityIds.clear(); marquee.snapshotActivities.forEach(v => selectedActivityIds.add(v));
        selectedSegmentKeys.clear(); marquee.snapshotSegments.forEach(v => selectedSegmentKeys.add(v));
    }
    newlySelectedActs.forEach(id => selectedActivityIds.add(id));
    newlySelectedSegs.forEach(key => selectedSegmentKeys.add(key));
    highlightSelection();
}

// Driv logiken för rektangelmarkering under musrörelse/uppsläpp
function handleMouseMove(e) {
    if (!marquee.active) return;
    marquee.currentX = e.clientX;
    marquee.currentY = e.clientY;
    updateMarqueeEl();
}

function handleMouseUp(e) {
    if (!marquee.active) return;
    applyMarqueeSelection();
    marquee.active = false;
    removeMarqueeEl();
}

function renderAndPreserveScroll() {
    if (searchRenderTimer) {
        clearTimeout(searchRenderTimer);
        searchRenderTimer = null;
    }
    return enqueueRenderRequest({ preserveScroll: true });
}

async function performRenderAndPreserveScroll() {
    const ganttContainer = document.querySelector('.gantt-container');
    const activityListEl = document.querySelector('#activityList');
    const scrollLeft = ganttContainer ? ganttContainer.scrollLeft : 0;
    const scrollTop = ganttContainer ? ganttContainer.scrollTop : 0;
    const activityScrollTop = activityListEl ? activityListEl.scrollTop : 0;
    const fallbackTop = ganttContainer ? scrollTop : activityScrollTop;

    // Capture an anchor row in the activity list (top-most visible) to restore by element, not only pixels
    let listAnchorId = null;
    let listAnchorOffset = 0;
    const getActivityRowOffsetTop = (activityId) => {
        if (!activityId) return null;
        const row = activityListEl?.querySelector(`.activity-item[data-id="${activityId}"]`);
        if (row) return row.offsetTop;
        if (typeof window.__getActivityRowOffsetTop === 'function') return window.__getActivityRowOffsetTop(activityId);
        return null;
    };
    if (activityListEl) {
        try {
            const curTop = activityListEl.scrollTop;
            // Priority 1: explicit anchor hint (e.g., toggled group)
            const forcedId = window.__forceListAnchorId || null;
            if (forcedId) {
                const forcedRowTop = getActivityRowOffsetTop(forcedId);
                if (forcedRowTop != null) {
                    listAnchorId = forcedId;
                    // If caller provided a precise precomputed offset, prefer it.
                    const forcedOffset = (typeof window.__forceListAnchorOffset === 'number') ? window.__forceListAnchorOffset : null;
                    listAnchorOffset = Math.round(forcedOffset != null ? forcedOffset : (curTop - forcedRowTop));
                }
                try { delete window.__forceListAnchorId; } catch(_) {}
                try { delete window.__forceListAnchorOffset; } catch(_) {}
            }
            // Priority 2: compute top-most visible row if no forced anchor
            if (!listAnchorId) {
                let best = null; let bestDelta = Infinity;
                activityListEl.querySelectorAll('.activity-item').forEach(row => {
                    const off = row.offsetTop;
                    const delta = curTop - off;
                    if (delta >= 0 && delta < bestDelta) { bestDelta = delta; best = row; }
                });
                if (!best) best = activityListEl.querySelector('.activity-item');
                if (best && best.dataset && best.dataset.id) {
                    listAnchorId = best.dataset.id;
                    listAnchorOffset = Math.round(curTop - best.offsetTop);
                }
            }
        } catch (_) {}
    }

    // Suppress "user scroll" detection caused by DOM changes during render
    window.__suppressScrollRestoreActive = true;

    if (typeof render === 'function') { await render(); }

    const now = Date.now();
    const userScrolling = (typeof window.__lastUserScrollAt === 'number' && (now - window.__lastUserScrollAt) < (window.__USER_SCROLL_DEBOUNCE_MS || 250));
    const recentInteraction = (typeof window.__lastUserInteractionAt === 'number' && (now - window.__lastUserInteractionAt) < (window.__USER_INTERACTION_DEBOUNCE_MS || 500));
    const pointerActive = !!window.__userPointerDown;
    const suppressForContext = !!window.__suppressScrollRestoreForContextMenu;
    const shouldRestore = !!ganttContainer && (listAnchorId || (!userScrolling && !pointerActive && !recentInteraction && !suppressForContext));

    const applySyncedScrollTop = (top) => {
        if (ganttContainer) ganttContainer.scrollTop = top;
        if (activityListEl) activityListEl.scrollTop = top;
        if (typeof window.__updateActivitySidebarViewport === 'function') window.__updateActivitySidebarViewport(top);
    };

    if (shouldRestore) {
        window.__suppressSyncDuringRestore = true;

        let targetTop = fallbackTop;
        if (listAnchorId && activityListEl) {
            const newRowTop = getActivityRowOffsetTop(listAnchorId);
            if (newRowTop != null) targetTop = Math.max(0, newRowTop + listAnchorOffset);
        }

        ganttContainer.scrollLeft = scrollLeft;
        applySyncedScrollTop(targetTop);

        await new Promise(resolve => {
            requestAnimationFrame(() => {
                let settledTop = targetTop;
                if (listAnchorId && activityListEl) {
                    const settledRowTop = getActivityRowOffsetTop(listAnchorId);
                    if (settledRowTop != null) settledTop = Math.max(0, settledRowTop + listAnchorOffset);
                }
                ganttContainer.scrollLeft = scrollLeft;
                applySyncedScrollTop(settledTop);
                window.__suppressSyncDuringRestore = false;
                resolve();
            });
        });
    } else {
        if (activityListEl && Math.abs(activityListEl.scrollTop - fallbackTop) > 1) {
            activityListEl.scrollTop = fallbackTop;
        }
        window.__suppressSyncDuringRestore = false;
    }
    
    // Sätt upp resizer efter render (ifall DOM har förändrats)
    setupSidebarResizer();
    // Återställ sidopanelens bredd från localStorage (om nödvändigt)
    const sidebar = document.getElementById('activitySidebar');
    const savedSidebarWidth = localStorage.getItem('activitySidebarWidth');
    if (sidebar && savedSidebarWidth) {
        sidebar.style.width = savedSidebarWidth;
    }

    // Re-enable user scroll detection after restoration completes
    window.__suppressScrollRestoreActive = false;
}

// === HÄNDELSELYSSNARE ===
function setupEventListeners() {
    // Spara datumintervall till localStorage vid ändring
    ['startDate', 'endDate'].forEach(id => {
        document.getElementById(id).addEventListener('input', (e) => {
            localStorage.setItem(id, e.target.value);
            renderAndPreserveScroll();
        });
    });
    // Övriga input render
    ['showHolidays', 'dateFormat', 'showManpower', 'countWeekends'].forEach(id => {
        document.getElementById(id).addEventListener('input', () => renderAndPreserveScroll());
    });
    const searchInput = document.getElementById('searchInput');
    if (searchInput) searchInput.addEventListener('input', scheduleSearchRender);
    const overviewEl = document.getElementById('overviewToggle');
    if (overviewEl) {
        overviewEl.addEventListener('change', (e) => {
            overviewMode = !!e.target.checked;
            if (overviewMode) {
                // Memorize current zoom; switch to ultra-zoom for coarse view
                prevZoomLevel = zoomLevel;
                zoomLevel = 0;
            } else if (prevZoomLevel != null) {
                zoomLevel = prevZoomLevel;
                prevZoomLevel = null;
            }
            renderAndPreserveScroll();
        });
    }
    document.getElementById('themeToggle').addEventListener('click', () => {
        const newTheme = document.documentElement.getAttribute('data-theme') === 'light' ? 'dark' : 'light';
        document.documentElement.setAttribute('data-theme', newTheme);
        localStorage.setItem('theme', newTheme);
    });

    // document.getElementById('toggleProjectPanel').addEventListener('click', toggleProjectPanel); // Flyttat till FAB

    document.getElementById('openFileInput').addEventListener('change', (e) => {
        if (e.target.files.length > 0) loadFromFile(e.target.files[0]);
    });
    
    document.getElementById('excelFileInput').addEventListener('change', (e) => {
        if (e.target.files.length > 0) handleExcelImport(e.target.files[0]);
    });
    
    document.addEventListener('keydown', handleGlobalKeyDown);
    document.addEventListener('click', handleGlobalClick);
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);

    // Track user scrolls to avoid render routines clobbering the user's scroll position
    window.__lastUserScrollAt = 0;
    window.__USER_SCROLL_DEBOUNCE_MS = 250;
    // Track broader user interactions (pointer down/up, context menu) so we don't restore
    // scroll while the user is interacting with the UI (dragging, right-click menus etc.)
    window.__userPointerDown = false;
    window.__lastUserInteractionAt = 0;
    window.__USER_INTERACTION_DEBOUNCE_MS = 500; // ms to treat recent interaction as active
    window.__suppressScrollRestoreForContextMenu = false;
    const ganttContainerEl = document.querySelector('.gantt-container');
    const activityListEl = document.querySelector('#activityList');
    const onUserScroll = () => { if (!window.__suppressScrollRestoreActive) { window.__lastUserScrollAt = Date.now(); } };
    ganttContainerEl?.addEventListener('scroll', onUserScroll, { passive: true });
    activityListEl?.addEventListener('scroll', onUserScroll, { passive: true });

    // Pointer events: mark when the user is pressing/holding the pointer (dragging etc.)
    document.addEventListener('pointerdown', (e) => {
        window.__userPointerDown = true;
        window.__lastUserInteractionAt = Date.now();
    }, { capture: true });
    document.addEventListener('pointerup', (e) => {
        window.__userPointerDown = false;
        window.__lastUserInteractionAt = Date.now();
    }, { capture: true });

    // Context menu (right-click): temporarily suppress scroll restore so the view doesn't snap
    document.addEventListener('contextmenu', (e) => {
        window.__lastUserInteractionAt = Date.now();
        window.__suppressScrollRestoreForContextMenu = true;
        // keep suppression for a short window to allow menu interaction without snapping
        setTimeout(() => { window.__suppressScrollRestoreForContextMenu = false; }, 600);
    }, { capture: true });

    // Marquee selection start
    const main = document.querySelector('.main-content');
    main?.addEventListener('mousedown', (e) => {
        // Starta inte om vi drar en stapel/resize eller klickar på inputs
        if (e.button !== 0) return;
        // Undvik marquee när vi startar drag på stapel, resize-handle, sidopanelens resizer eller via formulärkontroller
        if (e.target.closest('.bar') || e.target.closest('.resize-handle') || e.target.closest('#sidebarResizer') || e.target.closest('.sidebar-resizer') || e.target.closest('input, textarea, select, button') || e.target.closest('.drag-handle')) return;
        // Starta inte marquee om vi är vid sidopanelens resize-kant
        const sidebar = document.querySelector('.activity-sidebar');
        if (sidebar && sidebar.contains(e.target)) {
            const r = sidebar.getBoundingClientRect();
            const edge = 10; // px nära högerkanten för resize
            if (e.clientX > r.right - edge) return;
        }
        marquee.active = true;
        marquee.startX = e.clientX;
        marquee.startY = e.clientY;
        marquee.currentX = e.clientX;
        marquee.currentY = e.clientY;
        marquee.additive = e.ctrlKey || e.metaKey;
        marquee.snapshotActivities = new Set(selectedActivityIds);
        marquee.snapshotSegments = new Set(selectedSegmentKeys);
        ensureMarqueeEl();
        updateMarqueeEl();
        e.preventDefault();
    });

    // Zoom-knappar
    document.getElementById('zoomInBtn')?.addEventListener('click', () => changeZoom(1));
    document.getElementById('zoomOutBtn')?.addEventListener('click', () => changeZoom(-1));

    // Ctrl + Scroll för zoom
    const ganttContainer = document.querySelector('.gantt-container');
    ganttContainer?.addEventListener('wheel', (e) => {
        if (!e.ctrlKey) return; // endast med Ctrl
        e.preventDefault();
        const delta = e.deltaY > 0 ? -1 : 1;
        changeZoom(delta, e.clientX);
    }, { passive: false });
    
    document.getElementById('savePopupBtn').addEventListener('click', savePopup);
    document.getElementById('deleteActivityBtn').addEventListener('click', deleteFromPopup);
    document.getElementById('splitActivityBtn').addEventListener('click', splitFromPopup);

    setupSidebarResizer();

    // Bulk edit handlers
    const bulkApply = document.getElementById('bulkApplyBtn');
    const bulkClose = document.getElementById('bulkCloseBtn');
    const bulkDropdown = document.getElementById('bulkColorDropdown');
    const bulkColor = document.getElementById('bulkColorInput');
    if (bulkApply) bulkApply.addEventListener('click', applyBulkEdit);
    if (bulkClose) bulkClose.addEventListener('click', () => {
        const bar = document.getElementById('bulkEditBar');
        if (bar) bar.style.display = 'none';
    });
    if (bulkDropdown) bulkDropdown.addEventListener('change', () => {
        const v = bulkDropdown.value;
        if (v && bulkColor) bulkColor.value = v;
    });
}

// Deduped window listeners flag for resizer
let sidebarWindowListenersInstalled = false;

// === SIDEBAR RESIZER SETUP ===
function setupSidebarResizer() {
    if (window.developerMode) console.log('🔧 setupSidebarResizer() called');
    const resizer = document.getElementById('sidebarResizer');
    const sidebar = document.getElementById('activitySidebar');
    if (window.developerMode) console.log('🔧 Elements found:', { resizer: !!resizer, sidebar: !!sidebar });
    if (!resizer || !sidebar) {
    if (window.developerMode) console.warn('⚠️ Resizer or sidebar not found, skipping setup');
        return;
    }
    
    if (window.developerMode) console.log('🔧 Setting up event listeners for resizer...');
    
    if (resizer._pointerDownHandler) {
        resizer.removeEventListener('pointerdown', resizer._pointerDownHandler);
    }
    if (resizer._keyDownHandler) {
        resizer.removeEventListener('keydown', resizer._keyDownHandler);
    }
    
    const getMinSidebarWidth = () => {
        // Allow a narrower activity sidebar
        return 260;
    };

    const savedPx = localStorage.getItem('activitySidebarWidth');
    if (savedPx) {
        sidebar.style.width = savedPx;
    } else {
        // Set a narrower default width when none saved yet
        sidebar.style.width = '280px';
    }

    let dragging = false;
    let startX = 0;
    let startWidth = 0;
    let activePointerId = null;

    const onPointerMove = (ev) => {
        if (!dragging) return;
    if (window.developerMode) console.debug('[resizer] pointermove', { x: ev.clientX, y: ev.clientY });
    const dx = ev.clientX - startX;
    const minWidth = getMinSidebarWidth();
    const newWidth = Math.max(minWidth, startWidth + dx);
    if (window.developerMode) console.log('[resizer] Setting sidebar width:', newWidth);
    sidebar.style.width = newWidth + 'px';
    if (window.developerMode) console.log('[resizer] sidebar.style.width is now:', sidebar.style.width);
    showInteractionTooltip(`Bredd: ${newWidth}px`, ev.clientX, ev.clientY);
    };

    const finishResize = (ev) => {
        if (!dragging) return;
    if (window.developerMode) console.debug('[resizer] finishResize', { x: ev?.clientX, y: ev?.clientY, pointerId: ev?.pointerId });
        dragging = false;
        document.body.classList.remove('resizing-sidebar');
        try { if (activePointerId != null) resizer.releasePointerCapture(activePointerId); } catch(_) {}
        window.removeEventListener('pointermove', onPointerMove);
        window.removeEventListener('pointerup', finishResize);
        window.removeEventListener('pointercancel', finishResize);
        
        hideInteractionTooltip();
        
        localStorage.setItem('activitySidebarWidth', sidebar.style.width);
        activePointerId = null;
    };

    const pointerDownHandler = (e) => {
    if (window.developerMode) console.log('🖱️ Resizer pointerdown triggered!', e);
        e.preventDefault();
        e.stopPropagation();
    if (window.developerMode) console.debug('[resizer] pointerdown', { x: e.clientX, y: e.clientY, pointerId: e.pointerId });
        dragging = true;
        startX = e.clientX;
        startWidth = sidebar.getBoundingClientRect().width;
        activePointerId = e.pointerId;
        try { resizer.setPointerCapture(activePointerId); } catch(_) {}
        window.addEventListener('pointermove', onPointerMove, { passive: false });
        window.addEventListener('pointerup', finishResize);
        window.addEventListener('pointercancel', finishResize);
    };
    
    resizer._pointerDownHandler = pointerDownHandler;
    resizer.addEventListener('pointerdown', pointerDownHandler, { capture: true });

    resizer.setAttribute('tabindex', '0');
    const keyDownHandler = (e) => {
        const curW = sidebar.getBoundingClientRect().width;
        const minWidth = getMinSidebarWidth();
        
        if (e.key === 'ArrowLeft') { 
            const newWidth = Math.max(minWidth, curW - 20);
            sidebar.style.width = newWidth + 'px'; 
            localStorage.setItem('activitySidebarWidth', sidebar.style.width); 
        }
        if (e.key === 'ArrowRight') { 
            const newWidth = curW + 20;
            sidebar.style.width = newWidth + 'px'; 
            localStorage.setItem('activitySidebarWidth', sidebar.style.width); 
        }
        
        e.preventDefault();
    };
    
    resizer._keyDownHandler = keyDownHandler;
    resizer.addEventListener('keydown', keyDownHandler);
    
    if (window.developerMode) console.log('✅ Resizer setup complete!');
}

// Ändra zoomnivå (-1, 0, +1). Försök bevara fokus runt muspekarens datum.
async function changeZoom(step, focusClientX = null) {
    const oldZoom = zoomLevel;
    const maxZoom = (Array.isArray(zoomDayWidths) ? zoomDayWidths.length - 1 : 2);
    zoomLevel = Math.max(0, Math.min(maxZoom, zoomLevel + step));
    if (zoomLevel === oldZoom) return;

    const container = document.querySelector('.gantt-container');
    const grid = document.querySelector('.gantt-grid');
    if (!container || !grid) { 
        if (typeof render === 'function') await render(); 
        return; 
    }

    // Capture vertical scroll (scrollTop) to restore it exactly.
    // We will recalculate scrollLeft for the zoom focus.
    const scrollTop = container.scrollTop;

    // Beräkna fokusdatum före zoom
    const rect = grid.getBoundingClientRect();
    const viewStart = normalizeDate(document.getElementById('startDate').value);
    const oldDayWidth = interaction.dayWidth || (zoomDayWidths?.[oldZoom] ?? 35);
    let focusDayOffset = null;
    if (focusClientX != null) {
        const xInGrid = Math.max(0, focusClientX - rect.left + container.scrollLeft);
        focusDayOffset = Math.round(xInGrid / oldDayWidth);
    } else {
        // centrera kring mitten
        focusDayOffset = Math.round((container.scrollLeft + container.clientWidth / 2) / oldDayWidth);
    }
    
    // Använd direkt render med await, istället för renderAndPreserveScroll
    // som skulle försöka återställa gammal scrollLeft.
    if (typeof render === 'function') await render();

    // Efter render: scrolla så att samma dag hamnar under pekaren/mitt
    const newDayWidth = interaction.dayWidth || (zoomDayWidths?.[zoomLevel] ?? 35);
    const targetScrollLeft = Math.max(0, focusDayOffset * newDayWidth - (focusClientX ? (focusClientX - rect.left) : (container.clientWidth / 2)));
    
    // Återställ både vertikal och beräknad horisontell scroll
    if (container) {
        container.scrollTop = scrollTop;
        container.scrollLeft = targetScrollLeft;
    }
    
    // Säkerställ att resizer fungerar efter zoom-ändring
    setTimeout(() => {
    if (window.developerMode) console.log('🔧 Setting up resizer after zoom change');
        setupSidebarResizer();
    }, 0);
}

// === INTERAKTIONS-LOGIK ===

function handleGlobalKeyDown(e) {
    if (e.ctrlKey || e.metaKey) {
        if (e.key === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); }
        if (e.key === 'y') { e.preventDefault(); redo(); }
        if (e.key.toLowerCase() === 'c') { e.preventDefault(); try { copySelection(); } catch(err){ if(window.developerMode) console.error('[copySelection error]', err); } }
        if (e.key.toLowerCase() === 'v') { e.preventDefault(); try { pasteSelection(); } catch(err){ if(window.developerMode) console.error('[pasteSelection error]', err); } }
    }
    if (e.key === 'Escape') {
        closePopup();
        closeCenteredModal();
        clearSelection();
    }
    if (e.key === 'Delete') {
        try {
            if (selectedSegmentKeys.size > 0 || selectedActivityIds.size > 0) {
                deleteSelection();
            } else if (selectedActivityId) {
                deleteFromPopup();
            }
        } catch(err) { if(window.developerMode) console.error('[delete error]', err); }
    }
}

function handleGlobalClick(e) {
    const popup = document.getElementById('activityPopup');
    if (popup.style.display === 'block' && !popup.contains(e.target) && !e.target.closest('.bar')) {
        closePopup();
    }
    const modal = document.getElementById('centeredModal');
    if(modal.style.display === 'block' && e.target.id === 'modalBackdrop') {
        closeCenteredModal();
    }
}

// --- Dra-och-släpp i sidopanelen ---
function handleDragStart(e, id) {
    dragInfo.sourceId = id;
    e.dataTransfer.setData('text/plain', id);
    e.dataTransfer.effectAllowed = 'move';
    setTimeout(() => e.target.style.opacity = '0.5', 0);
    const src = findActivityById(id);
    if (src) {
        showInteractionTooltip(`Flyttar: ${src.name}`, e.clientX || 0, e.clientY || 0);
    }
}

function handleDragOver(e) {
    e.preventDefault();
    const targetItem = e.target.closest('.activity-item');
    if (!targetItem) return;
    
    const rect = targetItem.getBoundingClientRect();
    const offset = e.clientY - rect.top;
    
    document.querySelectorAll('.activity-item').forEach(item => {
        item.style.borderTop = item.style.borderBottom = '';
        item.classList.remove('drop-hover');
    });

    if (offset < rect.height / 3) {
        targetItem.style.borderTop = '2px solid var(--md-sys-color-primary)';
        dragInfo.position = 'above';
    } else if (offset > rect.height * 2 / 3) {
        targetItem.style.borderBottom = '2px solid var(--md-sys-color-primary)';
        dragInfo.position = 'below';
    } else {
        targetItem.classList.add('drop-hover');
        dragInfo.position = 'on';
    }

    // Tooltip message based on target and position
    const targetAct = findActivityById(targetItem.dataset.id);
    if (targetAct) {
        let msg = '';
        if (dragInfo.position === 'on') {
            msg = `Gruppera i: ${targetAct.name}`;
        } else if (dragInfo.position === 'above') {
            msg = `Flytta ovanför: ${targetAct.name}`;
        } else {
            msg = `Flytta under: ${targetAct.name}`;
        }
        updateInteractionTooltip(msg, e.clientX, e.clientY);
    }
}

function handleDragLeave(e) {
    e.target.closest('.activity-item')?.classList.remove('drop-hover');
}

function handleDrop(e, targetId) {
    e.preventDefault();
    document.querySelectorAll('.activity-item').forEach(item => {
        item.style.opacity = '1';
        item.style.borderTop = item.style.borderBottom = '';
        item.classList.remove('drop-hover');
    });
    
    if (dragInfo.sourceId === targetId) return;
    moveActivity(dragInfo.sourceId, targetId, dragInfo.position);
    hideInteractionTooltip();
}

// === Multi-selektion och urklipp ===
function toggleActivitySelection(id, additive) {
    if (!additive) { selectedActivityIds.clear(); selectedSegmentKeys.clear(); }
    if (selectedActivityIds.has(id)) selectedActivityIds.delete(id); else selectedActivityIds.add(id);
    selectedActivityId = id;
    highlightSelection();
}

function toggleSegmentSelection(activityId, segmentIndex, additive) {
    if (!additive) { selectedActivityIds.clear(); selectedSegmentKeys.clear(); }
    const key = `${activityId}:${segmentIndex}`;
    if (selectedSegmentKeys.has(key)) selectedSegmentKeys.delete(key); else selectedSegmentKeys.add(key);
    selectedActivityId = activityId;
    highlightSelection();
}

function clearSelection() {
    selectedActivityIds.clear();
    selectedSegmentKeys.clear();
    highlightSelection();
}

// Bygg färgpalett för Bulk Edit baserat på gemensamt projekt eller globala färger
function buildBulkPalette() {
    const dd = document.getElementById('bulkColorDropdown');
    if (!dd) return;
    dd.innerHTML = '';
    const projectIds = new Set();
    selectedActivityIds.forEach(id => { const a = findActivityById(id); const p = a ? findProjectForActivity(a) : null; if (p) projectIds.add(p.id); });
    selectedSegmentKeys.forEach(key => { const [aId] = key.split(':'); const a = findActivityById(aId); const p = a ? findProjectForActivity(a) : null; if (p) projectIds.add(p.id); });
    let palette = [];
    if (projectIds.size === 1) {
        const proj = findActivityById([...projectIds][0]);
        if (proj && Array.isArray(proj.colorPalette) && proj.colorPalette.length) palette = proj.colorPalette;
    }
    if (!palette.length) {
        const seen = new Map();
        (activities || []).forEach(a => {
            if (a.color) seen.set(a.color, a.colorName || a.color);
            if (Array.isArray(a.segments)) a.segments.forEach(s => { if (s.color) seen.set(s.color, s.color); });
        });
        palette = Array.from(seen.entries()).map(([color, name]) => ({ color, name }));
    }
    const emptyOpt = document.createElement('option');
    emptyOpt.value = '';
    emptyOpt.textContent = 'Välj från palett…';
    dd.appendChild(emptyOpt);
    palette.forEach(entry => {
        const opt = document.createElement('option');
        opt.value = entry.color;
        opt.textContent = entry.name ? `${entry.name} (${entry.color})` : entry.color;
        opt.style.background = entry.color;
        dd.appendChild(opt);
    });
}

// Verkställ bulkändring för namn och färg på markerade segment/aktiviteter
function applyBulkEdit() {
    const nameVal = (document.getElementById('bulkNameInput')?.value || '').trim();
    const colorVal = document.getElementById('bulkColorInput')?.value || '';
    const bulkHoursRaw = document.getElementById('bulkHoursPerDay')?.value || '';
    const bulkMenRaw = document.getElementById('bulkMenPerDay')?.value || '';
    if (!selectedActivityIds.size && !selectedSegmentKeys.size) return;

    const resolveHours = (act) => {
        if (bulkMenRaw !== '') {
            const men = parseFloat(bulkMenRaw);
            if (!isNaN(men)) return men * getStandardHoursPerDay();
        }
        if (bulkHoursRaw !== '') {
            const h = parseFloat(bulkHoursRaw);
            if (!isNaN(h)) return h;
        }
        return null;
    };

    const setSegFields = (seg, act) => {
        if (nameVal) seg.name = nameVal;
        if (colorVal) {
            seg.color = colorVal;
            try { const proj = findProjectForActivity(act); if (proj) ensureProjectHasColor(proj, colorVal); } catch(_) {}
        }
        const h = resolveHours(act);
        if (h !== null) seg.hoursPerDay = h;
    };

    // Segment som är markerade individuellt
    selectedSegmentKeys.forEach(key => {
        const [aId, idxStr] = key.split(':');
        const act = findActivityById(aId);
        const idx = parseInt(idxStr, 10);
        if (!act || !act.segments || isNaN(idx)) return;
        const seg = act.segments[idx];
        if (!seg) return;
        setSegFields(seg, act);
        updateActivityStartEndFromSegments(act);
        act.totalHours = calculateTotalHours(act);
        if (act.parent) updateGroupAndAncestors(act.parent);
    });

    // Alla segment i valda aktiviteter
    selectedActivityIds.forEach(aId => {
        const act = findActivityById(aId);
        if (!act || !Array.isArray(act.segments)) return;
        act.segments.forEach(seg => setSegFields(seg, act));
        updateActivityStartEndFromSegments(act);
        act.totalHours = calculateTotalHours(act);
        if (act.parent) updateGroupAndAncestors(act.parent);
    });

    pushHistory();
    renderAndPreserveScroll();
}

// === Panel visibility ===
function toggleProjectPanel() {
    const projectPanel = document.getElementById('projectPanel');
    if (!projectPanel) return;
    setProjectPanelVisibility(!isElementVisible(projectPanel));
    if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll();
}

function highlightSelection() {
    // Markera rader i sidopanelen
    document.querySelectorAll('.activity-item').forEach(el => {
        const id = el.dataset.id;
        if (id && selectedActivityIds.has(id)) el.classList.add('selected'); else el.classList.remove('selected');
    });
    // Markera staplar i gridden
    document.querySelectorAll('.bar').forEach(el => {
        const aId = el.dataset.id;
        const idx = el.dataset.segmentIndex;
        const isSegSel = idx != null && selectedSegmentKeys.has(`${aId}:${idx}`);
        if (isSegSel || (aId && selectedActivityIds.has(aId))) el.classList.add('selected'); else el.classList.remove('selected');
    });

    // Uppdatera Bulk Edit bar
    try {
        const bulkBar = document.getElementById('bulkEditBar');
        const countEl = document.getElementById('bulkSelectedCount');
        if (bulkBar && countEl) {
            const count = selectedActivityIds.size + selectedSegmentKeys.size;
            if (count > 1) {
                bulkBar.style.display = 'flex';
                countEl.textContent = String(count);
                buildBulkPalette();
            } else {
                bulkBar.style.display = 'none';
            }
        }
    } catch (_) {}
}

function copySelection() {
    clipboard.activities = [];
    clipboard.segments = [];
    const visited = new Set();
    function collect(act) {
        if (!act || visited.has(act.id)) return;
        visited.add(act.id);
        // Spara REFERENS (inte JSON) så vi behåller riktiga Date-objekt
        clipboard.activities.push(act);
        if (act.isGroup && act.children) act.children.forEach(cid => collect(findActivityById(cid)));
    }
    selectedActivityIds.forEach(id => collect(findActivityById(id)));
    // Segment som är markerade på aktiviteter som inte kopieras helt
    selectedSegmentKeys.forEach(key => {
        const [aId, idxStr] = key.split(':');
        if (visited.has(aId)) return;
        const a = findActivityById(aId);
        const idx = parseInt(idxStr, 10);
        const seg = a?.segments?.[idx];
        if (seg) {
            clipboard.segments.push({ activityId: aId, segment: { ...seg, start: new Date(seg.start), end: new Date(seg.end) } });
        }
    });
    if (window.developerMode) console.log('[copy] acts', clipboard.activities.length, 'segs', clipboard.segments.length);
}

// Raderar markerade aktiviteter och/eller segment (en historikpost)
function deleteSelection() {
    if (!selectedActivityIds.size && !selectedSegmentKeys.size) return;

    // 1. Hantera segmentraderingar för aktiviteter som inte är helt markerade
    const segmentRemovals = new Map(); // actId -> Set(indices)
    selectedSegmentKeys.forEach(key => {
        const [aId, idxStr] = key.split(':');
        if (selectedActivityIds.has(aId)) return; // hela aktiviteten raderas ändå
        if (!segmentRemovals.has(aId)) segmentRemovals.set(aId, new Set());
        segmentRemovals.get(aId).add(parseInt(idxStr, 10));
    });
    const activitiesBecomingEmpty = [];
    segmentRemovals.forEach((idxSet, aId) => {
        const act = findActivityById(aId);
        if (!act || !Array.isArray(act.segments)) return;
        const indices = [...idxSet].sort((a,b) => b-a); // ta bort baklänges
        indices.forEach(i => { if (i >=0 && i < act.segments.length) act.segments.splice(i,1); });
        if (!act.segments.length && !act.isGroup) {
            // Markera hela aktiviteten för radering
            selectedActivityIds.add(aId);
            activitiesBecomingEmpty.push(aId);
        } else {
            updateActivityStartEndFromSegments(act);
            act.totalHours = calculateTotalHours(act);
            if (act.parent) updateGroupAndAncestors(act.parent);
        }
    });

    // 2. Samla alla aktiviteter (inkl. grupper) som ska raderas
    const actsToDelete = new Set(selectedActivityIds);
    function collectDescendants(id) {
        const act = findActivityById(id);
        if (act?.isGroup && act.children) {
            act.children.forEach(ch => { actsToDelete.add(ch); collectDescendants(ch); });
        }
    }
    [...selectedActivityIds].forEach(collectDescendants);

    // 3. Samla föräldrar som måste uppdateras efteråt
    const parentsNeedingUpdate = new Set();
    actsToDelete.forEach(id => {
        const act = findActivityById(id);
        if (act?.parent) parentsNeedingUpdate.add(act.parent);
    });

    // 4. Ta bort child-referenser från föräldrar
    actsToDelete.forEach(id => {
        const act = findActivityById(id);
        if (act?.parent) removeChildFromParent(id, act.parent);
    });

    // 5. Filtrera bort aktiviteterna
    activities = activities.filter(a => !actsToDelete.has(a.id));

    // 6. Uppdatera berörda grupper (föräldrar)
    parentsNeedingUpdate.forEach(gid => updateGroupAndAncestors(gid));

    // 7. Rensa markering och logga historik
    // Capture an anchor before clearing selection (existing row id in OLD DOM)
    try {
        const anchorId = selectedActivityId || (selectedActivityIds.size ? [...selectedActivityIds][0] : null);
        if (anchorId) {
            window.__forceListAnchorId = anchorId;
            const listEl = document.getElementById('activityList');
            const rowEl = listEl ? listEl.querySelector(`.activity-item[data-id="${anchorId}"]`) : null;
            if (listEl && rowEl) {
                window.__forceListAnchorOffset = listEl.scrollTop - rowEl.offsetTop;
            }
        }
    } catch(_) {}
    clearSelection();
    pushHistory();
    if (window.developerMode) console.log('[delete] removed activities', actsToDelete.size);
    renderAndPreserveScroll();
}

function pasteSelection() {
    if (!clipboard.activities.length && !clipboard.segments.length) return;
    if (window.developerMode) console.log('[paste] start acts', clipboard.activities.length, 'segs', clipboard.segments.length, 'selectedActivityId', selectedActivityId);
    const viewStartInput = document.getElementById('startDate');
    const viewEndInput = document.getElementById('endDate');
    let viewStart = normalizeDate(viewStartInput.value);
    let viewEnd = normalizeDate(viewEndInput.value);

    const sourceActs = clipboard.activities.slice(); // referenser
    let earliest = null; let latest = null;
    sourceActs.forEach(a => {
        if (a.segments?.length) {
            a.segments.forEach(s => {
                const sS = new Date(s.start); const sE = new Date(s.end);
                if (!earliest || sS < earliest) earliest = sS;
                if (!latest || sE > latest) latest = sE;
            });
        } else if (a.start) {
            const sS = new Date(a.start); const sE = a.end ? new Date(a.end) : new Date(a.start);
            if (!earliest || sS < earliest) earliest = sS;
            if (!latest || sE > latest) latest = sE;
        }
    });
    clipboard.segments.forEach(item => {
        const sS = new Date(item.segment.start); const sE = new Date(item.segment.end);
        if (!earliest || sS < earliest) earliest = sS;
        if (!latest || sE > latest) latest = sE;
    });

    let dayOffset = 0;
    if (pasteAnchorDate && earliest) {
        const ms = normalizeDate(pasteAnchorDate).getTime() - normalizeDate(earliest).getTime();
        dayOffset = Math.round(ms / 86400000);
    }

    // ID-map
    const idMap = {};
    sourceActs.forEach(a => { idMap[a.id] = generateId(); });
    const nowBase = Date.now();
    const newCopies = [];

    function cloneActivity(src, orderIdx) {
        const copy = { ...src };
        copy.id = idMap[src.id];
        copy._order = nowBase + orderIdx;
        copy.children = [];
        copy.parent = null; // sätts senare om parent kopieras
        if (src.segments && src.segments.length) {
            copy.segments = src.segments.map(seg => {
                const st = new Date(seg.start); const en = new Date(seg.end);
                if (dayOffset) { st.setDate(st.getDate() + dayOffset); en.setDate(en.getDate() + dayOffset); }
                return { ...seg, start: st, end: en };
            });
        } else {
            copy.segments = [];
        }
        if (src.start) {
            copy.start = new Date(src.start); if (dayOffset) copy.start.setDate(copy.start.getDate() + dayOffset);
        }
        if (src.end) {
            copy.end = new Date(src.end); if (dayOffset) copy.end.setDate(copy.end.getDate() + dayOffset);
        }
        if (!copy.isGroup && copy.segments.length) {
            updateActivityStartEndFromSegments(copy);
            copy.totalHours = calculateTotalHours(copy);
        }
        return copy;
    }

    sourceActs.forEach((src, idx) => newCopies.push(cloneActivity(src, idx)));

    // Återskapa parent/children
    const newById = Object.fromEntries(newCopies.map(c => [c.id, c]));
    sourceActs.forEach(src => {
        if (src.isGroup && src.children?.length) {
            const newParent = newById[idMap[src.id]];
            src.children.forEach(chOld => { if (idMap[chOld]) newParent.children.push(idMap[chOld]); });
        }
    });
    // Sätt parent på barn som fick ny förälder
    sourceActs.forEach(src => {
        if (src.parent && idMap[src.parent]) {
            const newChild = newById[idMap[src.id]];
            newChild.parent = idMap[src.parent];
        }
    });

    // Bestäm målgrupp (om exakt en grupp vald)
    let targetGroupId = null;
    if (selectedActivityId) {
        const tgt = findActivityById(selectedActivityId);
        if (tgt?.isGroup) targetGroupId = tgt.id;
    } else if (selectedActivityIds.size === 1) {
        const only = [...selectedActivityIds][0];
        const tgt = findActivityById(only);
        if (tgt?.isGroup) targetGroupId = tgt.id;
    }
    if (targetGroupId) {
        const tgtGroup = findActivityById(targetGroupId);
        if (tgtGroup) {
            tgtGroup.children = tgtGroup.children || [];
            newCopies.forEach(c => { if (!c.parent) { c.parent = targetGroupId; tgtGroup.children.push(c.id); } });
            groupOpenState[targetGroupId] = true;
        }
    }

    // Lägg in
    newCopies.forEach(c => { activities.push(c); if (c.isGroup) groupOpenState[c.id] = true; });
    newCopies.filter(c => c.isGroup).forEach(g => updateGroupAndAncestors(g.id));
    if (targetGroupId) updateGroupAndAncestors(targetGroupId);

    // Fristående segment
    clipboard.segments.forEach(item => {
        const act = findActivityById(item.activityId);
        if (!act) return;
        const seg = { ...item.segment, start: new Date(item.segment.start), end: new Date(item.segment.end) };
        if (dayOffset) { seg.start.setDate(seg.start.getDate() + dayOffset); seg.end.setDate(seg.end.getDate() + dayOffset); }
        act.segments = act.segments || []; act.segments.push(seg);
        cleanupAndMergeSegments(act);
        updateActivityStartEndFromSegments(act);
        act.totalHours = calculateTotalHours(act);
        if (act.parent) updateGroupAndAncestors(act.parent);
    });

    // Auto-expand date span if needed
    if (autoExpandSpan) {
        let newMin = null, newMax = null;
        newCopies.forEach(c => {
            if (c.segments?.length) {
                c.segments.forEach(s => {
                    if (!newMin || s.start < newMin) newMin = s.start;
                    if (!newMax || s.end > newMax) newMax = s.end;
                });
            } else if (c.start) {
                if (!newMin || c.start < newMin) newMin = c.start;
                if (!newMax || c.end > newMax) newMax = c.end;
            }
        });
        if (newMin && viewStart && newMin < viewStart) { viewStart = newMin; viewStartInput.value = formatDateLocal(viewStart); }
    if (newMax && viewEnd && newMax > viewEnd) { viewEnd = newMax; viewEndInput.value = formatDateLocal(viewEnd); }
    }

    // Hint scroll restorer to anchor around the target area (existing row id in OLD DOM)
    try {
        const anchorId = targetGroupId || selectedActivityId || (selectedActivityIds.size ? [...selectedActivityIds][0] : null);
        if (anchorId) {
            window.__forceListAnchorId = anchorId;
            const listEl = document.getElementById('activityList');
            const rowEl = listEl ? listEl.querySelector(`.activity-item[data-id="${anchorId}"]`) : null;
            if (listEl && rowEl) {
                window.__forceListAnchorOffset = listEl.scrollTop - rowEl.offsetTop;
            }
        }
    } catch(_) {}

    pushHistory();
    selectedActivityIds.clear();
    selectedSegmentKeys.clear();
    newCopies.forEach(c => selectedActivityIds.add(c.id));
    highlightSelection();
    if (window.developerMode) console.log('[paste] inserted', newCopies.length, 'targetGroup', targetGroupId);
    renderAndPreserveScroll();
}

// Snabb kopia utan datumförskjutning eller clamp (bevarar exakt datum)
function duplicateSelectionSameDates() {
    if (!clipboard.activities.length && !selectedActivityIds.size) return;
    // Om ingen explicit kopia ännu: skapa en från aktuell selection
    if (!clipboard.activities.length) copySelection();
    // Temporärt spara anchor och disable clamp
    const prevAnchor = pasteAnchorDate;
    pasteAnchorDate = null; // ingen offset
    const prevClamp = enablePasteClamp;
    enablePasteClamp = false;
    pasteSelection();
    // återställ
    pasteAnchorDate = prevAnchor;
    enablePasteClamp = prevClamp;
}

// === POPUP-HANTERING (segment-medveten) ===

const popupDragState = {
    active: false,
    pointerId: null,
    offsetX: 0,
    offsetY: 0
};

function clampPopupPosition(popup, left, top) {
    const margin = 16;
    const rect = popup.getBoundingClientRect();
    const maxLeft = Math.max(margin, window.innerWidth - rect.width - margin);
    const maxTop = Math.max(margin, window.innerHeight - rect.height - margin);
    return {
        left: Math.min(Math.max(left, margin), maxLeft),
        top: Math.min(Math.max(top, margin), maxTop)
    };
}

function fitPopupToViewport(popup) {
    if (!popup) return 1;
    const margin = 16;
    popup.style.transformOrigin = 'top left';
    popup.style.transform = 'none';
    const rect = popup.getBoundingClientRect();
    const availableWidth = Math.max(320, window.innerWidth - margin * 2);
    const availableHeight = Math.max(320, window.innerHeight - margin * 2);
    const scale = Math.min(
        1,
        availableWidth / Math.max(rect.width, 1),
        availableHeight / Math.max(rect.height, 1)
    );
    popup.dataset.scale = String(scale);
    popup.style.transform = scale < 0.999 ? `scale(${scale})` : 'none';
    return scale;
}

function refreshPopupLayout() {
    const popup = document.getElementById('activityPopup');
    if (!popup || popup.style.display !== 'block') return;
    const currentLeft = parseFloat(popup.style.left);
    const currentTop = parseFloat(popup.style.top);
    fitPopupToViewport(popup);
    const { left, top } = clampPopupPosition(
        popup,
        Number.isFinite(currentLeft) ? currentLeft : 16,
        Number.isFinite(currentTop) ? currentTop : 16
    );
    popup.style.left = `${left}px`;
    popup.style.top = `${top}px`;
}

function showPopupAtPosition(popup, event) {
    if (!popup) return;
    popup.style.display = 'block';
    popup.style.visibility = 'hidden';
    popup.style.left = '16px';
    popup.style.top = '16px';
    fitPopupToViewport(popup);
    const rect = popup.getBoundingClientRect();
    const desiredLeft = event && Number.isFinite(event.clientX)
        ? event.clientX + 16
        : (window.innerWidth - rect.width) / 2;
    const desiredTop = event && Number.isFinite(event.clientY)
        ? event.clientY - 12
        : (window.innerHeight - rect.height) / 2;
    const { left, top } = clampPopupPosition(popup, desiredLeft, desiredTop);
    popup.style.left = `${left}px`;
    popup.style.top = `${top}px`;
    popup.style.visibility = '';
}

function setupPopupInteractions() {
    const popup = document.getElementById('activityPopup');
    const header = popup?.querySelector('.popup-header');
    if (!popup || !header || header.dataset.dragReady === 'true') return;
    header.dataset.dragReady = 'true';

    const stopDragging = (pointerId) => {
        if (!popupDragState.active) return;
        popupDragState.active = false;
        popupDragState.pointerId = null;
        popup.classList.remove('dragging');
        if (pointerId != null && typeof header.releasePointerCapture === 'function') {
            try { header.releasePointerCapture(pointerId); } catch (_) {}
        }
    };

    header.addEventListener('pointerdown', (e) => {
        if (e.button !== 0) return;
        if (e.target.closest('button, input, select, textarea, a')) return;
        const rect = popup.getBoundingClientRect();
        popupDragState.active = true;
        popupDragState.pointerId = e.pointerId;
        popupDragState.offsetX = e.clientX - rect.left;
        popupDragState.offsetY = e.clientY - rect.top;
        popup.classList.add('dragging');
        if (typeof header.setPointerCapture === 'function') {
            try { header.setPointerCapture(e.pointerId); } catch (_) {}
        }
        e.preventDefault();
    });

    header.addEventListener('pointermove', (e) => {
        if (!popupDragState.active || popupDragState.pointerId !== e.pointerId) return;
        const { left, top } = clampPopupPosition(
            popup,
            e.clientX - popupDragState.offsetX,
            e.clientY - popupDragState.offsetY
        );
        popup.style.left = `${left}px`;
        popup.style.top = `${top}px`;
    });

    header.addEventListener('pointerup', (e) => stopDragging(e.pointerId));
    header.addEventListener('pointercancel', (e) => stopDragging(e.pointerId));
    document.addEventListener('pointerup', () => stopDragging());
    window.addEventListener('resize', refreshPopupLayout);
}

function openPopup(activityId, event) {
    // Stäng eventuellt tidigare popup och välj aktivitet
    closePopup();
    selectedActivityId = activityId;

    const act = findActivityById(activityId);
    if (!act) return;

    const popup = document.getElementById('activityPopup');
    const titleEl = document.getElementById('popupTitle');
    const subtitleEl = document.getElementById('popupSubtitle');
    const eyebrowEl = document.getElementById('popupEyebrow');
    const metaEl = document.getElementById('popupMetaSummary');

    // Hjälpare för att visa/dölja element/sektioner
    const show = el => { if (el) el.style.display = ''; };
    const hide = el => { if (el) el.style.display = 'none'; };
    const rowOf = el => el ? el.closest('.popup-row') : null;

    // Gemensamma referenser
    const startRow = rowOf(document.getElementById('popupStartDate'));
    const endRow = rowOf(document.getElementById('popupEndDate'));
    const completedRow = rowOf(document.getElementById('popupCompletedValue'));
    const completedSlider = document.getElementById('popupCompleted');
    const splitRow = rowOf(document.getElementById('popupSplitOnWeekend'));
    const actColorRow = rowOf(document.getElementById('popupActivityColor'));
    const segColorRow = rowOf(document.getElementById('popupSegmentColor'));
    const groupColorRow = rowOf(document.getElementById('popupGroupColor'));
    const projectColorRow = document.getElementById('projectColorRow');
    const activitySettings = document.getElementById('popupActivitySettings');
    const groupSettings = document.getElementById('popupGroupSettings');
    const splitBtn = document.getElementById('splitActivityBtn');

    // Segmentfärgs-dropdown: bygg från projektpalett eller globala färger
    const segColorDropdown = document.getElementById('popupSegmentColorDropdown');
    if (segColorDropdown) {
        segColorDropdown.innerHTML = '';
        let palette = [];
        const proj = findProjectForActivity(act);
        if (proj && Array.isArray(proj.colorPalette) && proj.colorPalette.length > 0) {
            palette = proj.colorPalette;
        } else {
            const uniqueColors = new Map();
            (activities || []).forEach(a => { if (a.color) uniqueColors.set(a.color, a.colorName || a.color); });
            palette = Array.from(uniqueColors.entries()).map(([color, name]) => ({ color, name }));
        }
        // Lägg till en tom option för "egen" färg
        const emptyOpt = document.createElement('option');
        emptyOpt.value = '';
        emptyOpt.textContent = 'Välj från palett…';
        segColorDropdown.appendChild(emptyOpt);
        palette.forEach(entry => {
            const opt = document.createElement('option');
            opt.value = entry.color;
            opt.textContent = entry.name ? `${entry.name} (${entry.color})` : entry.color;
            opt.style.background = entry.color;
            segColorDropdown.appendChild(opt);
        });
        // När man väljer i listan uppdateras color-inputen
        segColorDropdown.onchange = () => {
            const v = segColorDropdown.value;
            const input = document.getElementById('popupSegmentColor');
            if (v) input.value = v;
        };
    }

    if (act.isGroup) {
        // Grupp eller projekt: minimal meny
        if (eyebrowEl) eyebrowEl.textContent = act.isProject ? 'Projekt' : 'Grupp';
        if (titleEl) titleEl.textContent = act.name || (act.isProject ? 'Projekt' : 'Grupp');
        if (subtitleEl) subtitleEl.textContent = act.isProject ? 'Justera projektgränser, färg och max timmar.' : 'Justera gruppens färg och övergripande timmar.';
        hide(startRow); hide(endRow); hide(completedRow); hide(completedSlider);
        hide(activitySettings); show(groupSettings);
    hide(splitRow); hide(segColorRow); hide(actColorRow);
        show(groupColorRow);
        if (act.isProject) {
            show(projectColorRow);
            const pc = document.getElementById('popupProjectColor');
            if (pc) pc.value = act.color || '#007bff';
            hide(groupColorRow);
        } else {
            hide(projectColorRow);
            const gc = document.getElementById('popupGroupColor');
            if (gc) gc.value = act.color || '#007bff';
        }
        hide(splitBtn);
        // Fyll fält
        const maxEl = document.getElementById('popupMaxHours');
        if (maxEl) maxEl.value = act.maxTotalHours || '';
        const ghi = document.getElementById('groupHoursInfo');
        if (ghi) ghi.textContent = `Totalt i grupp: ${getHoursForActivity(act).toFixed(1)}h`;
        if (metaEl) {
            const totalHours = getHoursForActivity(act).toFixed(1);
            metaEl.textContent = `Färdig: ${act.completed || 0}% • Totalt: ${totalHours}h${act.maxTotalHours ? ` • Max: ${act.maxTotalHours}h` : ''}`;
        }

        showPopupAtPosition(popup, event);
        return;
    }

    // Icke-grupp: normal aktivitet
    if (eyebrowEl) eyebrowEl.textContent = 'Aktivitet';
    const barElement = event && event.target ? event.target.closest('.bar') : null;
    const segmentIndex = barElement ? parseInt(barElement.dataset.segmentIndex, 10) : 0;
    const segment = (act.segments && act.segments.length) ? (act.segments[segmentIndex] || act.segments[0]) : null;
    popup.dataset.segmentIndex = String(segmentIndex); // Spara index för savePopup

    // --- Beräkna och spara delningsdatum ---
    const ganttGrid = document.querySelector('.gantt-grid');
    if (ganttGrid) {
        const dayWidth = (typeof interaction !== 'undefined' && interaction.dayWidth) ? interaction.dayWidth : 35;
        const viewStartDate = normalizeDate(document.getElementById('startDate').value);
        const gridRect = ganttGrid.getBoundingClientRect();
        const clickX = event.clientX - gridRect.left;
        const daysOffset = Math.floor(clickX / dayWidth);
        let splitDate = new Date(viewStartDate);
        splitDate.setDate(splitDate.getDate() + daysOffset);
        popup.dataset.splitDate = formatDateLocal(splitDate);
    }
    // --- Slut på beräkning ---

    // Fyll fält
    if (segment) {
        document.getElementById('popupStartDate').value = formatDateLocal(segment.start);
        document.getElementById('popupEndDate').value = formatDateLocal(segment.end);
        const startH = String(segment.start.getHours()).padStart(2, '0');
        const startM = String(segment.start.getMinutes()).padStart(2, '0');
        const endH = String(segment.end.getHours()).padStart(2, '0');
        const endM = String(segment.end.getMinutes()).padStart(2, '0');
        document.getElementById('popupStartTime').value = `${startH}:${startM}`;
        document.getElementById('popupEndTime').value = `${endH}:${endM}`;
        document.getElementById('popupInfo').value = segment.info || act.info || '';
        document.getElementById('popupSegmentName').value = segment.name || '';
        const respEl = document.getElementById('popupSegmentResponsibles');
        const respList = Array.isArray(segment.responsibles) ? segment.responsibles : [];
        respEl.value = respList.join(', ');
        // Segmentfärg
        document.getElementById('popupSegmentColor').value = (segment.color || act.color || '#007bff');
        // New: segment-level hours and men
        const segHpdEl = document.getElementById('popupSegmentHoursPerDay');
        const segMenEl = document.getElementById('popupSegmentMenPerDay');
        if (segHpdEl) segHpdEl.value = (typeof segment.hoursPerDay === 'number') ? segment.hoursPerDay : '';
        if (segMenEl) {
            const hoursPerMan = getStandardHoursPerDay();
            segMenEl.value = (typeof segment.hoursPerDay === 'number') ? (segment.hoursPerDay / hoursPerMan) : '';
        }
    } else {
        // Fallback om segment saknas
        document.getElementById('popupStartDate').value = '';
        document.getElementById('popupEndDate').value = '';
        document.getElementById('popupStartTime').value = '';
        document.getElementById('popupEndTime').value = '';
        document.getElementById('popupInfo').value = act.info || '';
        document.getElementById('popupSegmentName').value = '';
        document.getElementById('popupSegmentColor').value = act.color || '#007bff';
    }

    // Aktivitetsfärg och completion
    // Aktivitetsfärgfält borttaget från popup för normal aktivitet
    const completedInput = document.getElementById('popupCompleted');
    const completedValue = document.getElementById('popupCompletedValue');
    completedInput.value = act.completed || 0;
    completedValue.textContent = `${act.completed || 0}%`;
    completedInput.oninput = () => { completedValue.textContent = `${completedInput.value}%`; };

    if (titleEl) titleEl.textContent = act.name || 'Aktivitet';
    if (subtitleEl) {
        const segLabel = segment?.name ? `Segment ${segmentIndex + 1}: ${segment.name}` : `Segment ${segmentIndex + 1}`;
        subtitleEl.textContent = `Redigera ${segLabel.toLowerCase()} och uppdatera tid, färg och resurser.`;
    }

    // Uppdatera popup-meta-sammanfattning (completion, total hours, hours/day)
    try {
        const total = (act.totalHours || calculateTotalHours(act) || 0).toFixed(1);
        const hpd = (act.hoursPerDay || getStandardHoursPerDay());
        const comp = (act.completed || 0);
        if (metaEl) metaEl.textContent = `Färdig: ${comp}% • Total: ${total}h • ${hpd}h/d`;
    } catch (e) { /* ignore */ }

    // Synlighet för normal aktivitet
    show(startRow); show(endRow); show(completedRow); show(completedSlider);
    show(activitySettings); hide(groupSettings);
    show(splitRow); show(actColorRow); show(segColorRow);
    show(splitBtn);

    // Övriga fält
    document.getElementById('popupHoursPerDay').value = act.hoursPerDay;
    document.getElementById('popupSplitOnWeekend').checked = act.splitOnWeekend || false;
    document.getElementById('totalHoursInfo').textContent = `Summa: ${(act.totalHours || 0).toFixed(1)}h`;

    // Live-uppdatera timmar från tider
    const startTimeEl = document.getElementById('popupStartTime');
    const endTimeEl = document.getElementById('popupEndTime');
    const hoursPerDayEl = document.getElementById('popupHoursPerDay');
    const segHpdEl2 = document.getElementById('popupSegmentHoursPerDay');
    const segMenEl2 = document.getElementById('popupSegmentMenPerDay');
    const totalInfoEl = document.getElementById('totalHoursInfo');
    const recalcFromTimes = () => {
        const st = startTimeEl.value;
        const et = endTimeEl.value;
        if (!st || !et) return;
        const [sh, sm] = st.split(':').map(Number);
        const [eh, em] = et.split(':').map(Number);
        const startMin = (sh || 0) * 60 + (sm || 0);
        const endMin = (eh || 0) * 60 + (em || 0);
        let diffMin = endMin - startMin;
        if (diffMin < 0) diffMin = 0;
        const hours = Math.round((diffMin / 60) * 2) / 2;
        if (!isNaN(hours)) {
            if (segHpdEl2) segHpdEl2.value = hours;
            else hoursPerDayEl.value = hours;
            const totalDays = (act.segments || []).reduce((s, seg) => s + getWorkDays(seg.start, seg.end), 0);
            totalInfoEl.textContent = `Summa: ${(hours * totalDays).toFixed(1)}h`;
        }
    };
    startTimeEl.oninput = recalcFromTimes;
    endTimeEl.oninput = recalcFromTimes;

    // Live sync men<->hours for segment
    if (segHpdEl2 && segMenEl2) {
        const syncMenFromHours = () => {
            const val = parseFloat(segHpdEl2.value);
            if (!isNaN(val)) segMenEl2.value = (val / getStandardHoursPerDay()).toFixed(2);
            else segMenEl2.value = '';
        };
        const syncHoursFromMen = () => {
            const men = parseFloat(segMenEl2.value);
            if (!isNaN(men)) segHpdEl2.value = (men * getStandardHoursPerDay()).toFixed(2);
            else segHpdEl2.value = '';
        };
        segHpdEl2.addEventListener('input', syncMenFromHours);
        segMenEl2.addEventListener('input', syncHoursFromMen);
    }

    showPopupAtPosition(popup, event);
}

function closePopup() {
    const popup = document.getElementById('activityPopup');
    if (popup) {
        popup.style.display = 'none';
        popup.style.visibility = '';
        popup.classList.remove('dragging');
    }
    popupDragState.active = false;
    popupDragState.pointerId = null;
    selectedActivityId = null;
}

function savePopup() {
    if (!selectedActivityId) return;
    const act = findActivityById(selectedActivityId);
    const popup = document.getElementById('activityPopup');
    if (!act) return;

    // Spara grupp/prosjekt (minimal meny)
    if (act.isGroup) {
        const maxValRaw = document.getElementById('popupMaxHours').value;
        const maxVal = maxValRaw === '' ? undefined : Number(maxValRaw);
        if (maxVal === undefined || isFinite(maxVal)) {
            if (maxVal === undefined) delete act.maxTotalHours; else act.maxTotalHours = maxVal;
        }
        if (act.isProject) {
            act.color = document.getElementById('popupProjectColor').value || act.color;
        } else {
            act.color = document.getElementById('popupGroupColor').value || act.color;
        }
        pushHistory();
        renderAndPreserveScroll();
        closePopup();
        return;
    }

    const segmentIndex = parseInt(popup.dataset.segmentIndex, 10);
    const segment = act.segments ? act.segments[segmentIndex] : null;

    if (act && segment) {
        let newStart = normalizeDate(document.getElementById('popupStartDate').value);
        let newEnd = normalizeDate(document.getElementById('popupEndDate').value);
        // Applicera ev. tider
        const st = document.getElementById('popupStartTime').value;
        const et = document.getElementById('popupEndTime').value;
        if (st) {
            const [hh, mm] = st.split(':').map(Number);
            newStart.setHours(hh || 0, mm || 0, 0, 0);
        }
        if (et) {
            const [hh, mm] = et.split(':').map(Number);
            newEnd.setHours(hh || 0, mm || 0, 0, 0);
        }
        // Spara segmentfärg och säkerställ att den finns i projektpaletten
        const segColorVal = document.getElementById('popupSegmentColor').value;
        if (segColorVal) {
            segment.color = segColorVal;
            try {
                const projForPalette = findProjectForActivity(act);
                if (projForPalette) ensureProjectHasColor(projForPalette, segColorVal);
            } catch(_) {}
        } else {
            delete segment.color;
        }
        if (newStart && newEnd && newEnd < newStart) {
            newEnd = newStart;
        }
    segment.start = newStart;
    segment.end = newEnd;
        // Spara info anteckning på segmentnivå; fallback kan vara på aktivitetsnivå om så önskas
        const infoVal = document.getElementById('popupInfo').value;
        if (infoVal) segment.info = infoVal; else delete segment.info;
        // New: save segment hours/men
        const segHpdRaw = document.getElementById('popupSegmentHoursPerDay')?.value;
        const segMenRaw = document.getElementById('popupSegmentMenPerDay')?.value;
        let segHpd = segHpdRaw === undefined ? '' : segHpdRaw;
        if (segMenRaw && segMenRaw !== '') {
            const men = parseFloat(segMenRaw);
            if (!isNaN(men)) segHpd = (men * getStandardHoursPerDay());
        }
        if (segHpd === '' || isNaN(parseFloat(segHpd))) {
            delete segment.hoursPerDay;
        } else {
            segment.hoursPerDay = parseFloat(segHpd);
        }
    act.completed = parseInt(document.getElementById('popupCompleted').value, 10);
    // Aktivitetsfärg ändras inte här längre (endast segmentfärg)
    const segNameVal = document.getElementById('popupSegmentName').value;
    if (segNameVal) segment.name = segNameVal; else delete segment.name;
    // Spara ansvariga e‑postadresser
    const respStr = document.getElementById('popupSegmentResponsibles').value || '';
    const emails = respStr.split(',').map(s => s.trim()).filter(Boolean);
    if (emails.length) segment.responsibles = emails; else delete segment.responsibles;

        if (!act.isGroup) {
            act.hoursPerDay = parseFloat(document.getElementById('popupHoursPerDay').value);
            act.splitOnWeekend = document.getElementById('popupSplitOnWeekend').checked;
        }

        // Validate against project max before applying
        // Temporarily apply to compute
        updateActivityStartEndFromSegments(act);
        const tentativeTotal = calculateTotalHours(act);
        const proj = findProjectForActivity(act);
        const projMax = proj ? (proj.maxTotalHours || proj.projectInfo?.maxHours || null) : null;
        if (projMax && tentativeTotal > projMax + 1e-6) {
            const over = (tentativeTotal - projMax).toFixed(1);
            const ok = confirm(`Sparandet skulle överskrida projektets max timmar med ${over}h. Vill du fortsätta?`);
            if (!ok) {
                // revert segment changes
                segment.start = new Date(segment.start);
                segment.end = new Date(segment.end);
                // Recompute back to previous (we can reload from history or recompute but simplest is to cancel applying total change)
                // Here we reload last state by popping history or re-render without pushing.
                renderAndPreserveScroll();
                closePopup();
                return;
            }
        }
        act.totalHours = tentativeTotal;
        updateAllGroups();

    if (act.splitOnWeekend) {
            splitActivityAtWeekends(act.id);
        } else {
            pushHistory();
            renderAndPreserveScroll();
        }
        closePopup();
        return;
    }
}

// === PROJECT INFO MODAL ===
function openProjectInfo(projectId) {
    const act = findActivityById(projectId);
    if (!act) return;
    const backdrop = document.getElementById('modalBackdrop');
    const modal = document.getElementById('centeredModal');
    const content = document.getElementById('modalContent');
    // Build form
    const info = act.projectInfo || {};
    // compute planned hours
    const planned = typeof getProjectPlannedHours === 'function' ? getProjectPlannedHours(act) : getHoursForActivity(act);
    const maxSet = act.maxTotalHours || info.maxHours || '';
    const remaining = (maxSet ? (Number(maxSet) - planned) : null);
    content.innerHTML = `
        <div class="project-info-dialog">
            <h3>Planeringsinfo</h3>
            <div class="dialog-row"><input id="proj_planner" placeholder="Planerare" value="${escapeHtml(info.planner||'')}"></div>
            <div class="dialog-row"><input id="proj_company" placeholder="Företagsnamn" value="${escapeHtml(info.company||'')}"></div>
            <div class="dialog-row"><input id="proj_orgnr" placeholder="Organisationsnummer" value="${escapeHtml(info.orgnr||'')}"></div>
            <div class="dialog-row"><input id="proj_contact" placeholder="Kontaktuppgifter" value="${escapeHtml(info.contact||'')}"></div>
            <div class="dialog-row"><input id="proj_revision" placeholder="Revision" value="${escapeHtml(info.revision||'')}"></div>
            <div class="dialog-row"><input id="proj_stage" placeholder="Etapp" value="${escapeHtml(info.stage||'')}"></div>
            <div class="dialog-row"><textarea id="proj_comment" placeholder="Kommentar eller beskrivning">${escapeHtml(info.comment||'')}</textarea></div>
            <div class="dialog-row"><label for="proj_maxHours">Max timmar för projekt (valfritt)</label><input id="proj_maxHours" type="number" min="0" step="0.5" placeholder="Max timmar" value="${escapeHtml(maxSet||'')}"></div>
            <div class="project-dialog-metrics" id="proj_plannedRow">
                <div class="project-dialog-metric">
                    <span class="project-dialog-metric-label">Planerade timmar</span>
                    <strong class="project-dialog-metric-value" id="proj_plannedVal">${planned.toFixed(1)}h</strong>
                </div>
                <div class="project-dialog-metric">
                    <span class="project-dialog-metric-label">Återstående</span>
                    <strong class="project-dialog-metric-value${remaining!==null && remaining<0 ? ' status-error' : ''}" id="proj_remainingVal">${remaining===null?'-':remaining.toFixed(1)+'h'}</strong>
                </div>
                <div class="project-dialog-metric${remaining!==null && remaining<0 ? ' status-error' : ''}${remaining===null || remaining>=0 ? ' status-hidden' : ''}" id="proj_overrunRow">
                    <span class="project-dialog-metric-label">Under/överskott timmar</span>
                    <strong class="project-dialog-metric-value" id="proj_overrunVal">${remaining===null?'-':(remaining<0?('+'+(-remaining).toFixed(1)+'h'):('-'+(remaining).toFixed(1)+'h'))}</strong>
                </div>
            </div>
            <div class="project-dialog-actions">
                <button id="proj_save" class="filled-button" type="button">Spara</button>
                <button id="proj_clear" class="text-button" type="button">Rensa</button>
                <button id="proj_close" class="tonal-button" type="button">Stäng</button>
            </div>
        </div>`;
    backdrop.style.display = 'block';
    modal.style.display = 'block';

    // Handlers
    document.getElementById('proj_close').onclick = closeProjectModal;
    document.getElementById('proj_clear').onclick = () => {
        document.getElementById('proj_planner').value = '';
        document.getElementById('proj_company').value = '';
        document.getElementById('proj_orgnr').value = '';
        document.getElementById('proj_contact').value = '';
        document.getElementById('proj_revision').value = '';
        document.getElementById('proj_stage').value = '';
        document.getElementById('proj_comment').value = '';
    };
    document.getElementById('proj_save').onclick = () => {
        act.projectInfo = {
            planner: document.getElementById('proj_planner').value,
            company: document.getElementById('proj_company').value,
            orgnr: document.getElementById('proj_orgnr').value,
            contact: document.getElementById('proj_contact').value,
            revision: document.getElementById('proj_revision').value,
            stage: document.getElementById('proj_stage').value,
            comment: document.getElementById('proj_comment').value,
        };
    // save max hours if provided
    const maxVal = document.getElementById('proj_maxHours').value;
    if (maxVal === '') delete act.maxTotalHours;
    else act.maxTotalHours = Number(maxVal);
        pushHistory(); if (typeof renderAndPreserveScroll === 'function') renderAndPreserveScroll(); else render();
        closeProjectModal();
    };

    // Live update remaining when maxHours input changes
    const maxInput = document.getElementById('proj_maxHours');
    if (maxInput) {
        maxInput.addEventListener('input', () => {
            const v = maxInput.value === '' ? null : Number(maxInput.value);
            const plannedEl = document.getElementById('proj_plannedVal');
            const remEl = document.getElementById('proj_remainingVal');
            const plannedNum = plannedEl ? parseFloat(plannedEl.textContent.replace('h','')) : 0;
            const overrunRow = document.getElementById('proj_overrunRow');
            const overrunVal = document.getElementById('proj_overrunVal');
            if (v === null) {
                if (remEl) {
                    remEl.textContent = '-';
                    remEl.classList.remove('status-error');
                }
                if (overrunRow) {
                    overrunRow.classList.remove('status-error');
                    overrunRow.classList.add('status-hidden');
                    overrunVal.textContent = '';
                }
            } else {
                const rem = v - plannedNum;
                if (remEl) {
                    remEl.textContent = rem.toFixed(1) + 'h';
                    remEl.classList.toggle('status-error', rem < 0);
                }
                if (overrunRow) {
                    if (rem < 0) {
                        overrunRow.classList.add('status-error');
                        overrunRow.classList.remove('status-hidden');
                        overrunVal.textContent = (-rem).toFixed(1) + 'h';
                    } else {
                        overrunRow.classList.remove('status-error');
                        overrunRow.classList.add('status-hidden');
                        overrunVal.textContent = '';
                    }
                }
            }
        });
    }
}

function closeProjectModal() {
    const backdrop = document.getElementById('modalBackdrop');
    const modal = document.getElementById('centeredModal');
    backdrop.style.display = 'none';
    modal.style.display = 'none';
    document.getElementById('modalContent').innerHTML = '';
}

// small helper to avoid XSS when setting innerHTML
function escapeHtml(str) {
    return String(str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// === Popup actions: delete, split ===
function deleteFromPopup() {
    try {
        const popup = document.getElementById('activityPopup');
        const actId = selectedActivityId;
        if (!actId) return;
        const act = findActivityById(actId);
        if (!act) return;

        // Groups/projects: delete entire container (with children)
        if (act.isGroup) {
            const ok = confirm('Radera denna grupp/prosjekt och alla dess underaktiviteter?');
            if (!ok) return;
            deleteActivity(act.id);
            closePopup();
            return;
        }

        // Regular activity: if a specific segment is targeted, remove that segment;
        // if it was the last segment, remove the whole activity
        const segIdxRaw = popup?.dataset?.segmentIndex;
        const segIdx = segIdxRaw != null ? parseInt(segIdxRaw, 10) : -1;
        if (Array.isArray(act.segments) && act.segments.length > 0 && segIdx >= 0 && segIdx < act.segments.length) {
            // Remove only this segment
            act.segments.splice(segIdx, 1);
            if (act.segments.length === 0) {
                // No segments left -> delete activity
                deleteActivity(act.id);
            } else {
                // Recompute activity and groups
                updateActivityStartEndFromSegments(act);
                act.totalHours = calculateTotalHours(act);
                if (act.parent) updateGroupAndAncestors(act.parent);
                pushHistory();
                renderAndPreserveScroll();
            }
            closePopup();
            return;
        }

        // Fallback: delete whole activity
        const ok = confirm('Radera denna aktivitet?');
        if (!ok) return;
        deleteActivity(act.id);
        closePopup();
    } catch (err) {
        if (window.developerMode) console.error('[deleteFromPopup]', err);
    }
}

function splitFromPopup() {
    try {
        const popup = document.getElementById('activityPopup');
        const actId = selectedActivityId;
        if (!actId || !popup) return;
        const act = findActivityById(actId);
        if (!act || act.isGroup) return;
        const splitStr = popup.dataset.splitDate;
        if (!splitStr) {
            alert('Klicka på en dag i gridden för att välja delningsdatum.');
            return;
        }
        const splitDate = normalizeDate(splitStr);
        const segIdxRaw = popup?.dataset?.segmentIndex;
        const segIdx = segIdxRaw != null ? parseInt(segIdxRaw, 10) : undefined;
        splitActivity(act.id, splitDate, segIdx);
        closePopup();
    } catch (err) {
        if (window.developerMode) console.error('[splitFromPopup]', err);
    }
}

// === Drag/Resize handlers for bars ===
function startDrag(e, bar) {
    try {
        e.preventDefault();
        e.stopPropagation();
        const actId = bar?.dataset?.id;
        const segIdx = bar?.dataset?.segmentIndex != null ? parseInt(bar.dataset.segmentIndex, 10) : null;
        const act = findActivityById(actId);
        if (!act || segIdx == null || !Array.isArray(act.segments)) return;

        // Build targets: if multiple segments are selected, move all selected; otherwise move just this one
        const targets = [];
        if (selectedSegmentKeys.size > 0 && selectedSegmentKeys.has(`${actId}:${segIdx}`)) {
            selectedSegmentKeys.forEach(key => {
                const [aId, idxStr] = key.split(':');
                const a = findActivityById(aId);
                const i = parseInt(idxStr, 10);
                if (a && Array.isArray(a.segments) && a.segments[i]) {
                    const previewBar = (a.id === actId && i === segIdx) ? bar : document.querySelector(`.bar[data-id="${CSS.escape(a.id)}"][data-segment-index="${i}"]`);
                    targets.push({
                        act: a,
                        idx: i,
                        origStart: new Date(a.segments[i].start),
                        origEnd: new Date(a.segments[i].end),
                        previewWrapper: previewBar?.parentElement || null,
                        previewOriginalTransform: previewBar?.parentElement?.style.transform || ''
                    });
                }
            });
        } else {
            const seg = act.segments[segIdx];
            targets.push({
                act,
                idx: segIdx,
                origStart: new Date(seg.start),
                origEnd: new Date(seg.end),
                previewWrapper: bar?.parentElement || null,
                previewOriginalTransform: bar?.parentElement?.style.transform || ''
            });
        }

        const dayWidth = interaction.dayWidth || 35;
        interaction.type = 'drag';
        interaction.element = bar;
        interaction.activityId = actId;
        interaction.startX = e.clientX;
        interaction.originalStart = null;
        interaction.originalEnd = null;
        interaction.side = null;
        interaction._targets = targets;
        interaction._lastDelta = 0;

        const addDaysPreserveTime = (date, days) => {
            const d = new Date(date);
            d.setDate(d.getDate() + (days || 0));
            return d;
        };

        const onMove = (ev) => {
            if (interaction.type !== 'drag') return;
            const dx = (ev.clientX - interaction.startX) || 0;
            const deltaDays = Math.round(dx / dayWidth);
            if (deltaDays === interaction._lastDelta) return;
            interaction._lastDelta = deltaDays;
            const label = deltaDays === 0 ? 'Flytta: 0 d' : `Flytta: ${deltaDays > 0 ? '+' : ''}${deltaDays} d`;
            showInteractionTooltip(label, ev.clientX, ev.clientY);

            // Live update: shift bar wrappers with transform to avoid grid reflow on every mousemove
            try {
                const shiftPx = deltaDays * dayWidth;
                targets.forEach(t => {
                    if (!t.previewWrapper) return;
                    t.previewWrapper.style.transform = shiftPx ? `translateX(${shiftPx}px)` : t.previewOriginalTransform;
                });
            } catch(_) {}
        };

        const onUp = () => {
            window.removeEventListener('mousemove', onMove);
            window.removeEventListener('mouseup', onUp);
            const deltaDays = interaction._lastDelta || 0;
            hideInteractionTooltip();
            targets.forEach(t => {
                if (t.previewWrapper) t.previewWrapper.style.transform = t.previewOriginalTransform;
            });
            if (interaction.type !== 'drag') { interaction._targets = null; interaction.type = null; return; }
            if (deltaDays !== 0) {
                const seenParents = new Set();
                targets.forEach(t => {
                    const seg = t.act.segments[t.idx];
                    if (!seg) return;
                    seg.start = addDaysPreserveTime(t.origStart, deltaDays);
                    seg.end = addDaysPreserveTime(t.origEnd, deltaDays);
                    updateActivityStartEndFromSegments(t.act);
                    t.act.totalHours = calculateTotalHours(t.act);
                    if (t.act.parent) seenParents.add(t.act.parent);
                });
                seenParents.forEach(pid => updateGroupAndAncestors(pid));
                pushHistory();
                renderAndPreserveScroll();
            }
            interaction._targets = null;
            interaction.type = null;
        };

        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
    } catch (err) {
        if (window.developerMode) console.error('[startDrag]', err);
    }
}

function startResize(e, bar, side) {
    try {
        e.preventDefault();
        e.stopPropagation();
        const actId = bar?.dataset?.id;
        const segIdx = bar?.dataset?.segmentIndex != null ? parseInt(bar.dataset.segmentIndex, 10) : null;
        const act = findActivityById(actId);
        if (!act || segIdx == null || !Array.isArray(act.segments)) return;
        const seg = act.segments[segIdx];
        if (!seg) return;

        const dayWidth = interaction.dayWidth || 35;
        interaction.type = 'resize';
        interaction.element = bar;
        interaction.activityId = actId;
        interaction.startX = e.clientX;
        interaction.originalStart = new Date(seg.start);
        interaction.originalEnd = new Date(seg.end);
        interaction.side = side === 'left' ? 'left' : 'right';
        interaction._lastDelta = 0;

        const addDaysPreserveTime = (date, days) => {
            const d = new Date(date);
            d.setDate(d.getDate() + (days || 0));
            return d;
        };

        const baseLen = getWorkDays(interaction.originalStart, interaction.originalEnd);
        const originalBarTransform = bar.style.transform || '';
        const originalTransformOrigin = bar.style.transformOrigin || '';

        const onMove = (ev) => {
            if (interaction.type !== 'resize') return;
            const dx = (ev.clientX - interaction.startX) || 0;
            const deltaDays = Math.round(dx / dayWidth);
            if (deltaDays === interaction._lastDelta) return;
            interaction._lastDelta = deltaDays;
            const newLen = Math.max(1, baseLen + (interaction.side === 'right' ? deltaDays : -deltaDays));
            const label = `Längd: ${newLen} d` + (deltaDays ? ` • Δ${deltaDays > 0 ? '+' : ''}${deltaDays} d` : '');
            showInteractionTooltip(label, ev.clientX, ev.clientY);

            // Live update with transform scaling to avoid expensive grid layout changes during resize
            try {
                const scaleX = Math.max(1 / Math.max(baseLen, 1), newLen / Math.max(baseLen, 1));
                bar.style.transformOrigin = interaction.side === 'right' ? 'left center' : 'right center';
                bar.style.transform = scaleX === 1 ? originalBarTransform : `scaleX(${scaleX})`;
            } catch(_) {}
        };

        const onUp = () => {
            window.removeEventListener('mousemove', onMove);
            window.removeEventListener('mouseup', onUp);
            const deltaDays = interaction._lastDelta || 0;
            hideInteractionTooltip();
            bar.style.transform = originalBarTransform;
            bar.style.transformOrigin = originalTransformOrigin;
            if (interaction.type !== 'resize') { interaction.type = null; return; }
            if (deltaDays !== 0) {
                let newStart = new Date(interaction.originalStart);
                let newEnd = new Date(interaction.originalEnd);
                if (interaction.side === 'left') {
                    newStart = addDaysPreserveTime(interaction.originalStart, deltaDays);
                    // ensure at least 1 day
                    if (normalizeDate(newStart) > normalizeDate(newEnd)) newStart = new Date(newEnd);
                } else {
                    newEnd = addDaysPreserveTime(interaction.originalEnd, deltaDays);
                    if (normalizeDate(newEnd) < normalizeDate(newStart)) newEnd = new Date(newStart);
                }
                seg.start = newStart;
                seg.end = newEnd;
                updateActivityStartEndFromSegments(act);
                act.totalHours = calculateTotalHours(act);
                if (act.parent) updateGroupAndAncestors(act.parent);
                pushHistory();
                renderAndPreserveScroll();
            }
            interaction.type = null;
        };

        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
    } catch (err) {
        if (window.developerMode) console.error('[startResize]', err);
    }
}