// js/config.js

/**
 * Detta är applikationens centrala tillstånd (state).
 * Alla viktiga data som definierar det nuvarande projektet lagras här.
 */

// Huvudarrayen som innehåller alla aktivitets- och gruppobjekt.
// Varje objekt är en "rad" i Gantt-schemat.
let activities = [];

// Ett objekt för att hålla reda på vilka grupper som är öppna eller stängda i UI.
// nyckel = aktivitets-ID, värde = boolean (true=öppen, false=stängd)
const groupOpenState = {};

// Håller ID för den aktivitet som för närvarande är vald (t.ex. för att visa popup).
let selectedActivityId = null;

// Ett objekt som håller tillståndet för en pågående användarinteraktion,
// som att dra eller ändra storlek på en aktivitetsstapel.
let interaction = {
    type: null, // 'drag' eller 'resize'
    element: null, // HTML-elementet som interageras med
    activityId: null, // ID för aktiviteten som påverkas
    startX: 0, // Musens startposition (X-koordinat)
    dayWidth: 32, // Bredden på en dag i pixlar, måste synkas med CSS
    originalStart: null, // Aktivitetens startdatum innan interaktionen började
    originalEnd: null, // Aktivitetens slutdatum innan interaktionen började
    side: null // 'left' eller 'right' för resize
};

// Ett objekt för att hantera dra-och-släpp-logik i sidopanelen.
let dragInfo = {
    sourceId: null, // ID på aktiviteten som flyttas
    targetId: null, // ID på aktiviteten som den släpps på
    position: null // 'above', 'on', eller 'below'
};

// En array som lagrar information om filer som laddats upp till AI-assistenten.
let aiFiles = [];

// En global flagga för att aktivera/avaktivera utvecklarläge för AI-panelen.
window.developerMode = false;

// Zoom-inställningar: fyra nivåer (0=liten, 1=normal, 2=stor, 3=extra stor)
let zoomLevel = 1;
// Added ultra-zoom-out level at index 0; shows months-only header
const zoomDayWidths = [12, 22, 50, 150, 300];

// Radkapacitet: hur många rader gridden reserverar. Växer vid behov, krymper inte automatiskt.
let rowCapacity = 0; // 0 = initieras vid första render baserat på vyhöjd

// Översiktsläge: visar endast grupper/projekt och grov tidslinje
let overviewMode = false;
let prevZoomLevel = null;

// Multi-selektion: valda aktivitets-IDn och valda segmentnycklar ("activityId:segmentIndex")
const selectedActivityIds = new Set();
const selectedSegmentKeys = new Set();

// Urklipp för kopiering/inklistring (persistens ej nödvändig)
let clipboard = {
    activities: [], // fullständiga aktivitetobjekt för inklistring
    segments: []    // { activityId, segment } att klistra in (segment är en kopia)
};

// Datum-ankare för inklistring med offset (sätts via Alt-klick på en cell)
let pasteAnchorDate = null;

// Urvalsankare för Windows-liknande Shift-klick intervallmarkering i sidopanelen
let selectionAnchorActivityIndex = null;

// Marquee (gummiband) markering
let marquee = {
    active: false,
    startX: 0,
    startY: 0,
    currentX: 0,
    currentY: 0,
    additive: false,
    snapshotActivities: null,
    snapshotSegments: null
};

// Flagga för att slå på/av automatisk inklistrings-clamp (flyttar kopior in i datumspannet)
let enablePasteClamp = false;
// Automatisk utökning av datumspann vid inklistring om nya aktiviteter ligger utanför
let autoExpandSpan = true;

// Snabbindex för id -> aktivitet (byggs om vid render/lastningar)
let activityIndex = new Map();
function rebuildActivityIndex() {
    try {
        const map = new Map();
        (activities || []).forEach(a => { if (a && a.id) map.set(a.id, a); });
        activityIndex = map;
    } catch (_) { activityIndex = new Map(); }
}