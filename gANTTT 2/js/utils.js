// js/utils.js

/**
 * Innehåller en samling av pålitliga, fristående hjälpfunktioner (utilities)
 * som används på flera ställen i applikationen.
 */

const holidayCache = {};

/**
 * Normaliserar ett datum-input (sträng eller Date-objekt) till ett Date-objekt
 * som representerar början av dagen (midnatt) i lokal tidszon.
 * Detta är viktigt för att undvika tidszonsproblem vid beräkningar.
 * @param {string|Date} input - Datum att normalisera.
 * @returns {Date|null} Ett normaliserat Date-objekt eller null om input är ogiltigt.
 */
function normalizeDate(input) {
    if (!input) return null;
    let d;
    let hadTime = false;
    
    if (typeof input === 'string') {
        // Stöd både 'YYYY-MM-DD' och 'YYYY-MM-DDTHH:MM'
        if (input.includes('T')) {
            const [datePart, timePart] = input.split('T');
            const [y, m, d2] = datePart.split('-').map(Number);
            const [hh, mm] = timePart.split(':').map(Number);
            d = new Date(y, (m || 1) - 1, d2 || 1, hh || 0, mm || 0, 0, 0);
            hadTime = true;
        } else {
            const parts = input.split('-');
            if (parts.length < 3) return null;
            const [year, month, day] = parts.map(Number);
            d = new Date(year, month - 1, day);
        }
    } else if (input instanceof Date) {
        d = new Date(input);
        hadTime = (input.getHours() + input.getMinutes() + input.getSeconds() + input.getMilliseconds()) !== 0;
    } else {
        return null;
    }

    if (isNaN(d.getTime())) return null;
    if (!hadTime) d.setHours(0, 0, 0, 0);
    return d;
}

/**
 * Formaterar ett Date-objekt till 'YYYY-MM-DDTHH:MM' för <input type="datetime-local">.
 */
function formatDateTimeLocal(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) return '';
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    const hh = String(date.getHours()).padStart(2, '0');
    const mm = String(date.getMinutes()).padStart(2, '0');
    return `${y}-${m}-${d}T${hh}:${mm}`;
}

/**
 * Tolkar en sträng från <input type="datetime-local"> och returnerar ett Date-objekt i lokal tid.
 */
function parseDateTimeLocal(str) {
    if (!str) return null;
    const [datePart, timePart] = str.split('T');
    if (!datePart) return null;
    const [y, m, d] = datePart.split('-').map(Number);
    let hh = 0, mm = 0;
    if (timePart) {
        const tm = timePart.split(':').map(Number);
        hh = tm[0] || 0; mm = tm[1] || 0;
    }
    const dt = new Date(y, (m || 1) - 1, d || 1, hh, mm, 0, 0);
    return isNaN(dt.getTime()) ? null : dt;
}

/**
 * Returnerar true om datumet har tidsspecifik komponent (inte exakt 00:00:00.000)
 */
function isTimeSpecific(date) {
    if (!(date instanceof Date)) return false;
    return (date.getHours() + date.getMinutes() + date.getSeconds() + date.getMilliseconds()) !== 0;
}

/**
 * Formaterar ett Date-objekt till en sträng i formatet 'YYYY-MM-DD'.
 * Används för att sätta värden i <input type="date"> och för datalagring.
 * @param {Date} date - Datumobjektet som ska formateras.
 * @returns {string} Det formaterade datumet.
 */
function formatDateLocal(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) return '';
    return date.getFullYear() + '-' +
           String(date.getMonth() + 1).padStart(2, '0') + '-' +
           String(date.getDate()).padStart(2, '0');
}

/**
 * NYTT: Beräknar veckonumret för ett givet datum enligt ISO 8601-standarden.
 * @param {Date} date - Datumet att beräkna veckonummer för.
 * @returns {number} Veckonumret.
 */
function getWeekNumber(date) {
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    // Sätt till närmaste torsdag: nuvarande dag + 4 - veckodag
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
    // Årets första dag
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    // Beräkna antal dagar och dela med 7
    const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    return weekNo;
}


/**
 * Beräknar antal kalenderdagar mellan två datum (inklusive start- och slutdag).
 * @param {Date|string} startDate - Startdatum.
 * @param {Date|string} endDate - Slutdatum.
 * @returns {number} Antal dagar.
 */
function getWorkDays(startDate, endDate) {
    const start = normalizeDate(startDate);
    const end = normalizeDate(endDate);
    if (!start || !end || end < start) return 0;
    // Ignorera klockslag; räkna hela kalenderdagar (inklusive båda ändar)
    const s = new Date(start.getFullYear(), start.getMonth(), start.getDate());
    const e = new Date(end.getFullYear(), end.getMonth(), end.getDate());
    return Math.round((e.getTime() - s.getTime()) / (1000 * 60 * 60 * 24)) + 1;
}

/**
 * Beräknar antal arbetsdagar (mån–fre) mellan två datum, inkl. båda ändar.
 * Helger (lör/sön) räknas inte. Helgdagar hanteras inte här (kräver async fetch).
 * @param {Date|string} startDate
 * @param {Date|string} endDate
 * @returns {number}
 */
function getWorkingWeekdays(startDate, endDate) {
    const start = normalizeDate(startDate);
    const end = normalizeDate(endDate);
    if (!start || !end || end < start) return 0;
    let count = 0;
    for (let d = new Date(start.getFullYear(), start.getMonth(), start.getDate()); d <= end; d.setDate(d.getDate() + 1)) {
        const day = d.getDay();
        if (day !== 0 && day !== 6) count++;
    }
    return count;
}

/**
 * Ska helger räknas i timsummeringar? Styrs av UI-checkbox.
 * @returns {boolean}
 */
function shouldCountWeekends() {
    try { return !!document.getElementById('countWeekends')?.checked; } catch(_) { return false; }
}

/**
 * Ger effektivt antal dagar för timsummering baserat på inställning:
 * - Om helger ska räknas: använd kalenderdagar (getWorkDays)
 * - Annars: använd vardagar (getWorkingWeekdays)
 */
function getEffectiveDays(startDate, endDate) {
    return shouldCountWeekends() ? getWorkDays(startDate, endDate) : getWorkingWeekdays(startDate, endDate);
}

/**
 * Beräknar totala antalet timmar för en aktivitet baserat på dess period och timmar aktivitet.
 * @param {object} act - Aktivitetsobjektet.
 * @returns {number} Totalt antal beräknade timmar.
 */
function calculateTotalHours(act) {
    if (!act || act.isGroup || !act.segments || act.segments.length === 0) return 0;

    const total = act.segments.reduce((sum, segment) => {
        const hpd = (typeof segment.hoursPerDay === 'number') ? segment.hoursPerDay : (act.hoursPerDay ?? getStandardHoursPerDay());
        // Använd effektiv dagräkning beroende på om helger ska räknas eller inte
        const days = getEffectiveDays(segment.start, segment.end);
        return sum + (hpd * days);
    }, 0);

    return total;
}

/**
 * Hämtar standardantalet "timmar aktivitet" per dag baserat på valt land i UI.
 * @returns {number} Antal timmar (8 för SE, 7.5 för NO).
 */
function getStandardHoursPerDay() {
    const country = document.getElementById('dateFormat')?.value || "SE";
    return country === "NO" ? 7.5 : 8;
}

/**
 * Hämtar helgdagar för ett givet land och tidsspann.
 * Använder först en statisk lista som fallback och försöker sedan hämta från ett externt API.
 * Resultatet cachas för att undvika onödiga nätverksanrop.
 * @param {string} country - Landskod ('SE' eller 'NO').
 * @param {number} fromYear - Startår.
 * @param {number} toYear - Slutår.
 * @returns {Promise<Map<string,string>>} En promise som resolverar med en Map date->localName.
 */
async function getHolidays(country, fromYear, toYear) {
    const cacheKey = `${country}-${fromYear}-${toYear}`;
    if (holidayCache[cacheKey]) {
        return holidayCache[cacheKey];
    }

    const holidays = new Map();
    const STATIC_HOLIDAYS = {
        SE: {
            '2024': {
                '2024-01-01': 'Nyårsdagen',
                '2024-01-06': 'Trettondedag jul',
                '2024-03-29': 'Långfredagen',
                '2024-03-31': 'Påskdagen',
                '2024-04-01': 'Annandag påsk',
                '2024-05-01': 'Första maj',
                '2024-05-09': 'Kristi himmelsfärdsdag',
                '2024-05-19': 'Pingstdagen',
                '2024-06-06': 'Sveriges nationaldag',
                '2024-06-22': 'Midsommardagen',
                '2024-11-02': 'Alla helgons dag',
                '2024-12-25': 'Juldagen',
                '2024-12-26': 'Annandag jul'
            },
            '2025': {
                '2025-01-01': 'Nyårsdagen',
                '2025-01-06': 'Trettondedag jul',
                '2025-04-18': 'Långfredagen',
                '2025-04-20': 'Påskdagen',
                '2025-04-21': 'Annandag påsk',
                '2025-05-01': 'Första maj',
                '2025-05-29': 'Kristi himmelsfärdsdag',
                '2025-06-06': 'Sveriges nationaldag',
                '2025-06-08': 'Pingstdagen',
                '2025-06-21': 'Midsommardagen',
                '2025-11-01': 'Alla helgons dag',
                '2025-12-25': 'Juldagen',
                '2025-12-26': 'Annandag jul'
            }
        },
        NO: {
            '2024': {
                '2024-01-01': 'Første nyttårsdag',
                '2024-03-28': 'Skjærtorsdag',
                '2024-03-29': 'Langfredag',
                '2024-03-31': 'Første påskedag',
                '2024-04-01': 'Andre påskedag',
                '2024-05-01': 'Arbeidernes dag',
                '2024-05-09': 'Kristi himmelfartsdag',
                '2024-05-17': 'Grunnlovsdagen',
                '2024-05-19': 'Første pinsedag',
                '2024-05-20': 'Andre pinsedag',
                '2024-12-25': 'Første juledag',
                '2024-12-26': 'Andre juledag'
            },
            '2025': {
                '2025-01-01': 'Første nyttårsdag',
                '2025-04-17': 'Skjærtorsdag',
                '2025-04-18': 'Langfredag',
                '2025-04-20': 'Første påskedag',
                '2025-04-21': 'Andre påskedag',
                '2025-05-01': 'Arbeidernes dag',
                '2025-05-17': 'Grunnlovsdagen',
                '2025-05-29': 'Kristi himmelfartsdag',
                '2025-06-05': 'Helgedag',
                '2025-06-08': 'Første pinsedag',
                '2025-06-09': 'Andre pinsedag',
                '2025-12-25': 'Første juledag',
                '2025-12-26': 'Andre juledag'
            }
        }
    };

    let yearsToFetch = [];
    for (let year = fromYear; year <= toYear; year++) {
        if (STATIC_HOLIDAYS[country] && STATIC_HOLIDAYS[country][year]) {
            const entries = STATIC_HOLIDAYS[country][year];
            Object.keys(entries).forEach(dateStr => holidays.set(dateStr, entries[dateStr] || 'Helgdag'));
        } else {
            yearsToFetch.push(year);
        }
    }

    if (yearsToFetch.length > 0) {
        try {
            const fetchPromises = yearsToFetch.map(year =>
                fetch(`https://date.nager.at/api/v3/PublicHolidays/${year}/${country}`)
                .then(res => res.ok ? res.json() : Promise.reject('API fetch failed'))
            );
            const results = await Promise.all(fetchPromises);
            results.flat().forEach(h => holidays.set(h.date, h.localName || h.name || 'Helgdag'));
        } catch (error) {
            if (window.developerMode) console.warn("Kunde inte hämta helgdagar från API, använder endast statisk lista.", error);
        }
    }

    holidayCache[cacheKey] = holidays;
    return holidays;
}

/**
 * Genererar ett unikt ID för nya aktiviteter.
 * @returns {string} Ett unikt ID.
 */
function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substring(2, 9);
}

/**
 * Lägger till ett antal kalenderdagar (kan vara negativt) till ett datum och returnerar NYTT Date-objekt.
 * Muterar inte originalet.
 * @param {Date} date
 * @param {number} days
 * @returns {Date}
 */
function addDays(date, days) {
    if (!date) return null;
    const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    d.setDate(d.getDate() + (days||0));
    return d;
}

/**
 * Visar ett centrerat modalfönster med specificerat HTML-innehåll.
 * @param {string} contentHTML - HTML-innehållet som ska visas i modalen.
 */
function showCenteredModal(contentHTML) {
    const modalContent = document.getElementById('modalContent');
    const modalBackdrop = document.getElementById('modalBackdrop');
    const centeredModal = document.getElementById('centeredModal');
    
    if(modalContent) modalContent.innerHTML = contentHTML;
    if(modalBackdrop) modalBackdrop.style.display = 'block';
    if(centeredModal) centeredModal.style.display = 'block';
}

/**
 * Stänger det centrerade modalfönstret.
 */
function closeCenteredModal() {
    document.getElementById('modalBackdrop').style.display = 'none';
    document.getElementById('centeredModal').style.display = 'none';
    document.getElementById('modalContent').innerHTML = '';
}