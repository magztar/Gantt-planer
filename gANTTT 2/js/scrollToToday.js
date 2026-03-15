// Skrollar Gantt-vyn så att dagens datum är synlig och placerad längst till höger
function scrollToToday() {
    const ganttContainer = document.querySelector('.gantt-container');
    if (!ganttContainer) return;

    // Hämta start- och slutdatum från inputs
    const startInput = document.getElementById('startDate');
    const endInput = document.getElementById('endDate');
    if (!startInput || !endInput) return;
    const startDate = new Date(startInput.value);
    const endDate = new Date(endInput.value);
    const today = new Date();
    today.setHours(0,0,0,0);

    // Räkna ut index för dagens datum
    const totalDays = Math.floor((endDate - startDate) / (1000*60*60*24)) + 1;
    const todayIndex = Math.floor((today - startDate) / (1000*60*60*24));
    if (todayIndex < 0 || todayIndex >= totalDays) return; // Dagens datum utanför vyn

    // Hämta kolumnbredd (default 35px, kan vara dynamisk)
    let dayWidth = 35;
    if (window.interaction && window.interaction.dayWidth) dayWidth = window.interaction.dayWidth;

    // Skrolla så att dagens kolumn är längst till höger
    const scrollPos = (todayIndex + 1) * dayWidth - ganttContainer.clientWidth;
    ganttContainer.scrollLeft = Math.max(0, scrollPos);
}

window.scrollToToday = scrollToToday;
