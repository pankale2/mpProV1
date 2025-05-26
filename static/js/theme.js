// static/js/theme.js
document.addEventListener('DOMContentLoaded', () => {
    const themeToggle = document.getElementById('theme-toggle');
    // The worker previously added data-theme to 'main-container', let's find it.
    // If not, fall back to document.body as a common practice.
    const themeTargetElement = document.getElementById('main-container') || document.body;

    // Function to apply the theme
    function applyTheme(theme) {
        themeTargetElement.setAttribute('data-theme', theme);
        if (themeToggle) {
            themeToggle.textContent = theme === 'dark' ? 'Light Mode' : 'Dark Mode';
        }
    }

    // Function to toggle theme
    function toggleTheme() {
        const currentTheme = themeTargetElement.getAttribute('data-theme') || 'light';
        const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
        localStorage.setItem('theme', newTheme);
        applyTheme(newTheme);
    }

    // Apply saved theme on initial load
    const savedTheme = localStorage.getItem('theme') || 'light'; // Default to light
    applyTheme(savedTheme);

    // Add event listener to the toggle button
    if (themeToggle) {
        themeToggle.addEventListener('click', toggleTheme);
    }
});
