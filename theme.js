// Apply before the page paints, independently of OCR settings and API access.
(() => {
    const storageKey = 'register_scanner_theme';
    const systemTheme = window.matchMedia('(prefers-color-scheme: dark)');
    let preference = 'dark';
    try {
        const saved = localStorage.getItem(storageKey);
        if (['light', 'dark', 'system'].includes(saved)) preference = saved;
    } catch { /* Theme selection still works when storage is unavailable. */ }

    function applyTheme() {
        const theme = preference === 'system' ? (systemTheme.matches ? 'dark' : 'light') : preference;
        document.documentElement.dataset.theme = theme;
        document.documentElement.style.colorScheme = theme;
    }

    applyTheme();
    systemTheme.addEventListener('change', applyTheme);
    document.addEventListener('DOMContentLoaded', () => {
        const select = document.getElementById('theme-select');
        select.value = preference;
        select.addEventListener('change', () => {
            preference = select.value;
            try { localStorage.setItem(storageKey, preference); } catch { /* Keep the current session theme. */ }
            applyTheme();
        });
    });
})();
