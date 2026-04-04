(function () {
    const STORAGE_KEY = 'cheatcode-theme';
    const root = document.documentElement;
    let toggles = [];

    function normalizeTheme(theme) {
        return theme === 'light' ? 'light' : 'dark';
    }

    function getCurrentTheme() {
        return normalizeTheme(root.dataset.theme);
    }

    function updateToggleLabels(theme) {
        const nextTheme = theme === 'light' ? 'dark' : 'light';
        const icon = theme === 'light' ? 'fa-moon' : 'fa-sun';
        const label = nextTheme === 'light' ? 'Light Mode' : 'Dark Mode';

        toggles.forEach((toggle) => {
            toggle.setAttribute('aria-pressed', String(theme === 'light'));
            toggle.setAttribute('title', 'Switch to ' + label);

            const iconNode = toggle.querySelector('.theme-toggle-icon');
            if (iconNode) {
                iconNode.className = 'theme-toggle-icon fa-solid ' + icon;
            }

            const labelNode = toggle.querySelector('.theme-toggle-label');
            if (labelNode) {
                labelNode.textContent = label;
            }
        });
    }

    function applyTheme(theme, persist) {
        const normalized = normalizeTheme(theme);
        root.dataset.theme = normalized;
        root.style.colorScheme = normalized;

        if (persist) {
            window.localStorage.setItem(STORAGE_KEY, normalized);
        }

        updateToggleLabels(normalized);
    }

    document.addEventListener('DOMContentLoaded', function () {
        toggles = Array.from(document.querySelectorAll('[data-theme-toggle]'));

        toggles.forEach((toggle) => {
            toggle.addEventListener('click', function () {
                const nextTheme = getCurrentTheme() === 'light' ? 'dark' : 'light';
                applyTheme(nextTheme, true);
            });
        });

        applyTheme(getCurrentTheme(), false);
    });
})();
