/**
 * theme.js
 * Shared across all pages — handles light/dark mode toggle.
 * Uses Lucide icons (must load lucide before this runs, or after DOMContentLoaded).
 */

(function () {
  const STORAGE_KEY = "purchasing-tool-theme";

  function getPreference() {
    return localStorage.getItem(STORAGE_KEY) || "dark";
  }

  function applyTheme(theme) {
    if (theme === "light") {
      document.body.classList.add("light");
    } else {
      document.body.classList.remove("light");
    }
  }

  function updateIcon(theme) {
    const icon = document.getElementById("theme-icon");
    if (!icon) return;
    // Swap the lucide icon name and re-render
    icon.setAttribute("data-lucide", theme === "light" ? "moon" : "sun");
    if (window.lucide) lucide.createIcons();
  }

  function toggleTheme() {
    const current = getPreference();
    const next = current === "dark" ? "light" : "dark";
    localStorage.setItem(STORAGE_KEY, next);
    applyTheme(next);
    updateIcon(next);
  }

  // Apply saved theme immediately to avoid flash
  applyTheme(getPreference());

  document.addEventListener("DOMContentLoaded", () => {
    const btn = document.getElementById("theme-toggle");
    if (btn) btn.addEventListener("click", toggleTheme);
    updateIcon(getPreference());
  });
})();
