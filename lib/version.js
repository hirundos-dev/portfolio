// version.js 1.4.1 от 22.10.2025
document.addEventListener('DOMContentLoaded', function() {
    // Просто текст который нужно везде вставить
    const versionText = "1.6.1 от 29.01.2026";
    
    // Находит все элементы с классом .version-text и вставляет текст
    document.querySelectorAll('.version-text').forEach(el => {
        el.textContent = versionText;
    });
});