(function () {
    'use strict';

    const navButtons = document.querySelectorAll('.main-nav-btn[data-page]');
    const pages = document.querySelectorAll('.page');

    navButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const target = btn.dataset.page;
            navButtons.forEach(b => b.classList.toggle('active', b === btn));
            pages.forEach(p => {
                p.classList.toggle('active', p.id === 'page-' + target);
            });
            // Скролл к активному шагу выполняет отдельный скрипт в index.html.
            // Здесь ничего не делаем, чтобы не было скачка вверх.
        });
    });
})();