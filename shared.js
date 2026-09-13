/* shared.js — общие утилиты MyFiles */
(function () {
    'use strict';

    /* ============================================================
     * Экранирование
     * ============================================================ */
    function escapeHtml(str) {
        if (str === null || str === undefined) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function escapeAttr(str) {
        if (str === null || str === undefined) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    /* ============================================================
     * Slug
     * ============================================================ */
    function slugify(text) {
        if (!text) return '';
        let slug = text.replace(/[^a-zA-Z0-9а-яА-ЯёЁ\s\-]/g, '').trim().toLowerCase();
        return slug.replace(/[\s\-]+/g, '-');
    }

    /* ============================================================
     * Пагинация
     * ============================================================ */
    function getMaxPageButtons() {
        const w = window.innerWidth;
        if (w < 480) return 5;
        if (w < 768) return 7;
        return 11;
    }

    function getVisiblePages(currentPage, totalPages) {
        const MAX = getMaxPageButtons();
        const pages = [];
        if (totalPages <= MAX) {
            for (let i = 1; i <= totalPages; i++) pages.push(i);
            return pages;
        }
        const half = Math.floor(MAX / 2);
        let start = currentPage - half;
        let end = start + MAX - 1;
        if (start < 1) { start = 1; end = MAX; }
        if (end > totalPages) { end = totalPages; start = end - MAX + 1; }
        for (let i = start; i <= end; i++) pages.push(i);
        return pages;
    }

    function renderPagination(containerId, totalItems, totalPages, state) {
        const PAGE_SIZES = window.APP_CONFIG.PAGE_SIZES;
        const sizes = PAGE_SIZES.map(sz =>
            `<option value="${sz}"${state.pageSize === sz ? ' selected' : ''}>${sz}</option>`
        ).join('');
        const visible = getVisiblePages(state.currentPage, totalPages);
        const pageNums = visible.map(p =>
            `<button class="page-num${p === state.currentPage ? ' active' : ''}" data-page="${p}">${p}</button>`
        ).join('');
        return `
            <div class="pagination-bar" data-pagination-for="${escapeAttr(containerId)}">
                <div class="pagination-total">Всего: <strong>${totalItems}</strong> • Стр. <strong>${state.currentPage}</strong> из <strong>${totalPages}</strong></div>
                <div class="pagination-controls">
                    <label>Показывать по:
                        <select class="page-size-select" data-container="${escapeAttr(containerId)}">
                            ${sizes}
                        </select>
                    </label>
                    <button class="page-btn prev-btn" data-container="${escapeAttr(containerId)}" ${state.currentPage <= 1 ? 'disabled' : ''} title="Назад">
                        <i class="fas fa-chevron-left"></i>
                    </button>
                    <div class="page-numbers" data-container="${escapeAttr(containerId)}">
                        ${pageNums}
                    </div>
                    <button class="page-btn next-btn" data-container="${escapeAttr(containerId)}" ${state.currentPage >= totalPages ? 'disabled' : ''} title="Вперёд">
                        <i class="fas fa-chevron-right"></i>
                    </button>
                    <div class="page-jump">
                        <label>Стр.:</label>
                        <input type="number" class="page-jump-input" data-container="${escapeAttr(containerId)}" min="1" max="${totalPages}" placeholder="#">
                        <button class="page-jump-btn" data-container="${escapeAttr(containerId)}">Перейти</button>
                    </div>
                </div>
            </div>
        `;
    }

    function attachPaginationHandlers(container, containerId, state, onRender) {
        if (!state || !container) return;

        const sizeSelect = container.querySelector(`.page-size-select[data-container="${containerId}"]`);
        if (sizeSelect) sizeSelect.addEventListener('change', function () {
            state.pageSize = parseInt(this.value, 10);
            state.currentPage = 1;
            onRender();
        });

        const prevBtn = container.querySelector(`.prev-btn[data-container="${containerId}"]`);
        if (prevBtn) prevBtn.addEventListener('click', function () {
            if (state.currentPage > 1) { state.currentPage--; onRender(); }
        });

        const nextBtn = container.querySelector(`.next-btn[data-container="${containerId}"]`);
        if (nextBtn) nextBtn.addEventListener('click', function () {
            const totalPages = Math.max(1, Math.ceil(state.totalItems / state.pageSize));
            if (state.currentPage < totalPages) { state.currentPage++; onRender(); }
        });

        container.querySelectorAll(`.page-num[data-page]`).forEach(btn => {
            btn.addEventListener('click', function () {
                const p = parseInt(this.dataset.page, 10);
                if (!isNaN(p) && p !== state.currentPage) {
                    state.currentPage = p;
                    onRender();
                }
            });
        });

        const jumpBtn = container.querySelector(`.page-jump-btn[data-container="${containerId}"]`);
        const jumpInput = container.querySelector(`.page-jump-input[data-container="${containerId}"]`);
        if (jumpBtn && jumpInput) {
            const doJump = () => {
                const totalPages = Math.max(1, Math.ceil(state.totalItems / state.pageSize));
                let p = parseInt(jumpInput.value, 10);
                if (isNaN(p)) p = 1;
                if (p < 1) p = 1;
                if (p > totalPages) p = totalPages;
                state.currentPage = p;
                onRender();
            };
            jumpBtn.addEventListener('click', doJump);
            jumpInput.addEventListener('keydown', function (e) {
                if (e.key === 'Enter') { e.preventDefault(); doJump(); }
            });
        }
    }

    /* ============================================================
     * UI-хелперы
     * ============================================================ */
    function toggleDesc(btn) {
        const cell = btn.closest('td');
        if (!cell) return;
        const shortEl = cell.querySelector('.desc-short');
        const fullEl = cell.querySelector('.desc-full');
        if (!shortEl || !fullEl) return;
        if (fullEl.style.display === 'block') {
            fullEl.style.display = 'none';
            shortEl.style.display = '-webkit-box';
            btn.textContent = 'Развернуть';
        } else {
            fullEl.style.display = 'block';
            shortEl.style.display = 'none';
            btn.textContent = 'Свернуть';
        }
    }

    function triggerDownload(url) {
        const a = document.createElement('a');
        a.href = url;
        a.rel = 'noopener';
        a.style.display = 'none';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }

    function formatDateRu(date) {
        if (!date) return '—';
        const d = date instanceof Date ? date : new Date(date);
        if (isNaN(d.getTime())) return '—';
        return d.toLocaleDateString('ru-RU') + ' ' + d.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });
    }

    function debounce(fn, ms) {
        let t;
        return function (...args) {
            clearTimeout(t);
            t = setTimeout(() => fn.apply(this, args), ms);
        };
    }

    /* ============================================================
     * Supabase singleton
     * ============================================================ */
    let _supabaseClient = null;
    function getSupabaseClient() {
        if (_supabaseClient) return _supabaseClient;
        if (!window.supabase) throw new Error('Supabase SDK не загружен');
        const cfg = window.APP_CONFIG;
        if (!cfg) throw new Error('config.js не загружен');
        _supabaseClient = window.supabase.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY);
        return _supabaseClient;
    }

    /* ============================================================
     * Публикация
     * ============================================================ */
    window.MF = Object.freeze({
        escapeHtml,
        escapeAttr,
        slugify,
        getMaxPageButtons,
        getVisiblePages,
        renderPagination,
        attachPaginationHandlers,
        toggleDesc,
        triggerDownload,
        formatDateRu,
        debounce,
        getSupabaseClient
    });

    // Обратная совместимость для inline-onclick="toggleDesc(this)"
    window.toggleDesc = toggleDesc;
})();
