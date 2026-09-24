/* config.js — единый источник констант MyFiles */
(function () {
    'use strict';

    const DEFAULT_SECTIONS = {
        programs: {
            label: 'Программы', icon: 'fa-code',
            json: '_content/programs.json', container: 'programs-container',
            folderable: true,
            columns: ['folder','title','description','version','size','download_link']
        },
        books: {
            label: 'Книги', icon: 'fa-book',
            json: '_content/books.json', container: 'books-container',
            folderable: true,
            columns: ['cover','title','author','description','format','download_link']
        },
        movies: {
            label: 'Фильмы', icon: 'fa-film',
            json: '_content/movies.json', container: 'movies-container',
            folderable: true,
            columns: ['folder','title','year','description','download_link']
        },
        music: {
            label: 'Музыка', icon: 'fa-music',
            json: '_content/music.json', container: 'music-container',
            folderable: true,
            columns: ['title','artist','year','download_link']
        },
        games: {
            label: 'Игры', icon: 'fa-gamepad',
            json: '_content/games.json', container: 'games-container',
            folderable: true,
            columns: ['folder','title','platform','year','description','download_link']
        },
        misc: {
            label: 'Разное', icon: 'fa-ellipsis-h',
            json: '_content/misc.json', container: 'misc-container',
            folderable: false,
            columns: ['title','description','download_link']
        }
    };

    const APP_CONFIG = {
        SUPABASE_URL: 'https://rmoonebbvpmvthvpcmpt.supabase.co',
        SUPABASE_ANON_KEY: 'sb_publishable_tr2OeCNsnhEeOTw7-Y3yLw_OGCYtWHW',

        YANDEX_FOLDER_URL: 'https://disk.yandex.ru/d/zMxF4nXHPkIVCQ',

        GITHUB_OWNER: 'DimonMaxx',
        GITHUB_REPO: 'my-site',
        GITHUB_BRANCH: 'main',
        WORKFLOW_UPDATE: 'update-content.yml',
        WORKFLOW_DELETE: 'delete-content.yml',

        PAGE_SIZES: [5, 10, 15, 25, 50, 100],
        DEFAULT_PAGE_SIZE: 10,

        MAX_AVATAR_SIZE: 2 * 1024 * 1024,
        MAX_ATTACH_SIZE: 10 * 1024 * 1024,
        ALLOWED_IMAGE_TYPES: ['image/jpeg', 'image/png', 'image/webp', 'image/gif'],
        ALLOWED_ATTACH_TYPES: [
            'image/jpeg', 'image/png', 'image/webp', 'image/gif',
            'application/pdf',
            'application/zip'
        ],

        DEFAULT_SECTIONS,

        // SECTIONS — «живой» объект. Он заполняется сразу DEFAULT'ами,
        // а MF.loadSections() позже подменяет содержимое из Supabase.
        SECTIONS: Object.assign({}, DEFAULT_SECTIONS),

        COLUMN_LABELS: {
            title: 'Название', description: 'Описание', version: 'Версия',
            size: 'Размер (МБ)', download_link: 'Ссылка',
            author: 'Автор', format: 'Формат', date: 'Дата',
            body: 'Текст', year: 'Год', artist: 'Исполнитель',
            platform: 'Платформа', folder: 'Папка',
            file_name: 'Файл', username: 'Пользователь',
            downloaded_at: 'Дата и время'
        },

        SECTION_TO_SHEET: {
            programs: 'Программы', books: 'Книги', movies: 'Фильмы',
            music: 'Музыка', games: 'Игры', misc: 'Разное'
        }
    };

    // НЕ морозим APP_CONFIG — иначе SECTIONS будет невозможно обновить из Supabase.

    window.APP_CONFIG = APP_CONFIG;

    // Диагностика (можно убрать позже)
    console.log('[config.js] DEFAULT_SECTIONS keys:',
        Object.keys(DEFAULT_SECTIONS));
    console.log('[config.js] SECTIONS keys:',
        Object.keys(APP_CONFIG.SECTIONS));
})();
