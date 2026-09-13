/* config.js — единый источник констант MyFiles */
(function () {
    'use strict';

    const APP_CONFIG = Object.freeze({
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

        SECTIONS: {
            news:     { label: 'Новости',    icon: 'fa-newspaper',  json: '_content/news.json',     container: 'news-container',     columns: ['title', 'date', 'body'] },
            programs: { label: 'Программы',  icon: 'fa-code',       json: '_content/programs.json', container: 'programs-container', columns: ['title', 'description', 'version', 'size', 'download_link'] },
            books:    { label: 'Книги',      icon: 'fa-book',       json: '_content/books.json',    container: 'books-container',    columns: ['title', 'author', 'description', 'format', 'download_link'] },
            articles: { label: 'Статьи',     icon: 'fa-pen-fancy',  json: '_content/articles.json', container: 'articles-container', columns: ['title', 'date', 'body'] },
            movies:   { label: 'Фильмы',     icon: 'fa-film',       json: '_content/movies.json',   container: 'movies-container',   columns: ['title', 'year', 'description', 'download_link'] },
            music:    { label: 'Музыка',     icon: 'fa-music',      json: '_content/music.json',    container: 'music-container',    columns: ['title', 'artist', 'year', 'description', 'download_link'] },
            games:    { label: 'Игры',       icon: 'fa-gamepad',    json: '_content/games.json',    container: 'games-container',    columns: ['title', 'platform', 'year', 'description', 'download_link'] },
            misc:     { label: 'Разное',     icon: 'fa-ellipsis-h', json: '_content/misc.json',     container: 'misc-container',     columns: ['title', 'description', 'download_link'] }
        },

        COLUMN_LABELS: {
            title: 'Название',
            description: 'Описание',
            version: 'Версия',
            size: 'Размер (МБ)',
            download_link: 'Ссылка',
            author: 'Автор',
            format: 'Формат',
            date: 'Дата',
            body: 'Текст',
            year: 'Год',
            artist: 'Исполнитель',
            platform: 'Платформа',
            file_name: 'Файл',
            username: 'Пользователь',
            downloaded_at: 'Дата и время'
        },

        SECTION_TO_SHEET: {
            programs: 'Программы',
            books:    'Книги',
            news:     'Новости',
            articles: 'Статьи',
            movies:   'Фильмы',
            music:    'Музыка',
            games:    'Игры',
            misc:     'Разное'
        }
    });

    window.APP_CONFIG = APP_CONFIG;
})();
