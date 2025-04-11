/* --- ▼▼▼ 既存のCSSに追加 ▼▼▼ --- */

/* 印刷用スタイル */
@media print {
    body {
        background-color: #fff; /* 背景を白に */
        color: #000;
        font-size: 10pt; /* 印刷時のフォントサイズ調整 */
        padding-bottom: 0; /* フッター分パディング解除 */
    }

    /* 印刷時に非表示にする要素 */
    header,
    #data-input-section,
    .tab-nav,
    main > section:not(#print-report-content), /* 印刷用セクション以外 */
    footer,
    .controls-panel,
    .pagination-controls,
    #loading-indicator,
    #ai-advisor, /* AIセクションも非表示 */
    .settings-card, /* 設定カード非表示 */
    .charts-area, /* グラフエリア非表示 (テーブルで代替) */
    .alert-lists, /* 注目リスト非表示 (レポート内で表示) */
    #printReportButton, /* 印刷ボタン自体も非表示 */
    .tab-button, /* タブボタン非表示 */
    #resetFilters, /* リセットボタン非表示 */
    #itemsPerPageSelect, /* 件数選択非表示 */
    #searchInput, #categoryFilter, #makerFilter, #dangerRankFilter /* フィルター非表示 */
     {
        display: none !important;
    }

    /* 印刷用コンテンツエリアを表示 */
    main {
        display: block !important;
        padding: 0 !important;
        max-width: 100% !important;
        margin: 0 !important;
    }
    #print-report-content {
        display: block !important;
        box-shadow: none !important;
        border: none !important;
        padding: 1cm; /* 印刷余白 */
        margin: 0;
    }

    /* 印刷用テーブルスタイル */
    .print-table {
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 1.5em;
        font-size: 9pt;
    }
    .print-table th,
    .print-table td {
        border: 1px solid #ccc;
        padding: 0.4em 0.6em;
        text-align: left;
    }
    .print-table th {
        background-color: #eee;
        font-weight: bold;
    }
    .print-table caption {
        font-weight: bold;
        margin-bottom: 0.5em;
        text-align: left;
    }

    /* ページ区切り調整 (できる範囲で) */
    h2, h3 {
        page-break-after: avoid;
    }
    table {
        page-break-inside: auto;
    }
    tr {
        page-break-inside: avoid;
    }
    thead {
        display: table-header-group; /* 各ページでヘッダーを繰り返す */
    }

    /* リンクはテキストで表示 */
    a {
        text-decoration: none;
        color: #000;
    }
    a[href]:after {
        /* content: " (" attr(href) ")"; */ /* URL表示は不要ならコメントアウト */
    }
}

/* --- ▲▲▲ 既存のCSSに追加 ▲▲▲ --- */
