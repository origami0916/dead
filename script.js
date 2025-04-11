document.addEventListener('DOMContentLoaded', () => {
    // --- グローバル変数定義 ---
    let rawInventoryData = []; // パース直後のフィルタ前データ
    let inventoryData = [];    // 除外フィルタ後の分析対象データ
    let excludedDrugIdentifiers = new Set(); // 除外薬品コードリスト
    let filteredData = [];
    let currentData = [];
    let originalHeaders = [];

    // ページネーション関連
    let currentPage = 1;
    let itemsPerPage = 50;

    // ソート関連
    let currentSortColumn = null;
    let currentSortProcessedKey = null;
    let isSortAscending = true;

    // 設定値
    let expiryWeightParam = 50;
    let stagnantThresholdDays = 180;
    let dashboardListCount = 5;
    let analysisDate = new Date();

    // AI関連
    let userApiKey = null;

    // Chart.js インスタンス
    let categoryChartInstance = null;
    let makerChartInstance = null;
    let rankingChartInstance = null;

    // --- DOM要素取得 ---
    const dataInputSection = document.getElementById('data-input-section');
    const inventoryDataInput = document.getElementById('inventoryDataInput');
    const excludedDrugsInput = document.getElementById('excludedDrugsInput'); // New
    const loadDataButton = document.getElementById('loadDataButton');
    const dataErrorOutput = document.getElementById('data-error-output');

    const tabNav = document.querySelector('.tab-nav');
    const mainContent = document.querySelector('main');
    const footer = document.querySelector('footer');

    const tabButtons = document.querySelectorAll('.tab-button');
    const tabContents = document.querySelectorAll('.tab-content');

    // ダッシュボード要素
    const totalItemsEl = document.getElementById('total-items');
    const totalValueEl = document.getElementById('total-value');
    const cautionItemsEl = document.getElementById('caution-items');
    const cautionValueEl = document.getElementById('caution-value');
    const stagnantItemsEl = document.getElementById('stagnant-items');
    const stagnantValueEl = document.getElementById('stagnant-value');
    const nearExpiryItemsEl = document.getElementById('near-expiry-items');
    const nearExpiryValueEl = document.getElementById('near-expiry-value');
    const expiryWeightSlider = document.getElementById('expiryWeight');
    const weightLabel = document.getElementById('weightLabel');
    const stagnantThresholdInput = document.getElementById('stagnantThreshold');
    const dashboardListCountSelect = document.getElementById('dashboardListCount');
    const analysisDateInput = document.getElementById('analysisDate');
    const recalculateButton = document.getElementById('recalculateButton');
    const printReportButton = document.getElementById('printReportButton'); // New
    const printReportContent = document.getElementById('print-report-content'); // New
    const stagnantWorstListEl = document.getElementById('stagnant-worst-list');
    const nearExpiryWorstListEl = document.getElementById('near-expiry-worst-list');
    const highValueListEl = document.getElementById('high-value-list');
    const categoryChartCanvas = document.getElementById('categoryChart');
    const makerChartCanvas = document.getElementById('makerChart');

    // 在庫一覧要素
    const searchInput = document.getElementById('searchInput');
    const categoryFilter = document.getElementById('categoryFilter');
    const makerFilter = document.getElementById('makerFilter');
    const dangerRankFilter = document.getElementById('dangerRankFilter');
    const resetFiltersButton = document.getElementById('resetFilters');
    const inventoryTable = document.getElementById('inventory-table');
    const inventoryTableHeader = inventoryTable.querySelector('thead');
    const inventoryTableBody = inventoryTable.querySelector('tbody');
    const loadingIndicator = document.getElementById('loading-indicator');
    const itemCountEl = document.getElementById('itemCount');
    const prevPageButton = document.getElementById('prevPage');
    const nextPageButton = document.getElementById('nextPage');
    const pageInfoEl = document.getElementById('pageInfo');
    const itemsPerPageSelect = document.getElementById('itemsPerPage');

    // 分析レポート要素
    const rankingTypeSelect = document.getElementById('rankingType');
    const rankingChartCanvas = document.getElementById('rankingChart');

    // AIアドバイス要素
    const apiKeyInput = document.getElementById('apiKey');
    const saveApiKeyButton = document.getElementById('saveApiKey');
    const apiKeyStatus = document.getElementById('apiKeyStatus');
    const getAiAdviceButton = document.getElementById('getAiAdvice');
    const aiLoading = document.getElementById('ai-loading');
    const aiErrorOutput = document.getElementById('ai-error-output');
    const aiAdviceOutput = document.getElementById('ai-advice-output');
    const copyAiAdviceButton = document.getElementById('copyAiAdvice');

    // --- 初期化処理 ---
    // アプリのメインロジックを初期化（データロード後に呼ばれる）
    function initializeAppLogic() {
        showLoading();
        try {
            initializeSettings();
            // ★★★ 除外フィルタリングを追加 ★★★
            filterExcludedDrugs();
            // ★★★★★★★★★★★★★★★★★★★
            preprocessData();
            renderTableHeader();
            populateFilterOptions();
            applyFiltersAndSort();
            displayDashboard();

            // UI表示切り替え
            dataInputSection.style.display = 'none';
            tabNav.style.display = 'flex';
            mainContent.style.display = 'block';
            footer.style.display = 'block';

        } catch (error) {
            console.error('アプリケーションの初期化中にエラー:', error);
            showDataError('データの処理中にエラーが発生しました。データ形式を確認してください。');
            dataInputSection.style.display = 'block';
            tabNav.style.display = 'none';
            mainContent.style.display = 'none';
            footer.style.display = 'none';
        } finally {
             hideLoading();
        }
    }

    // 設定値の初期化とUI反映
    function initializeSettings() {
        const today = new Date();
        const jstOffset = 9 * 60;
        const localOffset = today.getTimezoneOffset();
        const jstNow = new Date(today.getTime() + (jstOffset + localOffset) * 60 * 1000);
        analysisDateInput.value = jstNow.toISOString().split('T')[0];
        analysisDate = new Date(analysisDateInput.value + 'T00:00:00');

        expiryWeightSlider.value = expiryWeightParam;
        updateWeightLabel();
        stagnantThresholdInput.value = stagnantThresholdDays;
        dashboardListCountSelect.value = dashboardListCount;
    }


    // --- データパーサー (TSV) ---
    function parseTSV(tsvString) {
        const lines = tsvString.trim().split('\n');
        if (lines.length < 2) { throw new Error("データが少なすぎます。ヘッダー行とデータ行が必要です。"); }
        originalHeaders = lines[0].split('\t').map(header => header.trim());
        const expectedHeaderCount = originalHeaders.length;
        const data = [];
        for (let i = 1; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue;
            const values = line.split('\t').map(value => value.trim());
            if (values.length !== expectedHeaderCount) { console.warn(`行 ${i + 1}: 列数不一致。スキップします。`); continue; }
            const item = {};
            for (let j = 0; j < expectedHeaderCount; j++) { item[originalHeaders[j]] = values[j]; }
            data.push(item);
        }
        if (data.length === 0) { throw new Error("有効なデータ行が見つかりませんでした。"); }
        return data;
    }

    // --- 除外薬品処理 ---
    function parseExcludedDrugs(excludedString) {
        return new Set(
            excludedString.trim().split('\n')
                .map(code => code.trim())
                .filter(Boolean) // 空行を除去
        );
    }

    function filterExcludedDrugs() {
        // 除外リストに該当する薬品を除外
        // '薬品コード' カラムが存在するか確認
        const drugCodeKey = originalHeaders.find(h => h.includes('薬品コード'));
        if (!drugCodeKey) {
            console.warn("除外処理スキップ: '薬品コード' カラムが見つかりません。");
            inventoryData = [...rawInventoryData]; // 除外せず全データを使用
            return;
        }

        if (excludedDrugIdentifiers.size > 0) {
            inventoryData = rawInventoryData.filter(item =>
                !excludedDrugIdentifiers.has(item[drugCodeKey])
            );
            console.log(`${rawInventoryData.length - inventoryData.length} 件の薬品を除外しました。`);
        } else {
            inventoryData = [...rawInventoryData]; // 除外リストが空なら全データを使用
        }
    }


    // --- データ前処理 ---
    function preprocessData() {
        const currentDate = analysisDate; // 分析基準日を使用
        const requiredKeys = ['有効期限', '最終出庫', '在庫金額(税別)', '在庫数量'];
        for (const key of requiredKeys) {
            if (!originalHeaders.includes(key)) { throw new Error(`必須カラム "${key}" がデータに含まれていません。`); }
        }
        const expiryDateKey = '有効期限';
        const lastOutDateKey = '最終出庫';
        const lastInDateKey = '最終入庫';
        const stagnationColKey = '(停滞)';
        const lastInStagnationKey = originalHeaders.includes(lastInDateKey) && originalHeaders.indexOf(lastInDateKey) + 1 < originalHeaders.length && originalHeaders[originalHeaders.indexOf(lastInDateKey) + 1] === stagnationColKey ? stagnationColKey : originalHeaders.find(h => h.includes('最終入庫') && h.includes('停滞'));
        const lastOutStagnationKey = originalHeaders.includes(lastOutDateKey) && originalHeaders.indexOf(lastOutDateKey) + 1 < originalHeaders.length && originalHeaders[originalHeaders.indexOf(lastOutDateKey) + 1] === stagnationColKey ? stagnationColKey : originalHeaders.find(h => h.includes('最終出庫') && h.includes('停滞'));

        // ★★★ inventoryData を直接更新 ★★★
        inventoryData = inventoryData.map(item => {
            const expiryDate = parseDate(item[expiryDateKey]);
            const lastOutDate = parseDate(item[lastOutDateKey]);
            const lastInDate = originalHeaders.includes(lastInDateKey) ? parseDate(item[lastInDateKey]) : null;
            let stagnationDays;
            if (lastOutStagnationKey && item[lastOutStagnationKey] && !isNaN(parseInt(item[lastOutStagnationKey]))) { stagnationDays = parseInt(item[lastOutStagnationKey]); }
            else { stagnationDays = calculateStagnationDays(currentDate, lastOutDate); }
            let lastInDays;
             if (lastInStagnationKey && item[lastInStagnationKey] && !isNaN(parseInt(item[lastInStagnationKey]))) { lastInDays = parseInt(item[lastInStagnationKey]); }
             else { lastInDays = calculateStagnationDays(currentDate, lastInDate); }
            const remainingDays = calculateRemainingDays(currentDate, expiryDate);
             const valueKey = originalHeaders.find(h => h.includes('在庫金額(税別)')) || '在庫金額(税別)';
             const quantityKey = originalHeaders.find(h => h.includes('在庫数量')) || '在庫数量';
            const itemValue = parseFloat(String(item[valueKey]).replace(/,/g, '')) || 0;
            const itemQuantity = parseFloat(String(item[quantityKey]).replace(/,/g, '')) || 0;
            const dangerRank = calculateDangerRank(remainingDays, stagnationDays, expiryWeightParam);
             const stagnationDisplay = isNaN(stagnationDays) || stagnationDays > 9000 ? '---' : stagnationDays;
             const remainingDisplay = isNaN(remainingDays) ? '---' : remainingDays;
             const lastInDisplay = isNaN(lastInDays) || lastInDays > 9000 ? '---' : lastInDays;
            const processedItem = {
                ...item, '有効期限Date': expiryDate, '最終出庫日Date': lastOutDate, '最終入庫日Date': lastInDate,
                '在庫金額Numeric': itemValue, '在庫数量Numeric': itemQuantity, '滞留日数Numeric': stagnationDays,
                '残り日数Numeric': remainingDays, '最終入庫日数Numeric': lastInDays, '危険度ランク': dangerRank,
                '滞留日数表示': stagnationDisplay, '有効期限表示': expiryDate ? formatDate(expiryDate) : '---', '最終入庫停滞表示': lastInDisplay
            };
             if (lastOutStagnationKey) processedItem[lastOutStagnationKey] = stagnationDisplay;
             if (originalHeaders.includes(expiryDateKey)) processedItem[expiryDateKey] = processedItem['有効期限表示'];
             if (lastInStagnationKey) processedItem[lastInStagnationKey] = lastInDisplay;
            return processedItem;
        });
    }

    // 日付文字列をDateオブジェクトに変換
    function parseDate(dateString) {
        if (!dateString || typeof dateString !== 'string') return null;
        const trimmed = dateString.trim();
        if (/^\d{4}$/.test(trimmed) || trimmed === '-') return null;
        const formattedDateString = trimmed.replace(/-/g, '/');
        const parts = formattedDateString.split('/');
        if (parts.length === 3) {
            const year = parseInt(parts[0], 10); const month = parseInt(parts[1], 10); const day = parseInt(parts[2], 10);
            if (!isNaN(year) && !isNaN(month) && !isNaN(day) && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                 const date = new Date(Date.UTC(year, month - 1, day));
                 if (date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day) { return date; }
            }
        }
        console.warn("Invalid date format:", dateString); return null;
    }
     // 日付オブジェクトをYYYY/MM/DD形式の文字列に変換
    function formatDate(date) {
        if (!date || !(date instanceof Date) || isNaN(date.getTime())) return '---';
        const y = date.getUTCFullYear(); const m = String(date.getUTCMonth() + 1).padStart(2, '0'); const d = String(date.getUTCDate()).padStart(2, '0');
        return `${y}/${m}/${d}`;
    }
    // 滞留日数計算
    function calculateStagnationDays(currentDate, lastOutDate) {
        if (!lastOutDate || !(lastOutDate instanceof Date) || isNaN(lastOutDate.getTime())) return 9999;
        const currentDayUTC = Date.UTC(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate());
        const lastOutDayUTC = Date.UTC(lastOutDate.getUTCFullYear(), lastOutDate.getUTCMonth(), lastOutDate.getUTCDate());
        const diffTime = currentDayUTC - lastOutDayUTC;
        return Math.max(0, Math.floor(diffTime / (1000 * 60 * 60 * 24)));
    }
    // 残り日数計算
    function calculateRemainingDays(currentDate, expiryDate) {
        if (!expiryDate || !(expiryDate instanceof Date) || isNaN(expiryDate.getTime())) return NaN;
        const currentDayUTC = Date.UTC(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate());
        const expiryDayUTC = Date.UTC(expiryDate.getUTCFullYear(), expiryDate.getUTCMonth(), expiryDate.getUTCDate());
        const diffTime = expiryDayUTC - currentDayUTC;
        return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    }


    // --- 危険度ランク計算 ---
    function calculateDangerRank(remainingDays, stagnationDays, weightParam) {
        // ★★★ 変更点: 期限切れは強制的にランク10 ★★★
        if (!isNaN(remainingDays) && remainingDays < 0) { return 10; }

        let scoreExpiry = 0;
        if (isNaN(remainingDays)) { scoreExpiry = 0; }
        else if (remainingDays <= 30) { scoreExpiry = 100; } else if (remainingDays <= 90) { scoreExpiry = 75; }
        else if (remainingDays <= 180) { scoreExpiry = 50; } else if (remainingDays <= 365) { scoreExpiry = 25; }
        else { scoreExpiry = 0; }

        let scoreStagnation = 0;
         if (isNaN(stagnationDays) || stagnationDays >= 730) { scoreStagnation = 100; }
         else if (stagnationDays >= 365) { scoreStagnation = 90; } else if (stagnationDays >= 180) { scoreStagnation = 60; }
         else if (stagnationDays >= 90) { scoreStagnation = 30; } else { scoreStagnation = 0; }

        const weightExpiry = weightParam / 100; const weightStagnation = 1.0 - weightExpiry;
        const totalScore = (scoreExpiry * weightExpiry) + (scoreStagnation * weightStagnation);
        return Math.max(1, Math.ceil(totalScore / 10));
    }

    // 全データの再計算と再表示
    function recalculateAll() {
        if (rawInventoryData.length === 0) return; // 元データがない場合は何もしない
        showLoading();
        // 新しい設定値をグローバル変数に反映
        const selectedDateStr = analysisDateInput.value;
        analysisDate = selectedDateStr ? new Date(selectedDateStr + 'T00:00:00') : new Date();
        analysisDateInput.value = analysisDate.toISOString().split('T')[0];

        expiryWeightParam = parseInt(expiryWeightSlider.value, 10);
        stagnantThresholdDays = parseInt(stagnantThresholdInput.value, 10) || 180;
        dashboardListCount = parseInt(dashboardListCountSelect.value, 10) || 5;

        console.log("Recalculating with settings:", { analysisDate, expiryWeightParam, stagnantThresholdDays, dashboardListCount });

        setTimeout(() => {
            try {
                filterExcludedDrugs(); // 除外フィルタを再適用
                preprocessData(); // データ再処理
                applyFiltersAndSort();
                displayDashboard();
                if (document.getElementById('analysis-reports').classList.contains('active')) {
                    displayRankingChart();
                }
            } catch (error) {
                 console.error('再計算中にエラー:', error);
                 showDataError('再計算中にエラーが発生しました。');
            } finally {
                hideLoading();
            }
        }, 50);
    }


    // 危険度スライダーのラベル更新
    function updateWeightLabel() {
        const expiryW = expiryWeightSlider.value; const stagnationW = 100 - expiryW;
        weightLabel.textContent = `期限重視: ${expiryW}% / 滞留重視: ${stagnationW}%`;
    }


    // --- フィルタリングとソート ---
    function applyFiltersAndSort() { /* 変更なし */
        const searchTerm = searchInput.value.toLowerCase().trim(); const selectedCategory = categoryFilter.value; const selectedMaker = makerFilter.value; const selectedRank = dangerRankFilter.value;
        filteredData = inventoryData.filter(item => {
            const rankMatch = selectedRank === "" || item.危険度ランク == selectedRank;
            const categoryKey = originalHeaders.find(h => h.includes('薬品種別')) || '薬品種別'; const makerKey = originalHeaders.find(h => h.includes('メーカー')) || 'メーカー';
            const categoryMatch = selectedCategory === "" || item[categoryKey] === selectedCategory; const makerMatch = selectedMaker === "" || item[makerKey] === selectedMaker;
            const nameKey = originalHeaders.find(h => h.includes('薬品名称')) || '薬品名称'; const furiganaKey = originalHeaders.find(h => h.includes('フリガナ')) || 'フリガナ';
            const yjKey = originalHeaders.find(h => h.includes('YJコード')) || 'YJコード'; const codeKey = originalHeaders.find(h => h.includes('薬品コード')) || '薬品コード';
            const searchMatch = searchTerm === "" || item[nameKey]?.toLowerCase().includes(searchTerm) || item[furiganaKey]?.toLowerCase().includes(searchTerm) || item[makerKey]?.toLowerCase().includes(searchTerm) || item[yjKey]?.includes(searchTerm) || item[codeKey]?.includes(searchTerm);
            return rankMatch && categoryMatch && makerMatch && searchMatch;
        });
        if (currentSortProcessedKey) {
            filteredData.sort((a, b) => {
                let valA = a[currentSortProcessedKey]; let valB = b[currentSortProcessedKey];
                if (valA instanceof Date && valB instanceof Date) { valA = isNaN(valA.getTime()) ? (isSortAscending ? Infinity : -Infinity) : valA.getTime(); valB = isNaN(valB.getTime()) ? (isSortAscending ? Infinity : -Infinity) : valB.getTime(); }
                else if (typeof valA === 'number' && typeof valB === 'number') { valA = isNaN(valA) ? (isSortAscending ? Infinity : -Infinity) : valA; valB = isNaN(valB) ? (isSortAscending ? Infinity : -Infinity) : valB; }
                else { valA = String(valA ?? '').toLowerCase(); valB = String(valB ?? '').toLowerCase(); }
                if (valA < valB) return isSortAscending ? -1 : 1; if (valA > valB) return isSortAscending ? 1 : -1;
                const nameKey = originalHeaders.find(h => h.includes('薬品名称')) || '薬品名称'; return String(a[nameKey] ?? '').localeCompare(String(b[nameKey] ?? ''));
            });
        } else if (currentSortColumn) {
             filteredData.sort((a, b) => {
                 let valA = a[currentSortColumn]; let valB = b[currentSortColumn]; valA = String(valA ?? '').toLowerCase(); valB = String(valB ?? '').toLowerCase();
                 if (valA < valB) return isSortAscending ? -1 : 1; if (valA > valB) return isSortAscending ? 1 : -1;
                 const nameKey = originalHeaders.find(h => h.includes('薬品名称')) || '薬品名称'; return String(a[nameKey] ?? '').localeCompare(String(b[nameKey] ?? ''));
             });
        }
        currentPage = 1; updatePagination(); displayCurrentPageData(); updateItemCount();
    }

    // --- ページネーション ---
    function updatePagination() { /* 変更なし */
         const totalItems = filteredData.length; const totalPages = itemsPerPage === 'all' ? 1 : Math.max(1, Math.ceil(totalItems / itemsPerPage)); currentPage = Math.max(1, Math.min(currentPage, totalPages));
         pageInfoEl.textContent = `ページ ${currentPage} / ${totalPages}`; prevPageButton.disabled = currentPage === 1; nextPageButton.disabled = currentPage === totalPages;
         if (itemsPerPage === 'all') { currentData = filteredData; } else { const start = (currentPage - 1) * itemsPerPage; const end = start + itemsPerPage; currentData = filteredData.slice(start, end); }
    }
     function displayCurrentPageData() { renderTable(currentData); }
     function updateItemCount() { itemCountEl.textContent = `全 ${filteredData.length} 件`; }
     function changeItemsPerPage() { /* 変更なし */
        const selectedValue = itemsPerPageSelect.value; itemsPerPage = (selectedValue === 'all') ? 'all' : parseInt(selectedValue, 10); currentPage = 1; updatePagination(); displayCurrentPageData();
    }


    // --- テーブル描画 ---
    // テーブルヘッダー生成
    function renderTableHeader() { /* 変更なし */
        const thead = inventoryTableHeader; thead.innerHTML = ''; const tr = document.createElement('tr');
        const displayColumns = { '危険度': { sortKey: '危険度ランク', processedKey: '危険度ランク', align: 'center', isNumeric: true }, '薬品名称': { sortKey: '薬品名称', processedKey: '薬品名称' }, 'フリガナ': { sortKey: 'フリガナ', processedKey: 'フリガナ' }, '在庫数量': { sortKey: '在庫数量', processedKey: '在庫数量Numeric', align: 'right', isNumeric: true }, '単位': { sortKey: '単位', processedKey: '単位', align: 'center' }, '在庫金額(税別)': { sortKey: '在庫金額(税別)', processedKey: '在庫金額Numeric', align: 'right', isNumeric: true, isCurrency: true }, '滞留日数': { sortKey: '(停滞)', processedKey: '滞留日数Numeric', align: 'right', isNumeric: true }, '有効期限': { sortKey: '有効期限', processedKey: '有効期限Date', align: 'center' }, '薬品種別': { sortKey: '薬品種別', processedKey: '薬品種別' }, 'メーカー': { sortKey: 'メーカー', processedKey: 'メーカー' }, '卸': { sortKey: '卸', processedKey: '卸' }, '最終入庫日': { sortKey: '最終入庫', processedKey: '最終入庫日Date', align: 'center' }, '最終入庫(停滞)': { sortKey: '(停滞)', processedKey: '最終入庫日数Numeric', align: 'right', isNumeric: true }, 'YJコード': { sortKey: 'YJコード', processedKey: 'YJコード' }, '薬品コード': { sortKey: '薬品コード', processedKey: '薬品コード' }, '包装コード': { sortKey: '包装コード', processedKey: '包装コード'}, '包装形状': { sortKey: '包装形状', processedKey: '包装形状'}, '出力設定': { sortKey: '出力設定（0：出庫なし、1：強制）', processedKey: '出力設定（0：出庫なし、1：強制）', align: 'center'}, '税額': { sortKey: '税額', processedKey: '税額', align: 'right'}, };
        const orderedHeaders = [ '危険度ランク', '薬品コード', 'YJコード', '薬品種別', 'フリガナ', '薬品名称', '包装コード', '包装形状', '在庫数量', '在庫金額(税別)', '最終入庫', '(停滞)', '最終出庫', '(停滞)', 'メーカー', '卸', '単位', '出力設定（0：出庫なし、1：強制）', '税額', '有効期限' ];
        orderedHeaders.forEach((headerName, index) => {
            let columnDef = null; let displayName = headerName; let sortKeyValue = headerName;
            if (headerName === '危険度ランク') { columnDef = displayColumns['危険度']; displayName = '危険度'; sortKeyValue = '危険度ランク'; }
            else if (headerName === '(停滞)' && orderedHeaders.indexOf('(停滞)') !== index) { columnDef = displayColumns['滞留日数']; displayName = '滞留日数'; sortKeyValue = '(停滞)'; }
            else if (headerName === '(停滞)') { columnDef = displayColumns['最終入庫(停滞)']; displayName = '最終入庫(停滞)'; sortKeyValue = '(停滞)'; }
            else if (headerName === '最終入庫') { columnDef = displayColumns['最終入庫日']; displayName = '最終入庫日'; sortKeyValue = '最終入庫'; }
            else if (headerName === '最終出庫') { displayName = '最終出庫日'; columnDef = { sortKey: '最終出庫', processedKey: '最終出庫日Date', align: 'center' }; sortKeyValue = '最終出庫'; }
            else { const foundEntry = Object.entries(displayColumns).find(([dName, def]) => def.sortKey === headerName); if (foundEntry) { displayName = foundEntry[0]; columnDef = foundEntry[1]; sortKeyValue = columnDef.sortKey; } else { columnDef = { sortKey: headerName, processedKey: headerName }; displayName = headerName; sortKeyValue = headerName; } }
             if (headerName !== '危険度ランク' && !originalHeaders.includes(sortKeyValue)) { console.warn(`Header "${sortKeyValue}" not found, skipping column "${displayName}".`); return; }
             const th = document.createElement('th'); th.textContent = displayName; th.dataset.sort = sortKeyValue; th.dataset.processedKey = columnDef.processedKey; if (columnDef.align) { th.style.textAlign = columnDef.align; } tr.appendChild(th);
        });
        thead.appendChild(tr);
    }
    // テーブルボディ描画
    function renderTable(data) { /* 変更なし (滞留クラス判定に閾値を使用する点は修正済み) */
        inventoryTableBody.innerHTML = '';
        if (!data || data.length === 0) { inventoryTableBody.innerHTML = `<tr><td colspan="${inventoryTableHeader.rows[0]?.cells.length || 1}" style="text-align: center;">該当するデータがありません。</td></tr>`; return; }
        const fragment = document.createDocumentFragment(); const displaySortKeys = Array.from(inventoryTableHeader.rows[0].cells).map(th => th.dataset.sort);
        data.forEach(item => {
            const row = document.createElement('tr'); row.classList.add(`rank-${item.危険度ランク}`); if (item.残り日数Numeric <= 30 && item.残り日数Numeric >= 0) row.classList.add('near-expiry'); if (item.滞留日数Numeric >= stagnantThresholdDays && item.滞留日数Numeric < 9000) row.classList.add('stagnant');
             displaySortKeys.forEach((sortKey, index) => {
                 const th = inventoryTableHeader.rows[0].cells[index]; const cell = document.createElement('td'); let cellValue = ''; let textAlign = th.style.textAlign || 'left';
                 switch (sortKey) {
                     case '危険度ランク': cellValue = item['危険度ランク']; cell.style.fontWeight = 'bold'; break;
                     case '在庫金額(税別)': cellValue = formatCurrency(item['在庫金額Numeric']); break;
                     case '在庫数量': cellValue = item['在庫数量']; break;
                     case '(停滞)': if (th.textContent === '滞留日数') { cellValue = item['滞留日数表示']; } else if (th.textContent === '最終入庫(停滞)') { cellValue = item['最終入庫停滞表示']; } else { cellValue = item[sortKey] ?? ''; } break;
                     case '有効期限': cellValue = item['有効期限表示']; break;
                     case '最終入庫': cellValue = item['最終入庫日Date'] ? formatDate(item['最終入庫日Date']) : '---'; break;
                     case '最終出庫': cellValue = item['最終出庫日Date'] ? formatDate(item['最終出庫日Date']) : '---'; break;
                     default: cellValue = item[sortKey] ?? '';
                 }
                 cell.textContent = cellValue; cell.style.textAlign = textAlign; row.appendChild(cell);
             });
            fragment.appendChild(row);
        });
        inventoryTableBody.appendChild(fragment);
    }
     // 数値をカンマ区切り通貨形式にフォーマット
    function formatCurrency(value) { if (isNaN(value) || value === null) return '---'; return Math.round(value).toLocaleString(); }
    // --- フィルター用選択肢生成 ---
    function populateFilterOptions() { /* 変更なし */
         const categoryKey = originalHeaders.find(h => h.includes('薬品種別')) || '薬品種別'; const makerKey = originalHeaders.find(h => h.includes('メーカー')) || 'メーカー'; categoryFilter.innerHTML = '<option value="">薬品種別 (すべて)</option>'; makerFilter.innerHTML = '<option value="">メーカー (すべて)</option>'; const categories = [...new Set(inventoryData.map(item => item[categoryKey]).filter(Boolean))].sort(); const makers = [...new Set(inventoryData.map(item => item[makerKey]).filter(Boolean))].sort(); categories.forEach(cat => { const option = document.createElement('option'); option.value = cat; option.textContent = cat; categoryFilter.appendChild(option); }); makers.forEach(maker => { const option = document.createElement('option'); option.value = maker; option.textContent = maker; makerFilter.appendChild(option); });
    }
     // --- テーブルヘッダーのソート処理 ---
     function handleSort(event) { /* 変更なし */
         const th = event.target.closest('th'); if (!th || !th.dataset.sort) return; const sortKey = th.dataset.sort; const processedKey = th.dataset.processedKey; if (currentSortColumn === sortKey) { isSortAscending = !isSortAscending; } else { currentSortColumn = sortKey; currentSortProcessedKey = processedKey; isSortAscending = true; } updateSortIndicators(th); applyFiltersAndSort();
     }
     // ソートインジケータ更新
     function updateSortIndicators(activeTh) { /* 変更なし */
        inventoryTableHeader.querySelectorAll('th').forEach(th => { th.classList.remove('sort-asc', 'sort-desc'); th.removeAttribute('aria-sort'); }); if (activeTh && currentSortColumn) { const sortClass = isSortAscending ? 'sort-asc' : 'sort-desc'; activeTh.classList.add(sortClass); activeTh.setAttribute('aria-sort', isSortAscending ? 'ascending' : 'descending'); }
    }

    // --- ダッシュボード表示 ---
    function displayDashboard() { /* 変更なし (内部でグローバル設定値を使用) */
         const categoryKey = originalHeaders.find(h => h.includes('薬品種別')) || '薬品種別'; const makerKey = originalHeaders.find(h => h.includes('メーカー')) || 'メーカー'; const nameKey = originalHeaders.find(h => h.includes('薬品名称')) || '薬品名称';
        const totalItems = inventoryData.length; const totalValue = inventoryData.reduce((sum, item) => sum + item.在庫金額Numeric, 0); const cautionData = inventoryData.filter(item => item.危険度ランク >= 7); const cautionItems = cautionData.length; const cautionValue = cautionData.reduce((sum, item) => sum + item.在庫金額Numeric, 0); const stagnantData = inventoryData.filter(item => item.滞留日数Numeric >= stagnantThresholdDays && item.滞留日数Numeric < 9000); const stagnantItems = stagnantData.length; const stagnantValue = stagnantData.reduce((sum, item) => sum + item.在庫金額Numeric, 0); const nearExpiryData = inventoryData.filter(item => !isNaN(item.残り日数Numeric) && item.残り日数Numeric <= 90 && item.残り日数Numeric >= 0); const nearExpiryItems = nearExpiryData.length; const nearExpiryValue = nearExpiryData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);
        totalItemsEl.textContent = totalItems.toLocaleString(); totalValueEl.textContent = `¥${formatCurrency(totalValue)}`; cautionItemsEl.textContent = cautionItems.toLocaleString(); cautionValueEl.textContent = `¥${formatCurrency(cautionValue)}`; stagnantItemsEl.textContent = `${stagnantItems.toLocaleString()} ( ${stagnantThresholdDays}日以上 )`; stagnantValueEl.textContent = `¥${formatCurrency(stagnantValue)}`; nearExpiryItemsEl.textContent = nearExpiryItems.toLocaleString(); nearExpiryValueEl.textContent = `¥${formatCurrency(nearExpiryValue)}`;
        updateRankingList(stagnantWorstListEl, inventoryData.filter(item => item.滞留日数Numeric < 9000), '滞留日数Numeric', false, dashboardListCount, item => `${item[nameKey]} (${item.滞留日数表示}日)`); updateRankingList(nearExpiryWorstListEl, inventoryData.filter(item => !isNaN(item.残り日数Numeric) && item.残り日数Numeric >= 0), '残り日数Numeric', true, dashboardListCount, item => `${item[nameKey]} (残${item.残り日数Numeric}日)`); updateRankingList(highValueListEl, inventoryData, '在庫金額Numeric', false, dashboardListCount, item => `${item[nameKey]} (¥${formatCurrency(item.在庫金額Numeric)})`);
        displayCategoryChart(categoryKey); displayMakerChart(makerKey);
    }
    // ランキングリスト更新ヘルパー
    function updateRankingList(listElement, data, sortKey, ascending, count, displayFormatter) { /* 変更なし */
        const sortedData = [...data].sort((a, b) => { const valA = a[sortKey]; const valB = b[sortKey]; const numA = parseFloat(valA); const numB = parseFloat(valB); const effA = isNaN(numA) ? (ascending ? Infinity : -Infinity) : numA; const effB = isNaN(numB) ? (ascending ? Infinity : -Infinity) : numB; if (effA < effB) return ascending ? -1 : 1; if (effA > effB) return ascending ? 1 : -1; const nameKey = originalHeaders.find(h => h.includes('薬品名称')) || '薬品名称'; return String(a[nameKey] ?? '').localeCompare(String(b[nameKey] ?? '')); }).slice(0, count);
        listElement.innerHTML = ''; if (sortedData.length === 0) { listElement.innerHTML = '<li>該当データなし</li>'; return; } sortedData.forEach(item => { const li = document.createElement('li'); li.textContent = displayFormatter(item); listElement.appendChild(li); });
    }

    // --- グラフ描画 (Chart.js) ---
    const chartOptions = { /* 変更なし */ };
    function displayCategoryChart(categoryKey) { /* 変更なし */ }
    function displayMakerChart(makerKey) { /* 変更なし */ }
    // ランキンググラフ (全件表示, 滞留は未出庫除外)
    function displayRankingChart() { /* 変更なし (内部で未出庫除外、slice削除済み) */
        const type = rankingTypeSelect.value; const nameKey = originalHeaders.find(h => h.includes('薬品名称')) || '薬品名称'; let sortKey = ''; let ascending = false; let labelKey = nameKey; let valueKey = ''; let chartLabel = ''; let chartUnit = ''; let indexAxis = 'y'; let chartType = 'bar'; let dataToUse = [...inventoryData];
         switch (type) {
            case 'amountDesc': sortKey = '在庫金額Numeric'; ascending = false; valueKey = '在庫金額Numeric'; chartLabel = '金額'; break;
            case 'stagnationDesc': sortKey = '滞留日数Numeric'; ascending = false; valueKey = '滞留日数Numeric'; chartLabel = '日数'; dataToUse = inventoryData.filter(item => item.滞留日数Numeric < 9000); break;
             case 'expiryAsc': sortKey = '残り日数Numeric'; ascending = true; valueKey = '残り日数Numeric'; chartLabel = '日数'; dataToUse = inventoryData.filter(item => !isNaN(item.残り日数Numeric) && item.残り日数Numeric >= 0); break;
             case 'quantityDesc': sortKey = '在庫数量Numeric'; ascending = false; valueKey = '在庫数量Numeric'; chartLabel = '数量'; break;
             case 'quantityAsc': sortKey = '在庫数量Numeric'; ascending = true; valueKey = '在庫数量Numeric'; chartLabel = '数量'; dataToUse = inventoryData.filter(item => item.在庫数量Numeric > 0); break;
            default: return;
        }
         const sortedData = dataToUse.sort((a, b) => { let valA = a[sortKey]; let valB = b[sortKey]; valA = isNaN(valA) ? (ascending ? Infinity : -Infinity) : valA; valB = isNaN(valB) ? (ascending ? Infinity : -Infinity) : valB; if (valA < valB) return ascending ? -1 : 1; if (valA > valB) return ascending ? 1 : -1; return String(a[labelKey] ?? '').localeCompare(String(b[labelKey] ?? '')); });
        const labels = sortedData.map(item => item[labelKey]); const dataValues = sortedData.map(item => item[valueKey]); const units = chartLabel === '数量' ? sortedData.map(item => item[originalHeaders.find(h => h.includes('単位')) || '単位'] || '') : null;
        if (sortedData.length > 50) { console.warn(`ランキンググラフ: ${sortedData.length}件描画試行中`); }
        if (rankingChartInstance) rankingChartInstance.destroy();
         try { rankingChartInstance = new Chart(rankingChartCanvas, { type: chartType, data: { labels: labels.reverse(), datasets: [{ label: chartLabel, data: dataValues.reverse(), backgroundColor: '#5C6BC0', borderRadius: 4, unit: units ? units.reverse() : '' }] }, options: { ...chartOptions, indexAxis: indexAxis } }); } catch(e) { console.error("Ranking Chart Error:", e); }
    }

    // --- AIアドバイス機能 ---
    async function fetchAiAdvice() { /* 変更なし */ }
    function generateInventorySummary() { /* 変更なし (内部で設定値を使用) */ }
    function createPrompt(summary) { /* 変更なし (内部で設定値を使用) */ }
    function saveApiKey() { /* 変更なし */ }
     function copyAdviceToClipboard() { /* 変更なし */ }
    // --- UI制御ヘルパー ---
    function showAiLoading() { aiLoading.style.display = 'block'; getAiAdviceButton.disabled = true; }
    function hideAiLoading() { aiLoading.style.display = 'none'; getAiAdviceButton.disabled = (userApiKey === null); }
    function showAiError(message) { aiErrorOutput.textContent = message; aiErrorOutput.style.display = 'block'; }
    function hideAiError() { aiErrorOutput.style.display = 'none'; }
    function showDataError(message) { dataErrorOutput.textContent = message; dataErrorOutput.style.display = 'block'; }
    function hideDataError() { dataErrorOutput.style.display = 'none'; }
    function showLoading() { if(loadingIndicator) loadingIndicator.style.display = 'block'; }
    function hideLoading() { if(loadingIndicator) loadingIndicator.style.display = 'none'; }

    // --- ▼▼▼ 印刷レポート生成 ▼▼▼ ---
    function generatePrintReport() {
        if (inventoryData.length === 0) {
            alert("印刷するデータがありません。");
            return;
        }

        // 印刷用コンテンツを生成
        const nameKey = originalHeaders.find(h => h.includes('薬品名称')) || '薬品名称';
        const valueKey = originalHeaders.find(h => h.includes('在庫金額(税別)')) || '在庫金額(税別)';
        const quantityKey = originalHeaders.find(h => h.includes('在庫数量')) || '在庫数量';
        const unitKey = originalHeaders.find(h => h.includes('単位')) || '単位';

        let reportHTML = `<h2>在庫レポート (${formatDate(analysisDate)}時点)</h2>`;

        // --- KPIサマリー ---
        reportHTML += `<h3><i class="fas fa-tachometer-alt"></i> 主要指標</h3>`;
        reportHTML += `<table class="print-table">
            <tr><th>総在庫品目数</th><td>${totalItemsEl.textContent}</td></tr>
            <tr><th>総在庫金額(税別)</th><td>${totalValueEl.textContent}</td></tr>
            <tr><th>要注意在庫(品目数)</th><td>${cautionItemsEl.textContent}</td></tr>
            <tr><th>要注意在庫(金額)</th><td>${cautionValueEl.textContent}</td></tr>
            <tr><th>滞留在庫(品目数)</th><td>${stagnantItemsEl.textContent}</td></tr>
            <tr><th>滞留在庫(金額)</th><td>${stagnantValueEl.textContent}</td></tr>
            <tr><th>期限切迫在庫(品目数)</th><td>${nearExpiryItemsEl.textContent}</td></tr>
            <tr><th>期限切迫在庫(金額)</th><td>${nearExpiryValueEl.textContent}</td></tr>
        </table>`;

        // --- 危険度ランク高いリスト ---
        const highRankItems = inventoryData.filter(item => item.危険度ランク >= 8).sort((a, b) => b.危険度ランク - a.危険度ランク);
        reportHTML += `<h3><i class="fas fa-exclamation-triangle"></i> 危険度高リスト (ランク8以上)</h3>`;
        if (highRankItems.length > 0) {
            reportHTML += `<table class="print-table"><thead><tr><th>ランク</th><th>薬品名</th><th>数量</th><th>金額</th><th>滞留日数</th><th>有効期限</th></tr></thead><tbody>`;
            highRankItems.slice(0, 30).forEach(item => { // 最大30件表示
                reportHTML += `<tr>
                    <td style="text-align:center;">${item.危険度ランク}</td>
                    <td>${item[nameKey]}</td>
                    <td style="text-align:right;">${item[quantityKey]} ${item[unitKey] || ''}</td>
                    <td style="text-align:right;">¥${formatCurrency(item.在庫金額Numeric)}</td>
                    <td style="text-align:right;">${item.滞留日数表示}</td>
                    <td style="text-align:center;">${item.有効期限表示}</td>
                </tr>`;
            });
            reportHTML += `</tbody></table>`;
            if (highRankItems.length > 30) reportHTML += `<p>...他 ${highRankItems.length - 30}件</p>`;
        } else {
            reportHTML += `<p>該当なし</p>`;
        }

        // --- 滞留在庫リスト ---
        const stagnantItemsList = inventoryData.filter(item => item.滞留日数Numeric >= stagnantThresholdDays && item.滞留日数Numeric < 9000).sort((a, b) => b.滞留日数Numeric - a.滞留日数Numeric);
        reportHTML += `<h3><i class="fas fa-hourglass-end"></i> 滞留在庫リスト (${stagnantThresholdDays}日以上)</h3>`;
         if (stagnantItemsList.length > 0) {
            reportHTML += `<table class="print-table"><thead><tr><th>滞留日数</th><th>薬品名</th><th>数量</th><th>金額</th><th>有効期限</th></tr></thead><tbody>`;
            stagnantItemsList.slice(0, 30).forEach(item => {
                reportHTML += `<tr>
                    <td style="text-align:right;">${item.滞留日数表示}</td>
                    <td>${item[nameKey]}</td>
                    <td style="text-align:right;">${item[quantityKey]} ${item[unitKey] || ''}</td>
                    <td style="text-align:right;">¥${formatCurrency(item.在庫金額Numeric)}</td>
                    <td style="text-align:center;">${item.有効期限表示}</td>
                </tr>`;
            });
            reportHTML += `</tbody></table>`;
             if (stagnantItemsList.length > 30) reportHTML += `<p>...他 ${stagnantItemsList.length - 30}件</p>`;
        } else {
            reportHTML += `<p>該当なし</p>`;
        }

        // --- 期限切迫リスト ---
        const nearExpiryList = inventoryData.filter(item => !isNaN(item.残り日数Numeric) && item.残り日数Numeric <= 90 && item.残り日数Numeric >= 0).sort((a, b) => a.残り日数Numeric - b.残り日数Numeric);
        reportHTML += `<h3><i class="fas fa-calendar-times"></i> 期限切迫リスト (90日以内)</h3>`;
        if (nearExpiryList.length > 0) {
            reportHTML += `<table class="print-table"><thead><tr><th>残り日数</th><th>薬品名</th><th>数量</th><th>金額</th><th>滞留日数</th></tr></thead><tbody>`;
            nearExpiryList.slice(0, 30).forEach(item => {
                reportHTML += `<tr>
                    <td style="text-align:right;">${item.残り日数Numeric}</td>
                    <td>${item[nameKey]}</td>
                    <td style="text-align:right;">${item[quantityKey]} ${item[unitKey] || ''}</td>
                    <td style="text-align:right;">¥${formatCurrency(item.在庫金額Numeric)}</td>
                    <td style="text-align:right;">${item.滞留日数表示}</td>
                </tr>`;
            });
            reportHTML += `</tbody></table>`;
             if (nearExpiryList.length > 30) reportHTML += `<p>...他 ${nearExpiryList.length - 30}件</p>`;
        } else {
            reportHTML += `<p>該当なし</p>`;
        }

        // 印刷用エリアにHTMLを挿入
        printReportContent.innerHTML = reportHTML;

        // 印刷ダイアログを開く
        window.print();
    }
    // --- ▲▲▲ 印刷レポート生成 ▲▲▲ ---


    // --- イベントリスナー設定 ---

    // データ読み込みボタン
    loadDataButton.addEventListener('click', () => {
        const tsvData = inventoryDataInput.value;
        const excludedData = excludedDrugsInput.value; // 除外リスト読み込み
        if (!tsvData.trim()) { showDataError('在庫データが入力されていません。'); return; }
        hideDataError();
        try {
            rawInventoryData = parseTSV(tsvData); // まず元データをパース
            excludedDrugIdentifiers = parseExcludedDrugs(excludedData); // 除外リストをパース
            initializeAppLogic(); // 初期化（内部で除外フィルタ実行）
        }
        catch (error) { console.error('データ読み込み/パースエラー:', error); showDataError(`データの読み込みに失敗しました: ${error.message}`); }
    });

    // タブ切り替え
    tabButtons.forEach(button => { /* 変更なし */ });

    // 在庫一覧: 検索、フィルタ、リセット
    searchInput.addEventListener('input', () => {if(inventoryData.length > 0) applyFiltersAndSort()});
    categoryFilter.addEventListener('change', () => {if(inventoryData.length > 0) applyFiltersAndSort()});
    makerFilter.addEventListener('change', () => {if(inventoryData.length > 0) applyFiltersAndSort()});
    dangerRankFilter.addEventListener('change', () => {if(inventoryData.length > 0) applyFiltersAndSort()});
    resetFiltersButton.addEventListener('click', () => { /* 変更なし */ });

     // 在庫一覧: テーブルヘッダーソート
    inventoryTableHeader.addEventListener('click', handleSort);

    // 在庫一覧: ページネーション
     prevPageButton.addEventListener('click', () => { /* 変更なし */ });
     nextPageButton.addEventListener('click', () => { /* 変更なし */ });
     itemsPerPageSelect.addEventListener('change', changeItemsPerPage);

     // --- ▼▼▼ 設定変更時のイベントリスナー ▼▼▼ ---
     expiryWeightSlider.addEventListener('input', updateWeightLabel); // ラベルのみリアルタイム更新
     // 以下の設定変更は再計算ボタンで反映
     // expiryWeightSlider.addEventListener('change', () => { if (inventoryData.length > 0) recalculateAll(); });
     // stagnantThresholdInput.addEventListener('change', () => { if (inventoryData.length > 0) recalculateAll(); });
     // analysisDateInput.addEventListener('change', () => { if (inventoryData.length > 0) recalculateAll(); }); // 自動再計算しない

     // ダッシュボードリスト件数 (変更時、これは即時反映)
     dashboardListCountSelect.addEventListener('change', () => {
         if (inventoryData.length > 0) {
             dashboardListCount = parseInt(dashboardListCountSelect.value, 10) || 5;
             displayDashboard(); // ダッシュボードのリスト部分のみ更新
         }
     });
     // 再計算ボタン
     recalculateButton.addEventListener('click', recalculateAll);
     // --- ▲▲▲ 設定変更時のイベントリスナー ▲▲▲ ---

     // 分析レポート: ランキング種類変更
     rankingTypeSelect.addEventListener('change', () => {if(inventoryData.length > 0) displayRankingChart()});

    // AIアドバイス
    saveApiKeyButton.addEventListener('click', saveApiKey);
    getAiAdviceButton.addEventListener('click', fetchAiAdvice);
     copyAiAdviceButton.addEventListener('click', copyAdviceToClipboard);

    // --- ▼▼▼ 印刷ボタンイベントリスナー ▼▼▼ ---
    printReportButton.addEventListener('click', generatePrintReport);
    // --- ▲▲▲ 印刷ボタンイベントリスナー ▲▲▲ ---


    // --- 初期状態設定 ---
    tabNav.style.display = 'none';
    mainContent.style.display = 'none';
    footer.style.display = 'none';

}); // End of DOMContentLoaded
