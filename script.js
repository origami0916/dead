document.addEventListener('DOMContentLoaded', () => {
    // --- グローバル変数定義 ---
    let rawInventoryData = [];
    let inventoryData = []; // 分析対象データ (除外フィルタ後)
    let excludedDrugIdentifiers = new Set();
    let filteredData = [];
    let currentData = [];
    let originalHeaders = [];
    let currentPage = 1;
    let itemsPerPage = 50;
    let currentSortColumn = null;
    let currentSortProcessedKey = null;
    let isSortAscending = true;
    let expiryWeightParam = 50;
    let stagnantThresholdDays = 180;
    let dashboardListCount = 5;
    let analysisDate = new Date();
    let userApiKey = null;
    let categoryChartInstance = null;
    let makerChartInstance = null;
    let rankingChartInstance = null;

    // --- DOM要素取得 ---
    const dataInputSection = document.getElementById('data-input-section');
    const inventoryDataInput = document.getElementById('inventoryDataInput');
    const excludedDrugsInput = document.getElementById('excludedDrugsInput');
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
    const unusedItemsEl = document.getElementById('unused-items'); // New KPI
    const unusedValueEl = document.getElementById('unused-value'); // New KPI
    const expiryWeightSlider = document.getElementById('expiryWeight');
    const weightLabel = document.getElementById('weightLabel');
    const stagnantThresholdInput = document.getElementById('stagnantThreshold');
    const dashboardListCountSelect = document.getElementById('dashboardListCount');
    const analysisDateInput = document.getElementById('analysisDate');
    const recalculateButton = document.getElementById('recalculateButton');
    const printReportButton = document.getElementById('printReportButton');
    const printReportContent = document.getElementById('print-report-content');
    const stagnantWorstListEl = document.getElementById('stagnant-worst-list');
    const nearExpiryWorstListEl = document.getElementById('near-expiry-worst-list');
    const highValueListEl = document.getElementById('high-value-list');
    const unusedListEl = document.getElementById('unused-list'); // New List
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
    function initializeAppLogic() {
        showLoading();
        try {
            initializeSettings();
            filterExcludedDrugs(); // 除外フィルタリング
            preprocessData(); // データ前処理
            renderTableHeader();
            populateFilterOptions();
            applyFiltersAndSort();
            displayDashboard(); // ダッシュボード表示

            // UI表示切り替え
            dataInputSection.style.display = 'none';
            tabNav.style.display = 'flex';
            mainContent.style.display = 'block';
            footer.style.display = 'block';
        } catch (error) {
            console.error('アプリケーションの初期化中にエラー:', error);
            showDataError('データの処理中にエラーが発生しました。データ形式を確認してください。');
            dataInputSection.style.display = 'block'; tabNav.style.display = 'none'; mainContent.style.display = 'none'; footer.style.display = 'none';
        } finally {
             hideLoading();
        }
    }

    // 設定値の初期化とUI反映
    function initializeSettings() {
        const today = new Date(); const jstOffset = 9 * 60; const localOffset = today.getTimezoneOffset();
        const jstNow = new Date(today.getTime() + (jstOffset + localOffset) * 60 * 1000);
        analysisDateInput.value = jstNow.toISOString().split('T')[0];
        analysisDate = new Date(analysisDateInput.value + 'T00:00:00');
        expiryWeightSlider.value = expiryWeightParam; updateWeightLabel();
        stagnantThresholdInput.value = stagnantThresholdDays;
        dashboardListCountSelect.value = dashboardListCount;
    }

    // --- データパーサー (TSV) ---
    function parseTSV(tsvString) { /* 変更なし */
        const lines = tsvString.trim().split('\n'); if (lines.length < 2) { throw new Error("データが少なすぎます。ヘッダー行とデータ行が必要です。"); }
        originalHeaders = lines[0].split('\t').map(header => header.trim()); const expectedHeaderCount = originalHeaders.length; const data = [];
        for (let i = 1; i < lines.length; i++) {
            const line = lines[i].trim(); if (!line) continue; const values = line.split('\t').map(value => value.trim());
            if (values.length !== expectedHeaderCount) { console.warn(`行 ${i + 1}: 列数不一致。スキップします。`); continue; }
            const item = {}; for (let j = 0; j < expectedHeaderCount; j++) { item[originalHeaders[j]] = values[j]; } data.push(item);
        } if (data.length === 0) { throw new Error("有効なデータ行が見つかりませんでした。"); } return data;
    }

    // --- 除外薬品処理 ---
    function parseExcludedDrugs(excludedString) { /* 変更なし */
        return new Set(excludedString.trim().split('\n').map(code => code.trim()).filter(Boolean));
    }
    function filterExcludedDrugs() { /* 変更なし */
        const drugCodeKey = originalHeaders.find(h => h.includes('薬品コード'));
        if (!drugCodeKey) { console.warn("除外処理スキップ: '薬品コード' カラムが見つかりません。"); inventoryData = [...rawInventoryData]; return; }
        if (excludedDrugIdentifiers.size > 0) { inventoryData = rawInventoryData.filter(item => !excludedDrugIdentifiers.has(item[drugCodeKey])); console.log(`${rawInventoryData.length - inventoryData.length} 件の薬品を除外しました。`); }
        else { inventoryData = [...rawInventoryData]; }
    }

    // --- データ前処理 ---
    function preprocessData() { /* 変更なし */
        const currentDate = analysisDate; const requiredKeys = ['有効期限', '最終出庫', '在庫金額(税別)', '在庫数量'];
        for (const key of requiredKeys) { if (!originalHeaders.includes(key)) { throw new Error(`必須カラム "${key}" がデータに含まれていません。`); } }
        const expiryDateKey = '有効期限'; const lastOutDateKey = '最終出庫'; const lastInDateKey = '最終入庫'; const stagnationColKey = '(停滞)';
        const lastInStagnationKey = originalHeaders.includes(lastInDateKey) && originalHeaders.indexOf(lastInDateKey) + 1 < originalHeaders.length && originalHeaders[originalHeaders.indexOf(lastInDateKey) + 1] === stagnationColKey ? stagnationColKey : originalHeaders.find(h => h.includes('最終入庫') && h.includes('停滞'));
        const lastOutStagnationKey = originalHeaders.includes(lastOutDateKey) && originalHeaders.indexOf(lastOutDateKey) + 1 < originalHeaders.length && originalHeaders[originalHeaders.indexOf(lastOutDateKey) + 1] === stagnationColKey ? stagnationColKey : originalHeaders.find(h => h.includes('最終出庫') && h.includes('停滞'));
        inventoryData = inventoryData.map(item => {
            const expiryDate = parseDate(item[expiryDateKey]); const lastOutDate = parseDate(item[lastOutDateKey]); const lastInDate = originalHeaders.includes(lastInDateKey) ? parseDate(item[lastInDateKey]) : null;
            let stagnationDays; if (lastOutStagnationKey && item[lastOutStagnationKey] && !isNaN(parseInt(item[lastOutStagnationKey]))) { stagnationDays = parseInt(item[lastOutStagnationKey]); } else { stagnationDays = calculateStagnationDays(currentDate, lastOutDate); }
            let lastInDays; if (lastInStagnationKey && item[lastInStagnationKey] && !isNaN(parseInt(item[lastInStagnationKey]))) { lastInDays = parseInt(item[lastInStagnationKey]); } else { lastInDays = calculateStagnationDays(currentDate, lastInDate); }
            const remainingDays = calculateRemainingDays(currentDate, expiryDate);
             const valueKey = originalHeaders.find(h => h.includes('在庫金額(税別)')) || '在庫金額(税別)'; const quantityKey = originalHeaders.find(h => h.includes('在庫数量')) || '在庫数量';
            const itemValue = parseFloat(String(item[valueKey]).replace(/,/g, '')) || 0; const itemQuantity = parseFloat(String(item[quantityKey]).replace(/,/g, '')) || 0;
            const dangerRank = calculateDangerRank(remainingDays, stagnationDays, expiryWeightParam);
             const stagnationDisplay = isNaN(stagnationDays) || stagnationDays > 9000 ? '---' : stagnationDays; const remainingDisplay = isNaN(remainingDays) ? '---' : remainingDays; const lastInDisplay = isNaN(lastInDays) || lastInDays > 9000 ? '---' : lastInDays;
            const processedItem = { ...item, '有効期限Date': expiryDate, '最終出庫日Date': lastOutDate, '最終入庫日Date': lastInDate, '在庫金額Numeric': itemValue, '在庫数量Numeric': itemQuantity, '滞留日数Numeric': stagnationDays, '残り日数Numeric': remainingDays, '最終入庫日数Numeric': lastInDays, '危険度ランク': dangerRank, '滞留日数表示': stagnationDisplay, '有効期限表示': expiryDate ? formatDate(expiryDate) : '---', '最終入庫停滞表示': lastInDisplay };
             if (lastOutStagnationKey) processedItem[lastOutStagnationKey] = stagnationDisplay; if (originalHeaders.includes(expiryDateKey)) processedItem[expiryDateKey] = processedItem['有効期限表示']; if (lastInStagnationKey) processedItem[lastInStagnationKey] = lastInDisplay;
            return processedItem;
        });
    }

    // 日付文字列をDateオブジェクトに変換
    function parseDate(dateString) { /* 変更なし */
        if (!dateString || typeof dateString !== 'string') return null; const trimmed = dateString.trim(); if (/^\d{4}$/.test(trimmed) || trimmed === '-') return null; const formattedDateString = trimmed.replace(/-/g, '/'); const parts = formattedDateString.split('/');
        if (parts.length === 3) { const year = parseInt(parts[0], 10); const month = parseInt(parts[1], 10); const day = parseInt(parts[2], 10); if (!isNaN(year) && !isNaN(month) && !isNaN(day) && month >= 1 && month <= 12 && day >= 1 && day <= 31) { const date = new Date(Date.UTC(year, month - 1, day)); if (date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day) { return date; } } } console.warn("Invalid date format:", dateString); return null;
    }
     // 日付オブジェクトをYYYY/MM/DD形式の文字列に変換
    function formatDate(date) { /* 変更なし */
        if (!date || !(date instanceof Date) || isNaN(date.getTime())) return '---'; const y = date.getUTCFullYear(); const m = String(date.getUTCMonth() + 1).padStart(2, '0'); const d = String(date.getUTCDate()).padStart(2, '0'); return `${y}/${m}/${d}`;
    }
    // 滞留日数計算
    function calculateStagnationDays(currentDate, lastOutDate) { /* 変更なし */
        if (!lastOutDate || !(lastOutDate instanceof Date) || isNaN(lastOutDate.getTime())) return 9999; const currentDayUTC = Date.UTC(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate()); const lastOutDayUTC = Date.UTC(lastOutDate.getUTCFullYear(), lastOutDate.getUTCMonth(), lastOutDate.getUTCDate()); const diffTime = currentDayUTC - lastOutDayUTC; return Math.max(0, Math.floor(diffTime / (1000 * 60 * 60 * 24)));
    }
    // 残り日数計算
    function calculateRemainingDays(currentDate, expiryDate) { /* 変更なし */
        if (!expiryDate || !(expiryDate instanceof Date) || isNaN(expiryDate.getTime())) return NaN; const currentDayUTC = Date.UTC(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate()); const expiryDayUTC = Date.UTC(expiryDate.getUTCFullYear(), expiryDate.getUTCMonth(), expiryDate.getUTCDate()); const diffTime = expiryDayUTC - currentDayUTC; return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    }

    // --- 危険度ランク計算 ---
    function calculateDangerRank(remainingDays, stagnationDays, weightParam) { /* 変更なし (期限切れランク10対応済み) */
        if (!isNaN(remainingDays) && remainingDays < 0) { return 10; }
        let scoreExpiry = 0; if (isNaN(remainingDays)) { scoreExpiry = 0; } else if (remainingDays <= 30) { scoreExpiry = 100; } else if (remainingDays <= 90) { scoreExpiry = 75; } else if (remainingDays <= 180) { scoreExpiry = 50; } else if (remainingDays <= 365) { scoreExpiry = 25; } else { scoreExpiry = 0; }
        let scoreStagnation = 0; if (isNaN(stagnationDays) || stagnationDays >= 730) { scoreStagnation = 100; } else if (stagnationDays >= 365) { scoreStagnation = 90; } else if (stagnationDays >= 180) { scoreStagnation = 60; } else if (stagnationDays >= 90) { scoreStagnation = 30; } else { scoreStagnation = 0; }
        const weightExpiry = weightParam / 100; const weightStagnation = 1.0 - weightExpiry; const totalScore = (scoreExpiry * weightExpiry) + (scoreStagnation * weightStagnation); return Math.max(1, Math.ceil(totalScore / 10));
    }

    // 全データの再計算と再表示
    function recalculateAll() { /* 変更なし (内部で設定値を使用) */
        if (rawInventoryData.length === 0) return; showLoading();
        const selectedDateStr = analysisDateInput.value; analysisDate = selectedDateStr ? new Date(selectedDateStr + 'T00:00:00') : new Date(); analysisDateInput.value = analysisDate.toISOString().split('T')[0];
        expiryWeightParam = parseInt(expiryWeightSlider.value, 10); stagnantThresholdDays = parseInt(stagnantThresholdInput.value, 10) || 180; dashboardListCount = parseInt(dashboardListCountSelect.value, 10) || 5;
        console.log("Recalculating with settings:", { analysisDate, expiryWeightParam, stagnantThresholdDays, dashboardListCount });
        setTimeout(() => {
            try { filterExcludedDrugs(); preprocessData(); applyFiltersAndSort(); displayDashboard(); if (document.getElementById('analysis-reports').classList.contains('active')) { displayRankingChart(); } }
            catch (error) { console.error('再計算中にエラー:', error); showDataError('再計算中にエラーが発生しました。'); }
            finally { hideLoading(); }
        }, 50);
    }

    // 危険度スライダーのラベル更新
    function updateWeightLabel() { /* 変更なし */ }

    // --- フィルタリングとソート ---
    function applyFiltersAndSort() { /* 変更なし */ }

    // --- ページネーション ---
    function updatePagination() { /* 変更なし */ }
     function displayCurrentPageData() { renderTable(currentData); }
     function updateItemCount() { itemCountEl.textContent = `全 ${filteredData.length} 件`; }
     function changeItemsPerPage() { /* 変更なし */ }

    // --- テーブル描画 ---
    function renderTableHeader() { /* 変更なし */ }
    function renderTable(data) { /* 変更なし (滞留クラス判定に閾値を使用する点は修正済み) */ }
     function formatCurrency(value) { /* 変更なし */ }
    // --- フィルター用選択肢生成 ---
    function populateFilterOptions() { /* 変更なし */ }
     // --- テーブルヘッダーのソート処理 ---
     function handleSort(event) { /* 変更なし */ }
     // ソートインジケータ更新
     function updateSortIndicators(activeTh) { /* 変更なし */ }

    // --- ダッシュボード表示 ---
    function displayDashboard() {
         const categoryKey = originalHeaders.find(h => h.includes('薬品種別')) || '薬品種別';
         const makerKey = originalHeaders.find(h => h.includes('メーカー')) || 'メーカー';
         const nameKey = originalHeaders.find(h => h.includes('薬品名称')) || '薬品名称';
         const lastInDateKey = originalHeaders.find(h => h.includes('最終入庫')) || '最終入庫'; // 最終入庫日キー

        // KPI計算
        const totalItems = inventoryData.length;
        const totalValue = inventoryData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);
        const cautionData = inventoryData.filter(item => item.危険度ランク >= 7);
        const cautionItems = cautionData.length;
        const cautionValue = cautionData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);
        const stagnantData = inventoryData.filter(item => item.滞留日数Numeric >= stagnantThresholdDays && item.滞留日数Numeric < 9000); // 未出庫(9999)除く
        const stagnantItems = stagnantData.length;
        const stagnantValue = stagnantData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);
        const nearExpiryData = inventoryData.filter(item => !isNaN(item.残り日数Numeric) && item.残り日数Numeric <= 90 && item.残り日数Numeric >= 0);
        const nearExpiryItems = nearExpiryData.length;
        const nearExpiryValue = nearExpiryData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);
        // ★★★ 未使用在庫KPI計算 ★★★
        const unusedData = inventoryData.filter(item => item.滞留日数Numeric === 9999);
        const unusedItems = unusedData.length;
        const unusedValue = unusedData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);
        // ★★★★★★★★★★★★★★★★★

        // KPI表示更新
        totalItemsEl.textContent = totalItems.toLocaleString();
        totalValueEl.textContent = `¥${formatCurrency(totalValue)}`;
        cautionItemsEl.textContent = cautionItems.toLocaleString();
        cautionValueEl.textContent = `¥${formatCurrency(cautionValue)}`;
        stagnantItemsEl.textContent = `${stagnantItems.toLocaleString()} ( ${stagnantThresholdDays}日以上 )`;
        stagnantValueEl.textContent = `¥${formatCurrency(stagnantValue)}`;
        nearExpiryItemsEl.textContent = nearExpiryItems.toLocaleString();
        nearExpiryValueEl.textContent = `¥${formatCurrency(nearExpiryValue)}`;
        // ★★★ 未使用在庫KPI表示更新 ★★★
        unusedItemsEl.textContent = unusedItems.toLocaleString();
        unusedValueEl.textContent = `¥${formatCurrency(unusedValue)}`;
        // ★★★★★★★★★★★★★★★★★★

        // 注目リスト更新
        updateRankingList(stagnantWorstListEl, inventoryData.filter(item => item.滞留日数Numeric < 9000), '滞留日数Numeric', false, dashboardListCount, item => `${item[nameKey]} (${item.滞留日数表示}日)`);
        updateRankingList(nearExpiryWorstListEl, inventoryData.filter(item => !isNaN(item.残り日数Numeric) && item.残り日数Numeric >= 0), '残り日数Numeric', true, dashboardListCount, item => `${item[nameKey]} (残${item.残り日数Numeric}日)`);
        updateRankingList(highValueListEl, inventoryData, '在庫金額Numeric', false, dashboardListCount, item => `${item[nameKey]} (¥${formatCurrency(item.在庫金額Numeric)})`);
        // ★★★ 未使用在庫リスト更新 (最終入庫日が古い順) ★★★
        updateRankingList(unusedListEl, unusedData, '最終入庫日Date', true, dashboardListCount, item => `${item[nameKey]} (最終入庫: ${item[lastInDateKey] || '不明'})`);
        // ★★★★★★★★★★★★★★★★★★★★★★★★★★★

        // グラフ更新
        displayCategoryChart(categoryKey);
        displayMakerChart(makerKey);
    }

    // ランキングリスト更新ヘルパー
    function updateRankingList(listElement, data, sortKey, ascending, count, displayFormatter) {
        const sortedData = [...data].sort((a, b) => {
            let valA = a[sortKey]; let valB = b[sortKey];
            // Dateオブジェクトの場合の比較を追加
            if (valA instanceof Date && valB instanceof Date) {
                valA = isNaN(valA.getTime()) ? (ascending ? Infinity : -Infinity) : valA.getTime();
                valB = isNaN(valB.getTime()) ? (ascending ? Infinity : -Infinity) : valB.getTime();
            } else {
                 const numA = parseFloat(valA); const numB = parseFloat(valB);
                 valA = isNaN(numA) ? (ascending ? Infinity : -Infinity) : numA;
                 valB = isNaN(numB) ? (ascending ? Infinity : -Infinity) : numB;
            }
            if (valA < valB) return ascending ? -1 : 1;
            if (valA > valB) return ascending ? 1 : -1;
            const nameKey = originalHeaders.find(h => h.includes('薬品名称')) || '薬品名称';
            return String(a[nameKey] ?? '').localeCompare(String(b[nameKey] ?? ''));
        }).slice(0, count);

        listElement.innerHTML = '';
        if (sortedData.length === 0) { listElement.innerHTML = '<li>該当データなし</li>'; return; }
        sortedData.forEach(item => { const li = document.createElement('li'); li.textContent = displayFormatter(item); listElement.appendChild(li); });
    }


    // --- グラフ描画 (Chart.js) ---
    const chartOptions = { /* 変更なし */ };
    function displayCategoryChart(categoryKey) { /* 変更なし */ }
    function displayMakerChart(makerKey) { /* 変更なし */ }
    // ランキンググラフ (全件表示, 滞留は未出庫除外)
    function displayRankingChart() { /* 変更なし (滞留未出庫除外済み) */ }

    // --- AIアドバイス機能 ---
    async function fetchAiAdvice() { /* 変更なし */ }
    // サマリー生成 (未使用在庫情報を追加)
    function generateInventorySummary() {
         const nameKey = originalHeaders.find(h => h.includes('薬品名称')) || '薬品名称';
         const unitKey = originalHeaders.find(h => h.includes('単位')) || '単位';
         const lastInDateKey = originalHeaders.find(h => h.includes('最終入庫')) || '最終入庫';

        const totalItems = inventoryData.length;
        const totalValue = inventoryData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);
        const cautionData = inventoryData.filter(item => item.危険度ランク >= 7);
        const cautionItems = cautionData.length;
        const cautionValue = cautionData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);
        const topCaution = cautionData.sort((a, b) => b.危険度ランク - a.危険度ランク).slice(0, 5).map(i => `- ${i[nameKey]} (ランク${i.危険度ランク}, 金額¥${formatCurrency(i.在庫金額Numeric)})`).join('\n');
        const stagnantData = inventoryData.filter(item => item.滞留日数Numeric >= stagnantThresholdDays && item.滞留日数Numeric < 9000);
        const stagnantItems = stagnantData.length;
        const stagnantValue = stagnantData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);
        const topStagnant = stagnantData.sort((a, b) => b.滞留日数Numeric - a.滞留日数Numeric).slice(0, 5).map(i => `- ${i[nameKey]} (${i.滞留日数表示}日, 金額¥${formatCurrency(i.在庫金額Numeric)})`).join('\n');
        const nearExpiryData = inventoryData.filter(item => !isNaN(item.残り日数Numeric) && item.残り日数Numeric <= 90 && item.残り日数Numeric >= 0);
        const nearExpiryItems = nearExpiryData.length;
        const nearExpiryValue = nearExpiryData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);
        const topNearExpiry = nearExpiryData.sort((a, b) => a.残り日数Numeric - b.残り日数Numeric).slice(0, 5).map(i => `- ${i[nameKey]} (残${i.残り日数Numeric}日, 金額¥${formatCurrency(i.在庫金額Numeric)})`).join('\n');
        const highValueData = inventoryData.filter(item => item.在庫金額Numeric > 0);
        const topHighValue = highValueData.sort((a,b) => b.在庫金額Numeric - a.在庫金額Numeric).slice(0,5).map(i => `- ${i[nameKey]} (¥${formatCurrency(i.在庫金額Numeric)})`).join('\n');
        const lowStockData = inventoryData.filter(item => item.在庫数量Numeric <= 5 && item.滞留日数Numeric <= 30 && item.在庫数量Numeric > 0);
        const topLowStock = lowStockData.sort((a,b) => a.在庫数量Numeric - b.在庫数量Numeric).slice(0,5).map(i => `- ${i[nameKey]} (残${i.在庫数量} ${i[unitKey] || ''})`).join('\n');
        // ★★★ 未使用在庫サマリー ★★★
        const unusedData = inventoryData.filter(item => item.滞留日数Numeric === 9999);
        const unusedItems = unusedData.length;
        const unusedValue = unusedData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);
        const topUnused = unusedData.sort((a, b) => { // 最終入庫日でソート
            const dateA = a.最終入庫日Date instanceof Date ? a.最終入庫日Date.getTime() : Infinity;
            const dateB = b.最終入庫日Date instanceof Date ? b.最終入庫日Date.getTime() : Infinity;
            return (isNaN(dateA) ? Infinity : dateA) - (isNaN(dateB) ? Infinity : dateB); // 古い順
        }).slice(0, 5).map(i => `- ${i[nameKey]} (最終入庫: ${i[lastInDateKey] || '不明'}, 金額¥${formatCurrency(i.在庫金額Numeric)})`).join('\n');
        // ★★★★★★★★★★★★★★★★★

         return `
## 在庫サマリー (${formatDate(analysisDate)}時点)

**基本情報:**
- 総在庫品目数: ${totalItems.toLocaleString()} 品目
- 総在庫金額 (税別): ¥${formatCurrency(totalValue)}

**要注意在庫 (危険度ランク7以上):**
- 該当品目数: ${cautionItems.toLocaleString()} 品目 / 合計金額: ¥${formatCurrency(cautionValue)}
- 上位リスト (抜粋):\n${topCaution || '- 該当なし'}

**滞留在庫 (${stagnantThresholdDays}日以上出庫なし):**
- 該当品目数: ${stagnantItems.toLocaleString()} 品目 / 合計金額: ¥${formatCurrency(stagnantValue)}
- 上位リスト (抜粋):\n${topStagnant || '- 該当なし'}

**期限切迫在庫 (残り90日以内):**
- 該当品目数: ${nearExpiryItems.toLocaleString()} 品目 / 合計金額: ¥${formatCurrency(nearExpiryValue)}
- 上位リスト (抜粋):\n${topNearExpiry || '- 該当なし'}

**未使用在庫 (一度も出庫記録なし):**
- 該当品目数: ${unusedItems.toLocaleString()} 品目 / 合計金額: ¥${formatCurrency(unusedValue)}
- 上位リスト (最終入庫が古い順抜粋):\n${topUnused || '- 該当なし'}

**高額在庫 (上位抜粋):**
- 上位リスト (抜粋):\n${topHighValue || '- 該当なし'}

**欠品注意在庫 (在庫僅少かつ最近出庫あり - 抜粋):**
- 上位リスト (抜粋):\n${topLowStock || '- 該当なし'}

**危険度ランク設定:**
- 現在の重み付け: 期限重視 ${expiryWeightParam}%, 滞留重視 ${100 - expiryWeightParam}%
`;
    }
    function createPrompt(summary) { /* 変更なし */ }
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
        if (inventoryData.length === 0) { alert("印刷するデータがありません。"); return; }
        const nameKey = originalHeaders.find(h => h.includes('薬品名称')) || '薬品名称';
        const valueKey = originalHeaders.find(h => h.includes('在庫金額(税別)')) || '在庫金額(税別)';
        const quantityKey = originalHeaders.find(h => h.includes('在庫数量')) || '在庫数量';
        const unitKey = originalHeaders.find(h => h.includes('単位')) || '単位';
        const lastInDateKey = originalHeaders.find(h => h.includes('最終入庫')) || '最終入庫';

        let reportHTML = `<h2>在庫レポート (${formatDate(analysisDate)}時点)</h2>`;
        // KPIサマリー
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
            <tr><th>未使用在庫(品目数)</th><td>${unusedItemsEl.textContent}</td></tr>
            <tr><th>未使用在庫(金額)</th><td>${unusedValueEl.textContent}</td></tr>
        </table>`;
        // 危険度高リスト
        const highRankItems = inventoryData.filter(item => item.危険度ランク >= 8).sort((a, b) => b.危険度ランク - a.危険度ランク);
        reportHTML += `<h3><i class="fas fa-exclamation-triangle"></i> 危険度高リスト (ランク8以上)</h3>`;
        if (highRankItems.length > 0) {
            reportHTML += `<table class="print-table"><thead><tr><th>ランク</th><th>薬品名</th><th>数量</th><th>金額</th><th>滞留日数</th><th>有効期限</th></tr></thead><tbody>`;
            highRankItems.slice(0, 30).forEach(item => { reportHTML += `<tr><td style="text-align:center;">${item.危険度ランク}</td><td>${item[nameKey]}</td><td style="text-align:right;">${item[quantityKey]} ${item[unitKey] || ''}</td><td style="text-align:right;">¥${formatCurrency(item.在庫金額Numeric)}</td><td style="text-align:right;">${item.滞留日数表示}</td><td style="text-align:center;">${item.有効期限表示}</td></tr>`; });
            reportHTML += `</tbody></table>`; if (highRankItems.length > 30) reportHTML += `<p>...他 ${highRankItems.length - 30}件</p>`;
        } else { reportHTML += `<p>該当なし</p>`; }
        // 滞留在庫リスト
        const stagnantItemsList = inventoryData.filter(item => item.滞留日数Numeric >= stagnantThresholdDays && item.滞留日数Numeric < 9000).sort((a, b) => b.滞留日数Numeric - a.滞留日数Numeric);
        reportHTML += `<h3><i class="fas fa-hourglass-end"></i> 滞留在庫リスト (${stagnantThresholdDays}日以上)</h3>`;
         if (stagnantItemsList.length > 0) {
            reportHTML += `<table class="print-table"><thead><tr><th>滞留日数</th><th>薬品名</th><th>数量</th><th>金額</th><th>有効期限</th></tr></thead><tbody>`;
            stagnantItemsList.slice(0, 30).forEach(item => { reportHTML += `<tr><td style="text-align:right;">${item.滞留日数表示}</td><td>${item[nameKey]}</td><td style="text-align:right;">${item[quantityKey]} ${item[unitKey] || ''}</td><td style="text-align:right;">¥${formatCurrency(item.在庫金額Numeric)}</td><td style="text-align:center;">${item.有効期限表示}</td></tr>`; });
            reportHTML += `</tbody></table>`; if (stagnantItemsList.length > 30) reportHTML += `<p>...他 ${stagnantItemsList.length - 30}件</p>`;
        } else { reportHTML += `<p>該当なし</p>`; }
        // ★★★ 未使用在庫リスト印刷 ★★★
        const unusedItemsList = inventoryData.filter(item => item.滞留日数Numeric === 9999).sort((a, b) => { const dateA = a.最終入庫日Date instanceof Date ? a.最終入庫日Date.getTime() : Infinity; const dateB = b.最終入庫日Date instanceof Date ? b.最終入庫日Date.getTime() : Infinity; return (isNaN(dateA) ? Infinity : dateA) - (isNaN(dateB) ? Infinity : dateB); });
        reportHTML += `<h3><i class="fas fa-box-open"></i> 未使用在庫リスト (最終入庫が古い順)</h3>`;
         if (unusedItemsList.length > 0) {
            reportHTML += `<table class="print-table"><thead><tr><th>最終入庫日</th><th>薬品名</th><th>数量</th><th>金額</th><th>有効期限</th></tr></thead><tbody>`;
            unusedItemsList.slice(0, 30).forEach(item => { reportHTML += `<tr><td style="text-align:center;">${item[lastInDateKey] || '不明'}</td><td>${item[nameKey]}</td><td style="text-align:right;">${item[quantityKey]} ${item[unitKey] || ''}</td><td style="text-align:right;">¥${formatCurrency(item.在庫金額Numeric)}</td><td style="text-align:center;">${item.有効期限表示}</td></tr>`; });
            reportHTML += `</tbody></table>`; if (unusedItemsList.length > 30) reportHTML += `<p>...他 ${unusedItemsList.length - 30}件</p>`;
        } else { reportHTML += `<p>該当なし</p>`; }
        // ★★★★★★★★★★★★★★★★★★★★
        // 期限切迫リスト
        const nearExpiryList = inventoryData.filter(item => !isNaN(item.残り日数Numeric) && item.残り日数Numeric <= 90 && item.残り日数Numeric >= 0).sort((a, b) => a.残り日数Numeric - b.残り日数Numeric);
        reportHTML += `<h3><i class="fas fa-calendar-times"></i> 期限切迫リスト (90日以内)</h3>`;
        if (nearExpiryList.length > 0) {
            reportHTML += `<table class="print-table"><thead><tr><th>残り日数</th><th>薬品名</th><th>数量</th><th>金額</th><th>滞留日数</th></tr></thead><tbody>`;
            nearExpiryList.slice(0, 30).forEach(item => { reportHTML += `<tr><td style="text-align:right;">${item.残り日数Numeric}</td><td>${item[nameKey]}</td><td style="text-align:right;">${item[quantityKey]} ${item[unitKey] || ''}</td><td style="text-align:right;">¥${formatCurrency(item.在庫金額Numeric)}</td><td style="text-align:right;">${item.滞留日数表示}</td></tr>`; });
            reportHTML += `</tbody></table>`; if (nearExpiryList.length > 30) reportHTML += `<p>...他 ${nearExpiryList.length - 30}件</p>`;
        } else { reportHTML += `<p>該当なし</p>`; }

        printReportContent.innerHTML = reportHTML;
        window.print(); // 印刷ダイアログ表示
    }
    // --- ▲▲▲ 印刷レポート生成 ▲▲▲ ---


    // --- イベントリスナー設定 ---
    loadDataButton.addEventListener('click', () => {
        const tsvData = inventoryDataInput.value; const excludedData = excludedDrugsInput.value;
        if (!tsvData.trim()) { showDataError('在庫データが入力されていません。'); return; } hideDataError();
        try { rawInventoryData = parseTSV(tsvData); excludedDrugIdentifiers = parseExcludedDrugs(excludedData); initializeAppLogic(); }
        catch (error) { console.error('データ読み込み/パースエラー:', error); showDataError(`データの読み込みに失敗しました: ${error.message}`); }
    });
    tabButtons.forEach(button => { /* 変更なし */ });
    searchInput.addEventListener('input', () => {if(inventoryData.length > 0) applyFiltersAndSort()});
    categoryFilter.addEventListener('change', () => {if(inventoryData.length > 0) applyFiltersAndSort()});
    makerFilter.addEventListener('change', () => {if(inventoryData.length > 0) applyFiltersAndSort()});
    dangerRankFilter.addEventListener('change', () => {if(inventoryData.length > 0) applyFiltersAndSort()});
    resetFiltersButton.addEventListener('click', () => { /* 変更なし */ });
    inventoryTableHeader.addEventListener('click', handleSort);
    prevPageButton.addEventListener('click', () => { /* 変更なし */ });
    nextPageButton.addEventListener('click', () => { /* 変更なし */ });
    itemsPerPageSelect.addEventListener('change', changeItemsPerPage);
    // --- ▼▼▼ 設定変更時のイベントリスナー ▼▼▼ ---
    expiryWeightSlider.addEventListener('input', updateWeightLabel);
    // 再計算ボタンですべての設定変更を反映
    recalculateButton.addEventListener('click', recalculateAll);
    // ダッシュボードリスト件数のみ即時反映
    dashboardListCountSelect.addEventListener('change', () => { if (inventoryData.length > 0) { dashboardListCount = parseInt(dashboardListCountSelect.value, 10) || 5; displayDashboard(); } });
    // --- ▲▲▲ 設定変更時のイベントリスナー ▲▲▲ ---
    rankingTypeSelect.addEventListener('change', () => {if(inventoryData.length > 0) displayRankingChart()});
    saveApiKeyButton.addEventListener('click', saveApiKey);
    getAiAdviceButton.addEventListener('click', fetchAiAdvice);
    copyAiAdviceButton.addEventListener('click', copyAdviceToClipboard);
    printReportButton.addEventListener('click', generatePrintReport); // 印刷ボタン

    // --- 初期状態設定 ---
    tabNav.style.display = 'none'; mainContent.style.display = 'none'; footer.style.display = 'none';

}); // End of DOMContentLoaded
