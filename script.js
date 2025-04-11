// Wrap the entire script in a try...catch
try {
    document.addEventListener('DOMContentLoaded', () => {
        console.log("DOM fully loaded and parsed");

        // --- Globals ---
        let inventoryData = []; let expiryData = []; let filteredData = [];
        let settings = { stagnantDays: 180, nearExpiryDays: 90 };
        let sortColumn = null; let sortDirection = 'asc';
        const today = new Date(); today.setHours(0, 0, 0, 0);
        let currentAlerts = { expired: 0, near_expiry: 0, stagnant: 0 };
        let statusPieChartInstance = null; let topValueBarChartInstance = null;
        let stagnationDistChartInstance = null; let makerBarChartInstance = null;
        let currentStatusFilter = null; // For alert filtering state

        // --- DOM Elements ---
        console.log("Getting DOM Elements...");
        const appContainer = document.getElementById('appContainer');
        const pasteDataInput = document.getElementById('pasteDataInput');
        const processPasteBtn = document.getElementById('processPasteBtn');
        const expiryDataInput = document.getElementById('expiryDataInput');
        const processExpiryBtn = document.getElementById('processExpiryBtn');
        const inventoryTable = document.getElementById('inventoryTable');
        const inventoryTableBody = inventoryTable?.querySelector('tbody');
        const inventoryTableHeader = inventoryTable?.querySelector('thead tr');
        const totalDeadStockValueEl = document.getElementById('totalDeadStockValue');
        const totalDeadStockItemsEl = document.getElementById('totalDeadStockItems');
        const alertOutputEl = document.getElementById('alertOutput');
        const filterMakerInput = document.getElementById('filterMaker');
        const filterWholesalerInput = document.getElementById('filterWholesaler');
        const filterLastOutDateStartInput = document.getElementById('filterLastOutDateStart');
        const filterLastOutDateEndInput = document.getElementById('filterLastOutDateEnd');
        const filterExpiryDateStartInput = document.getElementById('filterExpiryDateStart');
        const filterExpiryDateEndInput = document.getElementById('filterExpiryDateEnd');
        const filterKeywordInput = document.getElementById('filterKeyword');
        const applyFilterBtn = document.getElementById('applyFilterBtn');
        const resetFilterBtn = document.getElementById('resetFilterBtn');
        const exportCsvBtn = document.getElementById('exportCsvBtn');
        const exportPdfBtn = document.getElementById('exportPdfBtn');
        const fullscreenBtn = document.getElementById('fullscreenBtn');
        const currentDateEl = document.getElementById('currentDate');
        const settingsBtn = document.getElementById('settingsBtn');
        const settingsArea = document.getElementById('settingsArea');
        const stagnantDaysInput = document.getElementById('stagnantDays');
        const nearExpiryDaysInput = document.getElementById('nearExpiryDays');
        const saveSettingsBtn = document.getElementById('saveSettingsBtn');
        const closeSettingsBtn = document.getElementById('closeSettingsBtn');
        const aiAdvisorBtn = document.getElementById('aiAdvisorBtn');
        const aiAdvisorArea = document.getElementById('aiAdvisorArea');
        const apiKeyInput = document.getElementById('apiKey');
        const aiQuestionInput = document.getElementById('aiQuestion');
        const getAdviceBtn = document.getElementById('getAdviceBtn');
        const closeAiBtn = document.getElementById('closeAiBtn');
        const aiAdviceOutputEl = document.getElementById('aiAdviceOutput');
        const aiLoadingEl = document.getElementById('aiLoading');
        const statusPieChartCtx = document.getElementById('statusPieChart')?.getContext('2d');
        const topValueBarChartCtx = document.getElementById('topValueBarChart')?.getContext('2d');
        const stagnationDistChartCtx = document.getElementById('stagnationDistChart')?.getContext('2d');
        const makerBarChartCtx = document.getElementById('makerBarChart')?.getContext('2d');
        const adviceBoxEl = document.getElementById('adviceBox');
        const filterShowAllBtn = document.getElementById('filterShowAllBtn');
        const filterExpiredBtn = document.getElementById('filterExpiredBtn');
        const filterNearExpiryBtn = document.getElementById('filterNearExpiryBtn');
        const filterStagnantBtn = document.getElementById('filterStagnantBtn');
        const alertFilterButtons = [filterShowAllBtn, filterExpiredBtn, filterNearExpiryBtn, filterStagnantBtn];

        console.log("DOM Elements references acquired.");

        // --- Constants ---
        const HEADERS = [ "薬品コード", "YJコード", "薬品種別", "フリガナ", "薬品名称", "包装コード", "包装形状", "在庫数量", "在庫金額(税別)", "最終入庫(停滞)", "最終出庫(停滞)", "メーカー", "卸", "単位", "出力設定", "税額", "有効期限" ];
        const IDX_NAME = 4; const IDX_YJ = 1; const IDX_QTY = 7; const IDX_VALUE = 8;
        const IDX_LAST_IN_ST = 9; const IDX_LAST_OUT_ST = 10; const IDX_MAKER = 11;
        const IDX_WHOLESALER = 12; const IDX_EXPIRY = 16;
        const EXPIRY_HEADERS = [ "薬品名称", "包装形状", "有効期限", "在庫数", "在庫単価", "在庫金額", "最終入庫日", "入庫数量", "出庫数量", "薬価", "薬価金額", "印刷区分", "薬品種別", "税額" ];
        const EXP_IDX_NAME = 0; const EXP_IDX_EXPIRY = 2; const EXP_IDX_VALUE = 5; const EXP_IDX_YAKKA_VALUE = 10;

        // --- Initialization ---
        console.log("Initializing...");
        loadSettings(); updateCurrentDate(); setupEventListeners(); initializeCharts(); updateAdviceBox();
        console.log("Initialization complete.");

        // --- Core Functions ---
        function parsePastedData(dataString, expectedHeaders) {
             // 貼り付けデータをパース (変更なし)
            return new Promise((resolve, reject) => {
                if (!dataString || dataString.trim() === '') return reject("貼り付けるデータが空です。");
                if (typeof Papa === 'undefined') return reject("データ解析ライブラリ(PapaParse)が読み込まれていません。");
                Papa.parse(dataString, {
                    header: true, skipEmptyLines: true, escapeChar: '"',
                    delimitersToGuess: [',', '\t', '|', ';'],
                    complete: (results) => {
                        if (results.errors.length > 0) { console.error("Parse Errors:", results.errors); const firstError = results.errors[0]; const rowInfo = firstError.row ? ` (行: ${firstError.row + 2} 付近)` : ''; reject(`データパース中にエラーが発生しました${rowInfo}: ${firstError.message}. 貼り付けたデータの形式（特に引用符(")の使い方）を確認してください。`); }
                        else if (!results.data || results.data.length === 0) reject("貼り付けられたデータから有効な情報を抽出できませんでした。");
                        else if (!results.meta.fields || results.meta.fields.length < 3) { console.warn("データのヘッダーが期待値と異なるか、列数が少ないようです。", results.meta.fields); resolve(results.data); }
                        else { const actualHeaders = results.meta.fields.map(h => h.trim()); if (expectedHeaders && JSON.stringify(actualHeaders) !== JSON.stringify(expectedHeaders)) { console.warn("ヘッダー名が期待値と完全に一致しません。", "期待値:", expectedHeaders, "実際:", actualHeaders); } resolve(results.data); }
                    }, error: (error) => reject(`データパースエラー: ${error.message}`)
                });
            });
        }
        function parseDate(dateString) { /* ... (変更なし) ... */ }
        function parseStagnationDays(valueString) { /* ... (変更なし) ... */ }
        function daysDifference(date1, date2) { /* ... (変更なし) ... */ }
        function judgeDeadStock(item) { /* ... (変更なし) ... */ }

        // --- UI Update Functions ---
        function renderTable() {
            // テーブル描画 (クラス名をCSSファイルのものに合わせる)
            if (!inventoryTableBody || !inventoryTableHeader) {
                 console.error("Table body or header not found!");
                 return;
            }
            inventoryTableHeader.innerHTML = ''; inventoryTableBody.innerHTML = '';
            let displayHeaders = HEADERS;
            if (inventoryData && inventoryData.length > 0 && typeof inventoryData[0] === 'object' && inventoryData[0] !== null) {
                displayHeaders = Object.keys(inventoryData[0]);
            }
            // Render Header
            displayHeaders.forEach(header => {
                const th = document.createElement('th');
                th.textContent = header; th.dataset.column = header;
                th.addEventListener('click', handleSortClick);
                if (header === sortColumn) { th.textContent += sortDirection === 'asc' ? ' ▲' : ' ▼'; th.classList.add('sorted'); } // Add 'sorted' class maybe
                inventoryTableHeader.appendChild(th);
            });
            // Render Body
             if (filteredData.length === 0 && inventoryData.length > 0) {
                 inventoryTableBody.innerHTML = `<tr><td colspan="${displayHeaders.length}">フィルター条件に一致するデータがありません。</td></tr>`; return;
            } else if (inventoryData.length === 0) {
                 inventoryTableBody.innerHTML = `<tr><td colspan="${displayHeaders.length}">在庫データを貼り付けて「処理」ボタンを押してください。</td></tr>`; return;
            }
            filteredData.forEach(item => {
                const row = inventoryTableBody.insertRow();
                const status = judgeDeadStock(item);
                // ★ Use CSS classes defined in style.css ★
                row.classList.add('table-row'); // Add a general row class if needed
                switch(status) {
                    case 'expired': row.classList.add('status-expired'); break;
                    case 'near_expired': row.classList.add('status-near_expired'); break;
                    case 'stagnant': row.classList.add('status-stagnant'); break;
                }
                row.dataset.yjCode = item[HEADERS[IDX_YJ]] || '';
                displayHeaders.forEach(header => {
                    const cell = row.insertCell();
                    const value = item.hasOwnProperty(header) && item[header] !== undefined && item[header] !== null ? item[header] : '';
                    cell.textContent = value;
                    if (header === HEADERS[IDX_EXPIRY] && status === 'expired') cell.classList.add('important-cell'); // Example class
                });
            });
        }
        function handleSortClick(event) { /* ... (変更なし) ... */ }
        function sortData() { /* ... (変更なし) ... */ }

        function applyFiltersAndSearch() {
            // フィルター・検索適用 (アラートフィルター対応)
            const makerFilter = filterMakerInput?.value.toLowerCase() ?? '';
            const wholesalerFilter = filterWholesalerInput?.value.toLowerCase() ?? '';
            const lastOutStart = parseDate(filterLastOutDateStartInput?.value);
            const lastOutEnd = parseDate(filterLastOutDateEndInput?.value);
            const expiryStart = parseDate(filterExpiryDateStartInput?.value);
            const expiryEnd = parseDate(filterExpiryDateEndInput?.value);
            const keyword = filterKeywordInput?.value.toLowerCase() ?? '';

            filteredData = inventoryData.filter(item => {
                // General filters
                const itemMaker = String(item[HEADERS[IDX_MAKER]] || '').toLowerCase(); const itemWholesaler = String(item[HEADERS[IDX_WHOLESALER]] || '').toLowerCase();
                const itemLastOut = parseDate(item[HEADERS[IDX_LAST_OUT_ST]]); const itemExpiry = parseDate(item[HEADERS[IDX_EXPIRY]]);
                const itemName = String(item[HEADERS[IDX_NAME]] || '').toLowerCase(); const itemFurigana = String(item["フリガナ"] || '').toLowerCase();
                const itemCode = String(item[HEADERS[0]] || '').toLowerCase(); const itemYjCode = String(item[HEADERS[IDX_YJ]] || '').toLowerCase();
                if (makerFilter && !itemMaker.includes(makerFilter)) return false; if (wholesalerFilter && !itemWholesaler.includes(wholesalerFilter)) return false;
                if (lastOutStart && (!itemLastOut || itemLastOut < lastOutStart)) return false; if (lastOutEnd && (!itemLastOut || itemLastOut > lastOutEnd)) return false;
                if (expiryStart && (!itemExpiry || itemExpiry < expiryStart)) return false; if (expiryEnd && (!itemExpiry || itemExpiry > expiryEnd)) return false;
                if (keyword && !itemName.includes(keyword) && !itemFurigana.includes(keyword) && !itemCode.includes(keyword) && !itemYjCode.includes(keyword)) return false;

                // Alert status filter
                if (currentStatusFilter) {
                    const status = judgeDeadStock(item);
                    if (status !== currentStatusFilter) return false;
                }
                return true;
            });

            // Update UI
            sortData(); renderTable(); updateSummaryAndAlerts(); updateCharts();
            updateAdviceBox(); updateAlertFilterStyles();
        }

        function updateSummaryAndAlerts() {
            // サマリー・アラート更新 (アラートクラス名をCSSに合わせる)
            let deadStockValue = 0; let deadStockItemsCount = 0;
            currentAlerts = { expired: 0, near_expiry: 0, stagnant: 0 };
            const nearExpiryThresholdAlert = 30; const stagnantThresholdAlert = 365;
            const dataToScan = filteredData.length > 0 || (filterMakerInput?.value || filterWholesalerInput?.value || filterLastOutDateStartInput?.value || filterLastOutDateEndInput?.value || filterExpiryDateStartInput?.value || filterExpiryDateEndInput?.value || filterKeywordInput?.value) ? filteredData : inventoryData;
            dataToScan.forEach(item => {
                const status = judgeDeadStock(item);
                const itemValue = parseFloat(item[HEADERS[IDX_VALUE]]) || 0;
                if (status !== 'none') {
                    deadStockItemsCount++; deadStockValue += itemValue;
                     if (status === 'expired') currentAlerts.expired++;
                     if (status === 'near_expired') currentAlerts.near_expiry++;
                     if (status === 'stagnant') currentAlerts.stagnant++;
                }
            });
            if(totalDeadStockValueEl) totalDeadStockValueEl.textContent = deadStockValue.toLocaleString();
            if(totalDeadStockItemsEl) totalDeadStockItemsEl.textContent = deadStockItemsCount;

            if(alertOutputEl) {
                alertOutputEl.innerHTML = ''; // Clear previous alerts
                if (currentAlerts.expired > 0) {
                    const p = document.createElement('p');
                    p.className = 'alert alert-expired'; // ★ Use CSS class
                    p.textContent = `【！】期限切れ在庫が ${currentAlerts.expired} 品目あります。`; alertOutputEl.appendChild(p);
                }
                 if (currentAlerts.near_expiry > 0) {
                     const nearExpiryUrgent = dataToScan.filter(item => { const expiry = parseDate(item[HEADERS[IDX_EXPIRY]]); return expiry && judgeDeadStock(item) === 'near_expired' && daysDifference(expiry, today) <= nearExpiryThresholdAlert; }).length;
                     const p = document.createElement('p');
                     p.className = 'alert alert-near-expiry'; // ★ Use CSS class
                     let text = `【！】期限近接在庫 (${settings.nearExpiryDays}日以内) が ${currentAlerts.near_expiry} 品目あります。`;
                     if (nearExpiryUrgent > 0) text += ` (うち ${nearExpiryUrgent} 品目は ${nearExpiryThresholdAlert} 日以内)`;
                     p.textContent = text; alertOutputEl.appendChild(p);
                 }
                 if (currentAlerts.stagnant > 0) {
                     const stagnantLong = dataToScan.filter(item => { if (judgeDeadStock(item) !== 'stagnant') return false; const outputDays = parseStagnationDays(item[HEADERS[IDX_LAST_OUT_ST]]); const inputDays = parseStagnationDays(item[HEADERS[IDX_LAST_IN_ST]]); const days = outputDays !== null ? outputDays : inputDays; return days !== null && days >= stagnantThresholdAlert; }).length;
                     const p = document.createElement('p');
                     p.className = 'alert alert-stagnant'; // ★ Use CSS class
                     let text = `【！】長期停滞在庫 (${settings.stagnantDays}日以上) が ${currentAlerts.stagnant} 品目あります。`;
                     if (stagnantLong > 0) text += ` (うち ${stagnantLong} 品目は ${stagnantThresholdAlert} 日以上)`;
                     p.textContent = text; alertOutputEl.appendChild(p);
                 }
                 if (alertOutputEl.innerHTML === '') {
                      const p = document.createElement('p'); p.className = 'no-alerts'; // ★ Use CSS class
                      p.textContent = '現在、アラート対象の在庫はありません。'; alertOutputEl.appendChild(p);
                 }
            }
        }

        function updateAlertFilterStyles() {
            // アラートフィルターボタンのスタイル更新 (CSSクラス 'filter-active' を使用)
            alertFilterButtons.forEach(button => {
                if (button) {
                    const filterValue = button.dataset.statusFilter;
                    if ((!currentStatusFilter && !filterValue) || (currentStatusFilter === filterValue)) {
                        button.classList.add('filter-active');
                    } else {
                        button.classList.remove('filter-active');
                    }
                }
            });
        }

        // --- Export Functions ---
        function exportDeadStockCsv() { /* ... (変更なし) ... */ }
        function downloadCsv(csvContent, fileName) { /* ... (変更なし) ... */ }
        function getTimestamp() { /* ... (変更なし) ... */ }
        function exportToPdf() { /* ... (変更なし) ... */ }

        // --- Fullscreen Function ---
        function toggleFullScreen() { /* ... (変更なし) ... */ }

        // --- Settings Functions ---
        function loadSettings() { /* ... (変更なし, Nullチェックは前回追加済み) ... */ }
        function saveSettings() { /* ... (変更なし) ... */ }

        // --- Chart Functions ---
        function initializeCharts() { /* ... (変更なし) ... */ }
        function updateCharts() { /* ... (変更なし) ... */ }
        function createStatusChart(statusData) { /* ... (変更なし) ... */ }
        function createTopValueChart(topItemsData) { /* ... (変更なし) ... */ }
        function createStagnationDistributionChart(stagnationBins) { /* ... (変更なし) ... */ }
        function createMakerChart(makerData) { /* ... (変更なし) ... */ }

        // --- Static Advice Box Logic ---
        const adviceList = [ /* ... (変更なし) ... */ ];
        function updateAdviceBox() { /* ... (変更なし) ... */ }

        // --- AI Function ---
        async function getAiAdvice() { /* ... (変更なし, OpenAI gpt-4o-mini設定済み) ... */ }

        // --- Event Listeners Setup ---
        function setupEventListeners() {
            console.log("Setting up event listeners...");

            // Nullチェックを追加してリスナーを設定
            const checkAndListen = (element, event, handler, name) => {
                if (element) {
                    element.addEventListener(event, handler);
                } else {
                    console.error(`${name} element not found! Listener not attached.`);
                }
            };

            // Data Input Buttons
            checkAndListen(processPasteBtn, 'click', async () => { /* ... paste inventory ... */ }, 'Process Paste Button');
            checkAndListen(processExpiryBtn, 'click', async () => { /* ... paste expiry ... */ }, 'Process Expiry Button');

            // Filter/Search Buttons
            checkAndListen(applyFilterBtn, 'click', () => { currentStatusFilter = null; applyFiltersAndSearch(); }, 'Apply Filter Button');
            checkAndListen(resetFilterBtn, 'click', () => {
                if(filterMakerInput) filterMakerInput.value = ''; if(filterWholesalerInput) filterWholesalerInput.value = '';
                if(filterLastOutDateStartInput) filterLastOutDateStartInput.value = ''; if(filterLastOutDateEndInput) filterLastOutDateEndInput.value = '';
                if(filterExpiryDateStartInput) filterExpiryDateStartInput.value = ''; if(filterExpiryDateEndInput) filterExpiryDateEndInput.value = '';
                if(filterKeywordInput) filterKeywordInput.value = '';
                currentStatusFilter = null; applyFiltersAndSearch();
            }, 'Reset Filter Button');

            // Export Buttons
            checkAndListen(exportCsvBtn, 'click', exportDeadStockCsv, 'Export CSV Button');
            checkAndListen(exportPdfBtn, 'click', exportToPdf, 'Export PDF Button');

            // Fullscreen Button
            checkAndListen(fullscreenBtn, 'click', toggleFullScreen, 'Fullscreen Button');

            // Settings Area
            checkAndListen(settingsBtn, 'click', () => settingsArea?.classList.toggle('hidden'), 'Settings Button');
            checkAndListen(closeSettingsBtn, 'click', () => settingsArea?.classList.add('hidden'), 'Close Settings Button');
            checkAndListen(saveSettingsBtn, 'click', saveSettings, 'Save Settings Button');

            // AI Advisor Area
            checkAndListen(aiAdvisorBtn, 'click', () => { aiAdvisorArea?.classList.toggle('hidden'); if(getAdviceBtn) getAdviceBtn.disabled = !apiKeyInput?.value?.trim(); }, 'AI Advisor Button');
            checkAndListen(closeAiBtn, 'click', () => aiAdvisorArea?.classList.add('hidden'), 'Close AI Button');
            checkAndListen(getAdviceBtn, 'click', getAiAdvice, 'Get Advice Button');
            checkAndListen(apiKeyInput, 'input', () => { if(getAdviceBtn) getAdviceBtn.disabled = !apiKeyInput.value.trim(); }, 'API Key Input');

             // Alert Filter Buttons
             alertFilterButtons.forEach(button => {
                 if (button) {
                     button.addEventListener('click', (e) => {
                         const status = e.target.dataset.statusFilter;
                         currentStatusFilter = status || null;
                         console.log("Status filter set to:", currentStatusFilter);
                         applyFiltersAndSearch();
                     });
                 } else { console.error("One of the alert filter buttons not found!"); }
             });

            console.log("Event listeners setup complete.");
        }

    }); // End DOMContentLoaded
} catch (error) {
    console.error("重大な初期化エラーが発生しました:", error);
    alert("スクリプトの初期化中に重大なエラーが発生しました。開発者コンソールを確認してください。");
}
