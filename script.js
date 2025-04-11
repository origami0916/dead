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
        let currentStatusFilter = null;

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

        // --- Function Definitions (Moved before initialization calls) ---

        function updateCurrentDate() {
            // 現在の日付をフッターに表示
             if (currentDateEl) {
                const options = { year: 'numeric', month: 'long', day: 'numeric' };
                currentDateEl.textContent = today.toLocaleDateString('ja-JP', options);
             } else {
                 console.error("Current Date Element not found!");
             }
        }

        function parsePastedData(dataString, expectedHeaders) { /* ... (変更なし) ... */ }
        function parseDate(dateString) { /* ... (変更なし) ... */ }
        function parseStagnationDays(valueString) { /* ... (変更なし) ... */ }
        function daysDifference(date1, date2) { /* ... (変更なし) ... */ }
        function judgeDeadStock(item) { /* ... (変更なし) ... */ }
        function renderTable() { /* ... (変更なし) ... */ }
        function handleSortClick(event) { /* ... (変更なし) ... */ }
        function sortData() { /* ... (変更なし) ... */ }
        function applyFiltersAndSearch() { /* ... (変更なし) ... */ }
        function updateSummaryAndAlerts() { /* ... (変更なし) ... */ }
        function updateAlertFilterStyles() { /* ... (変更なし) ... */ }
        function exportDeadStockCsv() { /* ... (変更なし) ... */ }
        function downloadCsv(csvContent, fileName) { /* ... (変更なし) ... */ }
        function getTimestamp() { /* ... (変更なし) ... */ }
        function exportToPdf() { /* ... (変更なし) ... */ }
        function toggleFullScreen() { /* ... (変更なし) ... */ }
        function loadSettings() { /* ... (変更なし) ... */ }
        function saveSettings() { /* ... (変更なし) ... */ }
        function initializeCharts() { /* ... (変更なし) ... */ }
        function updateCharts() { /* ... (変更なし) ... */ }
        function createStatusChart(statusData) { /* ... (変更なし) ... */ }
        function createTopValueChart(topItemsData) { /* ... (変更なし) ... */ }
        function createStagnationDistributionChart(stagnationBins) { /* ... (変更なし) ... */ }
        function createMakerChart(makerData) { /* ... (変更なし) ... */ }
        const adviceList = [ /* ... (変更なし) ... */ ];
        function updateAdviceBox() { /* ... (変更なし) ... */ }
        async function getAiAdvice() { /* ... (変更なし) ... */ }

        // --- Initialization ---
        console.log("Initializing...");
        loadSettings();
        updateCurrentDate(); // ★ Call after function definition ★
        setupEventListeners(); // イベントリスナー設定呼び出し
        initializeCharts();
        updateAdviceBox(); // Initial advice box update
        console.log("Initialization complete.");


        // --- Event Listeners Setup ---
        function setupEventListeners() {
            console.log("Setting up event listeners...");

            const checkAndListen = (element, event, handler, name) => { /* ... (変更なし) ... */ };

            // Data Input Buttons
            checkAndListen(processPasteBtn, 'click', async () => { /* ... paste inventory ... */ }, 'Process Paste Button');
            checkAndListen(processExpiryBtn, 'click', async () => { /* ... paste expiry ... */ }, 'Process Expiry Button');

            // Filter/Search Buttons
            checkAndListen(applyFilterBtn, 'click', () => { currentStatusFilter = null; applyFiltersAndSearch(); }, 'Apply Filter Button');
            checkAndListen(resetFilterBtn, 'click', () => { /* ... reset inputs ... */ currentStatusFilter = null; applyFiltersAndSearch(); }, 'Reset Filter Button');

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
