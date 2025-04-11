document.addEventListener('DOMContentLoaded', () => {
    // --- グローバル変数定義 ---
    let inventoryData = []; // 元の在庫データ
    let filteredData = []; // フィルタリング/ソート後のデータ
    let currentData = []; // 現在ページに表示するデータ

    // ページネーション関連
    let currentPage = 1;
    let itemsPerPage = 50; // デフォルトは50件表示

    // ソート関連
    let currentSortColumn = null;
    let isSortAscending = true;

    // 危険度ランク関連
    let expiryWeightParam = 50; // デフォルトの有効期限重視度 (0-100)

    // AI関連
    let userApiKey = null;

    // Chart.js インスタンス
    let categoryChartInstance = null;
    let makerChartInstance = null;
    let rankingChartInstance = null;

    // --- DOM要素取得 ---
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
    const inventoryTableBody = document.getElementById('inventory-table').querySelector('tbody');
    const inventoryTableHeader = document.getElementById('inventory-table').querySelector('thead');
    const loadingIndicator = document.getElementById('loading-indicator');
    const itemCountEl = document.getElementById('itemCount');
    const prevPageButton = document.getElementById('prevPage');
    const nextPageButton = document.getElementById('nextPage');
    const pageInfoEl = document.getElementById('pageInfo');
    const itemsPerPageSelect = document.getElementById('itemsPerPage');

    // 分析レポート要素
    const rankingTypeSelect = document.getElementById('rankingType');
    const rankingTopNSelect = document.getElementById('rankingTopN');
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
    async function initializeApp() {
        showLoading();
        try {
            const response = await fetch('inventory_data.json'); // JSONファイルのパス
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            inventoryData = await response.json();

            // データの前処理（日付変換、危険度ランク計算など）
            preprocessData();

            // フィルタリング用選択肢を生成
            populateFilterOptions();

             // 危険度ランクスライダーの初期設定
            updateWeightLabel();

            // 初期表示
            applyFiltersAndSort(); // 初回は全データ表示
            displayDashboard();

        } catch (error) {
            console.error('データの読み込みに失敗しました:', error);
            alert('在庫データの読み込みに失敗しました。\nファイルが存在するか、形式が正しいか確認してください。');
            inventoryTableBody.innerHTML = `<tr><td colspan="11" class="error-message">データ読み込みエラー</td></tr>`; // Error message in table
        } finally {
            hideLoading();
        }
    }

    // --- データ前処理 ---
    function preprocessData() {
        const currentDate = new Date(); // 現在日時を取得

        inventoryData = inventoryData.map(item => {
            // 日付オブジェクトの生成 (不正な日付はnullに)
            const expiryDate = parseDate(item['有効期限']);
            const lastOutDate = parseDate(item['最終出庫日']);
            const lastInDate = parseDate(item['最終入庫(停滞)']); // カラム名修正

            // 滞留日数と残り日数の計算
            const stagnationDays = calculateStagnationDays(currentDate, lastOutDate);
            const remainingDays = calculateRemainingDays(currentDate, expiryDate);
            const lastInDays = calculateStagnationDays(currentDate, lastInDate); // 最終入庫からの経過日数


             // 在庫金額を数値に変換（失敗したら0）
            const itemValue = parseFloat(item['在庫金額(税別)']) || 0;
            const itemQuantity = parseFloat(item['在庫数量']) || 0;

            // 危険度ランク計算
            const dangerRank = calculateDangerRank(remainingDays, stagnationDays, expiryWeightParam);

            // 数値でない場合に備える
             const stagnationDisplay = isNaN(stagnationDays) || stagnationDays > 9000 ? '---' : stagnationDays;
             const remainingDisplay = isNaN(remainingDays) ? '---' : remainingDays;
             const lastInDisplay = isNaN(lastInDays) || lastInDays > 9000 ? '---' : lastInDays;


            return {
                ...item,
                '有効期限Date': expiryDate,
                '最終出庫日Date': lastOutDate,
                '最終入庫日Date': lastInDate, // カラム名修正
                '在庫金額Numeric': itemValue,
                '在庫数量Numeric': itemQuantity,
                '滞留日数Numeric': stagnationDays,
                '残り日数Numeric': remainingDays,
                '最終入庫日数Numeric': lastInDays, // カラム名修正
                '危険度ランク': dangerRank,
                 // 表示用の整形済みデータ
                '滞留日数': stagnationDisplay, // 元の"最終出庫(停滞)"列に上書き
                '有効期限': expiryDate ? formatDate(expiryDate) : '---', // 元の"有効期限"列に上書き
                '最終入庫(停滞)': lastInDisplay // 元の列に上書き
            };
        });
         //console.log("Preprocessed Data Sample:", inventoryData[0]);
    }

    // 日付文字列をDateオブジェクトに変換 (YYYY/MM/DD または YYYY-MM-DD を想定)
    function parseDate(dateString) {
        if (!dateString || typeof dateString !== 'string') return null;
        // 年だけのデータ("2019"など)は日付として無効とする
        if (/^\d{4}$/.test(dateString.trim())) return null;
        // '-' 区切りを '/' に統一
        const formattedDateString = dateString.trim().replace(/-/g, '/');
        const date = new Date(formattedDateString);
        // 不正な日付でないかチェック (例: "2023/02/30")
        return isNaN(date.getTime()) ? null : date;
    }

     // 日付オブジェクトをYYYY/MM/DD形式の文字列に変換
    function formatDate(date) {
        if (!date || !(date instanceof Date) || isNaN(date.getTime())) return '---';
        const y = date.getFullYear();
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const d = String(date.getDate()).padStart(2, '0');
        return `${y}/${m}/${d}`;
    }

    // 滞留日数計算
    function calculateStagnationDays(currentDate, lastOutDate) {
        if (!lastOutDate || !(lastOutDate instanceof Date) || isNaN(lastOutDate.getTime())) return 9999; // 出庫日なし or 不正な日付
        const diffTime = currentDate.getTime() - lastOutDate.getTime();
        return Math.max(0, Math.floor(diffTime / (1000 * 60 * 60 * 24)));
    }

    // 残り日数計算
    function calculateRemainingDays(currentDate, expiryDate) {
        if (!expiryDate || !(expiryDate instanceof Date) || isNaN(expiryDate.getTime())) return NaN; // 有効期限なし or 不正
        // 日付のみで比較するため、時刻部分をリセット
        const today = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate());
        const expiry = new Date(expiryDate.getFullYear(), expiryDate.getMonth(), expiryDate.getDate());
        const diffTime = expiry.getTime() - today.getTime();
        return Math.ceil(diffTime / (1000 * 60 * 60 * 24)); // 切り上げで残り日数を計算
    }


    // --- 危険度ランク計算 ---
    function calculateDangerRank(remainingDays, stagnationDays, weightParam) {
        // 有効期限スコア (0-100)
        let scoreExpiry = 0;
        if (isNaN(remainingDays)) { // 有効期限不明は中間のリスクとするか？ここでは0点
            scoreExpiry = 0; // または 50 などに設定も可能
        } else if (remainingDays < 0) { // 期限切れ
            scoreExpiry = 100;
        } else if (remainingDays <= 30) {
            scoreExpiry = 100;
        } else if (remainingDays <= 90) {
            scoreExpiry = 75;
        } else if (remainingDays <= 180) {
            scoreExpiry = 50;
        } else if (remainingDays <= 365) {
            scoreExpiry = 25;
        } else {
            scoreExpiry = 0;
        }

        // 滞留日数スコア (0-100)
        let scoreStagnation = 0;
         if (isNaN(stagnationDays) || stagnationDays >= 730) { // 2年以上 or 出庫日なし
            scoreStagnation = 100;
        } else if (stagnationDays >= 365) { // 1年～2年未満
            scoreStagnation = 90;
        } else if (stagnationDays >= 180) { // 半年～1年未満
            scoreStagnation = 60;
        } else if (stagnationDays >= 90) { // 3ヶ月～半年未満
            scoreStagnation = 30;
        } else { // 3ヶ月未満
            scoreStagnation = 0;
        }

        // 重み付け
        const weightExpiry = weightParam / 100;
        const weightStagnation = 1.0 - weightExpiry;

        // 総合スコア
        const totalScore = (scoreExpiry * weightExpiry) + (scoreStagnation * weightStagnation);

        // ランク変換 (1-10)
        return Math.max(1, Math.ceil(totalScore / 10));
    }

    // 危険度ランク再計算
    function recalculateAllDangerRanks() {
        inventoryData = inventoryData.map(item => ({
            ...item,
            危険度ランク: calculateDangerRank(item.残り日数Numeric, item.滞留日数Numeric, expiryWeightParam)
        }));
        applyFiltersAndSort(); // 表示を更新
        displayDashboard(); // ダッシュボードも更新
    }

    // 危険度スライダーのラベル更新
    function updateWeightLabel() {
        const expiryW = expiryWeightSlider.value;
        const stagnationW = 100 - expiryW;
        weightLabel.textContent = `期限重視: ${expiryW}% / 滞留重視: ${stagnationW}%`;
    }


    // --- フィルタリングとソート ---
    function applyFiltersAndSort() {
        const searchTerm = searchInput.value.toLowerCase().trim();
        const selectedCategory = categoryFilter.value;
        const selectedMaker = makerFilter.value;
        const selectedRank = dangerRankFilter.value;

        filteredData = inventoryData.filter(item => {
            const rankMatch = selectedRank === "" || item.危険度ランク == selectedRank;
            const categoryMatch = selectedCategory === "" || item.薬品種別 === selectedCategory;
            const makerMatch = selectedMaker === "" || item.メーカー === selectedMaker;

            const searchMatch = searchTerm === "" ||
                item.薬品名称?.toLowerCase().includes(searchTerm) ||
                item.フリガナ?.toLowerCase().includes(searchTerm) ||
                item.メーカー?.toLowerCase().includes(searchTerm) ||
                item.YJコード?.includes(searchTerm) ||
                item.薬品コード?.includes(searchTerm);


            return rankMatch && categoryMatch && makerMatch && searchMatch;
        });

        // ソート処理
        if (currentSortColumn) {
            filteredData.sort((a, b) => {
                let valA = a[currentSortColumn];
                let valB = b[currentSortColumn];

                // 特定のカラムは数値や日付として比較
                 const numericColumns = ['在庫数量Numeric', '在庫金額Numeric', '滞留日数Numeric', '残り日数Numeric','最終入庫日数Numeric', '危険度ランク'];
                 const dateColumns = ['有効期限Date', '最終出庫日Date', '最終入庫日Date'];


                 if (numericColumns.includes(currentSortColumn)) {
                    valA = parseFloat(valA) || 0;
                    valB = parseFloat(valB) || 0;
                } else if (dateColumns.includes(currentSortColumn)) {
                    // null や Invalid Date は比較のために最小値/最大値を与える
                    valA = (valA instanceof Date && !isNaN(valA)) ? valA.getTime() : (isSortAscending ? Infinity : -Infinity);
                    valB = (valB instanceof Date && !isNaN(valB)) ? valB.getTime() : (isSortAscending ? Infinity : -Infinity);
                 } else {
                    // 文字列比較 (null/undefinedを考慮)
                    valA = String(valA ?? '').toLowerCase();
                    valB = String(valB ?? '').toLowerCase();
                }


                if (valA < valB) {
                    return isSortAscending ? -1 : 1;
                }
                if (valA > valB) {
                    return isSortAscending ? 1 : -1;
                }
                return 0;
            });
        }

        currentPage = 1; // フィルタ/ソート後は1ページ目に戻る
        updatePagination();
        displayCurrentPageData();
        updateItemCount();
    }

    // --- ページネーション ---
    function updatePagination() {
         const totalItems = filteredData.length;
         const totalPages = itemsPerPage === 'all' ? 1 : Math.ceil(totalItems / itemsPerPage);

         // ページ番号が範囲外にならないように調整
         currentPage = Math.max(1, Math.min(currentPage, totalPages));

         pageInfoEl.textContent = `ページ ${currentPage} / ${totalPages}`;
         prevPageButton.disabled = currentPage === 1;
         nextPageButton.disabled = currentPage === totalPages || totalPages === 0;

         // 現在表示するデータのスライス
         if (itemsPerPage === 'all') {
             currentData = filteredData;
         } else {
             const start = (currentPage - 1) * itemsPerPage;
             const end = start + itemsPerPage;
             currentData = filteredData.slice(start, end);
         }
    }

     function displayCurrentPageData() {
         renderTable(currentData);
     }

     function updateItemCount() {
         itemCountEl.textContent = `全 ${filteredData.length} 件`;
     }

     function changeItemsPerPage() {
        const selectedValue = itemsPerPageSelect.value;
        if (selectedValue === 'all') {
            itemsPerPage = 'all';
        } else {
            itemsPerPage = parseInt(selectedValue, 10);
        }
        currentPage = 1; // ページネーションのリセット
        updatePagination();
        displayCurrentPageData(); // テーブル再描画
    }


    // --- テーブル描画 ---
    function renderTable(data) {
        inventoryTableBody.innerHTML = ''; // テーブル内容をクリア

        if (!data || data.length === 0) {
            inventoryTableBody.innerHTML = `<tr><td colspan="11" style="text-align: center;">該当するデータがありません。</td></tr>`;
            return;
        }

        const fragment = document.createDocumentFragment();
        data.forEach(item => {
            const row = document.createElement('tr');
            // 各危険度ランクに応じたクラスを追加
             row.classList.add(`rank-${item.危険度ランク}`);
             if (item.残り日数Numeric <= 30 && item.残り日数Numeric >= 0) row.classList.add('near-expiry');
             if (item.滞留日数Numeric >= 180 && item.滞留日数Numeric <= 9000) row.classList.add('stagnant'); // 9000は仮の上限

             // カラム定義: 表示したい順にキーを並べる
             // 危険度ランクはpreprocessDataで計算済み、滞留日数と有効期限も整形済み
             const columns = [
                '危険度ランク', '薬品名称', 'フリガナ', '在庫数量', '単位',
                '在庫金額(税別)', '滞留日数', '有効期限',
                '薬品種別', 'メーカー', '卸', '最終入庫(停滞)', 'YJコード', '薬品コード' // 後ろに追加
            ];

             columns.forEach(key => {
                const cell = document.createElement('td');
                // '在庫金額(税別)' は数値としてフォーマットする例
                if (key === '在庫金額(税別)') {
                     cell.textContent = formatCurrency(item['在庫金額Numeric']);
                     cell.style.textAlign = 'right';
                 } else if (key === '在庫数量') {
                    cell.textContent = item[key]; // 元データを使う
                    cell.style.textAlign = 'right';
                 } else if (key === '危険度ランク') {
                    cell.textContent = item[key];
                    cell.style.textAlign = 'center';
                    cell.style.fontWeight = 'bold';
                } else if (key === '滞留日数' || key === '最終入庫(停滞)') {
                    cell.textContent = item[key]; // 整形済みデータを使う
                    cell.style.textAlign = 'right';
                 } else {
                     cell.textContent = item[key] ?? ''; // null や undefined は空文字に
                 }
                row.appendChild(cell);
             });
            fragment.appendChild(row);
        });
        inventoryTableBody.appendChild(fragment);
    }

     // 数値をカンマ区切り通貨形式にフォーマット
    function formatCurrency(value) {
        if (isNaN(value) || value === null) return '---';
        return Math.round(value).toLocaleString(); // 小数点以下を四捨五入してカンマ区切り
    }

    // --- フィルター用選択肢生成 ---
    function populateFilterOptions() {
        const categories = [...new Set(inventoryData.map(item => item.薬品種別).filter(Boolean))].sort();
        const makers = [...new Set(inventoryData.map(item => item.メーカー).filter(Boolean))].sort();

        categories.forEach(cat => {
            const option = document.createElement('option');
            option.value = cat;
            option.textContent = cat;
            categoryFilter.appendChild(option);
        });

        makers.forEach(maker => {
            const option = document.createElement('option');
            option.value = maker;
            option.textContent = maker;
            makerFilter.appendChild(option);
        });
    }

     // --- テーブルヘッダーのソート処理 ---
     function handleSort(event) {
         const th = event.target.closest('th');
         if (!th || !th.dataset.sort) return;

         const sortKey = th.dataset.sort;

         if (currentSortColumn === sortKey) {
             isSortAscending = !isSortAscending;
         } else {
             currentSortColumn = sortKey;
             isSortAscending = true;
         }

         // ヘッダーのソート状態表示を更新 (オプション)
         updateSortIndicators(th);

         applyFiltersAndSort();
     }

     function updateSortIndicators(activeTh) {
        inventoryTableHeader.querySelectorAll('th').forEach(th => {
            th.classList.remove('sort-asc', 'sort-desc');
             th.removeAttribute('aria-sort');
        });
        if (activeTh) {
             const sortClass = isSortAscending ? 'sort-asc' : 'sort-desc';
             activeTh.classList.add(sortClass);
              activeTh.setAttribute('aria-sort', isSortAscending ? 'ascending' : 'descending');
        }
    }


    // --- ダッシュボード表示 ---
    function displayDashboard() {
        // KPI計算
        const totalItems = inventoryData.length;
        const totalValue = inventoryData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);

        // 注意在庫 (例: 危険度ランク 7以上)
         const cautionData = inventoryData.filter(item => item.危険度ランク >= 7);
         const cautionItems = cautionData.length;
         const cautionValue = cautionData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);


        // 滞留在庫 (例: 180日以上)
        const stagnantData = inventoryData.filter(item => item.滞留日数Numeric >= 180 && item.滞留日数Numeric < 9000);
        const stagnantItems = stagnantData.length;
        const stagnantValue = stagnantData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);

        // 期限切迫在庫 (例: 90日以内)
        const nearExpiryData = inventoryData.filter(item => !isNaN(item.残り日数Numeric) && item.残り日数Numeric <= 90 && item.残り日数Numeric >= 0);
        const nearExpiryItems = nearExpiryData.length;
        const nearExpiryValue = nearExpiryData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);

        // KPI表示更新
        totalItemsEl.textContent = totalItems.toLocaleString();
        totalValueEl.textContent = `¥${formatCurrency(totalValue)}`;
        cautionItemsEl.textContent = cautionItems.toLocaleString();
        cautionValueEl.textContent = `¥${formatCurrency(cautionValue)}`;
        stagnantItemsEl.textContent = stagnantItems.toLocaleString();
        stagnantValueEl.textContent = `¥${formatCurrency(stagnantValue)}`;
        nearExpiryItemsEl.textContent = nearExpiryItems.toLocaleString();
        nearExpiryValueEl.textContent = `¥${formatCurrency(nearExpiryValue)}`;

        // 注目リスト更新
        updateRankingList(stagnantWorstListEl, inventoryData, '滞留日数Numeric', false, 5, item => `${item.薬品名称} (${item.滞留日数}日)`);
        updateRankingList(nearExpiryWorstListEl, inventoryData.filter(item => !isNaN(item.残り日数Numeric) && item.残り日数Numeric >= 0), '残り日数Numeric', true, 5, item => `${item.薬品名称} (残${item.残り日数Numeric}日)`);
        updateRankingList(highValueListEl, inventoryData, '在庫金額Numeric', false, 5, item => `${item.薬品名称} (¥${formatCurrency(item.在庫金額Numeric)})`);


        // グラフ更新
        displayCategoryChart();
        displayMakerChart();
    }

    // ランキングリスト更新ヘルパー
    function updateRankingList(listElement, data, sortKey, ascending, count, displayFormatter) {
        const sortedData = [...data].sort((a, b) => {
            const valA = a[sortKey];
            const valB = b[sortKey];
            if (valA < valB) return ascending ? -1 : 1;
            if (valA > valB) return ascending ? 1 : -1;
            return 0;
        }).slice(0, count);

        listElement.innerHTML = '';
        if (sortedData.length === 0) {
            listElement.innerHTML = '<li>該当データなし</li>';
            return;
        }
        sortedData.forEach(item => {
            const li = document.createElement('li');
            li.textContent = displayFormatter(item);
            listElement.appendChild(li);
        });
    }


    // --- グラフ描画 (Chart.js) ---
    // グラフ共通オプション
    const chartOptions = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: {
                position: 'bottom',
            },
            tooltip: {
                callbacks: {
                    label: function(context) {
                        let label = context.dataset.label || '';
                        if (label) {
                            label += ': ';
                        }
                        if (context.parsed.y !== null) {
                             label += `¥${context.parsed.y.toLocaleString()}`; // 金額表示
                        } else if (context.parsed.x !== null && context.dataset.label === '金額') { // 横棒グラフの場合
                             label += `¥${context.parsed.x.toLocaleString()}`;
                         } else if (context.parsed.x !== null && context.dataset.label === '日数') {
                             label += `${context.parsed.x.toLocaleString()} 日`;
                         } else if (context.parsed.x !== null && context.dataset.label === '数量') {
                             label += `${context.parsed.x.toLocaleString()} ${context.dataset.unit || ''}`;
                        } else if (context.parsed !== null && typeof context.parsed === 'number') { // 円グラフ用
                            label += `¥${context.parsed.toLocaleString()}`;
                        }
                        return label;
                    }
                }
            }
        }
    };

    // 薬品種別グラフ
    function displayCategoryChart() {
        const categoryData = inventoryData.reduce((acc, item) => {
            const category = item.薬品種別 || '不明';
            acc[category] = (acc[category] || 0) + item.在庫金額Numeric;
            return acc;
        }, {});

        const labels = Object.keys(categoryData).sort((a,b) => categoryData[b] - categoryData[a]); // 金額が多い順
        const dataValues = labels.map(label => categoryData[label]);

        if (categoryChartInstance) {
            categoryChartInstance.destroy();
        }
        categoryChartInstance = new Chart(categoryChartCanvas, {
            type: 'doughnut', // ドーナツグラフに変更
            data: {
                labels: labels,
                datasets: [{
                    label: '在庫金額',
                    data: dataValues,
                    backgroundColor: [ // 色を適当に設定 (増やす必要あり)
                        '#4CAF50', '#2196F3', '#FFC107', '#E91E63', '#9C27B0',
                        '#00BCD4', '#FF5722', '#8BC34A', '#673AB7', '#FF9800'
                    ],
                    hoverOffset: 4
                }]
            },
            options: chartOptions
        });
    }

    // メーカー別グラフ
    function displayMakerChart() {
        const makerData = inventoryData.reduce((acc, item) => {
            const maker = item.メーカー || '不明';
            acc[maker] = (acc[maker] || 0) + item.在庫金額Numeric;
            return acc;
        }, {});

        const sortedMakers = Object.entries(makerData)
            .sort(([, a], [, b]) => b - a) // 金額が多い順
            .slice(0, 5); // 上位5件

        const labels = sortedMakers.map(([maker]) => maker);
        const dataValues = sortedMakers.map(([, value]) => value);

        if (makerChartInstance) {
            makerChartInstance.destroy();
        }
        makerChartInstance = new Chart(makerChartCanvas, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: '在庫金額',
                    data: dataValues,
                    backgroundColor: '#4CAF50',
                }]
            },
            options: { ...chartOptions, indexAxis: 'y' } // 横棒グラフ
        });
    }

    // ランキンググラフ
    function displayRankingChart() {
        const type = rankingTypeSelect.value;
        const topN = parseInt(rankingTopNSelect.value, 10);

        let sortKey = '';
        let ascending = false;
        let labelKey = '薬品名称';
        let valueKey = '';
        let chartLabel = '';
         let chartUnit = ''; // 単位表示用
         let indexAxis = 'y'; // デフォルトは横棒
         let chartType = 'bar';

         switch (type) {
            case 'amountDesc':
                sortKey = '在庫金額Numeric'; ascending = false; valueKey = '在庫金額Numeric'; chartLabel = '金額';
                break;
             case 'stagnationDesc':
                sortKey = '滞留日数Numeric'; ascending = false; valueKey = '滞留日数Numeric'; chartLabel = '日数';
                 // 滞留日数9999を除外してソート
                 filteredData = inventoryData.filter(item => item.滞留日数Numeric < 9000);
                 break;
             case 'expiryAsc':
                 sortKey = '残り日数Numeric'; ascending = true; valueKey = '残り日数Numeric'; chartLabel = '日数';
                 // 期限切れでないものだけを対象
                 filteredData = inventoryData.filter(item => !isNaN(item.残り日数Numeric) && item.残り日数Numeric >= 0);
                break;
             case 'quantityDesc':
                sortKey = '在庫数量Numeric'; ascending = false; valueKey = '在庫数量Numeric'; chartLabel = '数量'; labelKey='薬品名称'; chartUnit = ''; // 単位をデータから取れないので空
                break;
             case 'quantityAsc':
                sortKey = '在庫数量Numeric'; ascending = true; valueKey = '在庫数量Numeric'; chartLabel = '数量'; labelKey='薬品名称'; chartUnit = '';
                 // 在庫0を除外してソート
                filteredData = inventoryData.filter(item => item.在庫数量Numeric > 0);
                break;
            default:
                return;
        }


         const sortedData = [...filteredData].sort((a, b) => {
             let valA = a[sortKey];
             let valB = b[sortKey];
             // NaNや無効値を適切に扱う
              valA = isNaN(valA) ? (ascending ? Infinity : -Infinity) : valA;
              valB = isNaN(valB) ? (ascending ? Infinity : -Infinity) : valB;

             if (valA < valB) return ascending ? -1 : 1;
             if (valA > valB) return ascending ? 1 : -1;
             return 0;
         }).slice(0, topN);


        const labels = sortedData.map(item => item[labelKey]);
        const dataValues = sortedData.map(item => item[valueKey]);
         // 単位情報を付与（横棒グラフのツールチップ用）
        const units = chartLabel === '数量' ? sortedData.map(item => item['単位'] || '') : null;


        if (rankingChartInstance) {
            rankingChartInstance.destroy();
        }
        rankingChartInstance = new Chart(rankingChartCanvas, {
             type: chartType,
            data: {
                labels: labels.reverse(), // 横棒グラフ用にラベルを逆順に
                datasets: [{
                    label: chartLabel,
                    data: dataValues.reverse(), // データも逆順に
                    backgroundColor: '#2196F3',
                    unit: chartUnit // ツールチップ用
                }]
            },
             options: { ...chartOptions, indexAxis: indexAxis }
        });
    }


    // --- AIアドバイス機能 ---
    async function fetchAiAdvice() {
        if (!userApiKey) {
            showAiError("APIキーが設定されていません。");
            return;
        }
        if (!inventoryData || inventoryData.length === 0) {
             showAiError("分析対象の在庫データがありません。");
            return;
        }


        showAiLoading();
        hideAiError();
        aiAdviceOutput.textContent = ''; // 前のアドバイスをクリア
        copyAiAdviceButton.style.display = 'none';


        try {
            // データサマリー生成
            const summary = generateInventorySummary();

            // プロンプト生成
            const prompt = createPrompt(summary);

            // OpenAI API呼び出し
            const response = await fetch("https://api.openai.com/v1/chat/completions", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": `Bearer ${userApiKey}`
                },
                body: JSON.stringify({
                    model: "gpt-3.5-turbo", // or "gpt-4" if available
                    messages: [{ role: "user", content: prompt }],
                    temperature: 0.5, // 少し創造性を抑える
                    max_tokens: 1000 // 最大トークン数（適宜調整）
                }),
            });

            if (!response.ok) {
                const errorData = await response.json();
                console.error("OpenAI API Error:", errorData);
                throw new Error(`APIエラー: ${response.status} ${errorData.error?.message || '不明なエラー'}`);
            }

            const result = await response.json();
            const advice = result.choices[0]?.message?.content?.trim();

            if (advice) {
                aiAdviceOutput.textContent = advice; // 改行が反映されるようにtextContentを使用
                copyAiAdviceButton.style.display = 'inline-block'; // コピーボタン表示
            } else {
                showAiError("AIからのアドバイスを取得できませんでした。");
            }

        } catch (error) {
            console.error('AIアドバイス取得エラー:', error);
            showAiError(`エラーが発生しました: ${error.message}`);
        } finally {
            hideAiLoading();
        }
    }

    function generateInventorySummary() {
         // サマリー情報を生成するロジック
        const totalItems = inventoryData.length;
        const totalValue = inventoryData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);

        // 注意在庫 (例: 危険度ランク 7以上)
         const cautionData = inventoryData.filter(item => item.危険度ランク >= 7);
         const cautionItems = cautionData.length;
         const cautionValue = cautionData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);
         const topCaution = cautionData.sort((a, b) => b.危険度ランク - a.危険度ランク).slice(0, 5).map(i => `${i.薬品名称} (ランク${i.危険度ランク}, 金額¥${formatCurrency(i.在庫金額Numeric)})`).join('\n - ');


         // 滞留在庫 (例: 180日以上)
        const stagnantData = inventoryData.filter(item => item.滞留日数Numeric >= 180 && item.滞留日数Numeric < 9000);
        const stagnantItems = stagnantData.length;
        const stagnantValue = stagnantData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);
         const topStagnant = stagnantData.sort((a, b) => b.滞留日数Numeric - a.滞留日数Numeric).slice(0, 5).map(i => `${i.薬品名称} (${i.滞留日数}日, 金額¥${formatCurrency(i.在庫金額Numeric)})`).join('\n - ');

        // 期限切迫在庫 (例: 90日以内)
        const nearExpiryData = inventoryData.filter(item => !isNaN(item.残り日数Numeric) && item.残り日数Numeric <= 90 && item.残り日数Numeric >= 0);
        const nearExpiryItems = nearExpiryData.length;
        const nearExpiryValue = nearExpiryData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);
         const topNearExpiry = nearExpiryData.sort((a, b) => a.残り日数Numeric - b.残り日数Numeric).slice(0, 5).map(i => `${i.薬品名称} (残${i.残り日数Numeric}日, 金額¥${formatCurrency(i.在庫金額Numeric)})`).join('\n - ');

         // 高額在庫
         const highValueData = inventoryData.filter(item => item.在庫金額Numeric > 0); // 金額0を除外
         const topHighValue = highValueData.sort((a,b) => b.在庫金額Numeric - a.在庫金額Numeric).slice(0,5).map(i => `${i.薬品名称} (¥${formatCurrency(i.在庫金額Numeric)})`).join('\n - ');

         // 欠品注意 (例: 在庫数量 5以下で過去30日以内に出庫あり - 簡易版)
         const lowStockData = inventoryData.filter(item => item.在庫数量Numeric <= 5 && item.滞留日数Numeric <= 30 && item.在庫数量Numeric > 0);
         const topLowStock = lowStockData.sort((a,b) => a.在庫数量Numeric - b.在庫数量Numeric).slice(0,5).map(i => `${i.薬品名称} (残${i.在庫数量} ${i.単位 || ''})`).join('\n - ');


         return `
## 在庫サマリー (${formatDate(new Date())}時点)

**基本情報:**
- 総在庫品目数: ${totalItems.toLocaleString()} 品目
- 総在庫金額 (税別): ¥${formatCurrency(totalValue)}

**要注意在庫 (危険度ランク7以上):**
- 該当品目数: ${cautionItems.toLocaleString()} 品目
- 合計金額: ¥${formatCurrency(cautionValue)}
- 上位リスト (抜粋):\n - ${topCaution || '該当なし'}

**滞留在庫 (180日以上出庫なし):**
- 該当品目数: ${stagnantItems.toLocaleString()} 品目
- 合計金額: ¥${formatCurrency(stagnantValue)}
- 上位リスト (抜粋):\n - ${topStagnant || '該当なし'}

**期限切迫在庫 (残り90日以内):**
- 該当品目数: ${nearExpiryItems.toLocaleString()} 品目
- 合計金額: ¥${formatCurrency(nearExpiryValue)}
- 上位リスト (抜粋):\n - ${topNearExpiry || '該当なし'}

**高額在庫 (上位抜粋):**
- 上位リスト (抜粋):\n - ${topHighValue || '該当なし'}

**欠品注意在庫 (在庫僅少かつ最近出庫あり - 抜粋):**
- 上位リスト (抜粋):\n - ${topLowStock || '該当なし'}

**危険度ランク設定:**
- 現在の重み付け: 期限重視 ${expiryWeightParam}%, 滞留重視 ${100 - expiryWeightParam}%
`;
    }

    function createPrompt(summary) {
        return `
# あなたの役割
あなたは経験豊富な医薬品在庫管理コンサルタントです。提供された在庫データのサマリーを分析し、課題を特定し、具体的な改善アクションプランを提案してください。

# 指示
以下の在庫データサマリーに基づき、以下の観点から分析とアドバイスを行ってください。
1.  **全体的な在庫状況の評価:** 在庫量は適正か、偏りはないか。特に金額面での課題は？
2.  **滞留在庫の問題点:** 滞留の原因を推測し、削減のための具体的なアクション（例：使用促進、不動在庫化の検討、返品交渉の可能性など）を提案してください。リストアップされた品目にも触れてください。
3.  **有効期限管理:** 期限切れリスクの高い品目への対策、廃棄ロス削減策を具体的に提案してください。リストアップされた品目にも触れてください。
4.  **在庫コスト:** 高額在庫の適正化、コスト削減につながる提案をしてください。リストアップされた品目にも触れてください。
5.  **欠品リスク:** 在庫僅少品目に対する注意喚起や、欠品を防ぐための対策案があれば言及してください。
6.  **その他:** 危険度ランクの設定（重み付け）についてもコメントがあればお願いします。上記以外に気づいた点や改善提案があれば自由に記述してください。
アドバイスは具体的で、実行可能な内容を心がけてください。

# 在庫データサマリー
${summary}

# 出力形式
分析結果と具体的なアドバイスを明確に分けて、箇条書きなどで分かりやすく記述してください。
`;
    }


    function saveApiKey() {
        const key = apiKeyInput.value.trim();
        if (key && key.startsWith('sk-')) {
            userApiKey = key;
            apiKeyStatus.textContent = 'キーが一時的に保存されました。';
            apiKeyStatus.style.color = 'green';
            getAiAdviceButton.disabled = false; // アドバイスボタン有効化
            apiKeyInput.value = ''; // 入力欄をクリア
             //console.log("API Key temporarily saved.");
        } else {
            userApiKey = null;
            apiKeyStatus.textContent = '無効なキー形式です。';
            apiKeyStatus.style.color = 'red';
            getAiAdviceButton.disabled = true; // アドバイスボタン無効化
        }
    }

     function copyAdviceToClipboard() {
        if (aiAdviceOutput.textContent) {
            navigator.clipboard.writeText(aiAdviceOutput.textContent)
                .then(() => {
                    alert('アドバイスをクリップボードにコピーしました。');
                })
                .catch(err => {
                    console.error('コピーに失敗しました:', err);
                    alert('コピーに失敗しました。');
                });
        }
    }


    function showAiLoading() { aiLoading.style.display = 'block'; }
    function hideAiLoading() { aiLoading.style.display = 'none'; }
    function showAiError(message) { aiErrorOutput.textContent = message; aiErrorOutput.style.display = 'block'; }
    function hideAiError() { aiErrorOutput.style.display = 'none'; }


    // --- ローディング表示 ---
    function showLoading() { if(loadingIndicator) loadingIndicator.style.display = 'block'; }
    function hideLoading() { if(loadingIndicator) loadingIndicator.style.display = 'none'; }


    // --- イベントリスナー設定 ---

    // タブ切り替え
    tabButtons.forEach(button => {
        button.addEventListener('click', () => {
            const targetTab = button.dataset.tab;

            tabContents.forEach(content => {
                content.classList.remove('active');
            });
            document.getElementById(targetTab).classList.add('active');

            tabButtons.forEach(btn => {
                btn.classList.remove('active');
            });
            button.classList.add('active');

             // 分析レポートタブが選択されたら、ランキンググラフを初期表示
             if (targetTab === 'analysis-reports') {
                 displayRankingChart();
            }
             // ダッシュボードが表示されたら更新（必要なら）
            // if (targetTab === 'dashboard') {
            //     displayDashboard();
            // }
        });
    });

    // 在庫一覧: 検索、フィルタ、リセット
    searchInput.addEventListener('input', applyFiltersAndSort);
    categoryFilter.addEventListener('change', applyFiltersAndSort);
    makerFilter.addEventListener('change', applyFiltersAndSort);
    dangerRankFilter.addEventListener('change', applyFiltersAndSort);
    resetFiltersButton.addEventListener('click', () => {
        searchInput.value = '';
        categoryFilter.value = '';
        makerFilter.value = '';
        dangerRankFilter.value = '';
        currentSortColumn = null; // ソートもリセット
        isSortAscending = true;
         updateSortIndicators(null); // ソート表示もリセット
        applyFiltersAndSort();
    });

     // 在庫一覧: テーブルヘッダーソート
    inventoryTableHeader.addEventListener('click', handleSort);

    // 在庫一覧: ページネーション
     prevPageButton.addEventListener('click', () => {
        if (currentPage > 1) {
            currentPage--;
            updatePagination();
            displayCurrentPageData();
        }
    });
     nextPageButton.addEventListener('click', () => {
         const totalItems = filteredData.length;
         const totalPages = itemsPerPage === 'all' ? 1 : Math.ceil(totalItems / itemsPerPage);
        if (currentPage < totalPages) {
            currentPage++;
            updatePagination();
            displayCurrentPageData();
        }
    });
     itemsPerPageSelect.addEventListener('change', changeItemsPerPage);


     // 危険度ランクスライダー
     expiryWeightSlider.addEventListener('input', updateWeightLabel); // ラベルをリアルタイム更新
     expiryWeightSlider.addEventListener('change', () => { // スライダー操作完了時に再計算
         expiryWeightParam = parseInt(expiryWeightSlider.value, 10);
         recalculateAllDangerRanks();
     });

     // 分析レポート: ランキング種類/件数変更
     rankingTypeSelect.addEventListener('change', displayRankingChart);
     rankingTopNSelect.addEventListener('change', displayRankingChart);


    // AIアドバイス
    saveApiKeyButton.addEventListener('click', saveApiKey);
    getAiAdviceButton.addEventListener('click', fetchAiAdvice);
     copyAiAdviceButton.addEventListener('click', copyAdviceToClipboard);

    // --- アプリケーション初期化実行 ---
    initializeApp();

}); // End of DOMContentLoaded
