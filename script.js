document.addEventListener('DOMContentLoaded', () => {
    // --- グローバル変数定義 ---
    let inventoryData = []; // 元の在庫データ (パース後)
    let filteredData = []; // フィルタリング/ソート後のデータ
    let currentData = []; // 現在ページに表示するデータ
    let originalHeaders = []; // TSVから読み込んだヘッダー

    // ページネーション関連
    let currentPage = 1;
    let itemsPerPage = 50; // デフォルトは50件表示

    // ソート関連
    let currentSortColumn = null; // ソート対象のキー (ヘッダー名)
    let currentSortProcessedKey = null; // ソートに使う処理済みキー (Numeric/Date)
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
    const dataInputSection = document.getElementById('data-input-section');
    const inventoryDataInput = document.getElementById('inventoryDataInput');
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
    const inventoryTable = document.getElementById('inventory-table'); // テーブル全体
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
    // アプリのメインロジックを初期化（データロード後に呼ばれる）
    function initializeAppLogic() {
        showLoading(); // 処理中表示
        try {
            // データの前処理（日付変換、危険度ランク計算など）
            preprocessData();

            // テーブルヘッダー生成
            renderTableHeader();

            // フィルタリング用選択肢を生成
            populateFilterOptions();

             // 危険度ランクスライダーの初期設定
            updateWeightLabel();

            // 初期表示
            applyFiltersAndSort(); // 初回は全データ表示
            displayDashboard(); // ダッシュボード表示

            // UI表示切り替え
            dataInputSection.style.display = 'none'; // 入力エリア非表示
            tabNav.style.display = 'flex'; // タブ表示
            mainContent.style.display = 'block'; // メインコンテンツ表示
            footer.style.display = 'block'; // フッター表示

        } catch (error) {
            console.error('アプリケーションの初期化中にエラー:', error);
            showDataError('データの処理中にエラーが発生しました。データ形式を確認してください。');
            // エラー時はUIを戻すなどの処理が必要な場合がある
            dataInputSection.style.display = 'block';
            tabNav.style.display = 'none';
            mainContent.style.display = 'none';
            footer.style.display = 'none';
        } finally {
             hideLoading();
        }
    }

    // --- データパーサー (TSV) ---
    function parseTSV(tsvString) {
        const lines = tsvString.trim().split('\n');
        if (lines.length < 2) {
            throw new Error("データが少なすぎます。ヘッダー行とデータ行が必要です。");
        }

        // ヘッダー行を取得し、前後の空白を削除
        originalHeaders = lines[0].split('\t').map(header => header.trim());
        const expectedHeaderCount = originalHeaders.length;

        const data = [];
        for (let i = 1; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue; // 空行はスキップ

            const values = line.split('\t').map(value => value.trim());

            // 列数がヘッダーと一致しない行はエラーまたはスキップ
            if (values.length !== expectedHeaderCount) {
                console.warn(`行 ${i + 1}: 列数がヘッダーと一致しません (${values.length}列、期待値${expectedHeaderCount}列)。スキップします。 Line: "${line}"`);
                continue; // スキップする場合
                // throw new Error(`行 ${i + 1}: 列数がヘッダーと一致しません。`); // エラーにする場合
            }

            const item = {};
            for (let j = 0; j < expectedHeaderCount; j++) {
                item[originalHeaders[j]] = values[j];
            }
            data.push(item);
        }

        if (data.length === 0) {
             throw new Error("有効なデータ行が見つかりませんでした。");
        }

        return data;
    }

    // --- データ前処理 ---
    function preprocessData() {
        const currentDate = new Date(); // 現在日時を取得

        // 必須カラムの存在チェック (キー名はTSVヘッダーに依存)
        const requiredKeys = ['有効期限', '最終出庫日', '在庫金額(税別)', '在庫数量'];
        for (const key of requiredKeys) {
            if (!originalHeaders.includes(key)) {
                throw new Error(`必須カラム "${key}" がデータに含まれていません。`);
            }
        }


        inventoryData = inventoryData.map(item => {
            // 日付オブジェクトの生成 (不正な日付はnullに)
            const expiryDate = parseDate(item['有効期限']);
            const lastOutDate = parseDate(item['最終出庫日']);
            // '最終入庫(停滞)' カラムが存在すれば処理
            const lastInDate = originalHeaders.includes('最終入庫(停滞)') ? parseDate(item['最終入庫(停滞)']) : null;

            // 滞留日数と残り日数の計算
            const stagnationDays = calculateStagnationDays(currentDate, lastOutDate);
            const remainingDays = calculateRemainingDays(currentDate, expiryDate);
            const lastInDays = calculateStagnationDays(currentDate, lastInDate); // 最終入庫からの経過日数

             // 在庫金額と数量を数値に変換（失敗したら0）
            const itemValue = parseFloat(String(item['在庫金額(税別)']).replace(/,/g, '')) || 0; // カンマ除去
            const itemQuantity = parseFloat(String(item['在庫数量']).replace(/,/g, '')) || 0; // カンマ除去

            // 危険度ランク計算
            const dangerRank = calculateDangerRank(remainingDays, stagnationDays, expiryWeightParam);

            // 数値でない場合に備える & 表示用整形
             const stagnationDisplay = isNaN(stagnationDays) || stagnationDays > 9000 ? '---' : stagnationDays;
             const remainingDisplay = isNaN(remainingDays) ? '---' : remainingDays;
             const lastInDisplay = isNaN(lastInDays) || lastInDays > 9000 ? '---' : lastInDays;

            // 処理済みデータを新しいキーで追加
            const processedItem = {
                ...item, // 元のデータも保持
                '有効期限Date': expiryDate,
                '最終出庫日Date': lastOutDate,
                '最終入庫日Date': lastInDate,
                '在庫金額Numeric': itemValue,
                '在庫数量Numeric': itemQuantity,
                '滞留日数Numeric': stagnationDays,
                '残り日数Numeric': remainingDays,
                '最終入庫日数Numeric': lastInDays,
                '危険度ランク': dangerRank,
                 // 表示用の整形済みデータ (元のキーを上書き)
                '滞留日数表示': stagnationDisplay,
                '有効期限表示': expiryDate ? formatDate(expiryDate) : '---',
                '最終入庫停滞表示': lastInDisplay
            };
             // 元のカラム名が存在すれば上書き (表示用)
             if (originalHeaders.includes('最終出庫(停滞)')) processedItem['最終出庫(停滞)'] = stagnationDisplay;
             if (originalHeaders.includes('有効期限')) processedItem['有効期限'] = processedItem['有効期限表示'];
             if (originalHeaders.includes('最終入庫(停滞)')) processedItem['最終入庫(停滞)'] = processedItem['最終入庫停滞表示'];

            return processedItem;
        });
         //console.log("Preprocessed Data Sample:", inventoryData[0]);
    }

    // 日付文字列をDateオブジェクトに変換 (YYYY/MM/DD または YYYY-MM-DD を想定)
    function parseDate(dateString) {
        if (!dateString || typeof dateString !== 'string') return null;
        const trimmed = dateString.trim();
        // 年だけのデータ("2019"など)や '-' は日付として無効とする
        if (/^\d{4}$/.test(trimmed) || trimmed === '-') return null;
        // '-' 区切りを '/' に統一
        const formattedDateString = trimmed.replace(/-/g, '/');
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
        if (isNaN(remainingDays)) {
            scoreExpiry = 0; // 有効期限不明はリスク0扱い
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

        // フィルタリング
        filteredData = inventoryData.filter(item => {
            const rankMatch = selectedRank === "" || item.危険度ランク == selectedRank;
            // 薬品種別、メーカーは元のTSVヘッダー名を使用
            const categoryKey = originalHeaders.find(h => h.includes('薬品種別')) || '薬品種別';
            const makerKey = originalHeaders.find(h => h.includes('メーカー')) || 'メーカー';
            const categoryMatch = selectedCategory === "" || item[categoryKey] === selectedCategory;
            const makerMatch = selectedMaker === "" || item[makerKey] === selectedMaker;

            // 検索対象カラムを動的に設定 (フリガナ、YJコードなども元のヘッダー名を使用)
            const nameKey = originalHeaders.find(h => h.includes('薬品名称')) || '薬品名称';
            const furiganaKey = originalHeaders.find(h => h.includes('フリガナ')) || 'フリガナ';
            const yjKey = originalHeaders.find(h => h.includes('YJコード')) || 'YJコード';
            const codeKey = originalHeaders.find(h => h.includes('薬品コード')) || '薬品コード';

            const searchMatch = searchTerm === "" ||
                item[nameKey]?.toLowerCase().includes(searchTerm) ||
                item[furiganaKey]?.toLowerCase().includes(searchTerm) ||
                item[makerKey]?.toLowerCase().includes(searchTerm) ||
                item[yjKey]?.includes(searchTerm) ||
                item[codeKey]?.includes(searchTerm);

            return rankMatch && categoryMatch && makerMatch && searchMatch;
        });

        // ソート処理
        if (currentSortProcessedKey) { // 処理済みのキーでソート
            filteredData.sort((a, b) => {
                let valA = a[currentSortProcessedKey];
                let valB = b[currentSortProcessedKey];

                // Dateオブジェクトの場合
                if (valA instanceof Date && valB instanceof Date) {
                    valA = isNaN(valA.getTime()) ? (isSortAscending ? Infinity : -Infinity) : valA.getTime();
                    valB = isNaN(valB.getTime()) ? (isSortAscending ? Infinity : -Infinity) : valB.getTime();
                }
                // 数値の場合 (NaNは最後に)
                else if (typeof valA === 'number' && typeof valB === 'number') {
                    valA = isNaN(valA) ? (isSortAscending ? Infinity : -Infinity) : valA;
                    valB = isNaN(valB) ? (isSortAscending ? Infinity : -Infinity) : valB;
                }
                // 文字列の場合 (null/undefinedを考慮)
                else {
                    valA = String(valA ?? '').toLowerCase();
                    valB = String(valB ?? '').toLowerCase();
                }

                if (valA < valB) return isSortAscending ? -1 : 1;
                if (valA > valB) return isSortAscending ? 1 : -1;
                return 0;
            });
        } else if (currentSortColumn) { // 元のキーでソート (フォールバック)
             filteredData.sort((a, b) => {
                 let valA = a[currentSortColumn];
                 let valB = b[currentSortColumn];
                 valA = String(valA ?? '').toLowerCase();
                 valB = String(valB ?? '').toLowerCase();
                 if (valA < valB) return isSortAscending ? -1 : 1;
                 if (valA > valB) return isSortAscending ? 1 : -1;
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
         const totalPages = itemsPerPage === 'all' ? 1 : Math.max(1, Math.ceil(totalItems / itemsPerPage)); // 0件でも1ページ

         // ページ番号が範囲外にならないように調整
         currentPage = Math.max(1, Math.min(currentPage, totalPages));

         pageInfoEl.textContent = `ページ ${currentPage} / ${totalPages}`;
         prevPageButton.disabled = currentPage === 1;
         nextPageButton.disabled = currentPage === totalPages;

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
        itemsPerPage = (selectedValue === 'all') ? 'all' : parseInt(selectedValue, 10);
        currentPage = 1; // ページネーションのリセット
        updatePagination();
        displayCurrentPageData(); // テーブル再描画
    }


    // --- テーブル描画 ---
    // テーブルヘッダー生成
    function renderTableHeader() {
        const thead = inventoryTableHeader;
        thead.innerHTML = ''; // クリア
        const tr = document.createElement('tr');

        // 表示したいカラムとソートキーのマッピング
        // キー: 表示名, 値: 元のヘッダー名 or 処理済みキー
        const displayColumns = {
            '危険度': { sortKey: '危険度ランク', processedKey: '危険度ランク', align: 'center', isNumeric: true },
            '薬品名称': { sortKey: '薬品名称', processedKey: '薬品名称' },
            'フリガナ': { sortKey: 'フリガナ', processedKey: 'フリガナ' },
            '在庫数量': { sortKey: '在庫数量', processedKey: '在庫数量Numeric', align: 'right', isNumeric: true },
            '単位': { sortKey: '単位', processedKey: '単位', align: 'center' },
            '在庫金額(税別)': { sortKey: '在庫金額(税別)', processedKey: '在庫金額Numeric', align: 'right', isNumeric: true, isCurrency: true },
            '滞留日数': { sortKey: '最終出庫(停滞)', processedKey: '滞留日数Numeric', align: 'right', isNumeric: true }, // 元のヘッダー名でソートキー指定
            '有効期限': { sortKey: '有効期限', processedKey: '有効期限Date', align: 'center' }, // 元のヘッダー名でソートキー指定
            '薬品種別': { sortKey: '薬品種別', processedKey: '薬品種別' },
            'メーカー': { sortKey: 'メーカー', processedKey: 'メーカー' },
            '卸': { sortKey: '卸', processedKey: '卸' },
            '最終入庫(停滞)': { sortKey: '最終入庫(停滞)', processedKey: '最終入庫日数Numeric', align: 'right', isNumeric: true },
            'YJコード': { sortKey: 'YJコード', processedKey: 'YJコード' },
            '薬品コード': { sortKey: '薬品コード', processedKey: '薬品コード' },
        };

        // 元データに含まれるヘッダーに基づいて表示するカラムを決定
        originalHeaders.forEach(header => {
             // displayColumns に定義されているか、または基本的なカラムかチェック
             const columnDef = Object.entries(displayColumns).find(([_, def]) => def.sortKey === header)?.[1]
                             || Object.values(displayColumns).find(def => def.processedKey === header); // processedKey での検索も追加

             const displayName = Object.entries(displayColumns).find(([name, def]) => def.sortKey === header)?.[0] || header;

             if (columnDef) { // 表示対象カラムの場合
                const th = document.createElement('th');
                th.textContent = displayName;
                th.dataset.sort = header; // ソートの基準は元のヘッダー名
                th.dataset.processedKey = columnDef.processedKey; // 実際のソートに使うキー
                if (columnDef.align) {
                    th.style.textAlign = columnDef.align;
                }
                tr.appendChild(th);
             }
        });

         // 危険度ランクのカラムを手動で追加（TSVに無くても表示）
         if (!originalHeaders.includes('危険度ランク')) {
            const thRank = document.createElement('th');
            thRank.textContent = '危険度';
            thRank.dataset.sort = '危険度ランク'; // ソートキー
            thRank.dataset.processedKey = '危険度ランク'; // 処理済みキー
            thRank.style.textAlign = 'center';
            // 危険度ランクをどの位置に挿入するか？ 先頭に追加
            tr.insertBefore(thRank, tr.firstChild);
        }


        thead.appendChild(tr);
    }

    // テーブルボディ描画
    function renderTable(data) {
        inventoryTableBody.innerHTML = ''; // テーブル内容をクリア

        if (!data || data.length === 0) {
            inventoryTableBody.innerHTML = `<tr><td colspan="${inventoryTableHeader.rows[0]?.cells.length || 1}" style="text-align: center;">該当するデータがありません。</td></tr>`;
            return;
        }

        const fragment = document.createDocumentFragment();
        const displayHeaders = Array.from(inventoryTableHeader.rows[0].cells).map(th => th.dataset.sort || th.textContent); // 表示するヘッダーの順序を取得

        data.forEach(item => {
            const row = document.createElement('tr');
            // 各危険度ランクに応じたクラスを追加
             row.classList.add(`rank-${item.危険度ランク}`);
             if (item.残り日数Numeric <= 30 && item.残り日数Numeric >= 0) row.classList.add('near-expiry');
             if (item.滞留日数Numeric >= 180 && item.滞留日数Numeric < 9000) row.classList.add('stagnant');

             // ヘッダーの順序に合わせてセルを作成
             displayHeaders.forEach(headerKey => {
                 const cell = document.createElement('td');
                 let cellValue = '';
                 let textAlign = 'left';
                 let isCurrency = false;

                 // 表示する値とスタイルを決定
                 switch (headerKey) {
                     case '危険度ランク':
                         cellValue = item['危険度ランク'];
                         textAlign = 'center';
                         cell.style.fontWeight = 'bold';
                         break;
                     case '在庫金額(税別)':
                         cellValue = formatCurrency(item['在庫金額Numeric']);
                         textAlign = 'right';
                         isCurrency = true;
                         break;
                     case '在庫数量':
                         cellValue = item['在庫数量']; // 元の値を表示
                         textAlign = 'right';
                         break;
                     case '最終出庫(停滞)': // 表示用の滞留日数
                         cellValue = item['滞留日数表示'];
                         textAlign = 'right';
                         break;
                     case '有効期限': // 表示用の有効期限
                         cellValue = item['有効期限表示'];
                         textAlign = 'center';
                         break;
                     case '最終入庫(停滞)': // 表示用の最終入庫停滞日数
                         cellValue = item['最終入庫停滞表示'];
                         textAlign = 'right';
                         break;
                     default:
                         cellValue = item[headerKey] ?? ''; // 元のヘッダー名で値を取得
                 }

                 cell.textContent = cellValue;
                 cell.style.textAlign = textAlign;
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
         // 薬品種別とメーカーのキー名を動的に取得
         const categoryKey = originalHeaders.find(h => h.includes('薬品種別')) || '薬品種別';
         const makerKey = originalHeaders.find(h => h.includes('メーカー')) || 'メーカー';

        // 既存のオプションをクリア（「すべて」を除く）
        categoryFilter.innerHTML = '<option value="">薬品種別 (すべて)</option>';
        makerFilter.innerHTML = '<option value="">メーカー (すべて)</option>';


        const categories = [...new Set(inventoryData.map(item => item[categoryKey]).filter(Boolean))].sort();
        const makers = [...new Set(inventoryData.map(item => item[makerKey]).filter(Boolean))].sort();

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
         if (!th || !th.dataset.sort) return; // data-sort属性がない場合は無視

         const sortKey = th.dataset.sort; // 元のヘッダー名
         const processedKey = th.dataset.processedKey; // 処理済みキー

         if (currentSortColumn === sortKey) {
             isSortAscending = !isSortAscending;
         } else {
             currentSortColumn = sortKey;
             currentSortProcessedKey = processedKey; // ソートに使うキーを更新
             isSortAscending = true;
         }

         // ヘッダーのソート状態表示を更新
         updateSortIndicators(th);

         applyFiltersAndSort();
     }

     // ソートインジケータ更新
     function updateSortIndicators(activeTh) {
        inventoryTableHeader.querySelectorAll('th').forEach(th => {
            th.classList.remove('sort-asc', 'sort-desc');
             th.removeAttribute('aria-sort');
        });
        if (activeTh && currentSortColumn) { // currentSortColumnが設定されている場合のみ
             const sortClass = isSortAscending ? 'sort-asc' : 'sort-desc';
             activeTh.classList.add(sortClass);
              activeTh.setAttribute('aria-sort', isSortAscending ? 'ascending' : 'descending');
        }
    }


    // --- ダッシュボード表示 ---
    function displayDashboard() {
         // 薬品種別とメーカーのキー名を動的に取得
         const categoryKey = originalHeaders.find(h => h.includes('薬品種別')) || '薬品種別';
         const makerKey = originalHeaders.find(h => h.includes('メーカー')) || 'メーカー';
         const nameKey = originalHeaders.find(h => h.includes('薬品名称')) || '薬品名称';

        // KPI計算
        const totalItems = inventoryData.length;
        const totalValue = inventoryData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);

        // 注意在庫 (危険度ランク 7以上)
         const cautionData = inventoryData.filter(item => item.危険度ランク >= 7);
         const cautionItems = cautionData.length;
         const cautionValue = cautionData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);

        // 滞留在庫 (180日以上)
        const stagnantData = inventoryData.filter(item => item.滞留日数Numeric >= 180 && item.滞留日数Numeric < 9000);
        const stagnantItems = stagnantData.length;
        const stagnantValue = stagnantData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);

        // 期限切迫在庫 (90日以内)
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
        updateRankingList(stagnantWorstListEl, inventoryData, '滞留日数Numeric', false, 5, item => `${item[nameKey]} (${item.滞留日数表示}日)`);
        updateRankingList(nearExpiryWorstListEl, inventoryData.filter(item => !isNaN(item.残り日数Numeric) && item.残り日数Numeric >= 0), '残り日数Numeric', true, 5, item => `${item[nameKey]} (残${item.残り日数Numeric}日)`);
        updateRankingList(highValueListEl, inventoryData, '在庫金額Numeric', false, 5, item => `${item[nameKey]} (¥${formatCurrency(item.在庫金額Numeric)})`);

        // グラフ更新
        displayCategoryChart(categoryKey); // キー名を渡す
        displayMakerChart(makerKey); // キー名を渡す
    }

    // ランキングリスト更新ヘルパー
    function updateRankingList(listElement, data, sortKey, ascending, count, displayFormatter) {
        const sortedData = [...data].sort((a, b) => {
            const valA = a[sortKey];
            const valB = b[sortKey];
             // NaNや無効値を適切に扱う
             const numA = parseFloat(valA);
             const numB = parseFloat(valB);
             const effA = isNaN(numA) ? (ascending ? Infinity : -Infinity) : numA;
             const effB = isNaN(numB) ? (ascending ? Infinity : -Infinity) : numB;

            if (effA < effB) return ascending ? -1 : 1;
            if (effA > effB) return ascending ? 1 : -1;
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
    // Chart.js v4以降を想定
    // CSPエラーがChart.jsに起因する場合、設定やバージョンを見直す必要があるかもしれません。
    // グラフ共通オプション
    const chartOptions = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: {
                position: 'bottom',
                labels: { padding: 15 }
            },
            tooltip: {
                callbacks: {
                    label: function(context) {
                        let label = context.dataset.label || '';
                        if (label) { label += ': '; }
                        let value = 0;
                        let unit = context.dataset.unit || ''; // ランキンググラフ用

                        if (context.chart.config.type === 'doughnut' || context.chart.config.type === 'pie') {
                            value = context.parsed;
                        } else if (context.dataset.parsing === false) { // If data is pre-parsed object
                             value = context.raw;
                        } else if (context.parsed.x !== null && context.parsed.y === null) { // Horizontal bar
                            value = context.parsed.x;
                        } else if (context.parsed.y !== null) { // Vertical bar or other
                            value = context.parsed.y;
                        }

                        // Apply formatting based on label
                        if (label.includes('金額')) {
                             label += `¥${value.toLocaleString()}`;
                        } else if (label.includes('日数')) {
                             label += `${value.toLocaleString()} 日`;
                        } else if (label.includes('数量')) {
                             label += `${value.toLocaleString()} ${unit}`;
                        } else {
                            label += value.toLocaleString(); // Default formatting
                        }
                        return label;
                    }
                }
            }
        },
        scales: { // Add default scale options if needed, e.g., for bar charts
            x: { grid: { display: false } },
            y: { grid: { display: true, color: '#eee' } }
        }
    };

    // 薬品種別グラフ
    function displayCategoryChart(categoryKey) {
        const categoryData = inventoryData.reduce((acc, item) => {
            const category = item[categoryKey] || '不明';
            acc[category] = (acc[category] || 0) + item.在庫金額Numeric;
            return acc;
        }, {});

        const labels = Object.keys(categoryData).sort((a,b) => categoryData[b] - categoryData[a]); // 金額が多い順
        const dataValues = labels.map(label => categoryData[label]);

        if (categoryChartInstance) {
            categoryChartInstance.destroy();
        }
        try {
            categoryChartInstance = new Chart(categoryChartCanvas, {
                type: 'doughnut',
                data: {
                    labels: labels,
                    datasets: [{
                        label: '在庫金額',
                        data: dataValues,
                        backgroundColor: [ // 色を適当に設定 (必要に応じて増やす)
                            '#4CAF50', '#2196F3', '#FFC107', '#E91E63', '#9C27B0',
                            '#00BCD4', '#FF5722', '#8BC34A', '#673AB7', '#FF9800',
                            '#607D8B', '#795548', '#CDDC39', '#F44336', '#03A9F4'
                        ],
                        hoverOffset: 8, // 少し大きく
                        borderWidth: 1 // 境界線
                    }]
                },
                options: chartOptions
            });
        } catch(e) { console.error("Category Chart Error:", e); }
    }

    // メーカー別グラフ
    function displayMakerChart(makerKey) {
        const makerData = inventoryData.reduce((acc, item) => {
            const maker = item[makerKey] || '不明';
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
         try {
            makerChartInstance = new Chart(makerChartCanvas, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [{
                        label: '在庫金額',
                        data: dataValues,
                        backgroundColor: '#66BB6A', // 色変更
                        borderRadius: 4 // 角丸
                    }]
                },
                options: { ...chartOptions, indexAxis: 'y' } // 横棒グラフ
            });
        } catch(e) { console.error("Maker Chart Error:", e); }
    }

    // ランキンググラフ
    function displayRankingChart() {
        const type = rankingTypeSelect.value;
        const topN = parseInt(rankingTopNSelect.value, 10);
        const nameKey = originalHeaders.find(h => h.includes('薬品名称')) || '薬品名称'; // 薬品名称キー取得

        let sortKey = '';
        let ascending = false;
        let labelKey = nameKey; // デフォルトは薬品名称
        let valueKey = '';
        let chartLabel = '';
         let chartUnit = ''; // 単位表示用
         let indexAxis = 'y'; // デフォルトは横棒
         let chartType = 'bar';
         let dataToUse = [...inventoryData]; // 元データをコピーして使用

         switch (type) {
            case 'amountDesc':
                sortKey = '在庫金額Numeric'; ascending = false; valueKey = '在庫金額Numeric'; chartLabel = '金額';
                break;
             case 'stagnationDesc':
                sortKey = '滞留日数Numeric'; ascending = false; valueKey = '滞留日数Numeric'; chartLabel = '日数';
                 // 滞留日数9999を除外
                 dataToUse = inventoryData.filter(item => item.滞留日数Numeric < 9000);
                 break;
             case 'expiryAsc':
                 sortKey = '残り日数Numeric'; ascending = true; valueKey = '残り日数Numeric'; chartLabel = '日数';
                 // 期限切れでないものだけを対象
                 dataToUse = inventoryData.filter(item => !isNaN(item.残り日数Numeric) && item.残り日数Numeric >= 0);
                break;
             case 'quantityDesc':
                sortKey = '在庫数量Numeric'; ascending = false; valueKey = '在庫数量Numeric'; chartLabel = '数量';
                break;
             case 'quantityAsc':
                sortKey = '在庫数量Numeric'; ascending = true; valueKey = '在庫数量Numeric'; chartLabel = '数量';
                 // 在庫0を除外
                dataToUse = inventoryData.filter(item => item.在庫数量Numeric > 0);
                break;
            default:
                return;
        }


         const sortedData = dataToUse.sort((a, b) => {
             let valA = a[sortKey];
             let valB = b[sortKey];
             // NaNや無効値を適切に扱う
              valA = isNaN(valA) ? (ascending ? Infinity : -Infinity) : valA;
              valB = isNaN(valB) ? (ascending ? Infinity : -Infinity) : valB;

             if (valA < valB) return ascending ? -1 : 1;
             if (valA > valB) return ascending ? 1 : -1;
             // 同値の場合は薬品名でソート（安定ソート）
             return String(a[labelKey] ?? '').localeCompare(String(b[labelKey] ?? ''));
         }).slice(0, topN);


        const labels = sortedData.map(item => item[labelKey]);
        const dataValues = sortedData.map(item => item[valueKey]);
         // 単位情報を付与（横棒グラフのツールチップ用） - 数量の場合のみ
         const units = chartLabel === '数量' ? sortedData.map(item => item[originalHeaders.find(h => h.includes('単位')) || '単位'] || '') : null;

        if (rankingChartInstance) {
            rankingChartInstance.destroy();
        }
         try {
            rankingChartInstance = new Chart(rankingChartCanvas, {
                 type: chartType,
                data: {
                    labels: labels.reverse(), // 横棒グラフ用にラベルを逆順に
                    datasets: [{
                        label: chartLabel,
                        data: dataValues.reverse(), // データも逆順に
                        backgroundColor: '#5C6BC0', // 色変更 (Indigo)
                        borderRadius: 4,
                        unit: units ? units.reverse() : '' // 単位情報 (逆順)
                    }]
                },
                 options: { ...chartOptions, indexAxis: indexAxis }
            });
        } catch(e) { console.error("Ranking Chart Error:", e); }
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
                    model: "gpt-3.5-turbo", // or "gpt-4"
                    messages: [{ role: "user", content: prompt }],
                    temperature: 0.5,
                    max_tokens: 1000
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
                aiAdviceOutput.textContent = advice;
                copyAiAdviceButton.style.display = 'inline-block';
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
         // キー名を動的に取得
         const nameKey = originalHeaders.find(h => h.includes('薬品名称')) || '薬品名称';
         const unitKey = originalHeaders.find(h => h.includes('単位')) || '単位';

         // サマリー情報を生成するロジック (displayDashboardと類似)
        const totalItems = inventoryData.length;
        const totalValue = inventoryData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);

        const cautionData = inventoryData.filter(item => item.危険度ランク >= 7);
        const cautionItems = cautionData.length;
        const cautionValue = cautionData.reduce((sum, item) => sum + item.在庫金額Numeric, 0);
        const topCaution = cautionData.sort((a, b) => b.危険度ランク - a.危険度ランク).slice(0, 5).map(i => `- ${i[nameKey]} (ランク${i.危険度ランク}, 金額¥${formatCurrency(i.在庫金額Numeric)})`).join('\n');

        const stagnantData = inventoryData.filter(item => item.滞留日数Numeric >= 180 && item.滞留日数Numeric < 9000);
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

         return `
## 在庫サマリー (${formatDate(new Date())}時点)

**基本情報:**
- 総在庫品目数: ${totalItems.toLocaleString()} 品目
- 総在庫金額 (税別): ¥${formatCurrency(totalValue)}

**要注意在庫 (危険度ランク7以上):**
- 該当品目数: ${cautionItems.toLocaleString()} 品目
- 合計金額: ¥${formatCurrency(cautionValue)}
- 上位リスト (抜粋):\n${topCaution || '- 該当なし'}

**滞留在庫 (180日以上出庫なし):**
- 該当品目数: ${stagnantItems.toLocaleString()} 品目
- 合計金額: ¥${formatCurrency(stagnantValue)}
- 上位リスト (抜粋):\n${topStagnant || '- 該当なし'}

**期限切迫在庫 (残り90日以内):**
- 該当品目数: ${nearExpiryItems.toLocaleString()} 品目
- 合計金額: ¥${formatCurrency(nearExpiryValue)}
- 上位リスト (抜粋):\n${topNearExpiry || '- 該当なし'}

**高額在庫 (上位抜粋):**
- 上位リスト (抜粋):\n${topHighValue || '- 該当なし'}

**欠品注意在庫 (在庫僅少かつ最近出庫あり - 抜粋):**
- 上位リスト (抜粋):\n${topLowStock || '- 該当なし'}

**危険度ランク設定:**
- 現在の重み付け: 期限重視 ${expiryWeightParam}%, 滞留重視 ${100 - expiryWeightParam}%
`;
    }

    function createPrompt(summary) {
        // プロンプト内容は前回と同じ
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

    // --- UI制御ヘルパー ---
    function showAiLoading() { aiLoading.style.display = 'block'; getAiAdviceButton.disabled = true; }
    function hideAiLoading() { aiLoading.style.display = 'none'; getAiAdviceButton.disabled = (userApiKey === null); } // キーがあれば有効化
    function showAiError(message) { aiErrorOutput.textContent = message; aiErrorOutput.style.display = 'block'; }
    function hideAiError() { aiErrorOutput.style.display = 'none'; }
    function showDataError(message) { dataErrorOutput.textContent = message; dataErrorOutput.style.display = 'block'; }
    function hideDataError() { dataErrorOutput.style.display = 'none'; }
    function showLoading() { if(loadingIndicator) loadingIndicator.style.display = 'block'; }
    function hideLoading() { if(loadingIndicator) loadingIndicator.style.display = 'none'; }


    // --- イベントリスナー設定 ---

    // データ読み込みボタン
    loadDataButton.addEventListener('click', () => {
        const tsvData = inventoryDataInput.value;
        if (!tsvData.trim()) {
            showDataError('データが入力されていません。');
            return;
        }
        hideDataError();
        try {
            inventoryData = parseTSV(tsvData);
            initializeAppLogic(); // パース成功したらアプリ初期化
        } catch (error) {
            console.error('TSVパースエラー:', error);
            showDataError(`データの読み込みに失敗しました: ${error.message}`);
        }
    });


    // タブ切り替え
    tabButtons.forEach(button => {
        button.addEventListener('click', () => {
            // データがロードされていない場合は何もしない
            if (inventoryData.length === 0) return;

            const targetTab = button.dataset.tab;

            tabContents.forEach(content => content.classList.remove('active'));
            document.getElementById(targetTab).classList.add('active');

            tabButtons.forEach(btn => btn.classList.remove('active'));
            button.classList.add('active');

             // 分析レポートタブが選択されたら、ランキンググラフを初期表示
             if (targetTab === 'analysis-reports' && inventoryData.length > 0) {
                 displayRankingChart();
            }
        });
    });

    // 在庫一覧: 検索、フィルタ、リセット
    searchInput.addEventListener('input', () => {if(inventoryData.length > 0) applyFiltersAndSort()});
    categoryFilter.addEventListener('change', () => {if(inventoryData.length > 0) applyFiltersAndSort()});
    makerFilter.addEventListener('change', () => {if(inventoryData.length > 0) applyFiltersAndSort()});
    dangerRankFilter.addEventListener('change', () => {if(inventoryData.length > 0) applyFiltersAndSort()});
    resetFiltersButton.addEventListener('click', () => {
        if (inventoryData.length === 0) return;
        searchInput.value = '';
        categoryFilter.value = '';
        makerFilter.value = '';
        dangerRankFilter.value = '';
        currentSortColumn = null;
        currentSortProcessedKey = null;
        isSortAscending = true;
        updateSortIndicators(null);
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
         const totalPages = itemsPerPage === 'all' ? 1 : Math.max(1, Math.ceil(totalItems / itemsPerPage));
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
         if (inventoryData.length === 0) return;
         expiryWeightParam = parseInt(expiryWeightSlider.value, 10);
         recalculateAllDangerRanks();
     });

     // 分析レポート: ランキング種類/件数変更
     rankingTypeSelect.addEventListener('change', () => {if(inventoryData.length > 0) displayRankingChart()});
     rankingTopNSelect.addEventListener('change', () => {if(inventoryData.length > 0) displayRankingChart()});


    // AIアドバイス
    saveApiKeyButton.addEventListener('click', saveApiKey);
    getAiAdviceButton.addEventListener('click', fetchAiAdvice);
     copyAiAdviceButton.addEventListener('click', copyAdviceToClipboard);

    // --- 初期状態設定 ---
    // 最初はデータ入力エリアのみ表示
    tabNav.style.display = 'none';
    mainContent.style.display = 'none';
    footer.style.display = 'none';

}); // End of DOMContentLoaded
