const { createApp, ref, onMounted } = Vue;

createApp({
    setup() {
        // 
        const wallets = ref([]);
        const records = ref([]);
        const systemTotalBalance = ref(0);
        
        const filterType = ref("all");   
        const customStart = ref("");     
        const customEnd = ref("");       
        const filterWallet = ref(null);  
        
        const API_BASE = "http://127.0.0.1:8000/api";

        /**
         * 錢包樹狀重構演算法
         * 說明：將後端返回之無序錢包陣列，重新鏈結成「母錢包緊跟其附屬子錢包」之巢狀線性陣列結構。
         * @param {Array} rawWallets 原始錢包資料庫紀錄
         * @returns {Array} 已排序之樹狀結構平面化陣列
         */
        const buildWalletTree = (rawWallets) => {
            const parents = rawWallets.filter(w => w.parent_id === null);
            const children = rawWallets.filter(w => w.parent_id !== null);
            const sorted = [];

            parents.forEach(p => {
                sorted.push(p); // 寫入母錢包
                // 篩選出歸屬於該母錢包之子錢包
                const subWallets = children.filter(c => c.parent_id === p.id);
                sorted.push(...subWallets);
            });

            // 邊際效益防禦：回收並附加無法與現存母帳戶匹配之獨立帳戶，避免數據漏載
            const orphanChildren = children.filter(c => !parents.some(p => p.id === c.parent_id));
            sorted.push(...orphanChildren);

            return sorted;
        };

        /**
         * 全域同步異步數據擷取機制
         * 包含：1. 錢包結構獲取與重新編排 2. 多條件參數化交易紀錄抓取與資料正規化
         */
        const fetchData = async () => {
            try {
                // 1. 更新錢包資料與樹狀對應
                const resWallets = await fetch(`${API_BASE}/wallets`);
                const walletData = await resWallets.json();

                systemTotalBalance.value = walletData.total_system_balance || 0;
                wallets.value = buildWalletTree(walletData.wallets || []);

                // 2. 根據篩選控制項封裝 API Query String
                let url = `${API_BASE}/records?`;
                if (filterType.value === 'custom') {
                    if (customStart.value) url += `start_date=${customStart.value}&`;
                    if (customEnd.value) url += `end_date=${customEnd.value}&`;
                } else if (filterType.value !== 'all') {
                    url += `days=${filterType.value}&`;
                }
                if (filterWallet.value !== null) {
                    url += `wallet_id=${filterWallet.value}&`;
                }

                const resRecords = await fetch(url);
                const rawRecords = await resRecords.json();

                // 數據防禦層：對齊並清洗不同後端資料庫（如 ISOString、Timestamp）拋出之多型態日期欄位
                records.value = rawRecords.map(r => {
                    let formattedDate = "";
                    
                    if (r.date) {
                        const dateStr = String(r.date).trim();
                        // 標準 ISO 8601 或帶時區字串格式處理 (例: 2026-07-08T00:00:00)
                        if (dateStr.includes("T")) {
                            formattedDate = dateStr.split("T")[0];
                        } else if (dateStr.includes(" ")) {
                            formattedDate = dateStr.split(" ")[0];
                        } else {
                            formattedDate = dateStr; // 已符合標準 YYYY-MM-DD 規格
                        }
                    }
                    
                    // 異常值防禦：若數據損毀或為 null，預設補回當日前端系統時間，避免 HTML5 Input 渲染崩潰
                    if (!formattedDate || formattedDate === "null") {
                        formattedDate = new Date().toISOString().split('T')[0];
                    }

                    return {
                        ...r,
                        date: formattedDate // 寫回安全格式之字串型態
                    };
                });

            } catch (err) {
                console.error("遠端 API 同步失敗:", err);
            }
        };

        /**
         * 時間區間選擇器狀態變更事件回呼
         */
        const onFilterTypeChange = () => {
            if (filterType.value === 'custom') {
                const today = new Date();
                const pastMonth = new Date();
                pastMonth.setDate(today.getDate() - 30);
                customEnd.value = today.toISOString().split('T')[0];
                customStart.value = pastMonth.toISOString().split('T')[0];
            }
            fetchData();
        };

        /**
         * 範圍時間合法性驗證與數據重載
         */
        const validateAndFetch = () => {
            if (!customStart.value || !customEnd.value) return;
            const start = new Date(customStart.value);
            const end = new Date(customEnd.value);
            if (start > end) {
                alert("開始日期不能大於結束日期！");
                customEnd.value = customStart.value;
                return;
            }
            fetchData();
        };

        /**
         * 於本地資料響應陣列尾端初始化一筆新空白流水帳紀錄
         */
        const addNewRow = () => {
            records.value.push({
                id: null,
                date: new Date().toISOString().split('T')[0], 
                type: "expense",
                amount: 0,
                category: "餐飲",
                note: "",
                from_wallet_id: null,
                to_wallet_id: null,
                recon_status: "unreconciled"
            });
        };

        /**
         * 移除本地尚未存檔/尚未核對之特定交易列
         * @param {number} index 資料列索引
         */
        const removeRow = (index) => {
            records.value.splice(index, 1);
        };

        /**
         * 批次儲存
         */
        const submitBulkSave = async () => {
            try {
                const response = await fetch(`${API_BASE}/records/bulk-save`, {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify(records.value)
                });
                const result = await response.json();
                if (response.ok) {
                    alert(result.message);
                    await fetchData(); // 事務完成後，強制觸發伺服器快照同步
                } else {
                    alert(`【儲存失敗】\n原因：${result.detail}`);
                }
            } catch (err) {
                alert("網路連線異常，無法完成批次儲存！");
            }
        };

        // 生命週期掛載點：初始化數據載入
        onMounted(fetchData);

        return {
            wallets,
            records,
            systemTotalBalance,
            filterType,
            customStart,
            customEnd,
            filterWallet,
            fetchData,
            onFilterTypeChange,
            validateAndFetch,
            addNewRow,
            removeRow,
            submitBulkSave
        };
    }
}).mount('#app');