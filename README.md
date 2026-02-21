# 📊 QC 規格品質分析系統

🌐 **線上 App（直接點擊開啟）**：
👉 [https://qc-quality-analysis-jqgngzussrsjgmblu2dedr.streamlit.app/](https://qc-quality-analysis-jqgngzussrsjgmblu2dedr.streamlit.app/)

一個以 **Streamlit** 建立的互動式品質管理儀表板，用來分析產品規格與工廠驗收標準的符合情況。

---

## 📁 專案結構

```
python_02/
├── app02.py                        # 主程式
├── 0217_excel_python test.xlsx     # 原始資料（產品 × 工廠 × 規格）
├── 0218_color mapping.xlsx         # 顏色判定規則對照表
├── requirements.txt                # 套件版本清單
└── README.md                       # 本說明文件
```

---

## 🗂️ 資料說明

### `0217_excel_python test.xlsx`（工作表：`0217`）
| 欄位 | 說明 |
|---|---|
| Product | 產品編號（如 Product-0001） |
| Factory | 工廠編號（Factory-1 ~ Factory-8） |
| Customer | 客戶編號（Customer-1 ~ Customer-15） |
| Location | 地點 |
| Method | 驗收方式（Method-X / Method-Y / Method-Z） |
| A ~ L | 12種規格種類的驗收結果 |

### `0218_color mapping.xlsx`
顏色判定規則：根據 `(驗收方式, 規格種類, 規格代碼)` 決定每格屬於哪一類別。

---

## 🎨 顏色類別說明

| 顏色 | 類別 | 說明 |
|---|---|---|
| ⬜ 白色 | 符合中央標準 | 符合標準，正常 |
| 🟩 綠色 | 比中央更嚴格 | 比標準更嚴，品質較好 |
| 🩷 粉色 | 比中央寬鬆 | 比標準寬鬆，需注意 |
| ⬜ 灰色 | 不採納(NA) | 不適用此規格種類 |

---

## 🖥️ 四個分析頁籤

### Tab 1 — 📋 產品×規格 顏色表
- 每筆資料（Product × Factory × Method）橫向展開 A-L 共 12 欄
- 每格依類別顯示對應顏色
- 側邊欄可依「工廠」和「驗收方式」篩選
- 可下載 Excel 報表

### Tab 2 — 📈 XYZ vs A-L 分布
- 堆疊長條圖：X 軸為規格（A-L），每組按驗收方式（X/Y/Z）分欄
- 四種顏色各自累積，一眼看出哪個規格問題最多
- 附數字明細表（Method × Spec × 類別 數量）

### Tab 3 — 🔍 條件篩選名單
- 三個下拉選單：**規格（A-L）** ＋ **驗收方式** ＋ **顏色類別**
- 例：選「A ＋ Method-X ＋ 比中央寬鬆(粉色)」→ 顯示哪些工廠、哪些客戶符合此條件
- 同時顯示詳細明細表

### Tab 4 — ↔️ 跨工廠比較
- 選一個產品編號 → 列出所有生產此產品的工廠
- 橫向比較每家工廠在 A-L 各規格的顏色
- 進一步選定某規格欄 → 按顏色篩選出特定工廠與客戶

---

## 🚀 如何啟動

### 1. 安裝套件
```bash
pip install -r requirements.txt
```

### 2. 確認檔案位置
將以下兩個 Excel 檔案放在與 `app02.py` 同一個資料夾：
- `0217_excel_python test.xlsx`
- `0218_color mapping.xlsx`

### 3. 啟動 App
```bash
streamlit run app02.py
```

瀏覽器會自動開啟，或手動前往：
```
http://localhost:8501
```

---

## ⚙️ 技術細節

| 項目 | 說明 |
|---|---|
| 資料讀取 | `pandas.read_excel()` 讀取 `0217` 工作表 |
| 分類邏輯 | 先比對已分類的中文標籤，再查 color mapping 對照表 |
| 快取 | `@st.cache_data` 避免每次互動重新讀取 Excel |
| 顏色呈現 | Pandas Styler `apply()` 動態套用背景色 |
| 圖表 | Plotly Express 堆疊長條圖（`px.bar`） |

---

## 🔧 常見問題

**Q：App 無法啟動，顯示「找不到檔案」？**
> 確認 Excel 檔案與 `app02.py` 放在同一資料夾，並在該資料夾內執行 `streamlit run app02.py`。

**Q：某些格子顯示「未知」黃色？**
> 表示該格的原始值不在 `0218_color mapping.xlsx` 的規則中，請補充對應的規則列。

**Q：想新增更多產品或工廠資料？**
> 直接在 `0217_excel_python test.xlsx` 的 `0217` 工作表新增列即可，App 重啟後自動讀取。
