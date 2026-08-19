# 🎓 大學部 111-115 學年度錄取分數視覺化分析儀表板

本專案收錄並整理台灣大專院校 **111 至 115 學年度**（共 5 個學年度）之分發錄取平均標準與成績資料，針對 **醫學、護理、聽語、視光、醫檢** 五大醫療衛生學群（共 68 個系組），打造 90%+ 寬幅高對比、互動式折線圖視覺化儀表板與組內歷年排名分析。

🌐 **線上即時展示 (GitHub Pages)**：`https://<your-username>.github.io/<your-repo-name>/`

---

## 🌟 核心特色

1. **五大醫療衛生群組獨立對比**：
   - 🩺 **醫學組**（18 系組）：包含臺大、陽明交大、成大、北醫、長庚等公私立醫學及牙醫系組。
   - 💉 **護理組**（28 系組）：涵蓋全台公私立大學護理學系與各專項照護組。
   - 🗣️ **聽語組**（5 系組）：語言治療與聽力治療領域系組。
   - 👁️ **視光組**（4 系組）：中山醫大、亞洲大學、大葉大學、馬偕醫大視光學系。
   - 🔬 **醫檢組**（13 系組）：醫學檢驗生物技術、醫學生物技術暨檢驗學系。
2. **右側自選校系面板**：
   - 👑 **前 5 名校系** 一鍵聚焦
   - 🏛️ **僅公立大學** 快速篩選
   - 🚫 **全部取消 / 全部勾選** 自由配對
   - 🔍 群組內即時關鍵字搜尋
3. **組內當年度排名標註**：
   - 111～115 各年度精確計算組內排名，前三名附有金、銀、銅色高光徽章（`#1`、`#2`、`#3`）。
   - 下方表格支援按任一年度分數/名次進行升降冪排序。
4. **90%+ 寬幅高對比視覺設計**：
   - 採用深邃暗夜黑底色與純白天藍高對比排版，字體加粗加大，圖表高度達 650px。
   - 支援滑鼠懸停顯示排名 Tooltip、縮放區間（DataZoom）、平滑曲線切換與一鍵匯出高畫質 PNG。
5. **Excel 成果同步內嵌**：
   - 同步提供 [`大學部111-115成績_含折線圖.xlsx`](大學部111-115成績_含折線圖.xlsx)，各群組獨立分頁並內建 Excel 原生折線圖與組內排名。

---

## 📁 專案檔案架構

```text
├── index.html                   # 線上 GitHub Pages 入口網頁（互動式儀表板）
├── 大學部111-115成績_互動圖表.html # 本地互動式網頁備份
├── 大學部111-115成績_含折線圖.xlsx # 內建原生折線圖與排名的 Excel 成果檔
├── 大學部111-115成績.xlsx        # 已標註 5 大群組之成績資料庫（68 筆核心系組）
├── 111-115歷年各系組平均分數彙整.xlsx # 全台 2,576 筆系組歷年原始大表
├── 111_result_school_data.xlsx  # 111 學年度各校系原始分發資料
├── 112_result_school_data.xlsx  # 112 學年度各校系原始分發資料
├── 113_result_school_data.xlsx  # 113 學年度各校系原始分發資料
├── 114_result_school_data.xlsx  # 114 學年度各校系原始分發資料
├── 115_result_school_data.xlsx  # 115 學年度各校系原始分發資料
├── process_all_years.py         # 歷年 PDF 自動批次轉換程式
├── update_all_data.py           # 儀表板與 Excel 圖表自動更新腳本
├── PROJECT_SUMMARY.md           # 專案討論與技術執行完整紀錄
├── PROJECT_SUMMARY.html         # 專案紀錄網頁版
└── README.md                    # 本專案說明文件
```

---

## 🚀 如何上傳至 GitHub 並發布 GitHub Pages

### 方式 A：透過 GitHub 網頁版直接上傳（最簡單免裝指令）
1. 登入 [GitHub.com](https://github.com/)，點擊右上角 **「+」 ➔ 「New repository」**。
2. 設定 Repository 名稱（例如：`university-scores-dashboard`），選擇 **Public**，點擊 **Create repository**。
3. 在新頁面點擊 **「uploading an existing file」**。
4. 將專案資料夾內的所有檔案拖曳上傳至 GitHub，並點擊下方 **「Commit changes」**。
5. 點擊倉庫上方的 **Settings ➔ Pages**：
   - **Source** 選擇 `Deploy from a branch`
   - **Branch** 選擇 `main`，資料夾選擇 `/ (root)`，點擊 **Save**。
6. 等候 1～2 分鐘，即可獲得免費的線上互動網址！

---

### 方式 B：透過 Git 指令上傳
```bash
# 1. 初始化本地 Git 倉庫
git init

# 2. 加入所有檔案並提交
git add .
git commit -m "feat: initial commit of university admission scores dashboard 111-115"

# 3. 關聯遠端 GitHub 倉庫並推播 (請替換 <您的GitHub帳號> 與 <Repo名稱>)
git branch -M main
git remote add origin https://github.com/<您的GitHub帳號>/<Repo名稱>.git
git push -u origin main
```

---

## 🛠️ 技術棧
- **Frontend**：HTML5, JavaScript (ES6+), [Tailwind CSS](https://tailwindcss.com/), [Apache ECharts 5](https://echarts.apache.org/)
- **Data Processing**：Python, OpenPyXL, PDFPlumber, PyPDF
