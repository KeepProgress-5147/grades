# 🎓 大學部 111-115 學年度錄取分數視覺化分析儀表板

本專案收錄並整理台灣大專院校 **111 至 115 學年度**（共 5 個學年度）之分發錄取平均標準與成績資料，針對 **醫學、護理、聽語、視光、醫檢** 五大醫療衛生學群，打造 90%+ 寬幅高對比、互動式折線圖視覺化儀表板與組內歷年排名分析。

🌐 **線上即時展示 (GitHub Pages)**：\https://<your-username>.github.io/<your-repo-name>/\

---

## 🌟 核心特色

1. **五大群組獨立對比**：
   - 醫學組（18 系組）、護理組（28 系組）、聽語組（5 系組）、視光組（4 系組）、醫檢組（11 系組）各自獨立成圖，杜絕跨領域交叉混亂。
2. **右側自選校系面板**：
   - 針對護理與醫學等多系組群組，右側提供專屬勾選清單，支援：
     - 👑 **前 5 名校系** 一鍵聚焦
     - 🏛️ **僅公立大學** 快速篩選
     - 🚫 **全部取消 / 全部勾選** 自由配對
     - 🔍 群組內即時關鍵字搜尋
3. **組內當年度排名標註**：
   - 111～115 各年度精確計算組內排名，前三名附有金、銀、銅色高光徽章（\#1\, \#2\, \#3\）。
   - 下方表格支援按任一年度分數/名次進行升降冪排序。
4. **90%+ 寬幅高對比視覺設計**：
   - 採用深邃暗夜黑底色與純白天藍高對比排版，字體加粗加大，圖表高度達 650px。
   - 支援滑鼠懸停顯示排名 Tooltip、縮放區間（DataZoom）、平滑曲線切換與一鍵匯出高畫質 PNG。
5. **Excel 成果同步內嵌**：
   - 同步提供 [\大學部111-115成績_含折線圖.xlsx\](大學部111-115成績_含折線圖.xlsx)，各群組獨立分頁並內建 Excel 原生折線圖。

---

## 📁 專案檔案架構

\\\	ext
├── index.html                   # 線上 GitHub Pages 入口網頁（互動式儀表板）
├── 大學部111-115成績_互動圖表.html # 本地互動式網頁備份
├── 大學部111-115成績_含折線圖.xlsx # 內建原生折線圖與排名的 Excel 成果檔
├── 大學部111-115成績.xlsx        # 原始清洗後之 5 大群組成績資料庫
├── 111-115歷年各系組平均分數彙整.xlsx # 各系組歷年原始大表
├── 111_result_school_data.xlsx  # 111 學年度各校系原始分發資料
├── 112_result_school_data.xlsx  # 112 學年度各校系原始分發資料
├── 113_result_school_data.xlsx  # 113 學年度各校系原始分發資料
├── 114_result_school_data.xlsx  # 114 學年度各校系原始分發資料
├── 115_result_school_data.xlsx  # 115 學年度各校系原始分發資料
└── README.md                    # 專案說明文件
\\\

---

## 🚀 如何發布至 GitHub Pages

1. **建立 GitHub Repository**：
   - 登入 [GitHub](https://github.com/) 並點選 **New repository**。
   - Repository 名稱建議設定為：\university-admission-scores-dashboard\ 或自訂名稱。
   - 設為 **Public**，點擊 **Create repository**。

2. **推動本地程式碼至 GitHub**：
   在專案資料夾開啟 Terminal / PowerShell，執行以下指令：
   \\\ash
   git remote add origin https://github.com/<您的GitHub帳號>/<您的Repo名稱>.git
   git branch -M main
   git push -u origin main
   \\\

3. **啟用 GitHub Pages 免費網頁託管**：
   - 進入該 GitHub Repository 頁面，點擊上方 **Settings**。
   - 左側選單點擊 **Pages**。
   - 在 **Build and deployment** 下方的 **Source** 選擇 **Deploy from a branch**。
   - Branch 選擇 **\main\**，資料夾選擇 **\/ (root)\**，點擊 **Save**。
   - 等待 1~2 分鐘後，即可透過專屬網址線上瀏覽儀表板！

---

## 🛠️ 技術棧
- **Frontend**：HTML5, JavaScript (ES6+), [Tailwind CSS](https://tailwindcss.com/), [Apache ECharts 5](https://echarts.apache.org/)
- **Data Processing**：Python, Pandas, OpenPyXL
