import fs from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { XLSX_CONTENT } from "../xlsx-content.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, "..");
const outputDir = path.join(rootDir, "practice-data");
const manifestPath = path.join(rootDir, "lesson-downloads-data.js");
const execFileAsync = promisify(execFile);

const interLessonMeta = {
  "P1-02": {
    title: "基礎函式練習資料",
    intro: "這份資料適合直接練 SUM、AVERAGE、COUNT、MAX、MIN、LARGE、SMALL。先用同一張表把最常用統計函數練熟，後面所有報表課都會順很多。",
    fileName: "p1-02-basic-functions.csv",
    label: "基礎函式資料",
    description: "6 個月份的員工業績表，適合練加總、平均、計數、最大值、最小值與第 N 大。",
  },
  "P1-03": {
    title: "條件判斷練習資料",
    intro: "這份資料用來練 IF / IFS / SWITCH 的判斷邏輯。很適合拿來做績效等級、出勤提醒和薪資判斷。",
    fileName: "p1-03-if-logic.csv",
    label: "條件判斷資料",
    description: "員工、部門、月薪、績效與出勤資料，可直接練條件邏輯公式。",
  },
  "P2-01": {
    title: "條件統計練習資料",
    intro: "這課的重點是依部門、年資、績效條件去加總或計數。先用同一張資料表把 SUMIF / COUNTIF / AVERAGEIF 系列練順。",
    fileName: "p2-01-conditional-stats.csv",
    label: "條件統計資料",
    description: "可直接練 SUMIF / COUNTIF / AVERAGEIF / SUMIFS / COUNTIFS。",
  },
  "P2-02": {
    title: "查找比對練習資料",
    intro: "這份資料最適合搭配另一張對照表練 VLOOKUP、XLOOKUP、INDEX+MATCH。先把跨表查找的手感建立起來。",
    fileName: "p2-02-product-catalog.csv",
    label: "產品主檔資料",
    description: "產品代碼、名稱、類別、單價與庫存，可拿來做查表與補欄位。",
  },
  "P3-02": {
    title: "文字與日期清理練習資料",
    intro: "這份資料刻意保留常見的文字與日期欄位，適合練 LEFT、RIGHT、TEXT、DATE、EOMONTH 與清理格式。",
    fileName: "p3-02-text-date-cleanup.csv",
    label: "文字日期資料",
    description: "員工基本資料，適合練日期轉換、文字拆解與格式清理。",
  },
  "P4-01": {
    title: "動態陣列練習資料",
    intro: "這張產品表可以直接拿來練 FILTER、SORT、UNIQUE、SEQUENCE。先用它感受一個公式吐出整片結果的差別。",
    fileName: "p4-01-dynamic-array-catalog.csv",
    label: "動態陣列資料",
    description: "產品、類別、單價、銷量與評分，適合練篩選、排序與唯一值。",
  },
  "P4-02": {
    title: "進階函式練習資料",
    intro: "這份資料適合練 LET、LAMBDA、BYROW 等進階公式。可以直接拿來做年度業績彙總與條件判斷。",
    fileName: "p4-02-advanced-functions.csv",
    label: "進階函式資料",
    description: "員工年度業績資料，適合練 LET / LAMBDA / BYROW / REDUCE。",
  },
};

const curatedLessonAssets = {
  "P1-01": {
    title: "快捷鍵演練資料",
    intro: "這課不只是一張清單，所以我另外準備了鍵盤演練題。可以直接下載後照著一輪輪做，把常用操作練成反射動作。",
    files: [
      {
        fileName: "p1-01-keyboard-drills.csv",
        label: "快捷鍵演練清單",
        description: "日常最常用的 8 個 Mac Excel 操作情境，可照順序做一輪暖身。",
        header: ["練習順序", "場景", "目標操作", "Mac 快捷鍵", "完成後自評"],
        rows: [
          [1, "整理月報前先存檔", "儲存目前工作簿", "⌘S", ""],
          [2, "快速跳到資料底端", "往下跳到連續資料尾端", "⌘↓", ""],
          [3, "補上同一公式", "往下填滿", "⌃D", ""],
          [4, "直接編輯儲存格", "進入編輯模式", "F2", ""],
          [5, "一次改多格相同內容", "多格同時輸入", "⌘↩", ""],
          [6, "反悔剛剛的操作", "還原", "⌘Z", ""],
          [7, "再做一次剛才撤銷的步驟", "重做", "⌘Y", ""],
          [8, "切換工作表", "下一張 / 上一張工作表", "Fn+⌥↓ / Fn+⌥↑", ""],
        ],
      },
    ],
  },
  "P2-03": {
    title: "樞紐分析練習資料",
    intro: "這份原始交易資料是給樞紐分析用的。欄位已經整理成乾淨格式，直接匯入 Excel 後就能練列、欄、值與日期群組。",
    files: [
      {
        fileName: "p2-03-pivot-sales-transactions.csv",
        label: "樞紐原始交易資料",
        description: "跨部門、跨月份的交易流水帳，適合直接建立樞紐分析表。",
        header: ["日期", "部門", "產品類別", "業務", "地區", "數量", "金額"],
        rows: [
          ["2026-01-05", "電子", "耳機", "王小明", "北區", 18, 120000],
          ["2026-01-12", "家電", "清淨機", "李大明", "中區", 9, 86000],
          ["2026-01-18", "電子", "鍵盤", "張志偉", "南區", 24, 98000],
          ["2026-02-03", "家電", "除濕機", "李大明", "北區", 7, 91000],
          ["2026-02-11", "電子", "螢幕", "王小明", "中區", 11, 132000],
          ["2026-02-26", "電子", "滑鼠", "張志偉", "南區", 34, 67000],
          ["2026-03-04", "家電", "電風扇", "陳雅琪", "北區", 15, 74000],
          ["2026-03-10", "電子", "筆電支架", "張志偉", "中區", 26, 59000],
          ["2026-03-18", "家電", "烤箱", "李大明", "南區", 6, 88000],
          ["2026-03-27", "電子", "Webcam", "王小明", "北區", 19, 76000],
        ],
      },
    ],
  },
  "P2-04": {
    title: "條件式格式化練習資料",
    intro: "這份 KPI 月報就是拿來練條件式格式化的。你可以直接用它做資料橫條、負成長警示和高客訴提醒。",
    files: [
      {
        fileName: "p2-04-conditional-format-kpi.csv",
        label: "KPI 月報資料",
        description: "同時包含業績、達成率、月增率與客訴數，適合做多層規則。",
        header: ["姓名", "部門", "本月業績", "目標達成率", "月增率", "客訴數", "回款天數"],
        rows: [
          ["王小明", "業務", 128000, 1.08, 0.12, 2, 29],
          ["林美華", "財務", 0, 1.0, 0.0, 0, 0],
          ["張志偉", "業務", 91000, 0.76, -0.08, 5, 44],
          ["陳雅琪", "客服", 0, 1.0, 0.0, 8, 0],
          ["李大明", "業務", 146000, 1.15, 0.21, 1, 26],
          ["吳建志", "物流", 0, 1.0, 0.0, 3, 0],
        ],
      },
    ],
  },
  "P3-01": {
    title: "資料驗證練習資料",
    intro: "這課最適合用一張輸入表加一張來源清單來練。我把表單和下拉來源分開準備好了，直接下載就能做清單驗證與防呆規則。",
    files: [
      {
        fileName: "p3-01-onboarding-form.csv",
        label: "新進人員表單",
        description: "可直接拿來設定下拉選單、文字長度、日期與手機格式限制。",
        header: ["姓名", "部門", "到職日", "Email", "手機", "工作城市", "職等"],
        rows: [
          ["王小明", "", "", "", "", "", ""],
          ["林美華", "", "", "", "", "", ""],
          ["張志偉", "", "", "", "", "", ""],
          ["陳雅琪", "", "", "", "", "", ""],
        ],
      },
      {
        fileName: "p3-01-validation-lists.csv",
        label: "驗證來源清單",
        description: "做下拉選單用的來源資料，可當另一張工作表或命名範圍來源。",
        header: ["部門", "工作城市", "職等"],
        rows: [
          ["業務", "台北", "專員"],
          ["財務", "新竹", "資深專員"],
          ["資訊", "台中", "副理"],
          ["客服", "高雄", "經理"],
        ],
      },
    ],
  },
  "P3-03": {
    title: "圖表設計練習資料",
    intro: "這份月報資料很適合練圖表。你可以先插一張預設圖，再自己刪雜訊、改標題，做成更像能交付的版本。",
    files: [
      {
        fileName: "p3-03-chart-monthly-report.csv",
        label: "月報圖表資料",
        description: "月份、營收、目標與成長率都有，適合練柱狀圖、折線圖與組合圖。",
        header: ["月份", "營收", "目標", "成長率"],
        rows: [
          ["1月", 820000, 780000, 0.05],
          ["2月", 860000, 800000, 0.049],
          ["3月", 910000, 820000, 0.058],
          ["4月", 905000, 850000, -0.005],
          ["5月", 960000, 870000, 0.061],
          ["6月", 1015000, 900000, 0.057],
        ],
      },
    ],
  },
  "P3-04": {
    title: "表格與命名範圍練習資料",
    intro: "這課不是要你背理論，而是把原始清單真的轉成 Table。這份資料適合直接拿來做表格、命名範圍和結構化參照。",
    files: [
      {
        fileName: "p3-04-sales-table.csv",
        label: "銷售明細表",
        description: "適合先轉成 Table，再練結構化參照與自動擴展。",
        header: ["訂單編號", "訂單日期", "通路", "區域", "業務", "產品", "金額"],
        rows: [
          ["SO-1001", "2026-01-03", "官網", "北區", "王小明", "滑鼠", 12500],
          ["SO-1002", "2026-01-07", "經銷", "中區", "李大明", "螢幕", 46800],
          ["SO-1003", "2026-01-09", "官網", "南區", "張志偉", "鍵盤", 18400],
          ["SO-1004", "2026-01-14", "直營", "北區", "王小明", "Webcam", 22900],
          ["SO-1005", "2026-01-19", "經銷", "南區", "陳雅琪", "清淨機", 55200],
          ["SO-1006", "2026-01-26", "官網", "中區", "張志偉", "筆電支架", 9800],
        ],
      },
    ],
  },
  "P3-05": {
    title: "保護與安全練習資料",
    intro: "這份範本很適合拿來練哪些欄位該開放、哪些欄位該鎖。你可以直接把輸入區、計算區和輸出區切開來做工作表保護。",
    files: [
      {
        fileName: "p3-05-budget-template.csv",
        label: "預算範本資料",
        description: "適合練鎖定公式欄、保護工作表與區分輸入區。",
        header: ["部門", "項目", "負責人", "預算金額", "實際金額", "差異", "狀態"],
        rows: [
          ["業務", "廣告投放", "王小明", 120000, 108000, -12000, "正常"],
          ["財務", "顧問費", "林美華", 60000, 60000, 0, "正常"],
          ["資訊", "SaaS 訂閱", "陳雅琪", 90000, 96500, 6500, "超支"],
          ["客服", "教育訓練", "吳建志", 45000, 38500, -6500, "正常"],
        ],
      },
    ],
  },
  "P4-03": {
    title: "假設分析練習資料",
    intro: "這份資料適合練目標搜尋和情境反推。先把單價、成本、流量和轉換率放進模型，再去推要達標得改哪個數字。",
    files: [
      {
        fileName: "p4-03-what-if-pricing.csv",
        label: "價格情境資料",
        description: "可拿來做目標搜尋、情境分析與簡單 Solver 練習。",
        header: ["產品", "單價", "單位成本", "月流量", "轉換率", "固定成本"],
        rows: [
          ["入門課程包", 1900, 420, 5800, 0.024, 68000],
          ["進階課程包", 3200, 900, 4200, 0.018, 68000],
          ["企業內訓包", 8800, 2400, 1200, 0.011, 68000],
        ],
      },
    ],
  },
  "P4-04": {
    title: "Power Query 練習資料",
    intro: "這課最需要多個月份的原始檔，所以我直接幫你準備了可附加的 CSV。先把 1 到 3 月接起來，再新增 4 月做重新整理最有感。",
    files: [
      {
        fileName: "p4-04-sales-2026-01.csv",
        label: "2026-01 原始 CSV",
        description: "Power Query 匯入第一個月份資料。",
        header: ["訂單日期", "通路", "地區", "產品", "數量", "單價", "業績"],
        rows: [
          ["2026-01-03", "官網", "北區", "滑鼠", 18, 590, 10620],
          ["2026-01-08", "經銷", "中區", "螢幕", 6, 8900, 53400],
          ["2026-01-15", "官網", "南區", "鍵盤", 11, 2200, 24200],
          ["2026-01-26", "直營", "北區", "清淨機", 4, 12800, 51200],
        ],
      },
      {
        fileName: "p4-04-sales-2026-02.csv",
        label: "2026-02 原始 CSV",
        description: "第二個月份資料，可練追加與轉換。",
        header: ["訂單日期", "通路", "地區", "產品", "數量", "單價", "業績"],
        rows: [
          ["2026-02-05", "官網", "北區", "滑鼠", 21, 590, 12390],
          ["2026-02-13", "經銷", "中區", "螢幕", 5, 8900, 44500],
          ["2026-02-17", "官網", "南區", "鍵盤", 14, 2200, 30800],
          ["2026-02-24", "直營", "北區", "清淨機", 5, 12800, 64000],
        ],
      },
      {
        fileName: "p4-04-sales-2026-03.csv",
        label: "2026-03 原始 CSV",
        description: "第三個月份資料，可與前兩月附加成同一張表。",
        header: ["訂單日期", "通路", "地區", "產品", "數量", "單價", "業績"],
        rows: [
          ["2026-03-02", "官網", "北區", "滑鼠", 19, 590, 11210],
          ["2026-03-07", "經銷", "中區", "螢幕", 7, 8900, 62300],
          ["2026-03-16", "官網", "南區", "鍵盤", 12, 2200, 26400],
          ["2026-03-29", "直營", "北區", "清淨機", 6, 12800, 76800],
        ],
      },
      {
        fileName: "p4-04-sales-2026-04.csv",
        label: "2026-04 新增資料",
        description: "做完流程後再匯入這份，最適合拿來測重新整理。",
        header: ["訂單日期", "通路", "地區", "產品", "數量", "單價", "業績"],
        rows: [
          ["2026-04-04", "官網", "北區", "滑鼠", 24, 590, 14160],
          ["2026-04-12", "經銷", "中區", "螢幕", 6, 8900, 53400],
          ["2026-04-18", "官網", "南區", "鍵盤", 15, 2200, 33000],
          ["2026-04-26", "直營", "北區", "清淨機", 7, 12800, 89600],
        ],
      },
      {
        fileName: "p4-04-readme.txt",
        label: "Power Query 練習說明",
        description: "建議的匯入與重新整理順序。",
        text: [
          "P4-04 Power Query 練習順序",
          "",
          "1. 先匯入 p4-04-sales-2026-01.csv。",
          "2. 把欄位型別、欄名與日期格式整理好。",
          "3. 再匯入 2026-02、2026-03，做 Append / 合併。",
          "4. 最後新增 2026-04，按重新整理，確認流程能重跑。",
        ].join("\n"),
      },
    ],
  },
  "P4-05": {
    title: "資料模型練習資料",
    intro: "資料模型要有多張表才有感，所以我直接幫你拆成事實表和維度表。最適合拿來練關聯、量值和模型樞紐。",
    files: [
      {
        fileName: "p4-05-fact-sales.csv",
        label: "Fact Sales",
        description: "事實表，記錄每一筆銷售。",
        header: ["日期", "產品ID", "通路", "數量", "金額"],
        rows: [
          ["2026-01-05", "PRD-01", "官網", 18, 120000],
          ["2026-01-12", "PRD-03", "經銷", 9, 86000],
          ["2026-02-03", "PRD-02", "直營", 7, 91000],
          ["2026-02-18", "PRD-01", "官網", 11, 132000],
          ["2026-03-04", "PRD-04", "經銷", 15, 74000],
          ["2026-03-20", "PRD-02", "直營", 6, 88000],
        ],
      },
      {
        fileName: "p4-05-dim-products.csv",
        label: "Dim Products",
        description: "產品維度表，可跟 Fact Sales 透過產品ID 關聯。",
        header: ["產品ID", "產品名稱", "類別", "品牌"],
        rows: [
          ["PRD-01", "無線耳機", "電子", "Aurora"],
          ["PRD-02", "清淨機", "家電", "Breeze"],
          ["PRD-03", "機械鍵盤", "電子", "Forge"],
          ["PRD-04", "烤箱", "家電", "OvenLab"],
        ],
      },
      {
        fileName: "p4-05-dim-calendar.csv",
        label: "Dim Calendar",
        description: "日曆維度表，可延伸成年、季、月切片分析。",
        header: ["日期", "年", "季", "月", "月份標籤"],
        rows: [
          ["2026-01-05", 2026, "Q1", 1, "2026-01"],
          ["2026-01-12", 2026, "Q1", 1, "2026-01"],
          ["2026-02-03", 2026, "Q1", 2, "2026-02"],
          ["2026-02-18", 2026, "Q1", 2, "2026-02"],
          ["2026-03-04", 2026, "Q1", 3, "2026-03"],
          ["2026-03-20", 2026, "Q1", 3, "2026-03"],
        ],
      },
    ],
  },
  "P5-01": {
    title: "VBA 基礎練習資料",
    intro: "這份原始資料很適合拿來錄巨集和寫第一個 VBA 腳本。先做簡單整理與格式化，再慢慢接到自動化流程。",
    files: [
      {
        fileName: "p5-01-vba-sales-raw.csv",
        label: "巨集錄製原始資料",
        description: "適合練排序、加總、格式化與錄製第一個巨集。",
        header: ["日期", "業務", "產品", "數量", "單價", "金額"],
        rows: [
          ["2026-04-01", "王小明", "滑鼠", 12, 590, 7080],
          ["2026-04-02", "李大明", "鍵盤", 8, 2200, 17600],
          ["2026-04-03", "張志偉", "螢幕", 3, 8900, 26700],
          ["2026-04-03", "王小明", "Webcam", 6, 1500, 9000],
          ["2026-04-04", "陳雅琪", "清淨機", 2, 12800, 25600],
        ],
      },
    ],
  },
  "P5-02": {
    title: "VBA 進階練習資料",
    intro: "進階 VBA 比較需要一點批次處理感，所以我準備了任務清單和事件紀錄。很適合練陣列、Dictionary 和錯誤處理。",
    files: [
      {
        fileName: "p5-02-batch-tasks.csv",
        label: "批次處理任務",
        description: "可用來練迴圈、字典與批次狀態更新。",
        header: ["任務ID", "檔案名稱", "部門", "狀態", "優先級"],
        rows: [
          ["TASK-001", "sales_north.xlsx", "業務", "待處理", "高"],
          ["TASK-002", "sales_central.xlsx", "業務", "待處理", "中"],
          ["TASK-003", "budget_finance.xlsx", "財務", "待處理", "高"],
          ["TASK-004", "headcount_hr.xlsx", "人資", "待處理", "低"],
          ["TASK-005", "tickets_support.xlsx", "客服", "待處理", "中"],
        ],
      },
      {
        fileName: "p5-02-event-log.csv",
        label: "事件紀錄",
        description: "可練錯誤處理、事件偵測與執行紀錄彙總。",
        header: ["時間", "事件類型", "工作表", "訊息"],
        rows: [
          ["2026-04-20 09:00", "Open", "Dashboard", "使用者開啟工作簿"],
          ["2026-04-20 09:03", "Change", "Input", "B4 儲存格被修改"],
          ["2026-04-20 09:06", "Refresh", "Sales", "重新整理完成"],
          ["2026-04-20 09:08", "Error", "Macro", "找不到指定檔案"],
        ],
      },
    ],
  },
  "P5-03": {
    title: "VBA 實戰專案資料",
    intro: "這裡我直接準備一組比較像專案的資料。最適合練把讀資料、算結果、輸出報表串成完整流程。",
    files: [
      {
        fileName: "p5-03-orders.csv",
        label: "Orders",
        description: "訂單來源資料，可做月報彙總與批次輸出。",
        header: ["訂單編號", "日期", "客戶ID", "產品ID", "數量", "金額"],
        rows: [
          ["SO-2001", "2026-04-01", "C001", "PRD-01", 12, 7080],
          ["SO-2002", "2026-04-02", "C002", "PRD-02", 8, 17600],
          ["SO-2003", "2026-04-03", "C001", "PRD-03", 3, 26700],
          ["SO-2004", "2026-04-04", "C003", "PRD-01", 6, 3540],
          ["SO-2005", "2026-04-04", "C004", "PRD-04", 2, 25600],
        ],
      },
      {
        fileName: "p5-03-customers.csv",
        label: "Customers",
        description: "客戶主檔，可做補欄位與報表輸出。",
        header: ["客戶ID", "客戶名稱", "區域", "業務"],
        rows: [
          ["C001", "晨星科技", "北區", "王小明"],
          ["C002", "海峰商行", "中區", "李大明"],
          ["C003", "綠葉生活", "南區", "張志偉"],
          ["C004", "安和電器", "北區", "陳雅琪"],
        ],
      },
      {
        fileName: "p5-03-project-config.csv",
        label: "Config",
        description: "專案設定表，可拿來練參數化與輸出控制。",
        header: ["參數", "值"],
        rows: [
          ["輸出月份", "2026-04"],
          ["輸出資料夾", "reports"],
          ["報表標題", "四月銷售月報"],
          ["是否寄送通知", "N"],
        ],
      },
    ],
  },
  "P5-04": {
    title: "綜合挑戰資料包",
    intro: "綜合挑戰就該像真的在接任務，所以我準備了一組多來源資料。你可以拿這些檔案自己決定該用公式、樞紐、Power Query 還是 VBA 解。",
    files: [
      {
        fileName: "p5-04-capstone-sales.csv",
        label: "Sales",
        description: "主銷售資料，可做彙總、分析與報表。",
        header: ["日期", "產品ID", "區域", "業務", "金額"],
        rows: [
          ["2026-03-01", "PRD-01", "北區", "王小明", 128000],
          ["2026-03-05", "PRD-02", "中區", "李大明", 86000],
          ["2026-03-09", "PRD-03", "南區", "張志偉", 91000],
          ["2026-03-18", "PRD-04", "北區", "陳雅琪", 74000],
          ["2026-03-24", "PRD-01", "中區", "王小明", 132000],
        ],
      },
      {
        fileName: "p5-04-capstone-targets.csv",
        label: "Targets",
        description: "目標資料，可做達成率、差異分析與儀表板。",
        header: ["區域", "月目標"],
        rows: [
          ["北區", 210000],
          ["中區", 180000],
          ["南區", 160000],
        ],
      },
      {
        fileName: "p5-04-capstone-returns.csv",
        label: "Returns",
        description: "退貨資料，可加入異常與淨額分析。",
        header: ["日期", "產品ID", "區域", "退貨金額"],
        rows: [
          ["2026-03-07", "PRD-03", "南區", 8000],
          ["2026-03-17", "PRD-01", "北區", 5000],
          ["2026-03-26", "PRD-04", "北區", 3000],
        ],
      },
      {
        fileName: "p5-04-challenge-brief.txt",
        label: "挑戰說明",
        description: "適合當這份資料包的任務說明。",
        text: [
          "P5-04 綜合挑戰建議做法",
          "",
          "1. 先把 Sales、Targets、Returns 整理成可分析格式。",
          "2. 做出各區域達成率與淨銷售分析。",
          "3. 判斷哪一段最適合用公式、哪一段該交給樞紐或 Power Query。",
          "4. 如果你想挑戰進階版，再用 VBA 自動輸出月報。",
        ].join("\n"),
      },
    ],
  },
};

function trimTrailingEmptyColumns(header, rows) {
  let lastUsefulIndex = header.length - 1;
  const isEmpty = (value) => value === "" || value == null;
  while (lastUsefulIndex >= 0) {
    const headerEmpty = isEmpty(header[lastUsefulIndex]);
    const colEmpty = rows.every((row) => isEmpty(row[lastUsefulIndex]));
    if (!headerEmpty || !colEmpty) break;
    lastUsefulIndex -= 1;
  }
  const width = Math.max(lastUsefulIndex + 1, 1);
  return {
    header: header.slice(0, width),
    rows: rows.map((row) => row.slice(0, width)),
  };
}

function csvEscape(value) {
  const text = value == null ? "" : String(value);
  if (/[",\n]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}

function toCsv(header, rows) {
  const lines = [header.map(csvEscape).join(",")];
  rows.forEach((row) => {
    lines.push(row.map(csvEscape).join(","));
  });
  return `\ufeff${lines.join("\n")}\n`;
}

function formatFromName(fileName) {
  const ext = path.extname(fileName).slice(1).toUpperCase();
  return ext || "FILE";
}

async function writeCsvAsset(slug, file) {
  const filePath = path.join(outputDir, file.fileName);
  const csv = toCsv(file.header, file.rows);
  await fs.writeFile(filePath, csv, "utf8");
  return {
    label: file.label,
    description: file.description,
    href: `practice-data/${file.fileName}`,
    format: formatFromName(file.fileName),
    fileName: file.fileName,
  };
}

async function writeTextAsset(file) {
  const filePath = path.join(outputDir, file.fileName);
  await fs.writeFile(filePath, file.text, "utf8");
  return {
    label: file.label,
    description: file.description,
    href: `practice-data/${file.fileName}`,
    format: formatFromName(file.fileName),
    fileName: file.fileName,
  };
}

async function writeBundleAsset(slug, files) {
  if (!files.length || files.length === 1) return null;
  const fileName = `${slug.toLowerCase()}-practice-pack.zip`;
  await execFileAsync("zip", ["-q", "-j", fileName, ...files.map((file) => file.fileName)], {
    cwd: outputDir,
  });
  return {
    label: "本課資料包",
    description: `一次下載這課的 ${files.length} 份練習資料，比逐個點更省事。`,
    href: `practice-data/${fileName}`,
    format: "ZIP",
    fileName,
  };
}

async function buildAutoAssets(manifest) {
  for (const [slug, meta] of Object.entries(interLessonMeta)) {
    const lesson = XLSX_CONTENT.lessons?.[slug];
    if (!lesson?.inter?.dataHeader || !lesson?.inter?.dataRows) continue;
    const cleaned = trimTrailingEmptyColumns(lesson.inter.dataHeader, lesson.inter.dataRows);
    const files = [await writeCsvAsset(slug, {
      fileName: meta.fileName,
      label: meta.label,
      description: meta.description,
      header: cleaned.header,
      rows: cleaned.rows,
    })];
    const bundle = await writeBundleAsset(slug, files);
    manifest[slug] = {
      title: meta.title,
      intro: meta.intro,
      files,
      ...(bundle ? { bundle } : {}),
    };
  }
}

async function buildCuratedAssets(manifest) {
  for (const [slug, config] of Object.entries(curatedLessonAssets)) {
    const files = [];
    for (const file of config.files) {
      if (file.header && file.rows) files.push(await writeCsvAsset(slug, file));
      else if (typeof file.text === "string") files.push(await writeTextAsset(file));
    }
    const bundle = await writeBundleAsset(slug, files);
    manifest[slug] = {
      title: config.title,
      intro: config.intro,
      files,
      ...(bundle ? { bundle } : {}),
    };
  }
}

async function main() {
  await fs.rm(outputDir, { recursive: true, force: true });
  await fs.mkdir(outputDir, { recursive: true });

  const manifest = {};
  await buildAutoAssets(manifest);
  await buildCuratedAssets(manifest);

  const missing = Object.keys(XLSX_CONTENT.lessons).filter((slug) => !manifest[slug]);
  if (missing.length) {
    throw new Error(`Missing download manifest for lessons: ${missing.join(", ")}`);
  }

  const manifestSource = `export const LESSON_DOWNLOADS = ${JSON.stringify(manifest, null, 2)};\n`;
  await fs.writeFile(manifestPath, manifestSource, "utf8");
  console.log(`Generated ${Object.keys(manifest).length} lesson download entries in ${outputDir}`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
