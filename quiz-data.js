// ============================================================
// 測驗題庫 — 由 app.js 的 Quiz 模組讀取
// 目前僅在有 #quizBox 元素的頁面會渲染（cheatsheet 等未來功能）
// 每題格式：{ q: 題目, options: [選項], answer: 正解 index, explain: 解析 }
// ============================================================
window.QUIZ_CARDS = [
  // === Phase 1：操作效率 ===
  {
    q: "要一次選到表格的最底列，macOS 的快捷鍵是？",
    options: ["⌘+A", "⌘+↓", "⌘+Shift+↓", "⌘+End"],
    answer: 2,
    explain: "⌘+Shift+↓ 會從目前位置選到資料最底。⌘+↓ 只是跳過去不選取。"
  },
  {
    q: "絕對參照 $A$1 代表什麼？",
    options: ["欄列都鎖", "只鎖欄", "只鎖列", "都不鎖"],
    answer: 0,
    explain: "$ 在誰前面就鎖誰。$A$1 = 欄列都鎖，$A1 只鎖欄，A$1 只鎖列。"
  },
  {
    q: "要計算 A1:A100 但只加「大於 1000」的值，要用？",
    options: ["SUM", "SUMIF", "SUMIFS", "COUNTIF"],
    answer: 1,
    explain: "單一條件用 SUMIF；多個條件才用 SUMIFS。COUNTIF 是數個數不是加總。"
  },

  // === Phase 2：職場即戰力 ===
  {
    q: "VLOOKUP 跟 XLOOKUP 最大的差別是？",
    options: ["速度", "XLOOKUP 可以往左查、預設精確比對", "XLOOKUP 需要付費", "沒差"],
    answer: 1,
    explain: "XLOOKUP 預設精確比對、可雙向查找、可回傳陣列。舊版 Excel 不支援。"
  },
  {
    q: "樞紐分析的四大區域不包含？",
    options: ["列 Rows", "欄 Columns", "值 Values", "公式 Formula"],
    answer: 3,
    explain: "四大區域是：列、欄、值、篩選。公式是樞紐外另外計算的功能。"
  },

  // === Phase 3：專業打磨 ===
  {
    q: "類別超過 5 個要比較大小，最不適合用哪種圖？",
    options: ["橫條圖", "柱狀圖", "圓餅圖", "折線圖"],
    answer: 2,
    explain: "圓餅圖超過 5 片就難讀。改用橫條圖由大到小排序最清楚。"
  },
  {
    q: "要把 2024-03-15 轉成「3月15日」要用？",
    options: ["TEXT(A1, \"m月d日\")", "DATE(A1)", "FORMAT(A1)", "MONTH(A1) & DAY(A1)"],
    answer: 0,
    explain: "TEXT 函式可以把任何數值/日期轉成文字，第二個引數是格式碼。"
  },

  // === Phase 4：進階自動化 ===
  {
    q: "Power Query 主要解決什麼問題？",
    options: ["比公式更快", "自動化資料清整流程", "畫圖表", "寫巨集"],
    answer: 1,
    explain: "PQ 是 ETL 工具：抽資料→清資料→載入。錄一次流程，之後一鍵重跑。"
  },
  {
    q: "下列哪個是 Excel 365 的動態陣列函式？",
    options: ["VLOOKUP", "SUMIF", "UNIQUE", "COUNTA"],
    answer: 2,
    explain: "UNIQUE、FILTER、SORT、SEQUENCE 是動態陣列家族，結果會自動溢位填滿下方儲存格。"
  },

  // === Phase 5：VBA ===
  {
    q: "在 macOS Excel 打開 VBA 編輯器的方法？",
    options: ["Alt+F11", "⌥+F11", "⌘+F11", "工具→巨集→Visual Basic"],
    answer: 3,
    explain: "Mac 的 Alt+F11 有時不靈；最穩的是「工具 → 巨集 → Visual Basic 編輯器」。"
  },
  {
    q: "含巨集的 Excel 檔案要存成什麼副檔名？",
    options: [".xlsx", ".xlsm", ".xlsb", ".xls"],
    answer: 1,
    explain: ".xlsm = 啟用巨集的活頁簿。存成 .xlsx 會把巨集刪掉。"
  }
];
