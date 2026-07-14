// 由 xlsx-content.js + lesson-demos-data.js 拆分(2026-07-14);本檔是 P4-05 的唯一資料來源
export const CONTENT = {
 "knowledge": {
  "title": "🔗 Power Pivot 與資料模型",
  "subtitle": "難度：🔴 Lv.5  |  建模思維課  |  Mac 以觀念判斷與協作策略為主",
  "sections": [
   {
    "title": "📌 什麼時候真的需要資料模型",
    "items": [
     [
      "多張表",
      "訂單表、客戶表、商品表分開存放，需要共同分析"
     ],
     [
      "共用指標",
      "同一個毛利率、同期比較、部門業績，要給多張樞紐一起用"
     ],
     [
      "大量資料",
      "資料列數越來越大，用公式或單表樞紐開始難維護"
     ],
     [
      "團隊協作",
      "不想每個人都各自寫一套版本不同的公式"
     ]
    ]
   },
   {
    "title": "🧠 核心概念",
    "items": [
     [
      "Fact Table",
      "交易明細表，通常列數最多，像訂單、出貨、付款"
     ],
     [
      "Dimension Table",
      "提供分類與描述的表，像客戶、商品、日期、部門"
     ],
     [
      "Relationship",
      "用共同欄位把表連起來，例如 訂單[客戶ID] → 客戶[客戶ID]"
     ],
     [
      "Measure",
      "寫在模型裡的計算邏輯，例如 總業績、毛利率、去年同期"
     ],
     [
      "Calculated Column",
      "逐列算出的新欄位，適合列層級標記，不適合所有彙總場景"
     ]
    ]
   },
   {
    "title": "🎯 典型案例：訂單 / 客戶 / 商品三表分析",
    "items": [
     [
      "原始表",
      "訂單表有金額與商品ID，客戶表有客戶等級，商品表有類別與成本"
     ],
     [
      "想回答的問題",
      "高價值客戶在各商品類別的毛利貢獻是多少？"
     ],
     [
      "傳統做法",
      "多欄 VLOOKUP / XLOOKUP 串接，欄位一改就容易壞"
     ],
     [
      "資料模型做法",
      "先建關聯，再用 Measure 統一計算毛利與貢獻度，樞紐只負責展示"
     ]
    ]
   },
   {
    "title": "🍎 Mac 使用者的決策路線",
    "items": [
     [
      "先撐住",
      "如果只是兩表、三表的小型分析，先用 Power Query + 樞紐 + XLOOKUP"
     ],
     [
      "需要建模",
      "一旦需要共用量值或多表關聯，就要評估 Windows Excel 或 Power BI"
     ],
     [
      "團隊協作",
      "Mac 端最重要的能力不是硬做 Power Pivot，而是知道何時該交給適合的工具"
     ]
    ]
   }
  ],
  "handsTasks": [
   {
    "num": 1,
    "difficulty": "🟢 暖身",
    "desc": "拿一個你常見的報表題目，先判斷它是單表問題還是多表關聯問題"
   },
   {
    "num": 2,
    "difficulty": "🟢 暖身",
    "desc": "畫一張簡單關聯圖：訂單表連到客戶表、商品表，各自用哪個 ID 連接"
   },
   {
    "num": 3,
    "difficulty": "🔵 標準",
    "desc": "列出 3 個你常重複計算的指標，思考哪些適合做成 Measure，哪些只是一般欄位"
   },
   {
    "num": 4,
    "difficulty": "🔵 標準",
    "desc": "用 Power Query 先整理兩張表，再評估這個題目是否真的需要 Data Model"
   },
   {
    "num": 5,
    "difficulty": "🟡 變化",
    "desc": "把「傳統公式串表」和「資料模型建關聯」各寫出 3 個優缺點，建立工具判斷力"
   },
   {
    "num": 6,
    "difficulty": "🔴 挑戰",
    "desc": "找一個跨部門或跨月份分析需求，寫下：若在 Mac 上完成，你會選 Power Query + 樞紐、請 Windows 同事協作，還是直接用 Power BI，並說明理由"
   }
  ]
 },
 "meta": {
  "phase": "⚡ Phase 4：進階自動化    第 7~10 週  |  目標：處理大量資料、自動化流程",
  "stage": "第16階段",
  "topics": "資料模型 / 關聯 / Measure / DAX 判斷力",
  "difficulty": "🔴 Lv.5",
  "taskCount": "6 題",
  "time": "20 分鐘",
  "xp": "250 XP"
 }
};
export const DEMOS = [
 {
  "id": "data-model-relationship",
  "kind": "media",
  "badge": "概念演示",
  "title": "資料模型：把多張表串成同一套分析系統",
  "subtitle": "Power Query 清資料，資料模型管關聯，樞紐只負責把答案展示出來。",
  "media": {
   "kicker": "慢速循環短片",
   "title": "從多表、關聯、量值，到最後交給樞紐展示",
   "src": "/media/p4-05-data-model-demo.webp",
   "video": "/media/p4-05-data-model-demo.mp4",
   "poster": "/media/p4-05-data-model-demo-poster.jpg",
   "alt": "資料模型操作短片：先分開訂單表與產品表，建立產品ID關聯，寫量值公式，最後把結果交給樞紐分析表展示。",
   "note": "這一課本來就偏觀念，所以短片重點不是花俏拖曳，而是把『多表 → 關聯 → 量值 → 樞紐』這條路徑建立起來。對 Mac 使用者尤其重要，因為你更需要判斷何時該交給模型，而不是硬把所有邏輯塞進樞紐。"
  },
  "outcome": "結果：量值只定義一次，之後很多張報表都能共用同一套邏輯",
  "observations": [
   {
    "title": "先分清楚表角色",
    "text": "訂單表是事實表，產品表、日期表是維度表，角色清楚關聯才穩。"
   },
   {
    "title": "再看關聯線",
    "text": "用產品 ID 或日期欄建立關聯，樞紐才能跨表分析。"
   },
   {
    "title": "最後看量值",
    "text": "總業績這類指標定義成量值後，多張報表可以共用同一套算法。"
   }
  ],
  "followup": "學這課時先把表的角色和關聯想清楚，比急著背 DAX 函數更重要。"
 }
];
