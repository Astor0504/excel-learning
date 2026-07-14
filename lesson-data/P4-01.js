// 由 xlsx-content.js + lesson-demos-data.js 拆分(2026-07-14);本檔是 P4-01 的唯一資料來源
export const CONTENT = {
 "knowledge": {
  "title": "📊 動態陣列",
  "subtitle": "一個公式吐出整片結果，來源變動時自動更新",
  "sections": [
   {
    "title": "📌 三個核心觀念",
    "items": [
     [
      "Spill 溢出",
      "公式只輸入在左上角，結果會自動展開到多格",
      "輸出區被擋住時會出現 #SPILL!"
     ],
     [
      "D2#",
      "# 代表從 D2 開始的整個溢出範圍",
      "常用來接圖表、下拉選單或後續公式"
     ],
     [
      "條件陣列",
      "FILTER 的條件會產生 TRUE/FALSE 清單",
      "多條件 AND 用 *，OR 常用 +"
     ]
    ]
   },
   {
    "title": "🧰 常用組合",
    "items": [
     [
      "FILTER",
      "篩出符合條件的整列資料",
      "=FILTER(A:E,B:B=\"業務\",\"查無資料\")"
     ],
     [
      "SORT + FILTER",
      "先篩選，再排序",
      "=SORT(FILTER(A:E,B:B=\"業務\"),3,-1)"
     ],
     [
      "SORT + UNIQUE",
      "建立會自動更新的不重複清單",
      "=SORT(UNIQUE(B:B))"
     ]
    ]
   },
   {
    "title": "⚠️ 交付提醒",
    "items": [
     [
      "#CALC!",
      "FILTER 找不到資料又沒寫第 3 參數時容易出現",
      "正式報表請填「查無資料」"
     ],
     [
      "版本相容",
      "FILTER / SORT / UNIQUE 需要 Microsoft 365 或 Excel 2021+",
      "舊版使用者要準備替代方案"
     ],
     [
      "Table 邊界",
      "原始資料適合放 Table；溢出公式建議放在 Table 外",
      "避免表格擴張和溢出範圍互相干擾"
     ]
    ]
   }
  ],
  "handsTasks": [
   "用 FILTER 篩出業務部資料",
   "用 FILTER 做兩個條件",
   "用 SORT + UNIQUE 建立部門清單",
   "用 D2# 引用溢出範圍"
  ]
 },
 "pro": {
  "title": "📊 陣列動態  ─  FILTER / SORT / UNIQUE / SEQUENCE（365）",
  "subtitle": "難度：🟡 Lv.3  |  微任務數：8 題  |  建議時間：每題 2~3 分鐘",
  "dataHeader": [
   "員工",
   "部門",
   "月薪",
   "年資",
   "績效",
   "",
   ""
  ],
  "dataRows": [
   [
    "王小明",
    "業務",
    55000,
    7,
    88,
    "",
    ""
   ],
   [
    "林美華",
    "財務",
    48000,
    5,
    72,
    "",
    ""
   ],
   [
    "張志偉",
    "業務",
    68000,
    10,
    95,
    "",
    ""
   ],
   [
    "陳雅婷",
    "人事",
    42000,
    3,
    65,
    "",
    ""
   ],
   [
    "李建宏",
    "資訊",
    75000,
    8,
    91,
    "",
    ""
   ],
   [
    "劉淑芬",
    "財務",
    52000,
    6,
    79,
    "",
    ""
   ],
   [
    "黃志強",
    "業務",
    82000,
    11,
    98,
    "",
    ""
   ],
   [
    "吳雅琪",
    "資訊",
    63000,
    4,
    84,
    "",
    ""
   ]
  ],
  "tasks": [
   {
    "numLabel": "🟢1",
    "desc": "用 FILTER 篩選出業務部的員工，找不到時顯示「查無資料」",
    "time": "3m",
    "hint": "=FILTER(A6:E13,B6:B13=\"業務\",\"查無資料\")",
    "answer": "=FILTER(A6:E13,B6:B13=\"業務\",\"查無資料\")",
    "explain": "FILTER 自動展開結果，第 3 參數可避免空結果變成 #CALC!"
   },
   {
    "numLabel": "🟢2",
    "desc": "用 FILTER 篩選月薪 >= 60000 的員工",
    "time": "2m",
    "hint": "=FILTER(A6:E13,C6:C13>=60000,\"查無資料\")",
    "answer": "=FILTER(A6:E13,C6:C13>=60000,\"查無資料\")",
    "explain": "條件可以是任何比較運算"
   },
   {
    "numLabel": "🔵3",
    "desc": "依月薪由高到低排序",
    "time": "2m",
    "hint": "=SORT(A6:E13,3,-1)",
    "answer": "=SORT(A6:E13,3,-1)",
    "explain": "SORT(範圍,依哪欄,1升/-1降)"
   },
   {
    "numLabel": "🔵4",
    "desc": "用 SORT + UNIQUE 取出並排序所有不重複部門",
    "time": "2m",
    "hint": "=SORT(UNIQUE(B6:B13))",
    "answer": "=SORT(UNIQUE(B6:B13))",
    "explain": "UNIQUE 先去重，SORT 再排序，適合做動態下拉來源"
   },
   {
    "numLabel": "🟡5",
    "desc": "產生 1~10 的序號",
    "time": "1m",
    "hint": "=SEQUENCE(10)",
    "answer": "=SEQUENCE(10)",
    "explain": "SEQUENCE(列,欄,起始,間隔)"
   },
   {
    "numLabel": "🟡6",
    "desc": "篩選業務部且績效 >= 90",
    "time": "3m",
    "hint": "=FILTER(A6:E13,(B6:B13=\"業務\")*(E6:E13>=90),\"查無資料\")",
    "answer": "=FILTER(A6:E13,(B6:B13=\"業務\")*(E6:E13>=90),\"查無資料\")",
    "explain": "多條件用 * 相乘（AND 效果）"
   },
   {
    "numLabel": "🔴7",
    "desc": "假設 H6 是部門溢出清單，用 # 計算共有幾個部門",
    "time": "3m",
    "hint": "=COUNTA(H6#)",
    "answer": "=COUNTA(H6#)",
    "explain": "H6# 代表從 H6 開始的整個溢出範圍"
   },
   {
    "numLabel": "🔴8",
    "desc": "依績效排序後只取前 3 名",
    "time": "3m",
    "hint": "=TAKE(SORT(A6:E13,5,-1),3)",
    "answer": "=TAKE(SORT(A6:E13,5,-1),3)",
    "explain": "TAKE + SORT 組合，SORT 完再取前 N 列"
   }
  ]
 },
 "meta": {
  "phase": "⚡ Phase 4：進階自動化    第 7~10 週  |  目標：處理大量資料、自動化流程",
  "stage": "第13階段",
  "topics": "FILTER / SORT / UNIQUE / SEQUENCE / Spill",
  "difficulty": "🟡 Lv.3",
  "taskCount": "8 題",
  "time": "20 分鐘",
  "xp": "200 XP"
 },
 "inter": {
  "title": "🔢 陣列與動態函式 — SORT / FILTER / UNIQUE",
  "subtitle": "在黃色邊框儲存格輸入公式 → 答對變綠 ✅ | 答錯變紅 ❌ | F欄提示選取可見",
  "dataHeader": [
   "產品",
   "類別",
   "單價",
   "銷量",
   "評分",
   "",
   ""
  ],
  "dataRows": [
   [
    "無線滑鼠",
    "周邊",
    590,
    120,
    4.5,
    "",
    ""
   ],
   [
    "機械鍵盤",
    "周邊",
    2200,
    45,
    4.8,
    "",
    ""
   ],
   [
    "27吋螢幕",
    "顯示",
    8900,
    18,
    4.7,
    "",
    ""
   ],
   [
    "網路攝影機",
    "周邊",
    1500,
    67,
    4.2,
    "",
    ""
   ],
   [
    "USB集線器",
    "配件",
    450,
    200,
    4,
    "",
    ""
   ],
   [
    "外接硬碟",
    "儲存",
    2800,
    32,
    4.6,
    "",
    ""
   ]
  ],
  "tasks": [
   {
    "num": 1,
    "difficulty": "🟢",
    "desc": "用 FILTER 篩出周邊類產品",
    "hint": "=FILTER(A6:E13,B6:B13=\"周邊\",\"查無資料\")",
    "answer": "=FILTER(A6:E13,B6:B13=\"周邊\",\"查無資料\")"
   },
   {
    "num": 2,
    "difficulty": "🟢",
    "desc": "用 SORT 依單價由高到低排序",
    "hint": "=SORT(資料範圍,第3欄,-1)",
    "answer": "=SORT(A6:E13,3,-1)"
   },
   {
    "num": 3,
    "difficulty": "🔵",
    "desc": "列出不重複的類別清單",
    "hint": "=UNIQUE(類別欄)",
    "answer": "=UNIQUE(B6:B13)"
   },
   {
    "num": 4,
    "difficulty": "🔵",
    "desc": "把不重複類別清單排序",
    "hint": "=SORT(UNIQUE(類別欄))",
    "answer": "=SORT(UNIQUE(B6:B13))"
   },
   {
    "num": 5,
    "difficulty": "🟡",
    "desc": "篩選單價超過 1000 的產品",
    "hint": "=FILTER(A6:E13,C6:C13>1000,\"查無資料\")",
    "answer": "=FILTER(A6:E13,C6:C13>1000,\"查無資料\")"
   },
   {
    "num": 6,
    "difficulty": "🟡",
    "desc": "篩選周邊類且評分 >= 4.5 的產品",
    "hint": "多條件 AND 用 *",
    "answer": "=FILTER(A6:E13,(B6:B13=\"周邊\")*(E6:E13>=4.5),\"查無資料\")"
   },
   {
    "num": 7,
    "difficulty": "🔴",
    "desc": "假設 H6 是類別溢出清單，用 # 計算類別數",
    "hint": "H6# 代表 H6 開始的整個溢出結果",
    "answer": "=COUNTA(H6#)"
   },
   {
    "num": 8,
    "difficulty": "🔴",
    "desc": "依銷量排序後取前 3 名產品",
    "hint": "=TAKE(SORT(A6:E13,4,-1),3)",
    "answer": "=TAKE(SORT(A6:E13,4,-1),3)"
   }
  ]
 }
};
export const DEMOS = [
 {
  "id": "filter-multi-condition",
  "kind": "formula",
  "badge": "公式拆解",
  "title": "FILTER：用 TRUE/FALSE 陣列挑出整列資料",
  "subtitle": "FILTER 的重點不是只會篩選，而是理解每一列都會先被條件判斷成 TRUE 或 FALSE。",
  "outcome": "結果：只有業務部且績效 >= 90 的列會溢出成新清單",
  "formulaParts": [
   {
    "text": "=FILTER("
   },
   {
    "text": "A2:E7",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ", "
   },
   {
    "text": "(B2:B7=\"業務\")",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": "*"
   },
   {
    "text": "(E2:E7>=90)",
    "step": 3,
    "tone": "c3"
   },
   {
    "text": ", "
   },
   {
    "text": "\"查無資料\"",
    "step": 4,
    "tone": "c4"
   },
   {
    "text": ")"
   }
  ],
  "board": {
   "title": "員工績效表",
   "columns": [
    "員工",
    "部門",
    "月薪",
    "績效"
   ],
   "rows": [
    [
     {
      "text": "王小明",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "業務",
      "marks": {
       "2": "match"
      }
     },
     {
      "text": "55,000",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "88",
      "marks": {
       "3": "focus"
      }
     }
    ],
    [
     {
      "text": "張志偉",
      "marks": {
       "1": "focus",
       "4": "result"
      }
     },
     {
      "text": "業務",
      "marks": {
       "2": "match",
       "4": "result"
      }
     },
     {
      "text": "68,000",
      "marks": {
       "1": "focus",
       "4": "result"
      }
     },
     {
      "text": "95",
      "marks": {
       "3": "match",
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "李建宏",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "資訊",
      "marks": {
       "2": "focus"
      }
     },
     {
      "text": "75,000",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "91",
      "marks": {
       "3": "match"
      }
     }
    ],
    [
     {
      "text": "黃志強",
      "marks": {
       "1": "focus",
       "4": "result"
      }
     },
     {
      "text": "業務",
      "marks": {
       "2": "match",
       "4": "result"
      }
     },
     {
      "text": "82,000",
      "marks": {
       "1": "focus",
       "4": "result"
      }
     },
     {
      "text": "98",
      "marks": {
       "3": "match",
       "4": "result"
      }
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. 先圈出要回傳的整張資料",
    "tone": "c1",
    "text": "A2:E7 是 FILTER 最後要吐出的資料範圍。只要某一列通過條件，這一整列就會出現在結果裡。",
    "caption": "FILTER 不是只回傳單一值，而是能回傳整片資料。"
   },
   {
    "label": "2. 第一個條件：部門必須是業務",
    "tone": "c2",
    "text": "B2:B7=\"業務\" 會產生一串 TRUE/FALSE。業務部是 TRUE，其他部門是 FALSE。",
    "caption": "條件範圍的列數必須和資料範圍一致。"
   },
   {
    "label": "3. 第二個條件：績效必須 >= 90",
    "tone": "c3",
    "text": "E2:E7>=90 也會產生一串 TRUE/FALSE。兩個條件中間用 *，代表兩邊都要成立。",
    "caption": "多條件 AND 用 *；OR 情境通常用 +。"
   },
   {
    "label": "4. 找不到資料時要有可讀訊息",
    "tone": "c4",
    "text": "第 3 參數填「查無資料」後，如果沒有任何列符合條件，就不會直接丟出 #CALC!。",
    "caption": "交付報表時，錯誤訊息要變成使用者看得懂的文字。"
   }
  ]
 },
 {
  "id": "unique-sort-spill",
  "kind": "formula",
  "badge": "動態清單",
  "title": "SORT + UNIQUE：做一份會自動更新的下拉來源",
  "subtitle": "先去重，再排序，最後用 D2# 引用整個溢出結果。",
  "outcome": "結果：重複的部門只留下 1 次，並排序成可引用的動態清單",
  "formulaParts": [
   {
    "text": "=SORT(",
    "step": 3,
    "tone": "c3"
   },
   {
    "text": "UNIQUE(",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": "B2:B9",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": "))"
   }
  ],
  "board": {
   "title": "部門欄位會重複出現",
   "columns": [
    "原始部門",
    "UNIQUE 後",
    "SORT 後"
   ],
   "rows": [
    [
     {
      "text": "業務",
      "marks": {
       "1": "focus",
       "2": "match"
      }
     },
     {
      "text": "業務",
      "marks": {
       "2": "result"
      }
     },
     {
      "text": "人事",
      "marks": {
       "3": "result",
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "財務",
      "marks": {
       "1": "focus",
       "2": "match"
      }
     },
     {
      "text": "財務",
      "marks": {
       "2": "result"
      }
     },
     {
      "text": "業務",
      "marks": {
       "3": "result",
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "業務",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "人事",
      "marks": {
       "2": "result"
      }
     },
     {
      "text": "財務",
      "marks": {
       "3": "result",
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "資訊",
      "marks": {
       "1": "focus",
       "2": "match"
      }
     },
     {
      "text": "資訊",
      "marks": {
       "2": "result"
      }
     },
     {
      "text": "資訊",
      "marks": {
       "3": "result",
       "4": "result"
      }
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. 來源範圍：先抓會重複的欄",
    "tone": "c1",
    "text": "B2:B9 是原始部門欄。因為資料會新增或重複，所以不適合手動維護清單。",
    "caption": "實務上可改用 Table 欄位，例如 員工表[部門]。"
   },
   {
    "label": "2. UNIQUE 先去掉重複值",
    "tone": "c2",
    "text": "UNIQUE 會保留每個部門第一次出現的結果，做出一份不重複清單。",
    "caption": "這一步常拿來做資料驗證的來源。"
   },
   {
    "label": "3. SORT 再把清單排整齊",
    "tone": "c3",
    "text": "SORT 包在外層，讓去重後的部門清單按照順序排列。",
    "caption": "巢狀公式從內往外讀：先 UNIQUE，再 SORT。"
   },
   {
    "label": "4. 用 D2# 引用整片溢出結果",
    "tone": "c4",
    "text": "公式只寫在 D2，但結果會向下溢出。其他公式或下拉選單可以用 D2# 引用整個動態清單。",
    "caption": "# 是動態陣列最重要的引用符號。"
   }
  ]
 }
];
