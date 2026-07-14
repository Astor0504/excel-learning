// 由 xlsx-content.js + lesson-demos-data.js 拆分(2026-07-14);本檔是 P2-02 的唯一資料來源
export const CONTENT = {
 "knowledge": {
  "title": "📚 查找比對的重點不是背函數，是把資料接回來",
  "subtitle": "難度：🔵 Lv.2  |  先分清楚查找值、查找欄、回傳欄與找不到時的處理",
  "sections": [
   {
    "title": "📌 先問自己這 4 件事",
    "items": [
     [
      "你拿什麼去找？",
      "查找值",
      "例如訂單表裡的產品代碼 P003",
      "查找值要乾淨一致，最怕多空白或格式不一樣"
     ],
     [
      "去哪一欄找？",
      "查找範圍",
      "例如商品表的產品代碼欄",
      "XLOOKUP 和 INDEX+MATCH 都會把這一欄獨立寫出來"
     ],
     [
      "要回傳哪一欄？",
      "回傳範圍",
      "例如商品名稱、單價、庫存",
      "XLOOKUP 直接指定回傳欄；VLOOKUP 要手動數第幾欄"
     ],
     [
      "找不到怎麼辦？",
      "找不到時的顯示",
      "例如查無此商品",
      "不要讓 #N/A 直接出現在交付報表"
     ]
    ]
   },
   {
    "title": "⚠️ 3 個最常見的查找錯誤",
    "items": [
     [
      "VLOOKUP 省略第四參數",
      "省略時可能變成模糊比對",
      "職場代碼查找幾乎都要 FALSE 或 0",
      "這是很多錯誤報表的來源"
     ],
     [
      "把 VLOOKUP 當唯一答案",
      "VLOOKUP 只能從最左欄往右查",
      "已知名稱要反查代碼時，改用 XLOOKUP 或 INDEX+MATCH",
      "不要為了公式硬改資料欄位順序"
     ],
     [
      "不處理找不到資料",
      "#N/A 可能代表代碼不存在、格式不一致或資料還沒補",
      "新版用 XLOOKUP 第 4 參數，舊檔用 IFERROR",
      "正式報表要給人看得懂的訊息"
     ]
    ]
   },
   {
    "title": "🎯 專業現場怎麼選",
    "items": [
     [
      "新版環境",
      "優先 XLOOKUP",
      "公式可讀、可往左查、內建找不到處理",
      "這是 Microsoft 365 / Excel 2021 以上的主路線"
     ],
     [
      "舊版相容",
      "熟 INDEX + MATCH",
      "比 VLOOKUP 彈性高，也能處理反向查找",
      "維護 2016 / 2019 或混合環境時很重要"
     ],
     [
      "既有舊檔",
      "看懂並能修 VLOOKUP",
      "確認查找欄、欄位號與第四參數",
      "會維護舊公式，比只會寫新公式更實用"
     ]
    ]
   }
  ],
  "handsTasks": [
   {
    "num": 1,
    "difficulty": "🟢 暖身",
    "desc": "先用 XLOOKUP 查 P003 的產品名稱，再把回傳欄改成單價欄"
   },
   {
    "num": 2,
    "difficulty": "🔵 標準",
    "desc": "把同一題改成 VLOOKUP，指出第 4 個參數為什麼要填 0 或 FALSE"
   },
   {
    "num": 3,
    "difficulty": "🔵 標準",
    "desc": "已知「機械鍵盤」時，用 INDEX + MATCH 反查產品代碼"
   },
   {
    "num": 4,
    "difficulty": "🟡 變化",
    "desc": "查 P999 時不要出現 #N/A，改顯示「無此商品」"
   }
  ]
 },
 "pro": {
  "title": "🔍 查找比對  ─  XLOOKUP / VLOOKUP / INDEX+MATCH",
  "subtitle": "難度：🔵 Lv.2  |  微任務數：10 題  |  建議時間：每題 2~3 分鐘",
  "dataHeader": [
   "產品代碼",
   "產品名稱",
   "類別",
   "單價",
   "庫存",
   "",
   ""
  ],
  "dataRows": [
   [
    "P001",
    "無線滑鼠",
    "周邊",
    590,
    120,
    "",
    ""
   ],
   [
    "P002",
    "機械鍵盤",
    "周邊",
    2200,
    45,
    "",
    ""
   ],
   [
    "P003",
    "27吋螢幕",
    "顯示",
    8900,
    18,
    "",
    ""
   ],
   [
    "P004",
    "USB集線器",
    "周邊",
    380,
    200,
    "",
    ""
   ],
   [
    "P005",
    "網路攝影機",
    "周邊",
    1500,
    63,
    "",
    ""
   ],
   [
    "P006",
    "耳機麥克風",
    "音訊",
    750,
    85,
    "",
    ""
   ]
  ],
  "tasks": [
   {
    "numLabel": "🟢1",
    "desc": "用 XLOOKUP 查 P003 的產品名稱",
    "time": "2m",
    "hint": "=XLOOKUP(\"P003\",A6:A11,B6:B11,\"查無此商品\")",
    "answer": "=XLOOKUP(\"P003\",A6:A11,B6:B11,\"查無此商品\")",
    "explain": "XLOOKUP(查找值,查找欄,回傳欄,找不到時)"
   },
   {
    "numLabel": "🟢2",
    "desc": "用 XLOOKUP 查 P005 的單價",
    "time": "2m",
    "hint": "=XLOOKUP(\"P005\",A6:A11,D6:D11,\"查無此商品\")",
    "answer": "=XLOOKUP(\"P005\",A6:A11,D6:D11,\"查無此商品\")",
    "explain": "回傳欄改成 D 欄即可，不用重新數欄位"
   },
   {
    "numLabel": "🔵3",
    "desc": "用 VLOOKUP 查 P003 的產品名稱（舊檔寫法）",
    "time": "3m",
    "hint": "=VLOOKUP(\"P003\",A6:E11,2,0)",
    "answer": "=VLOOKUP(\"P003\",A6:E11,2,0)",
    "explain": "VLOOKUP 的查找欄必須在表格範圍最左邊，最後一個參數填 0 精確比對"
   },
   {
    "numLabel": "🔵4",
    "desc": "用 VLOOKUP 查 P005 的單價，並說出第 4 欄代表什麼",
    "time": "2m",
    "hint": "=VLOOKUP(\"P005\",A6:E11,4,0)",
    "answer": "=VLOOKUP(\"P005\",A6:E11,4,0)",
    "explain": "第 4 欄是單價欄；這也是 VLOOKUP 容易因插欄而壞掉的地方"
   },
   {
    "numLabel": "🔵5",
    "desc": "用 INDEX+MATCH 查 P003 名稱",
    "time": "3m",
    "hint": "=INDEX(B6:B11,MATCH(\"P003\",A6:A11,0))",
    "answer": "=INDEX(B6:B11,MATCH(\"P003\",A6:A11,0))",
    "explain": "MATCH 找位置，INDEX 取同位置的結果"
   },
   {
    "numLabel": "🟡6",
    "desc": "反向查找：已知名稱「機械鍵盤」查代碼",
    "time": "3m",
    "hint": "=INDEX(A6:A11,MATCH(\"機械鍵盤\",B6:B11,0))",
    "answer": "=INDEX(A6:A11,MATCH(\"機械鍵盤\",B6:B11,0))",
    "explain": "回傳欄 A 在查找欄 B 左邊，VLOOKUP 做不到，INDEX+MATCH 可以"
   },
   {
    "numLabel": "🟡7",
    "desc": "用 XLOOKUP 查不到 P999 時顯示「無此商品」",
    "time": "2m",
    "hint": "=XLOOKUP(\"P999\",A6:A11,B6:B11,\"無此商品\")",
    "answer": "=XLOOKUP(\"P999\",A6:A11,B6:B11,\"無此商品\")",
    "explain": "XLOOKUP 第 4 參數就是找不到時要顯示的內容"
   },
   {
    "numLabel": "🟡8",
    "desc": "舊檔 VLOOKUP 查不到 P999 時顯示「無此商品」",
    "time": "2m",
    "hint": "=IFERROR(VLOOKUP(\"P999\",A6:E11,2,0),\"無此商品\")",
    "answer": "=IFERROR(VLOOKUP(\"P999\",A6:E11,2,0),\"無此商品\")",
    "explain": "IFERROR 適合包住舊檔裡可能出錯的查找公式"
   },
   {
    "numLabel": "🔴9",
    "desc": "用 MATCH 找出「27吋螢幕」在第幾列",
    "time": "2m",
    "hint": "=MATCH(\"27吋螢幕\",B6:B11,0)",
    "answer": "=MATCH(\"27吋螢幕\",B6:B11,0)",
    "explain": "MATCH 回傳的是「相對位置」而非列號"
   },
   {
    "numLabel": "🔴10",
    "desc": "VLOOKUP 限制：為什麼不能用來查「已知名稱找代碼」？",
    "time": "2m",
    "hint": "查找欄必須在最左邊",
    "answer": "查找欄必須在最左邊",
    "explain": "VLOOKUP 只能往右查，這就是 INDEX+MATCH 存在的意義"
   }
  ]
 },
 "meta": {
  "phase": "💼 Phase 2：工作即戰力    第 3~4 週  |  目標：老闆交辦的任務都能搞定",
  "stage": "第 5 階段",
  "topics": "XLOOKUP / VLOOKUP / INDEX+MATCH / IFERROR",
  "difficulty": "🔵 Lv.2",
  "taskCount": "10 題",
  "time": "25 分鐘",
  "xp": "200 XP"
 },
 "inter": {
  "title": "🔍 查找比對互動練習",
  "subtitle": "在黃色格子輸入公式 → 答對自動變綠 ✅ 答錯變紅 ❌",
  "dataHeader": [
   "產品代碼",
   "產品名稱",
   "類別",
   "單價",
   "庫存",
   "",
   ""
  ],
  "dataRows": [
   [
    "P001",
    "無線滑鼠",
    "周邊",
    590,
    120,
    "",
    ""
   ],
   [
    "P002",
    "機械鍵盤",
    "周邊",
    2200,
    45,
    "",
    ""
   ],
   [
    "P003",
    "27吋螢幕",
    "顯示",
    8900,
    18,
    "",
    ""
   ],
   [
    "P004",
    "網路攝影機",
    "周邊",
    1500,
    67,
    "",
    ""
   ],
   [
    "P005",
    "USB集線器",
    "配件",
    450,
    200,
    "",
    ""
   ],
   [
    "P006",
    "外接硬碟",
    "儲存",
    2800,
    32,
    "",
    ""
   ]
  ],
  "tasks": [
   {
    "num": 1,
    "difficulty": "🟢",
    "desc": "XLOOKUP 查 P003 的產品名稱",
    "hint": "XLOOKUP(查找值,查找欄,回傳欄)",
    "answer": "=XLOOKUP(\"P003\",A6:A11,B6:B11)"
   },
   {
    "num": 2,
    "difficulty": "🟢",
    "desc": "XLOOKUP 查 P005 的單價",
    "hint": "回傳欄改成單價欄",
    "answer": "=XLOOKUP(\"P005\",A6:A11,D6:D11)"
   },
   {
    "num": 3,
    "difficulty": "🔵",
    "desc": "VLOOKUP 查 P003 的產品名稱",
    "hint": "VLOOKUP(查找值,表格,第幾欄,0)",
    "answer": "=VLOOKUP(\"P003\",A6:E11,2,0)"
   },
   {
    "num": 4,
    "difficulty": "🔵",
    "desc": "INDEX+MATCH 查 P003 名稱",
    "hint": "MATCH 找位置，INDEX 取值",
    "answer": "=INDEX(B6:B11,MATCH(\"P003\",A6:A11,0))"
   },
   {
    "num": 5,
    "difficulty": "🟡",
    "desc": "反向查找：已知「機械鍵盤」查代碼",
    "hint": "INDEX+MATCH 可以往左查",
    "answer": "=INDEX(A6:A11,MATCH(\"機械鍵盤\",B6:B11,0))"
   },
   {
    "num": 6,
    "difficulty": "🟡",
    "desc": "XLOOKUP 查不到 P999 時顯示「無此商品」",
    "hint": "XLOOKUP 第 4 參數=找不到時",
    "answer": "=XLOOKUP(\"P999\",A6:A11,B6:B11,\"無此商品\")"
   },
   {
    "num": 7,
    "difficulty": "🔴",
    "desc": "MATCH：「27吋螢幕」在第幾列",
    "hint": "MATCH 回傳相對位置",
    "answer": "=MATCH(\"27吋螢幕\",B6:B11,0)"
   },
   {
    "num": 8,
    "difficulty": "🔴",
    "desc": "舊檔 VLOOKUP 查不到 P999 時顯示「無此商品」",
    "hint": "IFERROR 包住公式",
    "answer": "=IFERROR(VLOOKUP(\"P999\",A6:E11,2,0),\"無此商品\")"
   }
  ]
 }
};
export const DEMOS = [
 {
  "id": "xlookup-direct-return",
  "kind": "formula",
  "badge": "公式拆解",
  "title": "XLOOKUP：查找欄和回傳欄分開寫",
  "subtitle": "新版 Excel 的主路線。你不用手算第幾欄，只要明確說出去哪裡找、回傳哪裡。",
  "outcome": "結果：P003 在產品代碼欄命中，所以回傳同列的 27吋螢幕",
  "formulaParts": [
   {
    "text": "=XLOOKUP("
   },
   {
    "text": "\"P003\"",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ", "
   },
   {
    "text": "A2:A6",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": ", "
   },
   {
    "text": "B2:B6",
    "step": 3,
    "tone": "c3"
   },
   {
    "text": ", "
   },
   {
    "text": "\"查無此商品\"",
    "step": 4,
    "tone": "c4"
   },
   {
    "text": ")"
   }
  ],
  "board": {
   "title": "商品表",
   "columns": [
    "產品代碼",
    "產品名稱",
    "單價"
   ],
   "rows": [
    [
     {
      "text": "P001",
      "marks": {
       "2": "focus"
      }
     },
     {
      "text": "無線滑鼠",
      "marks": {
       "3": "focus"
      }
     },
     {
      "text": "590"
     }
    ],
    [
     {
      "text": "P002",
      "marks": {
       "2": "focus"
      }
     },
     {
      "text": "機械鍵盤",
      "marks": {
       "3": "focus"
      }
     },
     {
      "text": "2,200"
     }
    ],
    [
     {
      "text": "P003",
      "marks": {
       "1": "focus",
       "2": "match"
      }
     },
     {
      "text": "27吋螢幕",
      "marks": {
       "3": "result"
      }
     },
     {
      "text": "8,900"
     }
    ],
    [
     {
      "text": "P004",
      "marks": {
       "2": "focus"
      }
     },
     {
      "text": "USB集線器",
      "marks": {
       "3": "focus"
      }
     },
     {
      "text": "380"
     }
    ],
    [
     {
      "text": "P005",
      "marks": {
       "2": "focus"
      }
     },
     {
      "text": "網路攝影機",
      "marks": {
       "3": "focus"
      }
     },
     {
      "text": "1,500"
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. 查找值：你手上拿什麼去找",
    "tone": "c1",
    "text": "這裡拿 P003 去查商品表。查找值可以直接寫在公式裡，也可以引用訂單表上的產品代碼儲存格。",
    "caption": "先確認查找值乾淨一致，像是 P003 不要混到多餘空白。"
   },
   {
    "label": "2. 查找範圍：去哪一欄找",
    "tone": "c2",
    "text": "A2:A6 是產品代碼欄。XLOOKUP 會在這一欄尋找 P003。",
    "caption": "查找範圍只需要放查找欄，不必把整張表都圈進去。"
   },
   {
    "label": "3. 回傳範圍：命中後拿哪一欄回來",
    "tone": "c3",
    "text": "B2:B6 是產品名稱欄。P003 命中第 3 列，所以回傳同一列的 27吋螢幕。",
    "caption": "這就是 XLOOKUP 比 VLOOKUP 清楚的地方：回傳欄直接指定，不用數第幾欄。"
   },
   {
    "label": "4. 找不到時的結果",
    "tone": "c4",
    "text": "如果找不到 P003，公式會顯示「查無此商品」，而不是讓 #N/A 直接出現在報表上。",
    "caption": "正式交付報表時，找不到資料也要有可讀訊息。"
   }
  ]
 },
 {
  "id": "index-match-reverse",
  "kind": "formula",
  "badge": "反向查找",
  "title": "INDEX + MATCH：舊版也能往左查",
  "subtitle": "MATCH 先找位置，INDEX 再依位置取值。這個拆法讓你不受 VLOOKUP 只能往右查的限制。",
  "outcome": "結果：機械鍵盤在名稱欄第 2 筆，所以 INDEX 從代碼欄取回 P002",
  "formulaParts": [
   {
    "text": "=INDEX("
   },
   {
    "text": "A2:A6",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ", MATCH("
   },
   {
    "text": "\"機械鍵盤\"",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": ", "
   },
   {
    "text": "B2:B6",
    "step": 3,
    "tone": "c3"
   },
   {
    "text": ", "
   },
   {
    "text": "0",
    "step": 4,
    "tone": "c4"
   },
   {
    "text": "))"
   }
  ],
  "board": {
   "title": "已知產品名稱，反查產品代碼",
   "columns": [
    "產品代碼",
    "產品名稱"
   ],
   "rows": [
    [
     {
      "text": "P001",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "無線滑鼠",
      "marks": {
       "3": "focus"
      }
     }
    ],
    [
     {
      "text": "P002",
      "marks": {
       "1": "focus",
       "4": "result"
      }
     },
     {
      "text": "機械鍵盤",
      "marks": {
       "2": "focus",
       "3": "match"
      }
     }
    ],
    [
     {
      "text": "P003",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "27吋螢幕",
      "marks": {
       "3": "focus"
      }
     }
    ],
    [
     {
      "text": "P004",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "USB集線器",
      "marks": {
       "3": "focus"
      }
     }
    ],
    [
     {
      "text": "P005",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "網路攝影機",
      "marks": {
       "3": "focus"
      }
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. INDEX 先指定要拿結果的欄",
    "tone": "c1",
    "text": "A2:A6 是產品代碼欄，也就是最後要回傳結果的範圍。",
    "caption": "和 VLOOKUP 不同，回傳欄可以在查找欄左邊。"
   },
   {
    "label": "2. MATCH 指定你要找的值",
    "tone": "c2",
    "text": "這裡要在產品名稱欄尋找「機械鍵盤」。",
    "caption": "MATCH 的工作不是回傳名稱，而是找出它在範圍中的第幾個。"
   },
   {
    "label": "3. MATCH 指定去哪一欄找",
    "tone": "c3",
    "text": "B2:B6 是產品名稱欄。機械鍵盤在這個範圍裡的第 2 個位置。",
    "caption": "MATCH 回傳的是相對位置，不是 Excel 的實際列號。"
   },
   {
    "label": "4. 0 代表精確比對，再交給 INDEX 取值",
    "tone": "c4",
    "text": "MATCH 回傳 2，INDEX 就從 A2:A6 的第 2 個位置取回 P002。",
    "caption": "這就是 INDEX + MATCH 的核心：先定位，再取值。"
   }
  ]
 }
];
