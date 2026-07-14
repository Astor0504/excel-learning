// 由 xlsx-content.js + lesson-demos-data.js 拆分(2026-07-14);本檔是 P1-02 的唯一資料來源
export const CONTENT = {
 "knowledge": {
  "title": "📚 基礎函式不是背語法，是先分清楚問題類型",
  "subtitle": "難度：🟢 Lv.1  |  先分清楚總和 / 平均 / 筆數 / 名次  |  這課穩了，後面所有函式都會順很多",
  "sections": [
   {
    "title": "📌 先問自己在找什麼答案",
    "items": [
     [
      "總和",
      "SUM",
      "把範圍內的數字全部加起來",
      "看到合計、小計、總業績時先想到它"
     ],
     [
      "平均",
      "AVERAGE",
      "空白與文字忽略，但 0 會算進去",
      "平均值怪怪的時，先檢查資料裡是不是有 0"
     ],
     [
      "筆數",
      "COUNT / COUNTA",
      "COUNT 數數字；COUNTA 數所有非空白",
      "報表最常錯在這裡，不是公式難，是問題問錯"
     ],
     [
      "名次",
      "MAX / MIN / LARGE / SMALL",
      "MAX / MIN 看最值；LARGE / SMALL 看第 N 名",
      "主管問前 3 名時，不要只會用 MAX"
     ],
     [
      "典型水準",
      "MEDIAN",
      "排序後取中間值，不容易被極端值帶歪",
      "薪資、房價、獎金常比 AVERAGE 更值得一起看"
     ]
    ]
   },
   {
    "title": "⚠️ 3 個最容易出錯的差別",
    "items": [
     [
      "COUNT vs COUNTA",
      "COUNT(B:B) 只數數字",
      "COUNTA(A:A) 只要不是空白就會算進去，包含文字、日期、0",
      "要算『有幾位員工有填資料』時，通常該用 COUNTA"
     ],
     [
      "AVERAGE 對 0 / 空白 / 文字",
      "0 = 會算進平均",
      "空白 / 文字 = 忽略",
      "所以 0 不是空值，真的會把平均拉低"
     ],
     [
      "MAX / MIN vs LARGE / SMALL",
      "MAX / MIN = 冠軍與最後一名",
      "LARGE / SMALL = 第 2 名、第 3 名…",
      "問名次就不要只用 MAX / MIN"
     ]
    ]
   },
   {
    "title": "🎯 這課真正要練熟的思路",
    "items": [
     [
      "先問問題，再寫公式",
      "你要的是總和、平均、筆數，還是名次？",
      "問題問對了，公式通常就只剩 2~3 個候選"
     ],
     [
      "先看資料內容，再信任結果",
      "平均值不合理時，先找 0、空白、文字",
      "筆數不合理時，先看你用的是 COUNT 還是 COUNTA"
     ],
     [
      "同時看最值與名次",
      "MAX 告訴你冠軍是多少",
      "LARGE 則能告訴你第 2 名、第 3 名差多少"
     ]
    ]
   }
  ],
  "handsTasks": [
   {
    "num": 1,
    "difficulty": "🟢 暖身",
    "desc": "在本課資料表先各寫一次 =COUNT(B6:B10) 與 =COUNTA(A6:A10)，說出兩者為什麼答案不同"
   },
   {
    "num": 2,
    "difficulty": "🟢 暖身",
    "desc": "在旁邊空白欄手動做一組資料：120000、0、空白、待補。寫 =AVERAGE(...) 算一次，再把 0 刪成空白重算一次"
   },
   {
    "num": 3,
    "difficulty": "🔵 標準",
    "desc": "先用 =MAX(B6:B10) 找最高值，再改成 =LARGE(B6:B10,2) 找第 2 名，感受『最值』和『名次』差別"
   },
   {
    "num": 4,
    "difficulty": "🟡 變化",
    "desc": "把 1 月業績用 =AVERAGE(B6:B10) 與 =MEDIAN(B6:B10) 各算一次，觀察兩個結果差多少，想想哪個更像『典型水準』"
   }
  ]
 },
 "pro": {
  "title": "📊 基礎函式  ─  統計、筆數、排名基本功",
  "subtitle": "難度：🟢 Lv.1  |  微任務數：8 題  |  建議時間：每題 2~3 分鐘",
  "dataHeader": [
   "姓名",
   "1月",
   "2月",
   "3月",
   "4月",
   "5月",
   "6月"
  ],
  "dataRows": [
   [
    "王小明",
    42000,
    38000,
    51000,
    47000,
    55000,
    60000
   ],
   [
    "林美華",
    35000,
    41000,
    39000,
    43000,
    40000,
    38000
   ],
   [
    "張志偉",
    58000,
    62000,
    57000,
    65000,
    70000,
    68000
   ],
   [
    "陳雅婷",
    29000,
    33000,
    31000,
    28000,
    35000,
    37000
   ],
   [
    "李建宏",
    47000,
    45000,
    50000,
    52000,
    49000,
    53000
   ]
  ],
  "tasks": [
   {
    "numLabel": "🟢1",
    "desc": "計算所有人 1 月的業績總和",
    "time": "2m",
    "hint": "=SUM(B6:B10)",
    "answer": "=SUM(B6:B10)",
    "explain": "SUM 把範圍內所有數字加起來"
   },
   {
    "numLabel": "🟢2",
    "desc": "計算 1 月的平均業績",
    "time": "2m",
    "hint": "=AVERAGE(B6:B10)",
    "answer": "=AVERAGE(B6:B10)",
    "explain": "AVERAGE = 總和 ÷ 個數"
   },
   {
    "numLabel": "🔵3",
    "desc": "找出 1 月最高業績",
    "time": "1m",
    "hint": "=MAX(B6:B10)",
    "answer": "=MAX(B6:B10)",
    "explain": "MAX 回傳範圍中的最大值"
   },
   {
    "numLabel": "🔵4",
    "desc": "找出 1 月最低業績",
    "time": "1m",
    "hint": "=MIN(B6:B10)",
    "answer": "=MIN(B6:B10)",
    "explain": "MIN 回傳範圍中的最小值"
   },
   {
    "numLabel": "🟡5",
    "desc": "計算有幾位業務員",
    "time": "2m",
    "hint": "=COUNTA(A6:A10)",
    "answer": "=COUNTA(A6:A10)",
    "explain": "COUNTA 計算非空白的儲存格數"
   },
   {
    "numLabel": "🟡6",
    "desc": "用 COUNT 計算 1 月業績欄有幾個數字",
    "time": "2m",
    "hint": "=COUNT(B6:B10)",
    "answer": "=COUNT(B6:B10)",
    "explain": "COUNT 只數數字；如果欄位裡混進文字或空白，它不會算進去"
   },
   {
    "numLabel": "🔴7",
    "desc": "用 LARGE 找出 1 月第 2 高的業績",
    "time": "2m",
    "hint": "=LARGE(B6:B10,2)",
    "answer": "=LARGE(B6:B10,2)",
    "explain": "第 N 大要用 LARGE；MAX 只會給你第 1 名"
   },
   {
    "numLabel": "🔴8",
    "desc": "計算 1 月業績的中位數",
    "time": "2m",
    "hint": "=MEDIAN(B6:B10)",
    "answer": "=MEDIAN(B6:B10)",
    "explain": "MEDIAN 排序後取中間值，不受極端值影響"
   }
  ]
 },
 "meta": {
  "phase": "🏃 Phase 1：操作效率基礎    第 1~2 週  |  目標：養成鍵盤操作習慣，複習基本功",
  "stage": "第 2 階段",
  "topics": "SUM / AVERAGE / COUNT / COUNTA / MAX / MIN / MEDIAN / LARGE / SMALL",
  "difficulty": "🟢 Lv.1",
  "taskCount": "8 題",
  "time": "15 分鐘",
  "xp": "100 XP"
 },
 "inter": {
  "title": "📊 基礎函式互動練習",
  "subtitle": "在黃色格子輸入公式 → 答對自動變綠 ✅ 答錯變紅 ❌ ── 答案完全隱藏！",
  "dataHeader": [
   "姓名",
   "1月",
   "2月",
   "3月",
   "4月",
   "5月",
   "6月"
  ],
  "dataRows": [
   [
    "王小明",
    42000,
    38000,
    51000,
    47000,
    55000,
    60000
   ],
   [
    "林美華",
    35000,
    41000,
    39000,
    43000,
    40000,
    38000
   ],
   [
    "張志偉",
    58000,
    62000,
    57000,
    65000,
    70000,
    68000
   ],
   [
    "陳雅琪",
    29000,
    33000,
    31000,
    36000,
    34000,
    37000
   ],
   [
    "李大明",
    45000,
    48000,
    52000,
    50000,
    46000,
    51000
   ]
  ],
  "tasks": [
   {
    "num": 1,
    "difficulty": "🟢",
    "desc": "計算所有人 1 月業績總和",
    "hint": "SUM(範圍)",
    "answer": "=SUM(B6:B10)"
   },
   {
    "num": 2,
    "difficulty": "🟢",
    "desc": "計算 1 月的平均業績",
    "hint": "AVERAGE(範圍)",
    "answer": "=AVERAGE(B6:B10)"
   },
   {
    "num": 3,
    "difficulty": "🔵",
    "desc": "找出 1 月最高業績",
    "hint": "MAX(範圍)",
    "answer": "=MAX(B6:B10)"
   },
   {
    "num": 4,
    "difficulty": "🔵",
    "desc": "找出 1 月最低業績",
    "hint": "MIN(範圍)",
    "answer": "=MIN(B6:B10)"
   },
   {
    "num": 5,
    "difficulty": "🔵",
    "desc": "計算有幾位業務員（非空白格數）",
    "hint": "COUNTA 數非空白",
    "answer": "=COUNTA(A6:A10)"
   },
   {
    "num": 6,
    "difficulty": "🟡",
    "desc": "用 COUNT 計算 1 月欄位裡有幾個數字",
    "hint": "COUNT 只數數字",
    "answer": "=COUNT(B6:B10)"
   },
   {
    "num": 7,
    "difficulty": "🟡",
    "desc": "計算 1 月業績中位數",
    "hint": "MEDIAN(範圍)",
    "answer": "=MEDIAN(B6:B10)"
   },
   {
    "num": 8,
    "difficulty": "🔴",
    "desc": "用 LARGE 找出 1 月第 2 高的業績",
    "hint": "LARGE(範圍,第幾大)",
    "answer": "=LARGE(B6:B10,2)"
   }
  ]
 }
};
export const DEMOS = [
 {
  "id": "average-zero-blank",
  "kind": "formula",
  "badge": "概念拆解",
  "title": "AVERAGE：0 會算進去，空白和文字會略過",
  "subtitle": "這是本課最容易誤會的地方。平均值看起來怪怪的，先不要怪公式，先檢查資料內容。",
  "outcome": "結果：(120,000 + 0 + 95,000) ÷ 3 = 71,667",
  "formulaParts": [
   {
    "text": "=AVERAGE("
   },
   {
    "text": "B2:B6",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ")"
   }
  ],
  "board": {
   "title": "同一欄裡混了數字、0、空白、文字",
   "columns": [
    "業績欄內容",
    "COUNT",
    "COUNTA",
    "AVERAGE"
   ],
   "rows": [
    [
     {
      "text": "120,000",
      "marks": {
       "1": "focus",
       "2": "match",
       "3": "match"
      }
     },
     {
      "text": "算",
      "marks": {
       "2": "result"
      }
     },
     {
      "text": "算",
      "marks": {
       "2": "result"
      }
     },
     {
      "text": "算",
      "marks": {
       "3": "result"
      }
     }
    ],
    [
     {
      "text": "0",
      "marks": {
       "1": "focus",
       "2": "match",
       "3": "match"
      }
     },
     {
      "text": "算",
      "marks": {
       "2": "result"
      }
     },
     {
      "text": "算",
      "marks": {
       "2": "result"
      }
     },
     {
      "text": "算",
      "marks": {
       "3": "result"
      }
     }
    ],
    [
     {
      "text": "（空白）",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "不算",
      "marks": {
       "2": "result"
      }
     },
     {
      "text": "不算",
      "marks": {
       "2": "result"
      }
     },
     {
      "text": "不算",
      "marks": {
       "3": "result"
      }
     }
    ],
    [
     {
      "text": "95,000",
      "marks": {
       "1": "focus",
       "2": "match",
       "3": "match"
      }
     },
     {
      "text": "算",
      "marks": {
       "2": "result"
      }
     },
     {
      "text": "算",
      "marks": {
       "2": "result"
      }
     },
     {
      "text": "算",
      "marks": {
       "3": "result"
      }
     }
    ],
    [
     {
      "text": "待補",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "不算",
      "marks": {
       "2": "result"
      }
     },
     {
      "text": "算",
      "marks": {
       "2": "result"
      }
     },
     {
      "text": "不算",
      "marks": {
       "3": "result"
      }
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. 先看整段範圍",
    "tone": "c1",
    "text": "AVERAGE 第一件事只是先圈出要檢查的範圍。Excel 會在這整段裡，自己判斷哪些值要參與平均。",
    "caption": "先記住：範圍裡的每格不一定都會真的被算進平均。"
   },
   {
    "label": "2. COUNT / COUNTA 幫你辨認資料型態",
    "tone": "c2",
    "text": "COUNT 只數數字；COUNTA 數所有非空白。這一步最重要的觀念是：0 是數字，文字「待補」不是，空白更不是。",
    "caption": "很多報表其實不是公式錯，而是你以為 0 跟空白是一樣的。"
   },
   {
    "label": "3. AVERAGE 只拿數字做平均，但 0 不會被排除",
    "tone": "c3",
    "text": "AVERAGE 會忽略空白和文字，但 0 會當成真正的數字算進去。所以這組資料是用 120,000、0、95,000 三個值去平均。",
    "caption": "只要格子裡是 0，它就會把平均拉低；想排除 0，要改用 AVERAGEIF。"
   }
  ]
 },
 {
  "id": "large-ranking",
  "kind": "formula",
  "badge": "公式拆解",
  "title": "LARGE：當你要的是名次，不是單純最大值",
  "subtitle": "MAX 只會給你冠軍；要第 2 名、第 3 名，才輪到 LARGE 上場。",
  "outcome": "結果：150,000 > 120,000 > 95,000，所以第 3 大是 95,000",
  "formulaParts": [
   {
    "text": "=LARGE("
   },
   {
    "text": "B2:B6",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ", "
   },
   {
    "text": "3",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": ")"
   }
  ],
  "board": {
   "title": "業績表",
   "columns": [
    "業務員",
    "業績"
   ],
   "rows": [
    [
     {
      "text": "王小明"
     },
     {
      "text": "120,000",
      "marks": {
       "1": "focus",
       "3": "match"
      }
     }
    ],
    [
     {
      "text": "林美玲"
     },
     {
      "text": "85,000",
      "marks": {
       "1": "focus"
      }
     }
    ],
    [
     {
      "text": "張志偉"
     },
     {
      "text": "95,000",
      "marks": {
       "1": "focus",
       "2": "result",
       "3": "result"
      }
     }
    ],
    [
     {
      "text": "陳雅婷"
     },
     {
      "text": "150,000",
      "marks": {
       "1": "focus",
       "3": "match"
      }
     }
    ],
    [
     {
      "text": "李大華"
     },
     {
      "text": "72,000",
      "marks": {
       "1": "focus"
      }
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. 範圍：先決定在哪一群數字裡排名",
    "tone": "c1",
    "text": "B2:B6 是參與排名的資料範圍。LARGE 只會在這群數字裡找名次，不會跨到別欄去想。",
    "caption": "非數字內容會被忽略，但你的範圍還是要抓對。"
   },
   {
    "label": "2. 名次：3 代表第 3 大，不是第 3 筆",
    "tone": "c2",
    "text": "第二個參數的意思是『第幾名』。填 1 = 最大值，填 2 = 第 2 大，填 3 = 第 3 大。",
    "caption": "問冠軍時用 MAX 就夠了；問前 3 名時才改用 LARGE。"
   },
   {
    "label": "3. 讀結果：第 3 大是排序後的第 3 個值",
    "tone": "c3",
    "text": "這組資料排序後依序是 150,000、120,000、95,000、85,000、72,000，所以公式回傳 95,000。",
    "caption": "如果你要看倒數第 2 名、倒數第 3 名，把 LARGE 換成 SMALL 即可。"
   }
  ]
 }
];
