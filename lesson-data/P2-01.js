// 由 xlsx-content.js + lesson-demos-data.js 拆分(2026-07-14);本檔是 P2-01 的唯一資料來源
export const CONTENT = {
 "pro": {
  "title": "📈 條件統計  ─  COUNTIFS / SUMIFS / AVERAGEIFS",
  "subtitle": "難度：🔵 Lv.2  |  微任務數：8 題  |  建議時間：每題 2~3 分鐘",
  "dataHeader": [
   "訂單日期",
   "業務員",
   "區域",
   "產品類別",
   "金額",
   "付款狀態",
   ""
  ],
  "dataRows": [
   [
    "2026-01-05",
    "王小明",
    "北區",
    "電子",
    15000,
    "已付",
    ""
   ],
   [
    "2026-01-12",
    "林美華",
    "南區",
    "服飾",
    8500,
    "未付",
    ""
   ],
   [
    "2026-01-18",
    "張志偉",
    "北區",
    "電子",
    22000,
    "已付",
    ""
   ],
   [
    "2026-02-03",
    "王小明",
    "北區",
    "服飾",
    12000,
    "已付",
    ""
   ],
   [
    "2026-02-10",
    "陳雅婷",
    "東區",
    "電子",
    9800,
    "未付",
    ""
   ],
   [
    "2026-02-15",
    "林美華",
    "南區",
    "電子",
    31000,
    "已付",
    ""
   ],
   [
    "2026-03-01",
    "張志偉",
    "北區",
    "電子",
    18500,
    "已付",
    ""
   ],
   [
    "2026-03-08",
    "陳雅婷",
    "東區",
    "服飾",
    7200,
    "已付",
    ""
   ],
   [
    "2026-03-20",
    "王小明",
    "北區",
    "電子",
    26000,
    "未付",
    ""
   ],
   [
    "2026-03-25",
    "林美華",
    "南區",
    "服飾",
    14000,
    "已付",
    ""
   ]
  ],
  "tasks": [
   {
    "numLabel": "🟢1",
    "desc": "北區共有幾筆訂單？",
    "time": "2m",
    "hint": "=COUNTIF(C6:C15,\"北區\")",
    "answer": "=COUNTIF(C6:C15,\"北區\")",
    "explain": "COUNTIF(範圍,條件) 計算符合條件的個數"
   },
   {
    "numLabel": "🟢2",
    "desc": "電子類別的訂單金額合計？",
    "time": "2m",
    "hint": "=SUMIF(D6:D15,\"電子\",E6:E15)",
    "answer": "=SUMIF(D6:D15,\"電子\",E6:E15)",
    "explain": "SUMIF(條件範圍,條件,加總範圍)"
   },
   {
    "numLabel": "🔵3",
    "desc": "電子 且 已付款的金額合計？",
    "time": "3m",
    "hint": "=SUMIFS(E6:E15,D6:D15,\"電子\",F6:F15,\"已付\")",
    "answer": "=SUMIFS(E6:E15,D6:D15,\"電子\",F6:F15,\"已付\")",
    "explain": "SUMIFS 多條件加總，注意加總欄放第一個"
   },
   {
    "numLabel": "🔵4",
    "desc": "北區電子類的平均金額？",
    "time": "3m",
    "hint": "=AVERAGEIFS(E6:E15,C6:C15,\"北區\",D6:D15,\"電子\")",
    "answer": "=AVERAGEIFS(E6:E15,C6:C15,\"北區\",D6:D15,\"電子\")",
    "explain": "AVERAGEIFS 語法同 SUMIFS"
   },
   {
    "numLabel": "🟡5",
    "desc": "王小明共有幾筆訂單？",
    "time": "2m",
    "hint": "=COUNTIF(B6:B15,\"王小明\")",
    "answer": "=COUNTIF(B6:B15,\"王小明\")",
    "explain": "COUNTIF 也可以用來計算某人的紀錄數"
   },
   {
    "numLabel": "🟡6",
    "desc": "金額>=20000 的訂單有幾筆？",
    "time": "2m",
    "hint": "=COUNTIF(E6:E15,\">=20000\")",
    "answer": "=COUNTIF(E6:E15,\">=20000\")",
    "explain": "條件可以用 >=、<=、<> 等運算子"
   },
   {
    "numLabel": "🔴7",
    "desc": "未付款的金額合計？",
    "time": "2m",
    "hint": "=SUMIF(F6:F15,\"未付\",E6:E15)",
    "answer": "=SUMIF(F6:F15,\"未付\",E6:E15)",
    "explain": "SUMIF 條件用文字也行"
   },
   {
    "numLabel": "🔴8",
    "desc": "北區且金額>=15000的訂單數？",
    "time": "3m",
    "hint": "=COUNTIFS(C6:C15,\"北區\",E6:E15,\">=15000\")",
    "answer": "=COUNTIFS(C6:C15,\"北區\",E6:E15,\">=15000\")",
    "explain": "COUNTIFS 多條件計數"
   }
  ]
 },
 "meta": {
  "phase": "💼 Phase 2：工作即戰力    第 3~4 週  |  目標：老闆交辦的任務都能搞定",
  "stage": "第 4 階段",
  "topics": "COUNTIFS / SUMIFS / AVERAGEIFS",
  "difficulty": "🔵 Lv.2",
  "taskCount": "8 題",
  "time": "20 分鐘",
  "xp": "150 XP"
 },
 "knowledge": {
  "title": "📚 條件統計選公式與驗算心法",
  "subtitle": "把題目翻成四件事：算什麼、條件幾個、範圍順序、哪些列被算進去。",
  "sections": [
   {
    "title": "選公式前先問",
    "items": [
     [
      "算什麼",
      "幾筆用 COUNT；合計用 SUM；平均用 AVERAGE。先決定家族，再看條件數。"
     ],
     [
      "條件幾個",
      "一個條件用 IF 版本；兩個以上條件用 IFS 版本。IFS 多條件預設都是 AND。"
     ],
     [
      "順序怎麼寫",
      "COUNTIFS 是條件範圍成對；SUMIFS / AVERAGEIFS 先放要計算的範圍，再放條件。"
     ]
    ]
   },
   {
    "title": "正式報表檢查點",
    "items": [
     [
      "先驗算命中列",
      "交付前用篩選或肉眼抽查 2~3 筆，確認被算進去的列就是你要的列。"
     ],
     [
      "範圍大小一致",
      "SUMIFS / COUNTIFS / AVERAGEIFS 裡每個條件範圍都要同樣列數，計算範圍也要對齊。"
     ],
     [
      "文字條件乾淨",
      "多餘空白、全形半形、狀態名稱差異，常常比公式本身更容易造成錯誤。"
     ],
     [
      "日期用真正日期",
      "跨 Excel / WPS 時，日期條件建議用 DATE() 或引用日期儲存格，避免文字日期判斷不一致。"
     ],
     [
      "OR 要另外處理",
      "多條件預設 AND。如果需求是 OR，通常要把多個 COUNTIFS 或 SUMIFS 相加。"
     ]
    ]
   }
  ]
 },
 "inter": {
  "title": "📈 條件統計練習 — 從單條件到多條件",
  "subtitle": "先做 COUNTIF / SUMIF / AVERAGEIF，再進到 COUNTIFS / AVERAGEIFS；每題都先判斷算什麼與條件幾個。",
  "dataHeader": [
   "員工",
   "部門",
   "月薪",
   "績效分數",
   "年資",
   "",
   ""
  ],
  "dataRows": [
   [
    "王小明",
    "業務",
    55000,
    88,
    3,
    "",
    ""
   ],
   [
    "林美華",
    "財務",
    48000,
    72,
    5,
    "",
    ""
   ],
   [
    "張志偉",
    "業務",
    68000,
    95,
    7,
    "",
    ""
   ],
   [
    "陳雅琪",
    "資訊",
    42000,
    65,
    2,
    "",
    ""
   ],
   [
    "李大明",
    "業務",
    51000,
    78,
    4,
    "",
    ""
   ],
   [
    "吳建志",
    "財務",
    52000,
    83,
    6,
    "",
    ""
   ]
  ],
  "tasks": [
   {
    "num": 1,
    "difficulty": "🟢",
    "desc": "計算業務部有幾個人",
    "hint": "COUNTIF(範圍,\"業務\")",
    "answer": "=COUNTIF(B6:B13,\"業務\")"
   },
   {
    "num": 2,
    "difficulty": "🟢",
    "desc": "計算月薪超過 50000 的人數",
    "hint": "COUNTIF(範圍,\">50000\")",
    "answer": "=COUNTIF(C6:C13,\">\"&50000)"
   },
   {
    "num": 3,
    "difficulty": "🔵",
    "desc": "計算業務部的薪資總和",
    "hint": "SUMIF(條件範圍,\"條件\",加總範圍)",
    "answer": "=SUMIF(B6:B13,\"業務\",C6:C13)"
   },
   {
    "num": 4,
    "difficulty": "🔵",
    "desc": "計算資訊部的平均績效",
    "hint": "AVERAGEIF(條件範圍,\"條件\",平均範圍)",
    "answer": "=AVERAGEIF(B6:B13,\"資訊\",D6:D13)"
   },
   {
    "num": 5,
    "difficulty": "🟡",
    "desc": "計算績效 ≥ 80 的人數",
    "hint": "COUNTIF(範圍,\">=80\")",
    "answer": "=COUNTIF(D6:D13,\">=\"&80)"
   },
   {
    "num": 6,
    "difficulty": "🟡",
    "desc": "計算年資 ≥ 5 的員工薪資總和",
    "hint": "SUMIF(年資欄,\">=5\",薪資欄)",
    "answer": "=SUMIF(E6:E13,\">=\"&5,C6:C13)"
   },
   {
    "num": 7,
    "difficulty": "🔴",
    "desc": "計算業務部且績效≥80 的人數（COUNTIFS）",
    "hint": "COUNTIFS(欄1,\"條件1\",欄2,\"條件2\")",
    "answer": "=COUNTIFS(B6:B13,\"業務\",D6:D13,\">=\"&80)"
   },
   {
    "num": 8,
    "difficulty": "🔴",
    "desc": "計算財務部且年資≥5 的平均薪資（AVERAGEIFS）",
    "hint": "AVERAGEIFS(平均欄,條件欄1,\"條件1\",條件欄2,\"條件2\")",
    "answer": "=AVERAGEIFS(C6:C13,B6:B13,\"財務\",E6:E13,\">=\"&5)"
   }
  ]
 }
};
export const DEMOS = [
 {
  "id": "conditional-stats-decision-path",
  "kind": "formula",
  "role": "primary",
  "badge": "主線示範",
  "title": "條件統計決策：先選公式，再檢查命中列",
  "subtitle": "不用一次背六個函數。先判斷要算什麼、條件幾個，再把命中的列加總或計數。",
  "outcome": "結果：這題是多條件加總，所以用 SUMIFS；命中 4 筆已付款電子訂單，合計 86,500。",
  "formulaParts": [
   {
    "text": "問題：電子且已付的金額合計"
   },
   {
    "text": " → 算合計",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": " → 2 個條件",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": " → SUMIFS",
    "step": 3,
    "tone": "c3"
   },
   {
    "text": " → 驗算命中列",
    "step": 4,
    "tone": "c4"
   }
  ],
  "board": {
   "title": "從問題到命中列",
   "columns": [
    "類別",
    "付款",
    "金額",
    "判斷"
   ],
   "rows": [
    [
     {
      "text": "電子",
      "marks": {
       "2": "candidate",
       "3": "candidate"
      }
     },
     {
      "text": "已付",
      "marks": {
       "2": "candidate",
       "3": "match"
      }
     },
     {
      "text": "15,000",
      "marks": {
       "1": "focus",
       "4": "result"
      }
     },
     {
      "text": "命中",
      "marks": {
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "服飾",
      "marks": {
       "2": "exclude",
       "3": "exclude"
      }
     },
     {
      "text": "未付",
      "marks": {
       "2": "focus",
       "3": "exclude"
      }
     },
     {
      "text": "8,500",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "排除",
      "marks": {
       "4": "exclude"
      }
     }
    ],
    [
     {
      "text": "電子",
      "marks": {
       "2": "candidate",
       "3": "candidate"
      }
     },
     {
      "text": "已付",
      "marks": {
       "2": "candidate",
       "3": "match"
      }
     },
     {
      "text": "22,000",
      "marks": {
       "1": "focus",
       "4": "result"
      }
     },
     {
      "text": "命中",
      "marks": {
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "服飾",
      "marks": {
       "2": "exclude",
       "3": "exclude"
      }
     },
     {
      "text": "已付",
      "marks": {
       "2": "focus",
       "3": "match"
      }
     },
     {
      "text": "12,000",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "排除",
      "marks": {
       "4": "exclude"
      }
     }
    ],
    [
     {
      "text": "電子",
      "marks": {
       "2": "candidate",
       "3": "candidate"
      }
     },
     {
      "text": "未付",
      "marks": {
       "2": "focus",
       "3": "exclude"
      }
     },
     {
      "text": "9,800",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "排除",
      "marks": {
       "4": "exclude"
      }
     }
    ],
    [
     {
      "text": "電子",
      "marks": {
       "2": "candidate",
       "3": "candidate"
      }
     },
     {
      "text": "已付",
      "marks": {
       "2": "candidate",
       "3": "match"
      }
     },
     {
      "text": "31,000",
      "marks": {
       "1": "focus",
       "4": "result"
      }
     },
     {
      "text": "命中",
      "marks": {
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "電子",
      "marks": {
       "2": "candidate",
       "3": "candidate"
      }
     },
     {
      "text": "已付",
      "marks": {
       "2": "candidate",
       "3": "match"
      }
     },
     {
      "text": "18,500",
      "marks": {
       "1": "focus",
       "4": "result"
      }
     },
     {
      "text": "命中",
      "marks": {
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "服飾",
      "marks": {
       "2": "exclude",
       "3": "exclude"
      }
     },
     {
      "text": "已付",
      "marks": {
       "2": "focus",
       "3": "match"
      }
     },
     {
      "text": "7,200",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "排除",
      "marks": {
       "4": "exclude"
      }
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. 先問：要算什麼",
    "tone": "c1",
    "text": "題目問的是「金額合計」，所以不是 COUNT，也不是 AVERAGE，而是 SUM 家族。",
    "caption": "幾筆用 COUNT；合計用 SUM；平均用 AVERAGE。"
   },
   {
    "label": "2. 再問：條件幾個",
    "tone": "c2",
    "text": "這題同時要求「電子」和「已付」，所以是多條件。多條件版本要用結尾有 S 的函數。",
    "caption": "COUNTIFS / SUMIFS / AVERAGEIFS 的條件預設都是 AND。"
   },
   {
    "label": "3. 選公式並確認順序",
    "tone": "c3",
    "text": "多條件加總用 SUMIFS，而且加總範圍要放第一個，後面才是一組一組條件範圍與條件值。",
    "caption": "SUMIF 是條件範圍在前；SUMIFS 是加總範圍在前。"
   },
   {
    "label": "4. 驗算命中列",
    "tone": "c4",
    "text": "最後回到資料表，只把同時符合電子與已付的金額加起來。四筆命中列合計 86,500。",
    "caption": "正式報表交付前，至少手動篩一次命中列確認公式沒有抓錯範圍。"
   }
  ]
 },
 {
  "id": "countif-region",
  "kind": "formula",
  "role": "supporting",
  "badge": "計數",
  "title": "COUNTIF：符合條件的筆數怎麼數出來",
  "subtitle": "問「有幾筆？」時，先用 COUNTIF 數條件欄裡命中的儲存格。",
  "outcome": "結果：北區命中 4 格，所以回傳 4",
  "formulaParts": [
   {
    "text": "=COUNTIF("
   },
   {
    "text": "C2:C9",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ", "
   },
   {
    "text": "\"北區\"",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": ")"
   }
  ],
  "board": {
   "title": "訂單資料",
   "columns": [
    "業務員",
    "區域",
    "類別",
    "金額"
   ],
   "rows": [
    [
     {
      "text": "王小明"
     },
     {
      "text": "北區",
      "marks": {
       "1": "focus",
       "2": "match",
       "3": "result"
      }
     },
     {
      "text": "電子"
     },
     {
      "text": "15,000"
     }
    ],
    [
     {
      "text": "林美華"
     },
     {
      "text": "南區",
      "marks": {
       "1": "focus",
       "2": "exclude"
      }
     },
     {
      "text": "服飾"
     },
     {
      "text": "8,500"
     }
    ],
    [
     {
      "text": "張志偉"
     },
     {
      "text": "北區",
      "marks": {
       "1": "focus",
       "2": "match",
       "3": "result"
      }
     },
     {
      "text": "電子"
     },
     {
      "text": "22,000"
     }
    ],
    [
     {
      "text": "王小明"
     },
     {
      "text": "北區",
      "marks": {
       "1": "focus",
       "2": "match",
       "3": "result"
      }
     },
     {
      "text": "服飾"
     },
     {
      "text": "12,000"
     }
    ],
    [
     {
      "text": "陳雅婷"
     },
     {
      "text": "東區",
      "marks": {
       "1": "focus",
       "2": "exclude"
      }
     },
     {
      "text": "電子"
     },
     {
      "text": "9,800"
     }
    ],
    [
     {
      "text": "林美華"
     },
     {
      "text": "南區",
      "marks": {
       "1": "focus",
       "2": "exclude"
      }
     },
     {
      "text": "電子"
     },
     {
      "text": "31,000"
     }
    ],
    [
     {
      "text": "張志偉"
     },
     {
      "text": "北區",
      "marks": {
       "1": "focus",
       "2": "match",
       "3": "result"
      }
     },
     {
      "text": "電子"
     },
     {
      "text": "18,500"
     }
    ],
    [
     {
      "text": "陳雅婷"
     },
     {
      "text": "東區",
      "marks": {
       "1": "focus",
       "2": "exclude"
      }
     },
     {
      "text": "服飾"
     },
     {
      "text": "7,200"
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. 範圍：檢查區域欄",
    "tone": "c1",
    "text": "第一個參數是要檢查的範圍。這裡只看區域欄 C2:C9，不需要把金額欄圈進來。",
    "caption": "COUNTIF 是數格數，不是加總。"
   },
   {
    "label": "2. 條件：北區",
    "tone": "c2",
    "text": "條件填「北區」後，Excel 逐格比對。等於北區的留下，不等於北區的排除。",
    "caption": "文字條件要放在引號裡。"
   },
   {
    "label": "3. 結果：命中幾格就回傳幾",
    "tone": "c3",
    "text": "北區一共命中 4 格，所以公式回傳 4。",
    "caption": "問幾筆、幾人、幾次時，先想到 COUNTIF / COUNTIFS。"
   }
  ]
 },
 {
  "id": "sumif-category",
  "kind": "formula",
  "role": "supporting",
  "badge": "加總",
  "title": "SUMIF：單一條件加總，條件範圍在前",
  "subtitle": "SUMIF 先找符合條件的列，再去加總同列的數字。",
  "outcome": "結果：15,000 + 22,000 + 9,800 + 31,000 + 18,500 = 96,300",
  "formulaParts": [
   {
    "text": "=SUMIF("
   },
   {
    "text": "D2:D9",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ", "
   },
   {
    "text": "\"電子\"",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": ", "
   },
   {
    "text": "E2:E9",
    "step": 3,
    "tone": "c3"
   },
   {
    "text": ")"
   }
  ],
  "board": {
   "title": "依產品類別加總",
   "columns": [
    "產品類別",
    "金額"
   ],
   "rows": [
    [
     {
      "text": "電子",
      "marks": {
       "1": "focus",
       "2": "match"
      }
     },
     {
      "text": "15,000",
      "marks": {
       "3": "result",
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "服飾",
      "marks": {
       "1": "focus",
       "2": "exclude"
      }
     },
     {
      "text": "8,500",
      "marks": {
       "3": "focus"
      }
     }
    ],
    [
     {
      "text": "電子",
      "marks": {
       "1": "focus",
       "2": "match"
      }
     },
     {
      "text": "22,000",
      "marks": {
       "3": "result",
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "服飾",
      "marks": {
       "1": "focus",
       "2": "exclude"
      }
     },
     {
      "text": "12,000",
      "marks": {
       "3": "focus"
      }
     }
    ],
    [
     {
      "text": "電子",
      "marks": {
       "1": "focus",
       "2": "match"
      }
     },
     {
      "text": "9,800",
      "marks": {
       "3": "result",
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "電子",
      "marks": {
       "1": "focus",
       "2": "match"
      }
     },
     {
      "text": "31,000",
      "marks": {
       "3": "result",
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "電子",
      "marks": {
       "1": "focus",
       "2": "match"
      }
     },
     {
      "text": "18,500",
      "marks": {
       "3": "result",
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "服飾",
      "marks": {
       "1": "focus",
       "2": "exclude"
      }
     },
     {
      "text": "7,200",
      "marks": {
       "3": "focus"
      }
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. 條件範圍",
    "tone": "c1",
    "text": "SUMIF 第一個參數是條件範圍。這裡先看產品類別欄。",
    "caption": "單條件 SUMIF 是先條件、後加總。"
   },
   {
    "label": "2. 條件值",
    "tone": "c2",
    "text": "條件填「電子」後，只有產品類別等於電子的列會留下。",
    "caption": "其他類別不參與最後加總。"
   },
   {
    "label": "3. 加總範圍",
    "tone": "c3",
    "text": "第三個參數才是要加總的金額欄。Excel 只會加總剛剛命中的同列金額。",
    "caption": "這是 SUMIF 和 COUNTIF 最大的不同。"
   },
   {
    "label": "4. 結果",
    "tone": "c4",
    "text": "把電子類別命中的金額加起來，得到 96,300。",
    "caption": "問合計多少時，用 SUMIF / SUMIFS。"
   }
  ]
 },
 {
  "id": "sumifs-category-status",
  "kind": "formula",
  "role": "supporting",
  "badge": "多條件加總",
  "title": "SUMIFS：加總範圍放第一個",
  "subtitle": "多條件加總先說要加總哪一欄，再一組一組寫條件。",
  "outcome": "結果：15,000 + 22,000 + 31,000 + 18,500 = 86,500",
  "formulaParts": [
   {
    "text": "=SUMIFS("
   },
   {
    "text": "E2:E9",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ", "
   },
   {
    "text": "D2:D9",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": ", "
   },
   {
    "text": "\"電子\"",
    "step": 3,
    "tone": "c3"
   },
   {
    "text": ", "
   },
   {
    "text": "F2:F9",
    "step": 4,
    "tone": "c4"
   },
   {
    "text": ", "
   },
   {
    "text": "\"已付\"",
    "step": 5,
    "tone": "c5"
   },
   {
    "text": ")"
   }
  ],
  "board": {
   "title": "電子且已付款才加總",
   "columns": [
    "類別",
    "金額",
    "付款"
   ],
   "rows": [
    [
     {
      "text": "電子",
      "marks": {
       "2": "focus",
       "3": "candidate"
      }
     },
     {
      "text": "15,000",
      "marks": {
       "1": "focus",
       "5": "result"
      }
     },
     {
      "text": "已付",
      "marks": {
       "4": "focus",
       "5": "match"
      }
     }
    ],
    [
     {
      "text": "服飾",
      "marks": {
       "2": "focus",
       "3": "exclude"
      }
     },
     {
      "text": "8,500",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "未付",
      "marks": {
       "4": "focus",
       "5": "exclude"
      }
     }
    ],
    [
     {
      "text": "電子",
      "marks": {
       "2": "focus",
       "3": "candidate"
      }
     },
     {
      "text": "22,000",
      "marks": {
       "1": "focus",
       "5": "result"
      }
     },
     {
      "text": "已付",
      "marks": {
       "4": "focus",
       "5": "match"
      }
     }
    ],
    [
     {
      "text": "服飾",
      "marks": {
       "2": "focus",
       "3": "exclude"
      }
     },
     {
      "text": "12,000",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "已付",
      "marks": {
       "4": "focus",
       "5": "match"
      }
     }
    ],
    [
     {
      "text": "電子",
      "marks": {
       "2": "focus",
       "3": "candidate"
      }
     },
     {
      "text": "9,800",
      "marks": {
       "1": "focus",
       "5": "exclude"
      }
     },
     {
      "text": "未付",
      "marks": {
       "4": "focus",
       "5": "exclude"
      }
     }
    ],
    [
     {
      "text": "電子",
      "marks": {
       "2": "focus",
       "3": "candidate"
      }
     },
     {
      "text": "31,000",
      "marks": {
       "1": "focus",
       "5": "result"
      }
     },
     {
      "text": "已付",
      "marks": {
       "4": "focus",
       "5": "match"
      }
     }
    ],
    [
     {
      "text": "電子",
      "marks": {
       "2": "focus",
       "3": "candidate"
      }
     },
     {
      "text": "18,500",
      "marks": {
       "1": "focus",
       "5": "result"
      }
     },
     {
      "text": "已付",
      "marks": {
       "4": "focus",
       "5": "match"
      }
     }
    ],
    [
     {
      "text": "服飾",
      "marks": {
       "2": "focus",
       "3": "exclude"
      }
     },
     {
      "text": "7,200",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "已付",
      "marks": {
       "4": "focus",
       "5": "match"
      }
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. 加總範圍",
    "tone": "c1",
    "text": "SUMIFS 的第一個參數永遠是要加總的數字欄。這點和 SUMIF 順序不同。",
    "caption": "先說最後要加哪一欄，再寫條件。"
   },
   {
    "label": "2. 第一個條件範圍",
    "tone": "c2",
    "text": "接著指定第一個條件要檢查的欄位：產品類別。",
    "caption": "條件範圍和條件值要成對出現。"
   },
   {
    "label": "3. 第一個條件：電子",
    "tone": "c3",
    "text": "產品類別等於電子的列先成為候選資料，服飾列排除。",
    "caption": "這一步只是第一輪篩選，還沒完成。"
   },
   {
    "label": "4. 第二個條件範圍",
    "tone": "c4",
    "text": "第二組條件改看付款狀態欄。",
    "caption": "多條件可以一直往後接，但每個範圍大小要一致。"
   },
   {
    "label": "5. 第二個條件：已付",
    "tone": "c5",
    "text": "同時是電子且已付的列才會被加總；電子但未付也會被排除。",
    "caption": "SUMIFS 的多條件預設是 AND，全部成立才算。"
   }
  ]
 },
 {
  "id": "countifs-region-amount",
  "kind": "formula",
  "role": "supporting",
  "badge": "多條件計數",
  "title": "COUNTIFS：多條件是 AND 關係",
  "subtitle": "COUNTIFS 不加總金額，只數同時滿足所有條件的列數。",
  "outcome": "結果：北區且金額 >= 15000 的訂單有 3 筆",
  "formulaParts": [
   {
    "text": "=COUNTIFS("
   },
   {
    "text": "C2:C9",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ", \"北區\", "
   },
   {
    "text": "E2:E9",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": ", \">=15000\")",
    "step": 3,
    "tone": "c3"
   }
  ],
  "board": {
   "title": "北區且達標",
   "columns": [
    "區域",
    "金額"
   ],
   "rows": [
    [
     {
      "text": "北區",
      "marks": {
       "1": "candidate"
      }
     },
     {
      "text": "15,000",
      "marks": {
       "2": "focus",
       "3": "result"
      }
     }
    ],
    [
     {
      "text": "南區",
      "marks": {
       "1": "exclude"
      }
     },
     {
      "text": "8,500",
      "marks": {
       "2": "focus",
       "3": "exclude"
      }
     }
    ],
    [
     {
      "text": "北區",
      "marks": {
       "1": "candidate"
      }
     },
     {
      "text": "22,000",
      "marks": {
       "2": "focus",
       "3": "result"
      }
     }
    ],
    [
     {
      "text": "北區",
      "marks": {
       "1": "candidate"
      }
     },
     {
      "text": "12,000",
      "marks": {
       "2": "focus",
       "3": "exclude"
      }
     }
    ],
    [
     {
      "text": "東區",
      "marks": {
       "1": "exclude"
      }
     },
     {
      "text": "9,800",
      "marks": {
       "2": "focus",
       "3": "exclude"
      }
     }
    ],
    [
     {
      "text": "南區",
      "marks": {
       "1": "exclude"
      }
     },
     {
      "text": "31,000",
      "marks": {
       "2": "focus",
       "3": "exclude"
      }
     }
    ],
    [
     {
      "text": "北區",
      "marks": {
       "1": "candidate"
      }
     },
     {
      "text": "18,500",
      "marks": {
       "2": "focus",
       "3": "result"
      }
     }
    ],
    [
     {
      "text": "東區",
      "marks": {
       "1": "exclude"
      }
     },
     {
      "text": "7,200",
      "marks": {
       "2": "focus",
       "3": "exclude"
      }
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. 條件一：北區",
    "tone": "c1",
    "text": "第一組條件先留下北區列。這些只是候選，還要通過下一個條件。",
    "caption": "COUNTIFS 每一組都是條件範圍、條件值。"
   },
   {
    "label": "2. 條件二：看金額欄",
    "tone": "c2",
    "text": "第二組條件檢查金額欄，準備判斷是否達到 15,000。",
    "caption": "比較條件通常寫成引號裡的字串，例如 \">=15000\"。"
   },
   {
    "label": "3. 同時成立才計數",
    "tone": "c3",
    "text": "只有北區且金額 >= 15,000 的列被計入，所以結果是 3。",
    "caption": "如果你想做 OR，通常要把多個 COUNTIFS 相加。"
   }
  ]
 },
 {
  "id": "averageifs-region-category",
  "kind": "formula",
  "role": "supporting",
  "badge": "多條件平均",
  "title": "AVERAGEIFS：只平均命中的數字",
  "subtitle": "AVERAGEIFS 的語法和 SUMIFS 很像，只是第一個參數換成要取平均的範圍。",
  "outcome": "結果：(15,000 + 22,000 + 18,500) / 3 = 18,500",
  "formulaParts": [
   {
    "text": "=AVERAGEIFS("
   },
   {
    "text": "E2:E9",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ", "
   },
   {
    "text": "C2:C9",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": ", \"北區\", "
   },
   {
    "text": "D2:D9",
    "step": 3,
    "tone": "c3"
   },
   {
    "text": ", \"電子\")",
    "step": 4,
    "tone": "c4"
   }
  ],
  "board": {
   "title": "北區電子類平均金額",
   "columns": [
    "區域",
    "類別",
    "金額"
   ],
   "rows": [
    [
     {
      "text": "北區",
      "marks": {
       "2": "candidate"
      }
     },
     {
      "text": "電子",
      "marks": {
       "3": "candidate"
      }
     },
     {
      "text": "15,000",
      "marks": {
       "1": "focus",
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "南區",
      "marks": {
       "2": "exclude"
      }
     },
     {
      "text": "服飾",
      "marks": {
       "3": "exclude"
      }
     },
     {
      "text": "8,500",
      "marks": {
       "1": "focus"
      }
     }
    ],
    [
     {
      "text": "北區",
      "marks": {
       "2": "candidate"
      }
     },
     {
      "text": "電子",
      "marks": {
       "3": "candidate"
      }
     },
     {
      "text": "22,000",
      "marks": {
       "1": "focus",
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "北區",
      "marks": {
       "2": "candidate"
      }
     },
     {
      "text": "服飾",
      "marks": {
       "3": "exclude"
      }
     },
     {
      "text": "12,000",
      "marks": {
       "1": "focus",
       "4": "exclude"
      }
     }
    ],
    [
     {
      "text": "東區",
      "marks": {
       "2": "exclude"
      }
     },
     {
      "text": "電子",
      "marks": {
       "3": "candidate"
      }
     },
     {
      "text": "9,800",
      "marks": {
       "1": "focus"
      }
     }
    ],
    [
     {
      "text": "南區",
      "marks": {
       "2": "exclude"
      }
     },
     {
      "text": "電子",
      "marks": {
       "3": "candidate"
      }
     },
     {
      "text": "31,000",
      "marks": {
       "1": "focus"
      }
     }
    ],
    [
     {
      "text": "北區",
      "marks": {
       "2": "candidate"
      }
     },
     {
      "text": "電子",
      "marks": {
       "3": "candidate"
      }
     },
     {
      "text": "18,500",
      "marks": {
       "1": "focus",
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "東區",
      "marks": {
       "2": "exclude"
      }
     },
     {
      "text": "服飾",
      "marks": {
       "3": "exclude"
      }
     },
     {
      "text": "7,200",
      "marks": {
       "1": "focus"
      }
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. 平均範圍",
    "tone": "c1",
    "text": "AVERAGEIFS 第一個參數是要取平均的數字欄。這和 SUMIFS 的位置邏輯一致。",
    "caption": "不是先寫條件欄，而是先寫最後要計算的欄。"
   },
   {
    "label": "2. 條件一：北區",
    "tone": "c2",
    "text": "第一組條件留下北區資料。",
    "caption": "非北區不會進入最後平均。"
   },
   {
    "label": "3. 條件二：電子",
    "tone": "c3",
    "text": "第二組條件再留下電子類別。北區但不是電子，也會被排除。",
    "caption": "多條件平均也是 AND 關係。"
   },
   {
    "label": "4. 結果：只平均命中的金額",
    "tone": "c4",
    "text": "最後只拿北區電子類的三筆金額平均，得到 18,500。",
    "caption": "問符合條件的平均時，用 AVERAGEIF / AVERAGEIFS。"
   }
  ]
 },
 {
  "id": "wildcards-comparisons",
  "kind": "formula",
  "role": "supporting",
  "badge": "條件寫法",
  "title": "萬用字元與比較條件：條件不只等於",
  "subtitle": "條件可以是文字片段、不等於，也可以是大於等於。",
  "outcome": "結果：姓王、非已付、金額達標都能用條件統計快速回答",
  "formulaParts": [
   {
    "text": "=COUNTIF(B2:B9,"
   },
   {
    "text": "\"王*\"",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ")  "
   },
   {
    "text": "=COUNTIF(F2:F9,"
   },
   {
    "text": "\"<>已付\"",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": ")  "
   },
   {
    "text": "=SUMIFS(E2:E9,C2:C9,\"北區\",E2:E9,"
   },
   {
    "text": "\">=15000\"",
    "step": 3,
    "tone": "c3"
   },
   {
    "text": ")"
   }
  ],
  "board": {
   "title": "常用條件寫法",
   "columns": [
    "業務員",
    "區域",
    "金額",
    "付款"
   ],
   "rows": [
    [
     {
      "text": "王小明",
      "marks": {
       "1": "match"
      }
     },
     {
      "text": "北區",
      "marks": {
       "3": "candidate"
      }
     },
     {
      "text": "15,000",
      "marks": {
       "3": "result"
      }
     },
     {
      "text": "已付",
      "marks": {
       "2": "exclude"
      }
     }
    ],
    [
     {
      "text": "林美華",
      "marks": {
       "1": "exclude"
      }
     },
     {
      "text": "南區"
     },
     {
      "text": "8,500"
     },
     {
      "text": "未付",
      "marks": {
       "2": "match"
      }
     }
    ],
    [
     {
      "text": "張志偉",
      "marks": {
       "1": "exclude"
      }
     },
     {
      "text": "北區",
      "marks": {
       "3": "candidate"
      }
     },
     {
      "text": "22,000",
      "marks": {
       "3": "result"
      }
     },
     {
      "text": "已付",
      "marks": {
       "2": "exclude"
      }
     }
    ],
    [
     {
      "text": "王小明",
      "marks": {
       "1": "match"
      }
     },
     {
      "text": "北區",
      "marks": {
       "3": "candidate"
      }
     },
     {
      "text": "12,000",
      "marks": {
       "3": "exclude"
      }
     },
     {
      "text": "已付",
      "marks": {
       "2": "exclude"
      }
     }
    ],
    [
     {
      "text": "陳雅婷",
      "marks": {
       "1": "exclude"
      }
     },
     {
      "text": "東區"
     },
     {
      "text": "9,800"
     },
     {
      "text": "未付",
      "marks": {
       "2": "match"
      }
     }
    ],
    [
     {
      "text": "林美華",
      "marks": {
       "1": "exclude"
      }
     },
     {
      "text": "南區"
     },
     {
      "text": "31,000"
     },
     {
      "text": "已付",
      "marks": {
       "2": "exclude"
      }
     }
    ],
    [
     {
      "text": "張志偉",
      "marks": {
       "1": "exclude"
      }
     },
     {
      "text": "北區",
      "marks": {
       "3": "candidate"
      }
     },
     {
      "text": "18,500",
      "marks": {
       "3": "result"
      }
     },
     {
      "text": "已付",
      "marks": {
       "2": "exclude"
      }
     }
    ],
    [
     {
      "text": "陳雅婷",
      "marks": {
       "1": "exclude"
      }
     },
     {
      "text": "東區"
     },
     {
      "text": "7,200"
     },
     {
      "text": "已付",
      "marks": {
       "2": "exclude"
      }
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. 王*：開頭是王",
    "tone": "c1",
    "text": "星號代表任意長度文字。王* 會命中王小明、王大華這類以王開頭的文字。",
    "caption": "用在姓名、產品名、公司名都很常見。"
   },
   {
    "label": "2. <>：不等於",
    "tone": "c2",
    "text": "\"<>已付\" 代表付款狀態不是已付。它可以拿來找未付、取消、待確認等例外資料。",
    "caption": "不等於很適合做追蹤清單。"
   },
   {
    "label": "3. >=：達到門檻",
    "tone": "c3",
    "text": "比較條件通常放在引號裡，例如 \">=15000\"。如果門檻在 H2，寫成 \">=\"&H2。",
    "caption": "條件統計常搭配 KPI 門檻使用。"
   }
  ]
 }
];
