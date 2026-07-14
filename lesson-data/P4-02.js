// 由 xlsx-content.js + lesson-demos-data.js 拆分(2026-07-14);本檔是 P4-02 的唯一資料來源
export const CONTENT = {
 "knowledge": {
  "title": "🧮 進階函式",
  "subtitle": "把長公式變得可讀、可測、可重用",
  "sections": [
   {
    "title": "📌 先判斷要解決什麼問題",
    "items": [
     [
      "公式太長",
      "用 LET 替中間結果命名",
      "讓 total、bonus、result 各自有角色"
     ],
     [
      "邏輯重複",
      "用 LAMBDA 封裝成自訂函式",
      "先測通，再放進名稱管理員"
     ],
     [
      "錯誤或隱藏列干擾統計",
      "用 AGGREGATE 選擇忽略方式",
      "比一般 SUM / AVERAGE 更能處理髒資料"
     ]
    ]
   },
   {
    "title": "🧰 常用語法",
    "items": [
     [
      "LET",
      "=LET(total,SUM(B2:E2),bonus,IF(total>=500000,total*5%,0),total+bonus)",
      "最後一段是回傳值"
     ],
     [
      "LAMBDA",
      "=LAMBDA(amount,rate,amount*(1+rate))(120000,5%)",
      "括號後面的值是立即測試輸入"
     ],
     [
      "SWITCH",
      "=SWITCH(A2,\"A\",\"優先\",\"B\",\"一般\",\"其他\")",
      "比多層 IF 更像對照表"
     ]
    ]
   },
   {
    "title": "⚠️ 高風險提醒",
    "items": [
     [
      "INDIRECT / OFFSET",
      "可以做動態參照，但屬於 volatile 函數",
      "大型檔案容易拖慢重算"
     ],
     [
      "版本相容",
      "LET / LAMBDA 需要較新的 Excel",
      "交付前先確認收件者環境"
     ],
     [
      "封裝前先測試",
      "LAMBDA 請先用立即執行版本驗證",
      "不要把錯誤公式藏進名稱管理員"
     ]
    ]
   }
  ],
  "handsTasks": [
   "用 LET 命名全年業績與獎金",
   "用 LAMBDA 測試可重用公式",
   "用 AGGREGATE 忽略錯誤值統計",
   "辨認 INDIRECT / OFFSET 的效能風險"
  ]
 },
 "pro": {
  "title": "🧮 進階函式  ─  LET / LAMBDA / AGGREGATE / SWITCH / INDIRECT",
  "subtitle": "難度：🟠 Lv.4  |  微任務數：8 題  |  建議時間：每題 2~3 分鐘",
  "dataHeader": [
   "季度",
   "Q1",
   "Q2",
   "Q3",
   "Q4",
   "",
   ""
  ],
  "dataRows": [
   [
    "2024",
    120000,
    135000,
    128000,
    145000,
    "",
    ""
   ],
   [
    "2025",
    150000,
    162000,
    158000,
    175000,
    "",
    ""
   ],
   [
    "2026",
    168000,
    180000,
    0,
    0,
    "",
    ""
   ]
  ],
  "tasks": [
   {
    "numLabel": "🟢1",
    "desc": "用 LET 計算 2025 全年業績，命名 total",
    "time": "3m",
    "hint": "=LET(total,SUM(B7:E7),total)",
    "answer": "=LET(total,SUM(B7:E7),total)",
    "explain": "LET 的最後一段是回傳值"
   },
   {
    "numLabel": "🟢2",
    "desc": "用 LET 計算 2025 全年業績，達標 600000 加 5% 獎金",
    "time": "3m",
    "hint": "=LET(total,SUM(B7:E7),bonus,IF(total>=600000,total*5%,0),total+bonus)",
    "answer": "=LET(total,SUM(B7:E7),bonus,IF(total>=600000,total*5%,0),total+bonus)",
    "explain": "把 total 與 bonus 分開命名，公式更容易讀"
   },
   {
    "numLabel": "🔵3",
    "desc": "用 LAMBDA 立即測試：金額 120000 加 5%",
    "time": "3m",
    "hint": "=LAMBDA(amount,rate,amount*(1+rate))(120000,5%)",
    "answer": "=LAMBDA(amount,rate,amount*(1+rate))(120000,5%)",
    "explain": "先用立即執行測通，再放到名稱管理員"
   },
   {
    "numLabel": "🔵4",
    "desc": "用 AGGREGATE 忽略錯誤值求平均",
    "time": "3m",
    "hint": "=AGGREGATE(1,6,B6:E8)",
    "answer": "=AGGREGATE(1,6,B6:E8)",
    "explain": "AGGREGATE(1,6,範圍) 代表忽略錯誤求平均"
   },
   {
    "numLabel": "🟡5",
    "desc": "用 SWITCH 把季度代碼轉成名稱",
    "time": "3m",
    "hint": "=SWITCH(\"Q2\",\"Q1\",\"第一季\",\"Q2\",\"第二季\",\"Q3\",\"第三季\",\"Q4\",\"第四季\",\"未知\")",
    "answer": "=SWITCH(\"Q2\",\"Q1\",\"第一季\",\"Q2\",\"第二季\",\"Q3\",\"第三季\",\"Q4\",\"第四季\",\"未知\")",
    "explain": "SWITCH 比多層 IF 更像清楚的對照表"
   },
   {
    "numLabel": "🟡6",
    "desc": "用 INDIRECT 把文字 B7 轉成儲存格參照",
    "time": "2m",
    "hint": "=INDIRECT(\"B7\")",
    "answer": "=INDIRECT(\"B7\")",
    "explain": "INDIRECT 可用，但大型檔案要留意重算效能"
   },
   {
    "numLabel": "🔴7",
    "desc": "用 OFFSET 取得 2025 年度從 Q1 開始的 4 欄總和",
    "time": "2m",
    "hint": "=SUM(OFFSET(B7,0,0,1,4))",
    "answer": "=SUM(OFFSET(B7,0,0,1,4))",
    "explain": "OFFSET 能動態位移與指定範圍大小，但也屬於 volatile 函數"
   },
   {
    "numLabel": "🔴8",
    "desc": "用 LET + SWITCH 將總額分成高 / 中 / 低",
    "time": "3m",
    "hint": "=LET(total,SUM(B7:E7),SWITCH(TRUE,total>=600000,\"高\",total>=500000,\"中\",\"低\"))",
    "answer": "=LET(total,SUM(B7:E7),SWITCH(TRUE,total>=600000,\"高\",total>=500000,\"中\",\"低\"))",
    "explain": "LET 先命名 total，SWITCH(TRUE,...) 再做門檻判斷"
   }
  ]
 },
 "meta": {
  "phase": "⚡ Phase 4：進階自動化    第 7~10 週  |  目標：處理大量資料、自動化流程",
  "stage": "第14階段",
  "topics": "LET / LAMBDA / AGGREGATE / SWITCH / INDIRECT",
  "difficulty": "🟠 Lv.4",
  "taskCount": "8 題",
  "time": "25 分鐘",
  "xp": "250 XP"
 },
 "inter": {
  "title": "🏆 進階綜合挑戰",
  "subtitle": "在黃色邊框儲存格輸入公式 → 答對變綠 ✅ | 答錯變紅 ❌ | F欄提示選取可見",
  "dataHeader": [
   "員工",
   "部門",
   "月薪",
   "Q1業績",
   "Q2業績",
   "Q3業績",
   "Q4業績"
  ],
  "dataRows": [
   [
    "王小明",
    "業務",
    55000,
    120000,
    135000,
    98000,
    145000
   ],
   [
    "林美華",
    "財務",
    48000,
    0,
    0,
    0,
    0
   ],
   [
    "張志偉",
    "業務",
    68000,
    210000,
    185000,
    220000,
    195000
   ],
   [
    "陳雅琪",
    "資訊",
    42000,
    0,
    0,
    0,
    0
   ],
   [
    "李大明",
    "業務",
    51000,
    95000,
    110000,
    88000,
    120000
   ],
   [
    "吳建志",
    "財務",
    52000,
    0,
    0,
    0,
    0
   ]
  ],
  "tasks": [
   {
    "num": 1,
    "difficulty": "🟢",
    "desc": "用 LET 命名王小明全年業績 total",
    "hint": "=LET(total,SUM(D6:G6),total)",
    "answer": "=LET(total,SUM(D6:G6),total)"
   },
   {
    "num": 2,
    "difficulty": "🟢",
    "desc": "用 LET 計算全年業績，達標 500000 加 5% 獎金",
    "hint": "total + bonus",
    "answer": "=LET(total,SUM(D6:G6),bonus,IF(total>=500000,total*5%,0),total+bonus)"
   },
   {
    "num": 3,
    "difficulty": "🔵",
    "desc": "用 SWITCH 把王小明的部門轉成英文標籤",
    "hint": "SWITCH(B6,\"業務\",\"Sales\",...)",
    "answer": "=SWITCH(B6,\"業務\",\"Sales\",\"財務\",\"Finance\",\"資訊\",\"IT\",\"其他\")"
   },
   {
    "num": 4,
    "difficulty": "🔵",
    "desc": "用 AGGREGATE 忽略錯誤求月薪平均",
    "hint": "=AGGREGATE(1,6,C6:C11)",
    "answer": "=AGGREGATE(1,6,C6:C11)"
   },
   {
    "num": 5,
    "difficulty": "🟡",
    "desc": "用 LAMBDA 立即測試：全年業績是否達標",
    "hint": "=LAMBDA(x,IF(x>=500000,\"達標\",\"追蹤\"))(SUM(D6:G6))",
    "answer": "=LAMBDA(x,IF(x>=500000,\"達標\",\"追蹤\"))(SUM(D6:G6))"
   },
   {
    "num": 6,
    "difficulty": "🟡",
    "desc": "用 INDIRECT 把文字 D6 轉成 Q1 業績值",
    "hint": "=INDIRECT(\"D6\")",
    "answer": "=INDIRECT(\"D6\")"
   },
   {
    "num": 7,
    "difficulty": "🔴",
    "desc": "用 OFFSET 加總王小明 Q1~Q4",
    "hint": "=SUM(OFFSET(D6,0,0,1,4))",
    "answer": "=SUM(OFFSET(D6,0,0,1,4))"
   },
   {
    "num": 8,
    "difficulty": "🔴",
    "desc": "用 LET + SWITCH 將全年業績分成高 / 中 / 低",
    "hint": "SWITCH(TRUE, 條件1, 結果1, 條件2, 結果2, 預設)",
    "answer": "=LET(total,SUM(D6:G6),SWITCH(TRUE,total>=500000,\"高\",total>=300000,\"中\",\"低\"))"
   }
  ]
 }
};
export const DEMOS = [
 {
  "id": "let-quarterly-bonus",
  "kind": "formula",
  "badge": "公式重構",
  "title": "LET：把季度獎金公式拆成可讀段落",
  "subtitle": "先命名全年業績 total，再命名獎金 bonus，最後回傳 total + bonus。",
  "outcome": "結果：810,000 達標，獎金 40,500，最後回傳 850,500",
  "formulaParts": [
   {
    "text": "=LET("
   },
   {
    "text": "total",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ", "
   },
   {
    "text": "SUM(B2:E2)",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": ", "
   },
   {
    "text": "bonus",
    "step": 3,
    "tone": "c3"
   },
   {
    "text": ", IF(total>=500000,total*5%,0), "
   },
   {
    "text": "total+bonus",
    "step": 4,
    "tone": "c4"
   },
   {
    "text": ")"
   }
  ],
  "board": {
   "title": "張志偉季度業績",
   "columns": [
    "Q1",
    "Q2",
    "Q3",
    "Q4",
    "公式結果"
   ],
   "rows": [
    [
     {
      "text": "210,000",
      "marks": {
       "2": "focus"
      }
     },
     {
      "text": "185,000",
      "marks": {
       "2": "focus"
      }
     },
     {
      "text": "220,000",
      "marks": {
       "2": "focus"
      }
     },
     {
      "text": "195,000",
      "marks": {
       "2": "focus"
      }
     },
     {
      "text": "850,500",
      "marks": {
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "total",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "810,000",
      "marks": {
       "2": "result"
      }
     },
     {
      "text": "bonus",
      "marks": {
       "3": "focus"
      }
     },
     {
      "text": "40,500",
      "marks": {
       "3": "result"
      }
     },
     {
      "text": "total + bonus",
      "marks": {
       "4": "match"
      }
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. 先替全年業績取名 total",
    "tone": "c1",
    "text": "LET 的變數名稱可以自己命名。total 讓後面的人一眼知道這段代表全年業績。",
    "caption": "命名不是為了省字，而是為了讓公式自我解釋。"
   },
   {
    "label": "2. total 的值只計算一次",
    "tone": "c2",
    "text": "SUM(B2:E2) 把 Q1 到 Q4 加總成 810,000。之後公式裡每次寫 total 都引用這個結果。",
    "caption": "長公式中重複出現的計算，很適合先用 LET 存起來。"
   },
   {
    "label": "3. 再命名 bonus，把判斷邏輯收起來",
    "tone": "c3",
    "text": "IF(total>=500000,total*5%,0) 代表達標才給 5% 獎金。把這段叫 bonus，比直接塞在最後更好讀。",
    "caption": "LET 可以定義多組名稱和值。"
   },
   {
    "label": "4. 最後一段就是回傳值",
    "tone": "c4",
    "text": "LET 的最後一個引數是最終答案。這裡回傳 total+bonus，也就是 850,500。",
    "caption": "讀 LET 時可以先找最後一段，知道這條公式最後要吐什麼。"
   }
  ]
 },
 {
  "id": "lambda-reusable-tax",
  "kind": "formula",
  "badge": "封裝邏輯",
  "title": "LAMBDA：把公式變成可重用的自訂函式",
  "subtitle": "先在儲存格立即執行測試，確認沒錯後再放進名稱管理員。",
  "outcome": "結果：120,000 加上 5% 後回傳 126,000",
  "formulaParts": [
   {
    "text": "=LAMBDA("
   },
   {
    "text": "amount, rate",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ", "
   },
   {
    "text": "amount*(1+rate)",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": ")("
   },
   {
    "text": "120000, 5%",
    "step": 3,
    "tone": "c3"
   },
   {
    "text": ")"
   }
  ],
  "board": {
   "title": "從測試公式到自訂函式",
   "columns": [
    "階段",
    "內容",
    "結果"
   ],
   "rows": [
    [
     {
      "text": "參數",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "amount, rate",
      "marks": {
       "1": "result"
      }
     },
     {
      "text": "等待輸入"
     }
    ],
    [
     {
      "text": "計算式",
      "marks": {
       "2": "focus"
      }
     },
     {
      "text": "amount*(1+rate)",
      "marks": {
       "2": "result"
      }
     },
     {
      "text": "加上比例"
     }
    ],
    [
     {
      "text": "測試值",
      "marks": {
       "3": "focus"
      }
     },
     {
      "text": "120000, 5%",
      "marks": {
       "3": "match"
      }
     },
     {
      "text": "126,000",
      "marks": {
       "3": "result",
       "4": "result"
      }
     }
    ],
    [
     {
      "text": "命名後",
      "marks": {
       "4": "focus"
      }
     },
     {
      "text": "=含稅價(120000,5%)",
      "marks": {
       "4": "match"
      }
     },
     {
      "text": "126,000",
      "marks": {
       "4": "result"
      }
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. 先列出這個邏輯需要哪些參數",
    "tone": "c1",
    "text": "amount 和 rate 是這個自訂函式需要的兩個輸入。參數名稱要清楚，之後才看得懂。",
    "caption": "LAMBDA 的第一段是參數清單。"
   },
   {
    "label": "2. 寫出要重用的計算邏輯",
    "tone": "c2",
    "text": "amount*(1+rate) 是實際要重複使用的公式邏輯。這裡示範把金額加上一個比例。",
    "caption": "這段可以很短，也可以包進 LET 讓它更清楚。"
   },
   {
    "label": "3. 先立即執行測試",
    "tone": "c3",
    "text": "最後括號裡的 120000, 5% 是測試輸入。立即執行能確認公式真的回傳 126,000。",
    "caption": "不要把沒測過的 LAMBDA 直接丟進名稱管理員。"
   },
   {
    "label": "4. 測通後再放進名稱管理員",
    "tone": "c4",
    "text": "命名成含稅價後，就能像內建函數一樣使用：=含稅價(120000,5%)。",
    "caption": "LAMBDA 的價值是讓常用邏輯可重用，而不是把公式藏起來。"
   }
  ]
 }
];
