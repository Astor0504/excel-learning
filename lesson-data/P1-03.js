// 由 xlsx-content.js + lesson-demos-data.js 拆分(2026-07-14);本檔是 P1-03 的唯一資料來源
export const CONTENT = {
 "knowledge": {
  "title": "📚 條件判斷不是硬塞 IF，是先分清楚判斷型態",
  "subtitle": "難度：🟢 Lv.1  |  先分清楚單一條件 / 多門檻 / 複合條件 / 固定值對應",
  "sections": [
   {
    "title": "📌 先問這是哪一種判斷",
    "items": [
     [
      "單一條件",
      "IF",
      "例如績效>=80 顯示優良，否則待改善",
      "最適合只有兩種結果的欄位判斷"
     ],
     [
      "多個門檻",
      "IFS",
      "例如 >=90 是 A、>=75 是 B、其他是 C",
      "門檻要由高到低寫，最後用 TRUE 當兜底"
     ],
     [
      "複合條件",
      "AND / OR",
      "AND 是全部成立；OR 是任一成立",
      "通常包在 IF 裡，用來判斷模範員工、優惠資格、風險提醒"
     ],
     [
      "固定值對應",
      "SWITCH",
      "例如部門名稱轉英文代碼、狀態代碼轉文字",
      "不是範圍門檻，而是一個值對一個結果"
     ]
    ]
   },
   {
    "title": "⚠️ 3 個常見錯誤",
    "items": [
     [
      "IFS 順序寫反",
      "先寫 >=70 再寫 >=90，90 分會先被判成 C",
      "多門檻永遠先放最嚴格的條件"
     ],
     [
      "AND / OR 放錯位置",
      "AND(D6>=80,E6>=22) 只會回傳 TRUE/FALSE",
      "若要顯示文字結果，要外面再包 IF"
     ],
     [
      "SWITCH 拿來判斷大小",
      "SWITCH 適合等於某個固定值",
      "若條件是 >=、<=、介於某區間，改用 IFS"
     ]
    ]
   },
   {
    "title": "🎯 這課真正要練熟的思路",
    "items": [
     [
      "先描述規則，再寫公式",
      "把規則念成一句話：如果績效大於等於 80，就顯示優良",
      "能說清楚，公式通常就能寫清楚"
     ],
     [
      "把兜底結果補上",
      "IFS 最後常放 TRUE；SWITCH 最後常放其他",
      "沒有兜底，資料一變就容易出錯或顯示空白"
     ],
     [
      "讓每列自己判斷",
      "P1-03 是欄位層級判斷，不是彙總報表",
      "要依條件加總或計數，下一課再交給 SUMIFS / COUNTIFS"
     ]
    ]
   }
  ],
  "handsTasks": [
   {
    "num": 1,
    "difficulty": "🟢 暖身",
    "desc": "先用 IF 判斷一位員工績效是否優良，再把門檻改成 75 觀察結果"
   },
   {
    "num": 2,
    "difficulty": "🟢 暖身",
    "desc": "用 IFS 寫 A / B / C 等第，並確認 TRUE 放在最後"
   },
   {
    "num": 3,
    "difficulty": "🔵 標準",
    "desc": "用 AND 判斷出勤與績效是否同時達標，再包進 IF 顯示「模範」"
   },
   {
    "num": 4,
    "difficulty": "🔵 標準",
    "desc": "用 SWITCH 把部門轉成英文代碼，例如業務=Sales、財務=Finance"
   }
  ]
 },
 "pro": {
  "title": "📐 條件判斷  ─  IF / IFS / AND / OR / SWITCH",
  "subtitle": "難度：🟢 Lv.1  |  微任務數：8 題  |  建議時間：每題 2~3 分鐘",
  "dataHeader": [
   "姓名",
   "部門",
   "月薪",
   "績效分數",
   "出勤天數",
   "",
   ""
  ],
  "dataRows": [
   [
    "王小明",
    "業務",
    55000,
    88,
    22,
    "",
    ""
   ],
   [
    "林美華",
    "財務",
    48000,
    72,
    20,
    "",
    ""
   ],
   [
    "張志偉",
    "業務",
    68000,
    95,
    23,
    "",
    ""
   ],
   [
    "陳雅婷",
    "人事",
    42000,
    65,
    18,
    "",
    ""
   ],
   [
    "李建宏",
    "資訊",
    75000,
    91,
    22,
    "",
    ""
   ],
   [
    "劉淑芬",
    "財務",
    52000,
    79,
    21,
    "",
    ""
   ]
  ],
  "tasks": [
   {
    "numLabel": "🟢1",
    "desc": "績效>=80 顯示「優良」否則「待改善」",
    "time": "2m",
    "hint": "=IF(D6>=80,\"優良\",\"待改善\")",
    "answer": "=IF(D6>=80,\"優良\",\"待改善\")",
    "explain": "IF(條件, 真, 假) 是最基本的判斷"
   },
   {
    "numLabel": "🟢2",
    "desc": "月薪>=60000 顯示「高薪」否則「一般」",
    "time": "2m",
    "hint": "=IF(C6>=60000,\"高薪\",\"一般\")",
    "answer": "=IF(C6>=60000,\"高薪\",\"一般\")",
    "explain": "把門檻值改成 60000 即可"
   },
   {
    "numLabel": "🔵3",
    "desc": "績效>=90「A」，>=75「B」，其他「C」",
    "time": "3m",
    "hint": "=IFS(D6>=90,\"A\",D6>=75,\"B\",TRUE,\"C\")",
    "answer": "=IFS(D6>=90,\"A\",D6>=75,\"B\",TRUE,\"C\")",
    "explain": "IFS 依序判斷，TRUE 是最後的「其他」"
   },
   {
    "numLabel": "🔵4",
    "desc": "同上，改用巢狀 IF 寫法",
    "time": "3m",
    "hint": "=IF(D6>=90,\"A\",IF(D6>=75,\"B\",\"C\"))",
    "answer": "=IF(D6>=90,\"A\",IF(D6>=75,\"B\",\"C\"))",
    "explain": "巢狀 IF = IF 裡面再放 IF"
   },
   {
    "numLabel": "🟡5",
    "desc": "出勤>=22 且 績效>=80 → 「模範」",
    "time": "2m",
    "hint": "=IF(AND(E6>=22,D6>=80),\"模範\",\"一般\")",
    "answer": "=IF(AND(E6>=22,D6>=80),\"模範\",\"一般\")",
    "explain": "AND 表示「兩個條件都要成立」"
   },
   {
    "numLabel": "🟡6",
    "desc": "部門是業務 或 績效>=90 → 「重點關注」",
    "time": "2m",
    "hint": "=IF(OR(B6=\"業務\",D6>=90),\"重點關注\",\"一般\")",
    "answer": "=IF(OR(B6=\"業務\",D6>=90),\"重點關注\",\"一般\")",
    "explain": "OR 表示「任一條件成立即可」"
   },
   {
    "numLabel": "🔴7",
    "desc": "績效>=90 獎金=月薪×3，>=75 獎金=×2，其他=×1",
    "time": "3m",
    "hint": "=IFS(D6>=90,C6*3,D6>=75,C6*2,TRUE,C6*1)",
    "answer": "=IFS(D6>=90,C6*3,D6>=75,C6*2,TRUE,C6*1)",
    "explain": "IFS 結合數學運算，直接算出金額"
   },
   {
    "numLabel": "🔴8",
    "desc": "用 SWITCH 把部門轉成英文代碼",
    "time": "2m",
    "hint": "=SWITCH(B6,\"業務\",\"Sales\",\"財務\",\"Finance\",\"資訊\",\"IT\",\"人事\",\"HR\",\"其他\")",
    "answer": "=SWITCH(B6,\"業務\",\"Sales\",\"財務\",\"Finance\",\"資訊\",\"IT\",\"人事\",\"HR\",\"其他\")",
    "explain": "SWITCH 適合固定值對應，比一串 IF 更好讀"
   }
  ]
 },
 "meta": {
  "phase": "🏃 Phase 1：操作效率基礎    第 1~2 週  |  目標：養成鍵盤操作習慣，複習基本功",
  "stage": "第 3 階段",
  "topics": "IF / IFS / AND / OR / SWITCH",
  "difficulty": "🟢 Lv.1",
  "taskCount": "8 題",
  "time": "15 分鐘",
  "xp": "100 XP"
 },
 "inter": {
  "title": "📐 條件判斷互動練習",
  "subtitle": "在黃色格子輸入公式 → 答對自動變綠 ✅ 答錯變紅 ❌",
  "dataHeader": [
   "姓名",
   "部門",
   "月薪",
   "績效分數",
   "出勤天數",
   "",
   ""
  ],
  "dataRows": [
   [
    "王小明",
    "業務",
    55000,
    88,
    22,
    "",
    ""
   ],
   [
    "林美華",
    "財務",
    48000,
    72,
    20,
    "",
    ""
   ],
   [
    "張志偉",
    "業務",
    68000,
    95,
    23,
    "",
    ""
   ],
   [
    "陳雅琪",
    "資訊",
    42000,
    65,
    18,
    "",
    ""
   ],
   [
    "李大明",
    "業務",
    51000,
    78,
    21,
    "",
    ""
   ]
  ],
  "tasks": [
   {
    "num": 1,
    "difficulty": "🟢",
    "desc": "績效>=80 → \"優良\"，否則 \"待改善\"",
    "hint": "IF(條件,真,假)",
    "answer": "=IF(D6>=80,\"優良\",\"待改善\")"
   },
   {
    "num": 2,
    "difficulty": "🟢",
    "desc": "月薪>=60000 → \"高薪\"，否則 \"一般\"",
    "hint": "把門檻改成 60000",
    "answer": "=IF(C6>=60000,\"高薪\",\"一般\")"
   },
   {
    "num": 3,
    "difficulty": "🔵",
    "desc": "績效>=90→\"A\"，>=75→\"B\"，其他→\"C\"",
    "hint": "IFS 依序判斷，TRUE=其他",
    "answer": "=IFS(D6>=90,\"A\",D6>=75,\"B\",TRUE,\"C\")"
   },
   {
    "num": 4,
    "difficulty": "🔵",
    "desc": "同上用巢狀 IF",
    "hint": "IF 裡面再放 IF",
    "answer": "=IF(D6>=90,\"A\",IF(D6>=75,\"B\",\"C\"))"
   },
   {
    "num": 5,
    "difficulty": "🟡",
    "desc": "出勤>=22 且 績效>=80 → \"模範\"",
    "hint": "AND(條件1,條件2)",
    "answer": "=IF(AND(E6>=22,D6>=80),\"模範\",\"一般\")"
   },
   {
    "num": 6,
    "difficulty": "🟡",
    "desc": "部門是業務 或 績效>=90 → \"重點關注\"",
    "hint": "OR(條件1,條件2)",
    "answer": "=IF(OR(B6=\"業務\",D6>=90),\"重點關注\",\"一般\")"
   },
   {
    "num": 7,
    "difficulty": "🔴",
    "desc": "績效>=90 獎金=月薪×3，>=75=×2，其他=×1",
    "hint": "IFS + 數學運算",
    "answer": "=IFS(D6>=90,C6*3,D6>=75,C6*2,TRUE,C6*1)"
   },
   {
    "num": 8,
    "difficulty": "🔴",
    "desc": "用 SWITCH 把部門轉成英文代碼",
    "hint": "SWITCH(值, 比對1,結果1, ..., 預設)",
    "answer": "=SWITCH(B6,\"業務\",\"Sales\",\"財務\",\"Finance\",\"資訊\",\"IT\",\"人事\",\"HR\",\"其他\")"
   }
  ]
 }
};
export const DEMOS = [
 {
  "id": "ifs-grade-flow",
  "kind": "formula",
  "badge": "公式拆解",
  "title": "IFS：多個門檻要由高到低檢查",
  "subtitle": "IFS 會從左到右找第一個 TRUE。一旦命中，就回傳結果，不會再看後面的條件。",
  "outcome": "結果：85 >= 90 是 FALSE，85 >= 80 是 TRUE，所以回傳 B",
  "formulaParts": [
   {
    "text": "=IFS("
   },
   {
    "text": "B2>=90,\"A\"",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ", "
   },
   {
    "text": "B2>=80,\"B\"",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": ", "
   },
   {
    "text": "B2>=70,\"C\"",
    "step": 3,
    "tone": "c3"
   },
   {
    "text": ", "
   },
   {
    "text": "TRUE,\"F\"",
    "step": 4,
    "tone": "c4"
   },
   {
    "text": ")"
   }
  ],
  "board": {
   "title": "85 分的判斷流程",
   "columns": [
    "分數",
    ">=90",
    ">=80",
    ">=70",
    "結果"
   ],
   "rows": [
    [
     {
      "text": "85",
      "marks": {
       "1": "focus",
       "2": "focus",
       "3": "focus",
       "4": "focus"
      }
     },
     {
      "text": "FALSE",
      "marks": {
       "1": "result"
      }
     },
     {
      "text": "TRUE",
      "marks": {
       "2": "match"
      }
     },
     {
      "text": "跳過",
      "marks": {
       "3": "result"
      }
     },
     {
      "text": "B",
      "marks": {
       "2": "result",
       "3": "result",
       "4": "result"
      }
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. 先檢查最高門檻",
    "tone": "c1",
    "text": "85 沒有大於等於 90，所以第一組條件不成立。Excel 會繼續往右看下一組條件。",
    "caption": "門檻判斷通常要從高到低寫，避免高分先被低門檻吃掉。"
   },
   {
    "label": "2. 第二個條件命中",
    "tone": "c2",
    "text": "85 大於等於 80，這一組條件成立，所以 IFS 立即回傳 B。",
    "caption": "IFS 的重點是一命中就結束，不會把後面所有條件都跑完。"
   },
   {
    "label": "3. 後面的條件不再執行",
    "tone": "c3",
    "text": "雖然 85 也大於等於 70，但前面已經命中 B，這一段不會再影響結果。",
    "caption": "這就是為什麼條件順序會直接改變答案。"
   },
   {
    "label": "4. TRUE 是最後兜底",
    "tone": "c4",
    "text": "最後的 TRUE 等於 else。只有前面全部沒命中時，才會回傳這裡的 F。",
    "caption": "沒有兜底條件時，IFS 找不到符合條件會出錯。"
   }
  ]
 },
 {
  "id": "switch-department-map",
  "kind": "formula",
  "badge": "概念拆解",
  "title": "SWITCH：固定值對應，比一串 IF 更清楚",
  "subtitle": "當你不是判斷範圍，而是把一個固定值翻成另一個固定結果，SWITCH 會比 IFS 更乾淨。",
  "outcome": "結果：B2 是「財務」，命中第二組對應，所以回傳 Finance",
  "formulaParts": [
   {
    "text": "=SWITCH("
   },
   {
    "text": "B2",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ", "
   },
   {
    "text": "\"業務\",\"Sales\"",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": ", "
   },
   {
    "text": "\"財務\",\"Finance\"",
    "step": 3,
    "tone": "c3"
   },
   {
    "text": ", "
   },
   {
    "text": "\"資訊\",\"IT\"",
    "step": 4,
    "tone": "c4"
   },
   {
    "text": ", \"其他\")"
   }
  ],
  "board": {
   "title": "部門名稱對應",
   "columns": [
    "B2",
    "業務?",
    "財務?",
    "後續",
    "結果"
   ],
   "rows": [
    [
     {
      "text": "財務",
      "marks": {
       "1": "focus",
       "2": "focus",
       "3": "focus",
       "4": "focus"
      }
     },
     {
      "text": "否",
      "marks": {
       "2": "result"
      }
     },
     {
      "text": "是",
      "marks": {
       "3": "match"
      }
     },
     {
      "text": "跳過",
      "marks": {
       "4": "result"
      }
     },
     {
      "text": "Finance",
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
    "label": "1. 先拿 B2 當比對值",
    "tone": "c1",
    "text": "SWITCH 的第一個參數是要被拿來比對的值。這裡 B2 的內容是「財務」。",
    "caption": "和 IFS 不同，SWITCH 不是每組都寫一個完整條件。"
   },
   {
    "label": "2. 第一組對應沒有命中",
    "tone": "c2",
    "text": "Excel 先看 B2 是否等於「業務」。因為不是，所以不回傳 Sales，繼續看下一組。",
    "caption": "每一組都是「要比對的值、命中後的結果」。"
   },
   {
    "label": "3. 第二組對應命中",
    "tone": "c3",
    "text": "B2 等於「財務」，所以公式回傳 Finance。",
    "caption": "這類部門、狀態、代碼對照就是 SWITCH 的舒適區。"
   },
   {
    "label": "4. 後面對應與預設值不再影響結果",
    "tone": "c4",
    "text": "因為已經命中財務，後面的資訊與其他都不會再改變結果。",
    "caption": "如果 B2 是完全沒列出的值，最後的「其他」才會出場。"
   }
  ]
 }
];
