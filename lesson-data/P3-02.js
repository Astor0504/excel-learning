// 由 xlsx-content.js + lesson-demos-data.js 拆分(2026-07-14);本檔是 P3-02 的唯一資料來源
export const CONTENT = {
 "knowledge": {
  "title": "📝 文字與日期清理",
  "subtitle": "先清髒字元，再拆欄位，最後轉成可計算日期",
  "sections": [
   {
    "title": "📌 清資料順序",
    "items": [
     [
      "清空白",
      "SUBSTITUTE(A2,CHAR(160),\"\") → CLEAN → TRIM",
      "先處理非斷行空白和不可見字元"
     ],
     [
      "拆欄位",
      "LEFT / MID / RIGHT 搭配 FIND",
      "適合訂單號、Email、地址、電話"
     ],
     [
      "轉日期",
      "DATE / DATEVALUE 讓文字變成真日期",
      "真日期才能排序、計算天數、做月份群組"
     ]
    ]
   },
   {
    "title": "🧰 常用公式",
    "items": [
     [
      "穩定清空白",
      "=TRIM(CLEAN(SUBSTITUTE(A2,CHAR(160),\"\")))",
      "處理外部系統貼上資料"
     ],
     [
      "Email 取帳號",
      "=LEFT(A2,FIND(\"@\",A2)-1)",
      "FIND 先找位置，再交給 LEFT / MID"
     ],
     [
      "訂單號轉日期",
      "=DATE(MID(A2,2,4),MID(A2,6,2),MID(A2,8,2))",
      "比只用 TEXT 更適合後續計算"
     ]
    ]
   },
   {
    "title": "⚠️ 常見誤區",
    "items": [
     [
      "TEXT 會回傳文字",
      "適合輸出顯示，不適合當日期來源",
      "後續要計算時請保留真日期欄"
     ],
     [
      "DATEDIF 是隱藏函數",
      "不會自動提示，但可以使用",
      "Y / M / D 分別代表完整年、月、天"
     ],
     [
      "MID 從 1 開始",
      "不是程式語言常見的 0-based index",
      "拆固定編碼時最容易差一碼"
     ]
    ]
   }
  ],
  "handsTasks": [
   "清掉姓名欄多餘空白",
   "從 Email 拆出帳號與網域",
   "從訂單號轉出真日期",
   "用 DATEDIF / EOMONTH 算期間"
  ]
 },
 "pro": {
  "title": "📝 文字日期  ─  TEXT / MID / TRIM / DATEDIF / WORKDAY",
  "subtitle": "難度：🟡 Lv.3  |  微任務數：10 題  |  建議時間：每題 2~3 分鐘",
  "dataHeader": [
   "原始資料",
   "說明",
   "",
   "",
   "",
   "",
   ""
  ],
  "dataRows": [
   [
    "王小明",
    "姓名（有空格）",
    "",
    "",
    "",
    "",
    ""
   ],
   [
    "A2026031101",
    "訂單號碼",
    "",
    "",
    "",
    "",
    ""
   ],
   [
    "alice@company.com",
    "Email",
    "",
    "",
    "",
    "",
    ""
   ],
   [
    "新北市板橋區文化路一段",
    "地址",
    "",
    "",
    "",
    "",
    ""
   ],
   [
    1234567.89,
    "金額數字",
    "",
    "",
    "",
    "",
    ""
   ],
   [
    "2025-06-01",
    "起始日期",
    "",
    "",
    "",
    "",
    ""
   ],
   [
    "2026-03-31",
    "結束日期",
    "",
    "",
    "",
    "",
    ""
   ],
   [
    "2026-01-15",
    "專案啟動日",
    "",
    "",
    "",
    "",
    ""
   ]
  ],
  "tasks": [
   {
    "numLabel": "🟢1",
    "desc": "用清理三連去除姓名的多餘空白與不可見字元",
    "time": "1m",
    "hint": "=TRIM(CLEAN(SUBSTITUTE(A6,CHAR(160),\"\")))",
    "answer": "=TRIM(CLEAN(SUBSTITUTE(A6,CHAR(160),\"\")))",
    "explain": "先處理 CHAR(160)，再清不可見字元，最後 TRIM"
   },
   {
    "numLabel": "🟢2",
    "desc": "從訂單號取出年份（第2~5碼）",
    "time": "2m",
    "hint": "=MID(A7,2,4)",
    "answer": "=MID(A7,2,4)",
    "explain": "MID(字串,起始位,取幾碼)"
   },
   {
    "numLabel": "🔵3",
    "desc": "把訂單號轉成真正日期",
    "time": "3m",
    "hint": "=DATE(MID(A7,2,4),MID(A7,6,2),MID(A7,8,2))",
    "answer": "=DATE(MID(A7,2,4),MID(A7,6,2),MID(A7,8,2))",
    "explain": "DATE 會把年月日組成可計算的真日期"
   },
   {
    "numLabel": "🔵4",
    "desc": "取出 Email 的帳號（@之前）",
    "time": "3m",
    "hint": "=LEFT(A8,FIND(\"@\",A8)-1)",
    "answer": "=LEFT(A8,FIND(\"@\",A8)-1)",
    "explain": "FIND 找 @ 的位置，LEFT 取左邊的字"
   },
   {
    "numLabel": "🔵5",
    "desc": "取出地址的縣市（前3個字）",
    "time": "1m",
    "hint": "=LEFT(A9,3)",
    "answer": "=LEFT(A9,3)",
    "explain": "LEFT(字串,幾個字)"
   },
   {
    "numLabel": "🟡6",
    "desc": "金額加千分位格式化",
    "time": "2m",
    "hint": "=TEXT(A10,\"#,##0.00\")",
    "answer": "=TEXT(A10,\"#,##0.00\")",
    "explain": "TEXT 把數字變成指定格式的文字，後續計算請保留原數字欄"
   },
   {
    "numLabel": "🟡7",
    "desc": "計算起始到結束相差幾天",
    "time": "2m",
    "hint": "=DATEDIF(A11,A12,\"D\")",
    "answer": "=DATEDIF(A11,A12,\"D\")",
    "explain": "DATEDIF 第三參數：\"D\"天,\"M\"月,\"Y\"年"
   },
   {
    "numLabel": "🟡8",
    "desc": "計算相差幾個月",
    "time": "1m",
    "hint": "=DATEDIF(A11,A12,\"M\")",
    "answer": "=DATEDIF(A11,A12,\"M\")",
    "explain": "改成 M 就是月份差"
   },
   {
    "numLabel": "🔴9",
    "desc": "專案啟動日後 30 個工作天是哪天",
    "time": "2m",
    "hint": "=WORKDAY(A13,30)",
    "answer": "=WORKDAY(A13,30)",
    "explain": "WORKDAY 自動跳過週末"
   },
   {
    "numLabel": "🔴10",
    "desc": "取得專案啟動日的月底日期",
    "time": "2m",
    "hint": "=EOMONTH(A13,0)",
    "answer": "=EOMONTH(A13,0)",
    "explain": "EOMONTH(日期,0)=當月底，1=下月底"
   }
  ]
 },
 "meta": {
  "phase": "📐 Phase 3：專業打磨    第 5~6 週  |  目標：做出讓人覺得「很專業」的報表",
  "stage": "第 9 階段",
  "topics": "TRIM / CLEAN / MID / DATE / TEXT / DATEDIF",
  "difficulty": "🟡 Lv.3",
  "taskCount": "10 題",
  "time": "20 分鐘",
  "xp": "150 XP"
 },
 "inter": {
  "title": "✂️ 文字與日期處理",
  "subtitle": "在黃色邊框儲存格輸入公式 → 答對變綠 ✅ | 答錯變紅 ❌ | F欄提示選取可見",
  "dataHeader": [
   "姓名",
   "身分證字號",
   "入職日期",
   "Email",
   "電話",
   "",
   ""
  ],
  "dataRows": [
   [
    "王小明",
    "A123456789",
    "2020/3/15",
    "wang@abc.com",
    "0912-345-678",
    "",
    ""
   ],
   [
    "林美華",
    "B234567890",
    "2018/7/1",
    "lin@abc.com",
    "0923-456-789",
    "",
    ""
   ],
   [
    "張志偉",
    "C345678901",
    "2022/1/20",
    "chang@xyz.com",
    "0934-567-890",
    "",
    ""
   ],
   [
    "陳雅琪",
    "D456789012",
    "2019/11/5",
    "chen@xyz.com",
    "0945-678-901",
    "",
    ""
   ],
   [
    "李大明",
    "E567890123",
    "2021/6/10",
    "lee@abc.com",
    "0956-789-012",
    "",
    ""
   ]
  ],
  "tasks": [
   {
    "num": 1,
    "difficulty": "🟢",
    "desc": "取出王小明身分證前 1 碼（英文字母）",
    "hint": "LEFT(文字,字數)",
    "answer": "=LEFT(B6,1)"
   },
   {
    "num": 2,
    "difficulty": "🟢",
    "desc": "取出王小明 Email 的 @ 後面域名",
    "hint": "MID + FIND(\"@\",文字)",
    "answer": "=MID(D6,FIND(\"@\",D6)+1,100)"
   },
   {
    "num": 3,
    "difficulty": "🔵",
    "desc": "計算王小明的年資（到今天的完整年數）",
    "hint": "DATEDIF(開始,結束,\"Y\")",
    "answer": "=DATEDIF(C6,TODAY(),\"Y\")"
   },
   {
    "num": 4,
    "difficulty": "🔵",
    "desc": "把王小明的姓和名串接，中間加空格",
    "hint": "LEFT取姓 & \" \" & MID取名",
    "answer": "=LEFT(A6,1)&\" \"&MID(A6,2,10)"
   },
   {
    "num": 5,
    "difficulty": "🟡",
    "desc": "取出王小明入職日期的月份",
    "hint": "MONTH(日期)",
    "answer": "=MONTH(C6)"
   },
   {
    "num": 6,
    "difficulty": "🟡",
    "desc": "將王小明姓名全轉大寫（英文測試用 Email）",
    "hint": "UPPER(文字)",
    "answer": "=UPPER(D6)"
   },
   {
    "num": 7,
    "difficulty": "🔴",
    "desc": "計算王小明電話號碼中去掉 - 後的純數字長度",
    "hint": "LEN(SUBSTITUTE(文字,\"-\",\"\"))",
    "answer": "=LEN(SUBSTITUTE(E6,\"-\",\"\"))"
   },
   {
    "num": 8,
    "difficulty": "🔴",
    "desc": "用 TEXT 把王小明入職日期格式化為 \"2020年03月\"",
    "hint": "TEXT(日期,\"yyyy年mm月\")",
    "answer": "=TEXT(C6,\"yyyy年mm月\")"
   }
  ]
 }
};
export const DEMOS = [
 {
  "id": "text-cleaning-stack",
  "kind": "formula",
  "badge": "清理順序",
  "title": "TRIM + CLEAN + SUBSTITUTE：先處理看不見的髒資料",
  "subtitle": "外部系統匯出的空白不一定是一般空白。先換掉 CHAR(160)，再清不可見字元，最後 TRIM。",
  "outcome": "結果：看起來正常但混亂的姓名，變成可比對、可查找的乾淨文字",
  "formulaParts": [
   {
    "text": "=TRIM("
   },
   {
    "text": "CLEAN(",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": "SUBSTITUTE("
   },
   {
    "text": "A2",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ", CHAR(160), \"\")",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ")",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": ")",
    "step": 3,
    "tone": "c3"
   },
   {
    "text": ")"
   }
  ],
  "board": {
   "title": "貼上來的姓名欄",
   "columns": [
    "原始值",
    "問題",
    "清理後"
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
      "text": "前後空白",
      "marks": {
       "3": "focus"
      }
     },
     {
      "text": "王小明",
      "marks": {
       "3": "result"
      }
     }
    ],
    [
     {
      "text": "林 美華",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "多個空格",
      "marks": {
       "3": "focus"
      }
     },
     {
      "text": "林 美華",
      "marks": {
       "3": "result"
      }
     }
    ],
    [
     {
      "text": "張志偉",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "CHAR(160)",
      "marks": {
       "1": "match"
      }
     },
     {
      "text": "張志偉",
      "marks": {
       "3": "result"
      }
     }
    ],
    [
     {
      "text": "陳雅琪",
      "marks": {
       "2": "focus"
      }
     },
     {
      "text": "不可見字元",
      "marks": {
       "2": "match"
      }
     },
     {
      "text": "陳雅琪",
      "marks": {
       "3": "result"
      }
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. 先處理非斷行空白 CHAR(160)",
    "tone": "c1",
    "text": "網頁或系統貼上來的空白常常不是一般空白，而是 CHAR(160)。TRIM 不一定能處理它，所以先用 SUBSTITUTE 換掉。",
    "caption": "如果 VLOOKUP 明明看起來一樣卻查不到，常常就是這種隱藏字元。"
   },
   {
    "label": "2. CLEAN 移除不可見控制字元",
    "tone": "c2",
    "text": "CLEAN 會處理換行、控制字元等肉眼看不出的髒資料，讓文字更穩定。",
    "caption": "它不會修正所有問題，但常是清資料流程裡很值得加的一層。"
   },
   {
    "label": "3. 最後才用 TRIM 整理空白",
    "tone": "c3",
    "text": "TRIM 會移除前後空白，並把連續空白壓成單一空白。放在最後，能把前面清完後剩下的空白整理乾淨。",
    "caption": "清資料要有順序：先換掉怪空白，再清不可見字元，最後整理空白。"
   }
  ]
 },
 {
  "id": "order-id-to-date",
  "kind": "formula",
  "badge": "日期轉換",
  "title": "DATE + MID：從訂單號拆出真正日期",
  "subtitle": "不要只用 TEXT 做顯示格式。若後面還要排序或計算天數，要把文字轉成真正日期。",
  "outcome": "結果：A2026031101 會被轉成 2026/03/11，可排序、分月與計算天數",
  "formulaParts": [
   {
    "text": "=DATE("
   },
   {
    "text": "MID(A2,2,4)",
    "step": 1,
    "tone": "c1"
   },
   {
    "text": ", "
   },
   {
    "text": "MID(A2,6,2)",
    "step": 2,
    "tone": "c2"
   },
   {
    "text": ", "
   },
   {
    "text": "MID(A2,8,2)",
    "step": 3,
    "tone": "c3"
   },
   {
    "text": ")"
   }
  ],
  "board": {
   "title": "訂單號碼 A2026031101",
   "columns": [
    "片段",
    "公式",
    "結果"
   ],
   "rows": [
    [
     {
      "text": "年份",
      "marks": {
       "1": "focus"
      }
     },
     {
      "text": "MID(A2,2,4)",
      "marks": {
       "1": "match"
      }
     },
     {
      "text": "2026",
      "marks": {
       "1": "result"
      }
     }
    ],
    [
     {
      "text": "月份",
      "marks": {
       "2": "focus"
      }
     },
     {
      "text": "MID(A2,6,2)",
      "marks": {
       "2": "match"
      }
     },
     {
      "text": "03",
      "marks": {
       "2": "result"
      }
     }
    ],
    [
     {
      "text": "日期",
      "marks": {
       "3": "focus"
      }
     },
     {
      "text": "MID(A2,8,2)",
      "marks": {
       "3": "match"
      }
     },
     {
      "text": "11",
      "marks": {
       "3": "result"
      }
     }
    ],
    [
     {
      "text": "真日期",
      "marks": {
       "4": "focus"
      }
     },
     {
      "text": "DATE(2026,03,11)",
      "marks": {
       "4": "match"
      }
     },
     {
      "text": "2026/03/11",
      "marks": {
       "4": "result"
      }
     }
    ]
   ]
  },
  "steps": [
   {
    "label": "1. 從第 2 碼開始取 4 碼作為年份",
    "tone": "c1",
    "text": "訂單號第一碼是字母 A，所以年份從第 2 碼開始。MID(A2,2,4) 會取出 2026。",
    "caption": "MID 的位置從 1 開始算，不是從 0。"
   },
   {
    "label": "2. 再取第 6 到第 7 碼作為月份",
    "tone": "c2",
    "text": "MID(A2,6,2) 會取出 03。DATE 可以接受這種文字型數字並轉成日期的一部分。",
    "caption": "月份如果是 03，不要先用 TEXT 格式化，先轉真日期。"
   },
   {
    "label": "3. 最後取日期並交給 DATE",
    "tone": "c3",
    "text": "MID(A2,8,2) 取出 11。DATE(年,月,日) 會把三段合成真正日期。",
    "caption": "真正日期才能排序、計算天數、接樞紐日期群組。"
   },
   {
    "label": "4. TEXT 是顯示格式，不是日期本身",
    "tone": "c4",
    "text": "如果只是要顯示成 2026-03-11，可以再用 TEXT；但請保留原本的真日期欄位供後續計算。",
    "caption": "這是很多報表日期軸怪掉的原因。"
   }
  ]
 }
];
