// Excel 課程知識點資料（手動維護）。原本由 build_xlsx_content.py 從 xlsx 產生，xlsx 已停用；直接編輯本檔即可。
export const XLSX_CONTENT = {
 "lessons": {
  "P1-01": {
   "shortcuts": {
    "title": "⌨️ 快捷鍵大全（macOS 版）",
    "subtitle": "難度：🟢 Lv.1  |  ⌘=Command  ⌥=Option  ⌃=Control  ⇧=Shift  fn=Function",
    "intro": "⚠️ 重要：先開啟 KeyTips！Excel → 偏好設定 → 輔助使用 → KeyTips → 啟動鍵選「Option (⌥)」。之後按 ⌥ 就能呼叫 Ribbon 快捷鍵（如同 Windows 的 Alt）",
    "groups": [
     {
      "title": "🏃 導航（移動光標）",
      "items": [
       {
        "key": "⌘ + ←/→/↑/↓",
        "desc": "跳到資料邊界",
        "note": "同 Windows 的 Ctrl+方向鍵功能"
       },
       {
        "key": "⌘ + Home (fn+←)",
        "desc": "跳到 A1",
        "note": "Mac 鍵盤用 fn+← 代替 Home"
       },
       {
        "key": "⌘ + End (fn+→)",
        "desc": "跳到有資料的最右下角",
        "note": "fn+→ 代替 End"
       },
       {
        "key": "⌃ + G 或 fn+F5",
        "desc": "到指定儲存格",
        "note": "同 Windows 的 F5/Ctrl+G 功能"
       },
       {
        "key": "⌃ + Page Up/Down",
        "desc": "切換工作表",
        "note": "fn+↑/fn+↓ 代替 Page Up/Down"
       },
       {
        "key": "⌘ + F",
        "desc": "搜尋",
        "note": "⌘+H = 取代"
       }
      ]
     },
     {
      "title": "✂️ 選取",
      "items": [
       {
        "key": "⌘ + ⇧ + ←/→/↑/↓",
        "desc": "跳選到資料邊界",
        "note": "最常用的選取方式"
       },
       {
        "key": "⌘ + ⇧ + End",
        "desc": "選取到有資料的右下角",
        "note": "快速選取整個資料區"
       },
       {
        "key": "⌃ + Space",
        "desc": "選取整欄",
        "note": "跟 Windows 一樣"
       },
       {
        "key": "⇧ + Space",
        "desc": "選取整列",
        "note": "跟 Windows 一樣"
       },
       {
        "key": "⌘ + A",
        "desc": "選取目前區域/全選",
        "note": "按一次=區域，再按=全部"
       },
       {
        "key": "⌘ + ⇧ + L",
        "desc": "開啟/關閉自動篩選",
        "note": "跟 Windows 一樣"
       }
      ]
     },
     {
      "title": "📋 編輯",
      "items": [
       {
        "key": "⌘ + D",
        "desc": "向下填滿",
        "note": "複製上方儲存格內容"
       },
       {
        "key": "⌘ + R",
        "desc": "向右填滿",
        "note": "複製左方儲存格內容"
       },
       {
        "key": "⌃ + ;",
        "desc": "插入今天日期（靜態）",
        "note": "不會自動更新，是固定值"
       },
       {
        "key": "⌘ + ;",
        "desc": "插入現在時間（靜態）",
        "note": "⚠️ 跟 Windows 不同！Windows 是 Ctrl+Shift+;"
       },
       {
        "key": "⌃ + ⌘ + Return",
        "desc": "同時在多格輸入相同內容",
        "note": "Windows 是 Ctrl+Enter"
       },
       {
        "key": "⌃ + ⌥ + Return",
        "desc": "儲存格內換行",
        "note": "Windows 是 Alt+Enter"
       },
       {
        "key": "⌃ + U 或 F2",
        "desc": "進入編輯模式",
        "note": "不用雙擊就能編輯"
       },
       {
        "key": "⌘ + T",
        "desc": "切換絕對/相對參照",
        "note": "⚠️ Mac 用 ⌘+T！Windows 是 F4"
       },
       {
        "key": "⌃ + `",
        "desc": "顯示/隱藏所有公式",
        "note": "反引號，在 Esc 下方"
       }
      ]
     },
     {
      "title": "⚡ 效率神技",
      "items": [
       {
        "key": "⌘ + T（選取資料時）",
        "desc": "轉換為表格",
        "note": "⚠️ 注意：⌘+T 在公式中=切換參照，在資料上=轉表格"
       },
       {
        "key": "⌘ + ⇧ + T",
        "desc": "自動 SUM",
        "note": "Windows 是 Alt+=，Mac 不同"
       },
       {
        "key": "⌃ + ⇧ + +",
        "desc": "插入列/欄",
        "note": "或 ⌘ + ⇧ + +"
       },
       {
        "key": "⌘ + -",
        "desc": "刪除列/欄",
        "note": "先選取列/欄再按"
       },
       {
        "key": "⌥ + fn + F1",
        "desc": "一鍵插入圖表",
        "note": "Windows 是 Alt+F1"
       },
       {
        "key": "⌘ + 1",
        "desc": "格式化儲存格對話框",
        "note": "最常用的格式化入口"
       },
       {
        "key": "⌘ + E",
        "desc": "快速填入（Flash Fill）",
        "note": "自動辨識模式填充，超好用"
       }
      ]
     },
     {
      "title": "🏆 高手必會",
      "items": [
       {
        "key": "⌥（按一下放開）",
        "desc": "啟動 KeyTips（需先開啟）",
        "note": "類似 Windows 的 Alt 鍵，259+ 快捷鍵"
       },
       {
        "key": "⌥ + N + V",
        "desc": "插入樞紐分析表（KeyTips）",
        "note": "跟 Windows Alt+N+V 一樣"
       },
       {
        "key": "⌘ + [",
        "desc": "追蹤公式的來源儲存格",
        "note": "按 ⌘+] 追蹤相依儲存格"
       },
       {
        "key": "⌃ + ⇧ + U",
        "desc": "展開/收合公式欄",
        "note": "看長公式時很方便"
       },
       {
        "key": "⌘ + ⇧ + F3",
        "desc": "依選取範圍建立命名",
        "note": "快速建立多個命名範圍"
       },
       {
        "key": "⌥ + H + O + I",
        "desc": "自動調整列高（KeyTips）",
        "note": "需要先啟用 KeyTips"
       }
      ]
     },
     {
      "title": "⚠️ macOS 特殊注意事項",
      "items": [
       {
        "key": "VBA 編輯器",
        "desc": "工具 → 巨集 → Visual Basic 編輯器",
        "note": "⚠️ Mac 沒有 Alt+F11！用選單開啟"
       },
       {
        "key": "UserForm",
        "desc": "Mac Excel 不支援 UserForm",
        "note": "⚠️ 需要替代方案（如 InputBox 或 對話框）"
       },
       {
        "key": "F 鍵問題",
        "desc": "確認系統偏好設定 → 鍵盤 → 「使用 F1/F2 做標準功能鍵」",
        "note": "否則 F2 會調螢幕亮度而非編輯模式"
       },
       {
        "key": "Power Query",
        "desc": "Mac 版可用但資料來源較少",
        "note": "不支援：Access、SQL Server 直連、部分 Web"
       },
       {
        "key": "KeyTips 限制",
        "desc": "Esc 會直接退出（不像 Windows 返回上一層）",
        "note": "Mac KeyTips 2024 年底才推出，仍在改善中"
       }
      ]
     },
     {
      "title": "⌨️ 全鍵盤操作訓練法（進階）",
      "items": []
     },
     {
      "title": "🍎 macOS 可達成度：80~90% 無滑鼠操作（2024 年 KeyTips 功能推出後大幅提升）",
      "items": []
     },
     {
      "title": "✅ 可 100% 鍵盤操作的任務（約佔日常 80~90%）",
      "items": [
       {
        "key": "導航與選取",
        "desc": "",
        "note": "⌘+方向鍵跳邊界、⇧+選取、⌘+⇧+方向鍵跳選到底"
       },
       {
        "key": "儲存格編輯",
        "desc": "",
        "note": "輸入/刪除/複製貼上、⌘+D 向下填滿、⌘+R 向右填滿"
       },
       {
        "key": "公式建立",
        "desc": "",
        "note": "所有公式輸入、⌘+T 切換絕對參照、Tab 自動完成函式名"
       },
       {
        "key": "表格操作",
        "desc": "",
        "note": "⌘+T 建立表格、⌘+⇧+L 開關篩選、排序（KeyTips 觸發）"
       },
       {
        "key": "格式化",
        "desc": "",
        "note": "⌘+1 格式對話框、⌘+B/I/U 粗體斜體底線、KeyTips 操作 Ribbon"
       },
       {
        "key": "工作表管理",
        "desc": "",
        "note": "⌃+Page Up/Down 切換工作表、⌥+⇧+→ 群組/取消群組"
       },
       {
        "key": "檔案操作",
        "desc": "",
        "note": "⌘+S 儲存、⌘+N 新建、⌘+O 開啟、⌘+P 列印"
       },
       {
        "key": "VBA 編輯",
        "desc": "",
        "note": "Tools → Macro → VBE 開啟編輯器（無快捷鍵但可 KeyTips 觸發）"
       }
      ]
     },
     {
      "title": "⚠️ 仍需要滑鼠的操作（約佔 10~20%）",
      "items": [
       {
        "key": "樞紐分析表",
        "desc": "",
        "note": "將欄位拖拉到列/欄/值/篩選四大區域，目前無法純鍵盤完成"
       },
       {
        "key": "圖表微調",
        "desc": "",
        "note": "資料標籤位置、趨勢線設定、圖表元素的精確拖曳"
       },
       {
        "key": "條件式格式化",
        "desc": "",
        "note": "自訂規則對話框中的色彩選取器"
       },
       {
        "key": "交叉分析篩選器",
        "desc": "",
        "note": "多選按鈕的點選操作"
       },
       {
        "key": "Power Query",
        "desc": "",
        "note": "編輯器中部分進階 UI 操作（但基本操作可用鍵盤）"
       },
       {
        "key": "繪圖/SmartArt",
        "desc": "",
        "note": "形狀繪製、SmartArt 圖形調整"
       }
      ]
     },
     {
      "title": "🧠 練習計畫",
      "items": [
       {
        "key": "第 1 週",
        "desc": "",
        "note": "⌘+方向鍵導航 + ⇧ 選取 + ⌘+C/V 複製貼上（每天 3 個）"
       },
       {
        "key": "第 2 週",
        "desc": "",
        "note": "⌘+T 表格 + ⌘+⇧+T 自動SUM + ⌘+1 格式化 + ⌘+⇧+L 篩選"
       },
       {
        "key": "第 3 週",
        "desc": "",
        "note": "⌃+; 日期 + ⌘+E Flash Fill + ⌃+` 顯示公式 + ⌘+D/R 填滿"
       },
       {
        "key": "第 4 週",
        "desc": "",
        "note": "KeyTips（⌥）Ribbon 全鍵盤操作 + Tab 切換對話框按鈕"
       },
       {
        "key": "Tips",
        "desc": "",
        "note": "每天挑 3 個快捷鍵貼在螢幕旁，強迫自己用一天 → 兩週變肌肉記憶"
       }
      ]
     },
     {
      "title": "📊 WPS Office vs Microsoft Excel 差異對照表（macOS）",
      "items": []
     },
     {
      "title": "建議以 Microsoft Excel 為主力學習工具。進階功能（KeyTips / VBA / Power Query / Flash Fill）請用 Excel。",
      "items": [
       {
        "key": "功能",
        "desc": "Microsoft Excel macOS",
        "note": "WPS macOS"
       },
       {
        "key": "KeyTips (Ribbon 鍵盤導航)",
        "desc": "⌥ 按住啟動（需在偏好設定開啟）",
        "note": "macOS 版未確認支援"
       },
       {
        "key": "基本快捷鍵 ⌘+C/V/X/Z/S",
        "desc": "完整支援",
        "note": "完整支援，與 Excel 相同"
       },
       {
        "key": "⌘+T 建立表格",
        "desc": "⌘+T",
        "note": "⌘+T（大致相同）"
       },
       {
        "key": "切換絕對參照（$）",
        "desc": "⌘+T（公式編輯中）",
        "note": "F4 或 Fn+F4（較接近 Windows）"
       },
       {
        "key": "自動 SUM",
        "desc": "⌘+⇧+T",
        "note": "可嘗試 ⌥+= 或手動輸入"
       },
       {
        "key": "Flash Fill 快速填入",
        "desc": "⌘+E（365 版）",
        "note": "不支援"
       },
       {
        "key": "⌃+; 插入日期",
        "desc": "⌃+;",
        "note": "⌃+;（相同）"
       },
       {
        "key": "⌃+` 顯示公式",
        "desc": "⌃+`",
        "note": "⌃+`（相同）"
       },
       {
        "key": "⌘+1 格式化對話框",
        "desc": "⌘+1",
        "note": "⌘+1"
       },
       {
        "key": "⌘+⇧+L 篩選",
        "desc": "⌘+⇧+L",
        "note": "⌘+⇧+L（大致相同）"
       },
       {
        "key": "VBA / 巨集",
        "desc": "支援（macOS 限制：無 UserForm）",
        "note": "支援更低，部分物件不可用"
       },
       {
        "key": "Power Query",
        "desc": "支援（資料來源較少）",
        "note": "不支援"
       },
       {
        "key": "樞紐分析表",
        "desc": "完整支援",
        "note": "支援基本功能"
       },
       {
        "key": "條件式格式化",
        "desc": "完整支援",
        "note": "支援基本規則"
       },
       {
        "key": "自訂快捷鍵",
        "desc": "有限",
        "note": "選項→自訂功能區→自訂"
       }
      ]
     },
     {
      "title": "✅ 相同 = 行為一致    ⚠️ 差異 = 可用但方式不同    ❌ 不支援 = WPS 無此功能",
      "items": []
     }
    ]
   }
  },
  "P1-02": {
   "pro": {
    "title": "📊 基礎函式  ─  SUM / AVERAGE / COUNT / MAX / MIN",
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
      "desc": "計算王小明 6 個月總業績",
      "time": "2m",
      "hint": "=SUM(B6:G6)",
      "answer": "=SUM(B6:G6)",
      "explain": "SUM 也可以橫向加總"
     },
     {
      "numLabel": "🔴7",
      "desc": "用 COUNT 計算 B 欄有幾個數字",
      "time": "2m",
      "hint": "=COUNT(B6:B10)",
      "answer": "=COUNT(B6:B10)",
      "explain": "COUNT 只數數字，COUNTA 數所有非空白"
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
    "topics": "SUM / AVERAGE / COUNT / MAX / MIN",
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
      "desc": "計算王小明 6 個月總業績",
      "hint": "SUM 也可以橫向",
      "answer": "=SUM(B6:G6)"
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
  },
  "P1-03": {
   "pro": {
    "title": "📐 條件判斷  ─  IF / IFS / AND / OR / 巢狀 IF",
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
      "desc": "判斷績效是否「偶數」（用 MOD）",
      "time": "2m",
      "hint": "=IF(MOD(D6,2)=0,\"偶數\",\"奇數\")",
      "answer": "=IF(MOD(D6,2)=0,\"偶數\",\"奇數\")",
      "explain": "MOD(數字,2) 回傳除以2的餘數"
     }
    ]
   },
   "meta": {
    "phase": "🏃 Phase 1：操作效率基礎    第 1~2 週  |  目標：養成鍵盤操作習慣，複習基本功",
    "stage": "第 3 階段",
    "topics": "IF / IFS / 巢狀 IF / AND / OR",
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
      "desc": "績效是偶數→\"偶數\"，否則→\"奇數\"",
      "hint": "MOD(數字,2) 餘數",
      "answer": "=IF(MOD(D6,2)=0,\"偶數\",\"奇數\")"
     }
    ]
   }
  },
  "P2-01": {
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
   "inter": {
    "title": "📈 條件統計 — COUNTIF / SUMIF / AVERAGEIF",
    "subtitle": "在黃色邊框儲存格輸入公式 → 答對變綠 ✅ | 答錯變紅 ❌ | F欄提示選取可見",
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
  },
  "P2-02": {
   "pro": {
    "title": "🔍 查找比對  ─  VLOOKUP / XLOOKUP / INDEX+MATCH",
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
      "desc": "用 VLOOKUP 查 P003 的產品名稱",
      "time": "2m",
      "hint": "=VLOOKUP(\"P003\",A6:E11,2,0)",
      "answer": "=VLOOKUP(\"P003\",A6:E11,2,0)",
      "explain": "VLOOKUP(查找值,表格,第幾欄,0=精確)"
     },
     {
      "numLabel": "🟢2",
      "desc": "用 VLOOKUP 查 P005 的單價",
      "time": "2m",
      "hint": "=VLOOKUP(\"P005\",A6:E11,4,0)",
      "answer": "=VLOOKUP(\"P005\",A6:E11,4,0)",
      "explain": "第 4 欄 = 單價欄位"
     },
     {
      "numLabel": "🔵3",
      "desc": "用 INDEX+MATCH 查 P003 名稱",
      "time": "3m",
      "hint": "=INDEX(B6:B11,MATCH(\"P003\",A6:A11,0))",
      "answer": "=INDEX(B6:B11,MATCH(\"P003\",A6:A11,0))",
      "explain": "MATCH 找位置，INDEX 取值，比 VLOOKUP 更靈活"
     },
     {
      "numLabel": "🔵4",
      "desc": "用 INDEX+MATCH 查 P001 庫存",
      "time": "2m",
      "hint": "=INDEX(E6:E11,MATCH(\"P001\",A6:A11,0))",
      "answer": "=INDEX(E6:E11,MATCH(\"P001\",A6:A11,0))",
      "explain": "換成 E 欄即可查庫存"
     },
     {
      "numLabel": "🔵5",
      "desc": "反向查找：已知名稱「機械鍵盤」查代碼",
      "time": "3m",
      "hint": "=INDEX(A6:A11,MATCH(\"機械鍵盤\",B6:B11,0))",
      "answer": "=INDEX(A6:A11,MATCH(\"機械鍵盤\",B6:B11,0))",
      "explain": "VLOOKUP 只能往右查，INDEX+MATCH 可以往左"
     },
     {
      "numLabel": "🟡6",
      "desc": "XLOOKUP 查 P003 名稱（365 版）",
      "time": "2m",
      "hint": "=XLOOKUP(\"P003\",A6:A11,B6:B11)",
      "answer": "=XLOOKUP(\"P003\",A6:A11,B6:B11)",
      "explain": "XLOOKUP 最簡潔，但僅 365 版支援"
     },
     {
      "numLabel": "🟡7",
      "desc": "如果查不到代碼 P999，顯示「無此商品」",
      "time": "2m",
      "hint": "=IFERROR(VLOOKUP(\"P999\",A6:E11,2,0),\"無此商品\")",
      "answer": "=IFERROR(VLOOKUP(\"P999\",A6:E11,2,0),\"無此商品\")",
      "explain": "IFERROR 包住可能出錯的公式"
     },
     {
      "numLabel": "🟡8",
      "desc": "用 MATCH 找出「27吋螢幕」在第幾列",
      "time": "2m",
      "hint": "=MATCH(\"27吋螢幕\",B6:B11,0)",
      "answer": "=MATCH(\"27吋螢幕\",B6:B11,0)",
      "explain": "MATCH 回傳的是「相對位置」而非列號"
     },
     {
      "numLabel": "🔴9",
      "desc": "查 P002 的名稱+單價（一次兩欄）",
      "time": "3m",
      "hint": "=VLOOKUP(\"P002\",A6:E11,{2,4},0)",
      "answer": "=VLOOKUP(\"P002\",A6:E11,{2,4},0)",
      "explain": "陣列常數 {2,4} 可一次取多欄（需 ⌘+⇧+Return（舊版需要，365 自動））"
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
    "topics": "VLOOKUP / XLOOKUP / INDEX+MATCH",
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
      "desc": "VLOOKUP 查 P003 的產品名稱",
      "hint": "VLOOKUP(查找值,表格,第幾欄,0)",
      "answer": "=VLOOKUP(\"P003\",A6:E11,2,0)"
     },
     {
      "num": 2,
      "difficulty": "🟢",
      "desc": "VLOOKUP 查 P005 的單價",
      "hint": "第 4 欄=單價",
      "answer": "=VLOOKUP(\"P005\",A6:E11,4,0)"
     },
     {
      "num": 3,
      "difficulty": "🔵",
      "desc": "INDEX+MATCH 查 P003 名稱",
      "hint": "MATCH 找位置，INDEX 取值",
      "answer": "=INDEX(B6:B11,MATCH(\"P003\",A6:A11,0))"
     },
     {
      "num": 4,
      "difficulty": "🔵",
      "desc": "XLOOKUP 查 P003 名稱（365 版）",
      "hint": "XLOOKUP(查,查找欄,回傳欄)",
      "answer": "=XLOOKUP(\"P003\",A6:A11,B6:B11)"
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
      "desc": "查不到 P999 時顯示「無此商品」",
      "hint": "IFERROR 包住公式",
      "answer": "=IFERROR(VLOOKUP(\"P999\",A6:E11,2,0),\"無此商品\")"
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
      "desc": "XLOOKUP 查 P001，找不到回傳「N/A」",
      "hint": "XLOOKUP 第 4 參數=找不到時",
      "answer": "=XLOOKUP(\"P001\",A6:A11,B6:B11,\"N/A\")"
     }
    ]
   }
  },
  "P2-03": {
   "knowledge": {
    "title": "📊 樞紐分析表",
    "subtitle": "難度：🔵 Lv.2  |  資料分析的核心武器  |  面試必考",
    "sections": [
     {
      "title": "📌 建立步驟",
      "items": [
       [
        "Step 1",
        "選取資料範圍（建議先轉為表格 ⌘+T）"
       ],
       [
        "Step 2",
        "插入 → 樞紐分析表 → 選擇放置位置"
       ],
       [
        "Step 3",
        "拖拉欄位到四個區域：列/欄/值/篩選"
       ],
       [
        "Step 4",
        "值欄位設定：加總、計數、平均等"
       ]
      ]
     },
     {
      "title": "🔧 四大區域",
      "items": [
       [
        "列（Rows）",
        "放分類欄位（如：部門、產品類別）",
        "決定表格「垂直方向」怎麼分組",
        "通常放最主要的分析維度"
       ],
       [
        "欄（Columns）",
        "放第二層分類（如：月份、年度）",
        "決定「水平方向」怎麼展開",
        "不要放太多，超過 5 個會太擠"
       ],
       [
        "值（Values）",
        "放數值欄位（如：金額、數量）",
        "可選擇加總/計數/平均/最大/最小",
        "右鍵值 → 值欄位設定 → 改計算方式"
       ],
       [
        "篩選（Filters）",
        "放全域篩選條件",
        "出現在樞紐表上方的下拉選單",
        "適合放年度、部門等大範圍篩選"
       ]
      ]
     },
     {
      "title": "⚡ 進階技巧",
      "items": [
       [
        "群組化日期",
        "右鍵日期欄位 → 群組 → 選月/季/年",
        "把逐日資料自動彙整為月或季報",
        "最常被忽略但最實用的功能"
       ],
       [
        "計算欄位",
        "分析 → 欄位/項目/集 → 計算欄位",
        "在樞紐表內新增公式欄（如利潤率）",
        "=營收-成本 直接在樞紐內算"
       ],
       [
        "顯示值為百分比",
        "右鍵值 → 值顯示方式 → 欄總計百分比",
        "看每個項目佔總體的比例",
        "老闆最愛看的格式"
       ],
       [
        "交叉分析篩選器",
        "插入 → 交叉分析篩選器",
        "漂亮的按鈕式篩選，適合做儀表板",
        "多個樞紐表可以共用同一個篩選器"
       ],
       [
        "資料模型",
        "勾選「新增至資料模型」",
        "可以跨表建立關聯，不用 VLOOKUP",
        "Power Pivot 的入口"
       ]
      ]
     },
     {
      "title": "❌ 常見錯誤",
      "items": [
       [
        "資料有空列",
        "樞紐表會在空列處截斷 → 清除空列再建"
       ],
       [
        "欄位名稱重複",
        "每欄必須有唯一的標題名稱"
       ],
       [
        "忘記重新整理",
        "來源資料更新後要按「重新整理」按鈕"
       ],
       [
        "混合資料型態",
        "同一欄不要混數字和文字，否則計算會出錯"
       ]
      ]
     }
    ],
    "handsTasks": [
     {
      "num": 1,
      "difficulty": "🟢 暖身",
      "desc": "在「📈 條件統計」工作表的訂單資料上，按 ⌘+T 轉為表格，再插入樞紐分析表"
     },
     {
      "num": 2,
      "difficulty": "🟢 暖身",
      "desc": "把「業務員」拖到列區域，「金額」拖到值區域 → 看到各業務員的總金額"
     },
     {
      "num": 3,
      "difficulty": "🔵 標準",
      "desc": "加上「區域」到欄區域 → 變成交叉表：每位業務員在各區域的業績"
     },
     {
      "num": 4,
      "difficulty": "🔵 標準",
      "desc": "在值欄位設定中，把「加總」改成「計數」→ 看出每人的訂單筆數"
     },
     {
      "num": 5,
      "difficulty": "🟡 變化",
      "desc": "對日期欄位右鍵 → 群組 → 依月份群組 → 看到月度彙總"
     },
     {
      "num": 6,
      "difficulty": "🟡 變化",
      "desc": "新增計算欄位：平均單價 = 金額 / 計數。觀察哪位業務員的單價最高"
     },
     {
      "num": 7,
      "difficulty": "🔴 挑戰",
      "desc": "插入交叉分析篩選器，用「付款狀態」篩選 → 對比已付/未付的業績分佈"
     },
     {
      "num": 8,
      "difficulty": "🔴 挑戰",
      "desc": "建立第二個樞紐表：產品類別 × 區域的金額矩陣，並讓兩個樞紐共用同一個篩選器"
     }
    ]
   },
   "meta": {
    "phase": "💼 Phase 2：工作即戰力    第 3~4 週  |  目標：老闆交辦的任務都能搞定",
    "stage": "第 6 階段",
    "topics": "建立/四大區域/日期群組/計算欄位",
    "difficulty": "🔵 Lv.2",
    "taskCount": "8 題",
    "time": "25 分鐘",
    "xp": "200 XP"
   }
  },
  "P2-04": {
   "knowledge": {
    "title": "🎨 條件式格式化",
    "subtitle": "難度：🔵 Lv.2  |  讓數字自己說故事  |  報表視覺化必備",
    "sections": [
     {
      "title": "📌 基本操作路徑",
      "items": [
       [
        "醒目提示規則",
        "常用 → 條件式格式化 → 醒目提示儲存格規則",
        "大於/小於/等於/介於/包含文字",
        "最快上手的功能"
       ],
       [
        "頂端/底端規則",
        "常用 → 條件式格式化 → 頂端/底端規則",
        "前 10 名、後 10%、高於平均",
        "快速找出前段班/後段班"
       ],
       [
        "資料橫條",
        "常用 → 條件式格式化 → 資料橫條",
        "在儲存格內顯示橫條圖",
        "一眼看出大小差異"
       ],
       [
        "色階",
        "常用 → 條件式格式化 → 色階",
        "用漸層色彩表示高低",
        "熱力圖效果，適合矩陣型資料"
       ],
       [
        "圖示集",
        "常用 → 條件式格式化 → 圖示集",
        "用箭頭/星星/旗幟表示狀態",
        "紅黃綠燈號最常用"
       ]
      ]
     },
     {
      "title": "🔧 自訂公式規則（進階）",
      "items": [
       [
        "整列標色",
        "公式：=$E2>50000（注意 $ 鎖欄不鎖列）",
        "套用到 $A$2:$H$100",
        "$ 放在欄前面 = 鎖住參考欄"
       ],
       [
        "隔列變色",
        "公式：=MOD(ROW(),2)=0",
        "偶數列自動上色",
        "最簡單的斑馬紋效果"
       ],
       [
        "標記重複",
        "公式：=COUNTIF($A:$A,$A2)>1",
        "重複值自動標紅",
        "資料清洗必用"
       ],
       [
        "到期提醒",
        "公式：=($D2-TODAY())<7",
        "7 天內到期的標紅",
        "專案管理常用"
       ],
       [
        "空白警告",
        "公式：=ISBLANK($C2)",
        "必填欄位為空時標黃",
        "表單驗證視覺化"
       ]
      ]
     },
     {
      "title": "💡 專業技巧",
      "items": [
       [
        "規則優先順序",
        "常用 → 條件式格式化 → 管理規則 → 調整順序",
        "上面的規則優先"
       ],
       [
        "停止為 True",
        "勾選「如果為 True 則停止」",
        "避免多條規則互相衝突"
       ],
       [
        "複製格式",
        "格式刷（或 ⌘+C → 選擇性貼上 → 格式）",
        "把條件格式套到其他範圍"
       ]
      ]
     }
    ],
    "handsTasks": [
     {
      "num": 1,
      "difficulty": "🟢 暖身",
      "desc": "在「📊 基礎函式」的業績資料上，選取整個月份欄 → 條件式格式化 → 資料橫條"
     },
     {
      "num": 2,
      "difficulty": "🟢 暖身",
      "desc": "選取同一欄 → 條件式格式化 → 色階 → 觀察紅綠漸層代表高低"
     },
     {
      "num": 3,
      "difficulty": "🔵 標準",
      "desc": "設定醒目提示：業績 > 60000 標為綠色底，< 40000 標為紅色底"
     },
     {
      "num": 4,
      "difficulty": "🔵 標準",
      "desc": "使用「頂端/底端規則」→ 標記前 3 名業績為粗體金色"
     },
     {
      "num": 5,
      "difficulty": "🟡 變化",
      "desc": "自訂公式規則：=MOD(ROW(),2)=0 → 隔列加底色（斑馬紋效果）"
     },
     {
      "num": 6,
      "difficulty": "🟡 變化",
      "desc": "自訂公式規則：=$E2>50000 → 整列（不是單格）標色。注意 $ 的位置！"
     },
     {
      "num": 7,
      "difficulty": "🔴 挑戰",
      "desc": "用圖示集（紅黃綠燈號）標記績效分數：≥90 綠燈、≥60 黃燈、其他紅燈"
     },
     {
      "num": 8,
      "difficulty": "🔴 挑戰",
      "desc": "自訂公式：=COUNTIF($A:$A,$A2)>1 → 標記重複的姓名。再試試用在訂單資料找重複"
     }
    ]
   },
   "meta": {
    "phase": "💼 Phase 2：工作即戰力    第 3~4 週  |  目標：老闆交辦的任務都能搞定",
    "stage": "第 7 階段",
    "topics": "資料橫條/色階/自訂公式規則",
    "difficulty": "🔵 Lv.2",
    "taskCount": "8 題",
    "time": "20 分鐘",
    "xp": "150 XP"
   }
  },
  "P3-01": {
   "knowledge": {
    "title": "✅ 資料驗證",
    "subtitle": "難度：🔵 Lv.2  |  防止錯誤輸入，保護資料品質",
    "sections": [
     {
      "title": "📌 操作路徑：資料 → 資料驗證",
      "items": [
       [
        "下拉選單",
        "清單 → 輸入：業務,財務,資訊,人事",
        "最常用！限制只能選這些值",
        "用 INDIRECT 可以做連動下拉"
       ],
       [
        "整數限制",
        "整數 → 介於 → 1 到 100",
        "限制只能輸入 1~100 的整數",
        "適合分數、百分比欄位"
       ],
       [
        "日期限制",
        "日期 → 大於 → =TODAY()",
        "只允許輸入未來的日期",
        "適合到期日、預約日欄位"
       ],
       [
        "文字長度",
        "文字長度 → 等於 → 10",
        "限制只能輸入 10 個字元",
        "適合身分證號、電話號碼"
       ],
       [
        "自訂公式",
        "自訂 → =AND(LEN(A2)=10, LEFT(A2,1)=\"A\")",
        "自訂任意驗證條件",
        "最靈活但需要公式基礎"
       ]
      ]
     },
     {
      "title": "🔗 連動下拉選單（進階必學）",
      "items": [
       [
        "Step 1",
        "建立命名範圍：每個主分類對應一組子選項"
       ],
       [
        "Step 2",
        "主下拉：資料驗證 → 清單 → 業務,財務,資訊"
       ],
       [
        "Step 3",
        "子下拉：資料驗證 → 清單 → =INDIRECT(A2)"
       ],
       [
        "原理",
        "INDIRECT 把 A2 的文字轉成命名範圍的參照"
       ]
      ]
     },
     {
      "title": "⚠️ 輸入訊息 & 錯誤提醒",
      "items": [
       [
        "輸入訊息",
        "資料驗證 → 輸入訊息 tab → 填寫提示文字",
        "滑鼠移到儲存格時顯示提示"
       ],
       [
        "錯誤提醒",
        "資料驗證 → 錯誤提醒 tab → 選擇停止/警告/資訊",
        "停止=不允許，警告=提醒但可繼續"
       ],
       [
        "圈選無效資料",
        "資料 → 資料驗證 → 圈選無效資料",
        "自動標出不符合規則的現有資料"
       ]
      ]
     }
    ],
    "handsTasks": [
     {
      "num": 1,
      "difficulty": "🟢 暖身",
      "desc": "在空白儲存格建立下拉選單：業務, 財務, 資訊, 人事（逗號分隔直接輸入）"
     },
     {
      "num": 2,
      "difficulty": "🟢 暖身",
      "desc": "建立整數驗證：限制只能輸入 1~100 的分數，並設定錯誤提醒訊息"
     },
     {
      "num": 3,
      "difficulty": "🔵 標準",
      "desc": "建立日期驗證：限制只能輸入今天以後的日期（公式：=TODAY()）"
     },
     {
      "num": 4,
      "difficulty": "🔵 標準",
      "desc": "在另一個儲存格建立「文字長度 = 10」的驗證 → 模擬手機號碼限制"
     },
     {
      "num": 5,
      "difficulty": "🟡 變化",
      "desc": "用命名範圍做下拉選單：先把部門清單命名為「部門列表」→ 驗證來源填 =部門列表"
     },
     {
      "num": 6,
      "difficulty": "🟡 變化",
      "desc": "建立連動下拉：主選單選「區域」→ 子選單用 INDIRECT 顯示該區域的業務員"
     },
     {
      "num": 7,
      "difficulty": "🔴 挑戰",
      "desc": "自訂公式驗證：=AND(LEN(A2)=10, LEFT(A2,1)=\"0\") → 限制只能輸入 0 開頭的 10 碼"
     },
     {
      "num": 8,
      "difficulty": "🔴 挑戰",
      "desc": "在已有資料的欄位加上驗證後 → 資料 → 圈選無效資料 → 找出不符規則的舊資料"
     }
    ]
   },
   "meta": {
    "phase": "📐 Phase 3：專業打磨    第 5~6 週  |  目標：做出讓人覺得「很專業」的報表",
    "stage": "第 8 階段",
    "topics": "下拉選單/連動選單/自訂驗證",
    "difficulty": "🔵 Lv.2",
    "taskCount": "8 題",
    "time": "15 分鐘",
    "xp": "150 XP"
   }
  },
  "P3-02": {
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
      "desc": "去除姓名的多餘空格",
      "time": "1m",
      "hint": "=TRIM(A6)",
      "answer": "=TRIM(A6)",
      "explain": "TRIM 移除前後空格及連續空格"
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
      "desc": "取出 Email 的帳號（@之前）",
      "time": "3m",
      "hint": "=LEFT(A8,FIND(\"@\",A8)-1)",
      "answer": "=LEFT(A8,FIND(\"@\",A8)-1)",
      "explain": "FIND 找 @ 的位置，LEFT 取左邊的字"
     },
     {
      "numLabel": "🔵4",
      "desc": "取出地址的縣市（前3個字）",
      "time": "1m",
      "hint": "=LEFT(A9,3)",
      "answer": "=LEFT(A9,3)",
      "explain": "LEFT(字串,幾個字)"
     },
     {
      "numLabel": "🔵5",
      "desc": "金額加千分位格式化",
      "time": "2m",
      "hint": "=TEXT(A10,\"#,##0.00\")",
      "answer": "=TEXT(A10,\"#,##0.00\")",
      "explain": "TEXT 把數字變成指定格式的文字"
     },
     {
      "numLabel": "🟡6",
      "desc": "計算起始到結束相差幾天",
      "time": "2m",
      "hint": "=DATEDIF(A11,A12,\"D\")",
      "answer": "=DATEDIF(A11,A12,\"D\")",
      "explain": "DATEDIF 第三參數：\"D\"天,\"M\"月,\"Y\"年"
     },
     {
      "numLabel": "🟡7",
      "desc": "計算相差幾個月",
      "time": "1m",
      "hint": "=DATEDIF(A11,A12,\"M\")",
      "answer": "=DATEDIF(A11,A12,\"M\")",
      "explain": "改成 M 就是月份差"
     },
     {
      "numLabel": "🟡8",
      "desc": "專案啟動日後 30 個工作天是哪天",
      "time": "2m",
      "hint": "=WORKDAY(A13,30)",
      "answer": "=WORKDAY(A13,30)",
      "explain": "WORKDAY 自動跳過週末"
     },
     {
      "numLabel": "🔴9",
      "desc": "取得專案啟動日的月底日期",
      "time": "2m",
      "hint": "=EOMONTH(A13,0)",
      "answer": "=EOMONTH(A13,0)",
      "explain": "EOMONTH(日期,0)=當月底，1=下月底"
     },
     {
      "numLabel": "🔴10",
      "desc": "合併姓名+部門為「王小明-業務」格式",
      "time": "2m",
      "hint": "=TRIM(A6)&\"-業務\"",
      "answer": "=TRIM(A6)&\"-業務\"",
      "explain": "& 串接文字，先 TRIM 去空格再串"
     }
    ]
   },
   "meta": {
    "phase": "📐 Phase 3：專業打磨    第 5~6 週  |  目標：做出讓人覺得「很專業」的報表",
    "stage": "第 9 階段",
    "topics": "TEXT / MID / TRIM / DATEDIF / WORKDAY",
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
  },
  "P3-03": {
   "knowledge": {
    "title": "📈 圖表設計",
    "subtitle": "難度：🟡 Lv.3  |  把數字變成故事  |  簡報必備技能",
    "sections": [
     {
      "title": "📌 選對圖表類型",
      "items": [
       [
        "直條圖/橫條圖",
        "比較不同項目的數值大小",
        "最泛用，適合類別比較",
        "項目多用橫條，少用直條"
       ],
       [
        "折線圖",
        "顯示趨勢變化（時間序列）",
        "X 軸通常是時間",
        "超過 5 條線就太雜了"
       ],
       [
        "圓餅圖",
        "顯示比例構成",
        "只適合 3~6 個類別",
        "專業人士其實很少用圓餅圖"
       ],
       [
        "散佈圖",
        "顯示兩個變數的相關性",
        "需要 X/Y 兩組數值資料",
        "加趨勢線：右鍵 → 新增趨勢線"
       ],
       [
        "組合圖",
        "不同量級的資料放同一圖（如營收+利潤率）",
        "插入 → 組合圖 → 選擇副座標軸",
        "營收用直條、利潤率用折線"
       ],
       [
        "走勢圖(Sparklines)",
        "插入 → 走勢圖 → 選範圍",
        "迷你圖放在儲存格裡，適合儀表板",
        "折線/直條/盈虧 三種"
       ]
      ]
     },
     {
      "title": "🎨 專業圖表設計原則",
      "items": [
       [
        "刪除多餘元素",
        "去掉格線、圖例（如果只有一組）、框線"
       ],
       [
        "直接標記數據",
        "在資料點上標數字，不要讓讀者去對座標軸"
       ],
       [
        "使用品牌色彩",
        "統一 2~3 個色系，避免彩虹配色"
       ],
       [
        "標題要有觀點",
        "不要「月營收」→ 改成「營收連續 3 月成長 15%」"
       ],
       [
        "字型一致",
        "圖表字型與報表一致，建議 10~12pt Arial"
       ]
      ]
     },
     {
      "title": "⚡ 高手技巧",
      "items": [
       [
        "動態圖表",
        "資料轉表格 (⌘+T) → 圖表會自動更新範圍"
       ],
       [
        "圖表範本",
        "右鍵圖表 → 另存為範本 → 下次直接套用"
       ],
       [
        "快速鍵",
        "Fn+⌥+F1 一鍵插入圖表；Fn+F11 插入到新工作表"
       ],
       [
        "樞紐圖表",
        "從樞紐分析表直接建圖，篩選連動",
        "分析 → 樞紐圖表"
       ]
      ]
     }
    ],
    "handsTasks": [
     {
      "num": 1,
      "difficulty": "🟢 暖身",
      "desc": "選取「📊 基礎函式」的業績資料 → 插入直條圖 → 觀察自動產生的圖表"
     },
     {
      "num": 2,
      "difficulty": "🟢 暖身",
      "desc": "把同一份資料改成折線圖 → 對比直條 vs 折線，哪個更適合看趨勢？"
     },
     {
      "num": 3,
      "difficulty": "🔵 標準",
      "desc": "刪除不必要的格線和框線 → 加上資料標籤 → 讓數字直接顯示在圖上"
     },
     {
      "num": 4,
      "difficulty": "🔵 標準",
      "desc": "修改標題：從「業績」改成有觀點的標題，如「Q2 業績成長 15%，張志偉持續領先」"
     },
     {
      "num": 5,
      "difficulty": "🟡 變化",
      "desc": "建立組合圖：左軸=金額（直條），右軸=成長率（折線），用在不同量級的資料"
     },
     {
      "num": 6,
      "difficulty": "🟡 變化",
      "desc": "在每位業務員旁邊加走勢圖（Sparklines）：插入 → 走勢圖 → 折線"
     },
     {
      "num": 7,
      "difficulty": "🔴 挑戰",
      "desc": "做一張「儀表板風格」圖表：只用 2 個顏色、無格線、有觀點標題、字體一致"
     },
     {
      "num": 8,
      "difficulty": "🔴 挑戰",
      "desc": "嘗試建立瀑布圖（Waterfall Chart）來展示收入和支出的累計效果"
     }
    ]
   },
   "meta": {
    "phase": "📐 Phase 3：專業打磨    第 5~6 週  |  目標：做出讓人覺得「很專業」的報表",
    "stage": "第10階段",
    "topics": "圖表選擇/設計原則/走勢圖",
    "difficulty": "🟡 Lv.3",
    "taskCount": "8 題",
    "time": "25 分鐘",
    "xp": "200 XP"
   }
  },
  "P3-04": {
   "knowledge": {
    "title": "🔗 命名範圍&表格",
    "subtitle": "難度：🟡 Lv.3  |  讓公式易讀、範圍自動擴展",
    "sections": [
     {
      "title": "📌 命名範圍（Named Range）",
      "items": [
       [
        "建立方式 1",
        "選取範圍 → 在名稱方塊（左上角）直接輸入名稱",
        "最快的方式"
       ],
       [
        "建立方式 2",
        "公式 → 定義名稱 → 輸入名稱與參照",
        "可以加註解"
       ],
       [
        "建立方式 3",
        "⌘+⇧+F3 → 依選取範圍自動建立",
        "欄標題自動變成名稱"
       ],
       [
        "管理",
        "公式 → 名稱管理員 → 檢視/修改/刪除",
        "定期清理不再使用的名稱"
       ]
      ]
     },
     {
      "title": "💡 命名範圍的好處",
      "items": [
       [
        "公式更易讀",
        "=SUM(月薪) 比 =SUM(E2:E100) 容易理解"
       ],
       [
        "全域可用",
        "任何工作表都可以使用同一個名稱"
       ],
       [
        "資料驗證",
        "下拉選單的來源可以用命名範圍"
       ],
       [
        "搭配 INDIRECT",
        "=INDIRECT(A1) 把文字轉成命名範圍的參照"
       ]
      ]
     },
     {
      "title": "📊 表格功能（⌘+T）",
      "items": [
       [
        "建立",
        "選取資料 → ⌘+T → 確認範圍",
        "自動加上篩選、格式、結構化參照"
       ],
       [
        "結構化參照",
        "=SUM(銷售表[金額]) 而非 =SUM(E2:E100))",
        "公式自動參照整欄，超好讀"
       ],
       [
        "自動擴展",
        "在最下方新增資料 → 表格自動包含新列",
        "公式、圖表、樞紐全部自動更新"
       ],
       [
        "彙總列",
        "表格設計 → 勾選「彙總列」→ 自動加總/平均/計數",
        "一鍵加上合計列"
       ],
       [
        "去除重複",
        "表格設計 → 移除重複項",
        "快速去重"
       ]
      ]
     },
     {
      "title": "⚠️ 表格 vs 一般範圍",
      "items": [
       [
        "建議",
        "所有資料都應該先轉表格，再做分析"
       ],
       [
        "限制",
        "表格不能跨工作表合併，不能有合併儲存格"
       ],
       [
        "命名",
        "建立表格後在「表格設計」改名為有意義的名稱"
       ]
      ]
     }
    ],
    "handsTasks": [
     {
      "num": 1,
      "difficulty": "🟢 暖身",
      "desc": "在「📊 基礎函式」選取業績資料的金額欄 → 在名稱方塊輸入「月薪」→ 測試 =SUM(月薪)"
     },
     {
      "num": 2,
      "difficulty": "🟢 暖身",
      "desc": "選取任意資料範圍 → ⌘+T 轉為表格 → 觀察公式欄自動變成結構化參照"
     },
     {
      "num": 3,
      "difficulty": "🔵 標準",
      "desc": "用 ⌘+⇧+F3 → 依選取範圍的欄標題自動建立多個命名範圍 → 公式→名稱管理員檢查"
     },
     {
      "num": 4,
      "difficulty": "🔵 標準",
      "desc": "在表格中新增一列資料 → 觀察公式和範圍是否自動擴展"
     },
     {
      "num": 5,
      "difficulty": "🟡 變化",
      "desc": "寫出結構化參照公式：=SUMIFS(銷售表[金額], 銷售表[區域], \"北區\")"
     },
     {
      "num": 6,
      "difficulty": "🟡 變化",
      "desc": "用 @ 符號：在表格內某列寫 =[@金額]*1.1 → 觀察公式自動填充到整欄"
     },
     {
      "num": 7,
      "difficulty": "🔴 挑戰",
      "desc": "建立動態命名範圍：公式→定義名稱→參照用 OFFSET+COUNTA 讓範圍隨資料自動增長"
     },
     {
      "num": 8,
      "difficulty": "🔴 挑戰",
      "desc": "用命名範圍做表格間的跨表引用：Sheet2 的公式直接用名稱而非 Sheet1!A1:A10"
     }
    ]
   },
   "meta": {
    "phase": "📐 Phase 3：專業打磨    第 5~6 週  |  目標：做出讓人覺得「很專業」的報表",
    "stage": "第11階段",
    "topics": "命名/表格/結構化參照",
    "difficulty": "🟡 Lv.3",
    "taskCount": "8 題",
    "time": "20 分鐘",
    "xp": "150 XP"
   }
  },
  "P3-05": {
   "knowledge": {
    "title": "🛡️ 保護與安全",
    "subtitle": "難度：🔵 Lv.2  |  防止誤改、保護公式、控制權限",
    "sections": [
     {
      "title": "📌 三層保護",
      "items": [
       [
        "儲存格鎖定",
        "選取要「允許」編輯的儲存格 → 右鍵 → 儲存格格式 → 保護 tab → 取消「鎖定」",
        "預設所有儲存格都是鎖定的，保護工作表後才生效"
       ],
       [
        "工作表保護",
        "校閱 → 保護工作表 → 設密碼 → 勾選允許的操作",
        "保護公式不被修改，但允許輸入資料"
       ],
       [
        "活頁簿保護",
        "校閱 → 保護活頁簿結構 → 設密碼",
        "防止新增/刪除/重新命名工作表"
       ]
      ]
     },
     {
      "title": "🔧 常見保護情境",
      "items": [
       [
        "保護公式，開放輸入",
        "把輸入格取消鎖定 → 保護工作表",
        "最常見的需求"
       ],
       [
        "隱藏公式",
        "儲存格格式 → 保護 tab → 勾「隱藏」→ 保護工作表",
        "公式欄不會顯示公式內容"
       ],
       [
        "允許特定範圍",
        "校閱 → 允許使用者編輯範圍 → 為不同範圍設不同密碼",
        "不同人可以編輯不同區域"
       ],
       [
        "防止複製",
        "保護工作表 → 不勾「選取鎖定的儲存格」",
        "完全無法選取=無法複製"
       ],
       [
        "🔒 檔案安全"
       ],
       [
        "開啟密碼",
        "檔案 → 另存新檔 → 工具 → 一般選項 → 開啟密碼",
        "沒密碼打不開檔案"
       ],
       [
        "修改密碼",
        "同上 → 修改密碼",
        "可以開啟但不能修改"
       ],
       [
        "標示為完稿",
        "檔案 → 資訊 → 保護活頁簿 → 標示為完稿",
        "提示性保護，可解除"
       ],
       [
        "數位簽章",
        "檔案 → 資訊 → 保護活頁簿 → 新增數位簽章",
        "驗證檔案來源和完整性"
       ]
      ]
     },
     {
      "title": "⚠️ 重要提醒",
      "items": [
       [
        "Excel 密碼不安全",
        "工作表保護密碼很容易被破解，不要當成真正的安全措施"
       ],
       [
        "備份",
        "保護前一定要留備份，忘記密碼就打不開了"
       ],
       [
        "VBA 可繞過",
        "VBA 可以解除工作表保護，真正安全需要開啟密碼"
       ]
      ]
     }
    ],
    "handsTasks": [
     {
      "num": 1,
      "difficulty": "🟢 暖身",
      "desc": "選取要讓使用者編輯的儲存格 → 右鍵 → 儲存格格式 → 保護 → 取消「鎖定」"
     },
     {
      "num": 2,
      "difficulty": "🟢 暖身",
      "desc": "校閱 → 保護工作表 → 設一個簡單密碼 → 測試：公式格不能改，輸入格可以改"
     },
     {
      "num": 3,
      "difficulty": "🔵 標準",
      "desc": "選取含公式的儲存格 → 勾選「隱藏」→ 再保護工作表 → 觀察公式欄是否顯示空白"
     },
     {
      "num": 4,
      "difficulty": "🔵 標準",
      "desc": "校閱 → 保護活頁簿結構 → 測試：嘗試新增/刪除工作表，應該被擋住"
     },
     {
      "num": 5,
      "difficulty": "🟡 變化",
      "desc": "取消保護 → 重新設定：允許排序和篩選，但不允許修改格式和插入欄列"
     },
     {
      "num": 6,
      "difficulty": "🟡 變化",
      "desc": "檔案 → 另存新檔 → 工具 → 一般選項 → 設定「開啟密碼」和「修改密碼」"
     },
     {
      "num": 7,
      "difficulty": "🔴 挑戰",
      "desc": "綜合保護：建立一張「客戶填寫用」的表格，公式全鎖定+隱藏，輸入格解鎖，設下拉選單，保護工作表。測試別人能不能正常填寫又不能改公式"
     },
     {
      "num": 8,
      "difficulty": "🔴 挑戰",
      "desc": "搭配 VBA 保護：用 VBA 在開啟檔案時自動保護工作表，關閉時解除保護（Workbook_Open / BeforeClose 事件）"
     }
    ]
   },
   "meta": {
    "phase": "📐 Phase 3：專業打磨    第 5~6 週  |  目標：做出讓人覺得「很專業」的報表",
    "stage": "第12階段",
    "topics": "鎖定/保護/密碼/隱藏公式",
    "difficulty": "🔵 Lv.2",
    "taskCount": "6 題",
    "time": "15 分鐘",
    "xp": "100 XP"
   }
  },
  "P4-01": {
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
      "desc": "篩選出業務部的員工（動態陣列）",
      "time": "3m",
      "hint": "=FILTER(A6:E13,B6:B13=\"業務\")",
      "answer": "=FILTER(A6:E13,B6:B13=\"業務\")",
      "explain": "FILTER 自動展開結果，不需輔助欄"
     },
     {
      "numLabel": "🟢2",
      "desc": "篩選月薪>=60000 的員工",
      "time": "2m",
      "hint": "=FILTER(A6:E13,C6:C13>=60000)",
      "answer": "=FILTER(A6:E13,C6:C13>=60000)",
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
      "desc": "取出所有不重複的部門名稱",
      "time": "2m",
      "hint": "=UNIQUE(B6:B13)",
      "answer": "=UNIQUE(B6:B13)",
      "explain": "UNIQUE 自動去重，適合做下拉選單"
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
      "desc": "篩選業務部且績效>=90",
      "time": "3m",
      "hint": "=FILTER(A6:E13,(B6:B13=\"業務\")*(E6:E13>=90))",
      "answer": "=FILTER(A6:E13,(B6:B13=\"業務\")*(E6:E13>=90))",
      "explain": "多條件用 * 相乘（AND 效果）"
     },
     {
      "numLabel": "🔴7",
      "desc": "計算每個部門有幾人（用 UNIQUE+COUNTIF）",
      "time": "3m",
      "hint": "先 =UNIQUE(B6:B13)，再 =COUNTIF(B6:B13,結果儲存格)",
      "answer": "先 =UNIQUE(B6:B13)，再 =COUNTIF(B6:B13,結果儲存格)",
      "explain": "組合技：先取不重複，再對每個值計數"
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
    "topics": "FILTER / SORT / UNIQUE / SEQUENCE",
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
      "desc": "計算所有產品的營收總和（單價×銷量 的總和）",
      "hint": "SUMPRODUCT(單價欄,銷量欄)",
      "answer": "=SUMPRODUCT(C6:C13,D6:D13)"
     },
     {
      "num": 2,
      "difficulty": "🟢",
      "desc": "列出不重複的類別清單（第一個）",
      "hint": "UNIQUE(範圍) 動態溢出",
      "answer": "=UNIQUE(B6:B13)"
     },
     {
      "num": 3,
      "difficulty": "🔵",
      "desc": "計算周邊類產品的平均評分",
      "hint": "AVERAGEIF(類別欄,\"周邊\",評分欄)",
      "answer": "=AVERAGEIF(B6:B13,\"周邊\",E6:E13)"
     },
     {
      "num": 4,
      "difficulty": "🔵",
      "desc": "找出銷量最高的產品名稱",
      "hint": "INDEX+MATCH+MAX 三層組合",
      "answer": "=INDEX(A6:A13,MATCH(MAX(D6:D13),D6:D13,0))"
     },
     {
      "num": 5,
      "difficulty": "🟡",
      "desc": "計算單價超過 1000 的產品數量",
      "hint": "COUNTIF(範圍,\">1000\")",
      "answer": "=COUNTIF(C6:C13,\">\"&1000)"
     },
     {
      "num": 6,
      "difficulty": "🟡",
      "desc": "用 SUMPRODUCT 計算周邊類的總銷量",
      "hint": "SUMPRODUCT((條件)*數值欄)",
      "answer": "=SUMPRODUCT((B6:B13=\"周邊\")*D6:D13)"
     },
     {
      "num": 7,
      "difficulty": "🔴",
      "desc": "計算評分最高的產品單價",
      "hint": "INDEX(單價欄,MATCH(MAX(評分),評分欄,0))",
      "answer": "=INDEX(C6:C13,MATCH(MAX(E6:E13),E6:E13,0))"
     },
     {
      "num": 8,
      "difficulty": "🔴",
      "desc": "計算所有產品的加權平均評分（以銷量為權重）",
      "hint": "SUMPRODUCT(銷量,評分)/SUM(銷量)",
      "answer": "=SUMPRODUCT(D6:D13,E6:E13)/SUM(D6:D13)"
     }
    ]
   }
  },
  "P4-02": {
   "pro": {
    "title": "🧮 進階函式  ─  INDIRECT / OFFSET / LET / AGGREGATE / LAMBDA",
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
      "desc": "用 INDIRECT 動態引用：輸入「B6」的文字取出對應值",
      "time": "3m",
      "hint": "=INDIRECT(\"B6\")",
      "answer": "=INDIRECT(\"B6\")",
      "explain": "INDIRECT 把文字轉成儲存格參照"
     },
     {
      "numLabel": "🟢2",
      "desc": "用 OFFSET 取得 A6 往右2格往下1格的值",
      "time": "3m",
      "hint": "=OFFSET(A6,1,2)",
      "answer": "=OFFSET(A6,1,2)",
      "explain": "OFFSET(基準,下移,右移) 動態偏移"
     },
     {
      "numLabel": "🔵3",
      "desc": "用 OFFSET+COUNTA 自動擴展的 SUM 範圍",
      "time": "3m",
      "hint": "=SUM(OFFSET(B6,0,0,COUNTA(A6:A20),1))",
      "answer": "=SUM(OFFSET(B6,0,0,COUNTA(A6:A20),1))",
      "explain": "用 COUNTA 計算列數，讓範圍自動增長"
     },
     {
      "numLabel": "🔵4",
      "desc": "用 LET 命名變數簡化公式",
      "time": "3m",
      "hint": "=LET(avg,AVERAGE(B6:B8),max,MAX(B6:B8),max-avg)",
      "answer": "=LET(avg,AVERAGE(B6:B8),max,MAX(B6:B8),max-avg)",
      "explain": "LET 給中間結果取名，公式更易讀"
     },
     {
      "numLabel": "🟡5",
      "desc": "用 AGGREGATE 忽略錯誤值求平均",
      "time": "3m",
      "hint": "=AGGREGATE(1,6,B6:E8)",
      "answer": "=AGGREGATE(1,6,B6:E8)",
      "explain": "AGGREGATE(函式碼,選項,範圍) 超強統計函式"
     },
     {
      "numLabel": "🟡6",
      "desc": "用 CHOOSE 根據數字選月份名稱",
      "time": "2m",
      "hint": "=CHOOSE(2,\"一月\",\"二月\",\"三月\")",
      "answer": "=CHOOSE(2,\"一月\",\"二月\",\"三月\")",
      "explain": "CHOOSE(索引,值1,值2,...) 按編號選"
     },
     {
      "numLabel": "🔴7",
      "desc": "用 SWITCH 把部門代碼轉名稱（1=業務,2=財務,3=資訊）",
      "time": "2m",
      "hint": "=SWITCH(1,1,\"業務\",2,\"財務\",3,\"資訊\",\"未知\")",
      "answer": "=SWITCH(1,1,\"業務\",2,\"財務\",3,\"資訊\",\"未知\")",
      "explain": "SWITCH 比巢狀 IF 乾淨很多"
     },
     {
      "numLabel": "🔴8",
      "desc": "用 LAMBDA 建立自訂函式：華氏轉攝氏",
      "time": "3m",
      "hint": "=LAMBDA(f,(f-32)*5/9)(100)",
      "answer": "=LAMBDA(f,(f-32)*5/9)(100)",
      "explain": "LAMBDA 建立可重用的自訂函式（365）"
     }
    ]
   },
   "meta": {
    "phase": "⚡ Phase 4：進階自動化    第 7~10 週  |  目標：處理大量資料、自動化流程",
    "stage": "第14階段",
    "topics": "INDIRECT / OFFSET / LET / AGGREGATE",
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
      "desc": "計算王小明的全年業績總和",
      "hint": "SUM 橫向加總",
      "answer": "=SUM(D6:G6)"
     },
     {
      "num": 2,
      "difficulty": "🟢",
      "desc": "找出所有業務員中 Q1 業績最高者的名字",
      "hint": "INDEX+MATCH+MAX+IF 陣列公式",
      "answer": "=INDEX(A6:A11,MATCH(MAX(IF(B6:B11=\"業務\",D6:D11)),IF(B6:B11=\"業務\",D6:D11),0))"
     },
     {
      "num": 3,
      "difficulty": "🔵",
      "desc": "計算業務部全年業績平均（含Q1~Q4）",
      "hint": "4個AVERAGEIFS相加或用SUMPRODUCT",
      "answer": "=AVERAGEIFS(D6:D11,B6:B11,\"業務\")+AVERAGEIFS(E6:E11,B6:B11,\"業務\")+AVERAGEIFS(F6:F11,B6:B11,\"業務\")+AVERAGEIFS(G6:G11,B6:B11,\"業務\")"
     },
     {
      "num": 4,
      "difficulty": "🔵",
      "desc": "用 RANK 計算王小明 Q1 業績在所有人中的排名",
      "hint": "RANK(值,範圍)",
      "answer": "=RANK(D6,D6:D11)"
     },
     {
      "num": 5,
      "difficulty": "🟡",
      "desc": "計算有業績（Q1>0）的人數",
      "hint": "COUNTIF(範圍,\">0\")",
      "answer": "=COUNTIF(D6:D11,\">\"&0)"
     },
     {
      "num": 6,
      "difficulty": "🟡",
      "desc": "計算業務部平均月薪與全公司平均月薪的差額",
      "hint": "AVERAGEIF - AVERAGE",
      "answer": "=AVERAGEIF(B6:B11,\"業務\",C6:C11)-AVERAGE(C6:C11)"
     },
     {
      "num": 7,
      "difficulty": "🔴",
      "desc": "用 IF+MAX 找出王小明最佳季度業績",
      "hint": "MAX 橫向取最大值",
      "answer": "=MAX(D6:G6)"
     },
     {
      "num": 8,
      "difficulty": "🔴",
      "desc": "計算全公司薪資的標準差",
      "hint": "STDEV(範圍) 樣本標準差",
      "answer": "=STDEV(C6:C11)"
     }
    ]
   }
  },
  "P4-03": {
   "knowledge": {
    "title": "🧩 假設分析",
    "subtitle": "難度：🟠 Lv.4  |  商業決策的分析工具  |  資料 → 假設分析",
    "sections": [
     {
      "title": "📌 三大工具",
      "items": [
       [
        "目標搜尋 (Goal Seek)",
        "已知結果，反推輸入值",
        "例：利潤要達 100 萬，售價應為多少？",
        "資料 → 假設分析 → 目標搜尋"
       ],
       [
        "資料表 (Data Table)",
        "一次看多個假設情境的結果",
        "例：不同利率 × 不同年期的月付款",
        "資料 → 假設分析 → 資料表"
       ],
       [
        "分析藍本 (Scenario Manager)",
        "儲存多組假設值，隨時切換",
        "例：樂觀/中性/悲觀三種情境",
        "資料 → 假設分析 → 分析藍本管理員"
       ]
      ]
     },
     {
      "title": "🔧 目標搜尋 (Goal Seek)",
      "items": [
       [
        "Step 1",
        "在某個儲存格設好公式（如：=售價×數量-成本）"
       ],
       [
        "Step 2",
        "資料 → 假設分析 → 目標搜尋"
       ],
       [
        "Step 3",
        "目標儲存格=利潤格、目標值=1000000、變更儲存格=售價格"
       ],
       [
        "用途",
        "反推損益兩平、目標達成所需的數值"
       ]
      ]
     },
     {
      "title": "📊 單/雙變數資料表",
      "items": [
       [
        "單變數",
        "一個輸入值（如利率）變化 → 看結果如何變"
       ],
       [
        "雙變數",
        "兩個輸入值同時變化 → 矩陣式呈現"
       ],
       [
        "操作",
        "左欄或上列放假設值 → 交叉格放結果公式 → 資料 → 資料表"
       ],
       [
        "用途",
        "敏感度分析、定價策略、貸款試算"
       ],
       [
        "🧮 規劃求解 (Solver)"
       ],
       [
        "什麼是",
        "最佳化工具：在限制條件下找出最佳解"
       ],
       [
        "啟用",
        "檔案 → 選項 → 增益集 → 規劃求解 → 啟用"
       ],
       [
        "用途",
        "最大化利潤、最小化成本、資源分配最佳化"
       ],
       [
        "設定",
        "目標儲存格 + 可變動儲存格 + 限制條件"
       ]
      ]
     }
    ],
    "handsTasks": [
     {
      "num": 1,
      "difficulty": "🟢 暖身",
      "desc": "建立一個簡單公式：利潤 = 售價 × 數量 - 成本。然後手動改售價觀察利潤變化"
     },
     {
      "num": 2,
      "difficulty": "🟢 暖身",
      "desc": "資料 → 假設分析 → 目標搜尋：讓利潤 = 100000，變更售價 → 看 Excel 反推出多少"
     },
     {
      "num": 3,
      "difficulty": "🔵 標準",
      "desc": "建立單變數資料表：左欄放 5 種不同利率 → 看每種利率下的月付款金額"
     },
     {
      "num": 4,
      "difficulty": "🔵 標準",
      "desc": "資料 → 假設分析 → 分析藍本：建立「樂觀/中性/悲觀」三種情境並切換觀察"
     },
     {
      "num": 5,
      "difficulty": "🟡 變化",
      "desc": "建立雙變數資料表：左欄=利率（3%~7%），上列=年期（10~30年）→ 交叉矩陣"
     },
     {
      "num": 6,
      "difficulty": "🟡 變化",
      "desc": "在分析藍本中 → 摘要報表 → 自動產生三種情境的對比表"
     },
     {
      "num": 7,
      "difficulty": "🔴 挑戰",
      "desc": "綜合運用：建立一個定價模型，用目標搜尋找損益兩平點，用資料表做敏感度分析"
     },
     {
      "num": 8,
      "difficulty": "🔴 挑戰",
      "desc": "嘗試規劃求解（Solver）：在限制條件下最佳化結果（需先啟用增益集）"
     }
    ]
   },
   "meta": {
    "phase": "⚡ Phase 4：進階自動化    第 7~10 週  |  目標：處理大量資料、自動化流程",
    "stage": "第16階段",
    "topics": "目標搜尋/資料表/分析藍本",
    "difficulty": "🟠 Lv.4",
    "taskCount": "8 題",
    "time": "20 分鐘",
    "xp": "200 XP"
   }
  },
  "P4-04": {
   "knowledge": {
    "title": "⚡ Power Query",
    "subtitle": "難度：🟠 Lv.4  |  現代 ETL 工具  |  自動清洗、合併、轉換資料",
    "sections": [
     {
      "title": "📌 什麼是 Power Query",
      "items": [
       [
        "定位",
        "Excel 內建的資料轉換引擎（資料 → 取得資料）"
       ],
       [
        "核心概念",
        "建立「查詢步驟」→ 資料更新時一鍵重跑"
       ],
       [
        "優勢",
        "不用公式、不用 VBA，滑鼠點點就能處理資料"
       ]
      ]
     },
     {
      "title": "🔧 常用操作",
      "items": [
       [
        "匯入資料",
        "資料 → 取得資料 → 從 Excel/CSV/資料夾/Web",
        "支援幾十種資料來源"
       ],
       [
        "移除欄/列",
        "選取欄 → 右鍵移除（或移除其他欄）",
        "只保留需要的資料"
       ],
       [
        "篩選",
        "點欄標題的下拉箭頭 → 篩選",
        "跟 Excel 篩選一樣直覺"
       ],
       [
        "分割欄位",
        "轉換 → 分割資料行 → 依分隔符號",
        "把「姓名-部門」拆成兩欄"
       ],
       [
        "合併欄位",
        "新增資料行 → 合併資料行",
        "把多欄合成一欄"
       ],
       [
        "取代值",
        "轉換 → 取代值",
        "批次替換，不怕改到公式"
       ],
       [
        "樞紐/取消樞紐",
        "轉換 → 取消資料行樞紐",
        "把寬表變長表（分析用）"
       ],
       [
        "群組",
        "轉換 → 群組依據",
        "類似 SQL 的 GROUP BY"
       ]
      ]
     },
     {
      "title": "🚀 進階技巧",
      "items": [
       [
        "合併查詢",
        "常用 → 合併查詢（類似 VLOOKUP 但更強）",
        "可選 Inner/Left/Full 等 Join 方式"
       ],
       [
        "附加查詢",
        "常用 → 附加查詢（垂直合併多張表）",
        "把結構相同的表上下接在一起"
       ],
       [
        "從資料夾匯入",
        "取得資料 → 從資料夾 → 自動合併所有檔案",
        "每月報表自動合併的神技"
       ],
       [
        "參數化查詢",
        "管理參數 → 新增參數 → 在查詢中引用",
        "讓使用者輸入篩選條件"
       ],
       [
        "M 語言",
        "進階編輯器可以寫 M 語言自訂邏輯",
        "Power Query 的底層語言"
       ]
      ]
     },
     {
      "title": "💡 實務建議",
      "items": [
       [
        "先 PQ 再分析",
        "用 Power Query 清洗 → 載入到表格 → 建樞紐表"
       ],
       [
        "保留步驟",
        "每個步驟都有紀錄，可以隨時回溯或修改"
       ],
       [
        "重新整理",
        "資料更新後 → 右鍵查詢 → 重新整理"
       ],
       [
        "效能",
        "大量資料（>10萬列）用 PQ 比公式快很多倍"
       ]
      ]
     }
    ],
    "handsTasks": [
     {
      "num": 1,
      "difficulty": "🟢 暖身",
      "desc": "資料 → 取得資料 → 從表格/範圍 → 進入 Power Query 編輯器 → 觀察介面"
     },
     {
      "num": 2,
      "difficulty": "🟢 暖身",
      "desc": "在 PQ 編輯器中：移除空白列 → 變更欄位類型（文字/數字/日期）→ 關閉並載入"
     },
     {
      "num": 3,
      "difficulty": "🔵 標準",
      "desc": "匯入一個 CSV 檔案 → 在 PQ 中：分割欄位（依分隔符號）→ 移除不需要的欄位"
     },
     {
      "num": 4,
      "difficulty": "🔵 標準",
      "desc": "新增自訂欄位：用 PQ 公式新增「含稅金額 = [金額] * 1.05」"
     },
     {
      "num": 5,
      "difficulty": "🟡 變化",
      "desc": "合併查詢：把兩個表格依共同欄位（如產品代碼）做 Left Join 合併"
     },
     {
      "num": 6,
      "difficulty": "🟡 變化",
      "desc": "附加查詢：把兩個相同結構的月報表上下合併（垂直合併）"
     },
     {
      "num": 7,
      "difficulty": "🔴 挑戰",
      "desc": "從資料夾匯入：把多個 Excel 檔案放同一個資料夾 → 取得資料 → 從資料夾 → 自動合併"
     },
     {
      "num": 8,
      "difficulty": "🔴 挑戰",
      "desc": "建立參數化查詢：讓使用者可以在 Excel 中輸入月份 → PQ 自動篩選該月資料"
     }
    ]
   },
   "meta": {
    "phase": "⚡ Phase 4：進階自動化    第 7~10 週  |  目標：處理大量資料、自動化流程",
    "stage": "第15階段",
    "topics": "匯入/清洗/合併/從資料夾匯入",
    "difficulty": "🟠 Lv.4",
    "taskCount": "8 題",
    "time": "30 分鐘",
    "xp": "300 XP"
   }
  },
  "P5-01": {
   "vba": {
    "title": "⚙️ VBA 基礎  ─  變數 / 迴圈 / 條件 / 儲存格操作",
    "subtitle": "難度：🟡 Lv.3  |  複製程式碼到 VBE（Alt+F11）執行",
    "warning": "⚠️ macOS 注意：VBA 編輯器從「工具→巨集→Visual Basic 編輯器」開啟（沒有 Alt+F11）。Mac 不支援 UserForm，需用 InputBox 或 MsgBox 替代。",
    "quickOps": [
     {
      "op": "開啟 VBE",
      "how": "Alt + F11"
     },
     {
      "op": "插入模組",
      "how": "VBE → 插入 → 模組"
     },
     {
      "op": "執行巨集",
      "how": "F5（在 VBE）或 工具→巨集（在 Excel）"
     },
     {
      "op": "逐行除錯",
      "how": "F8（逐步執行，觀察變數）"
     }
    ],
    "microTasks": [
     {
      "title": "微任務 1：Hello World（你的第一個 VBA）",
      "time": "⏱ 2 分鐘",
      "code": [
       "Sub HelloWorld()",
       "MsgBox \"Hello, Excel VBA！\"",
       "End Sub"
      ],
      "tip": "💡 MsgBox 彈出對話框，最基本的輸出方式"
     },
     {
      "title": "微任務 2：變數與資料型態",
      "time": "⏱ 3 分鐘",
      "code": [
       "Sub LearnVariables()",
       "Dim name As String    ' 文字型態",
       "Dim age As Integer    ' 整數",
       "Dim salary As Double  ' 含小數",
       "Dim isActive As Boolean  ' 布林",
       "name = \"Astor\"",
       "age = 30",
       "salary = 55000.5",
       "isActive = True",
       "MsgBox name & \" 年齡 \" & age & \" 薪資 \" & salary",
       "End Sub"
      ],
      "tip": "💡 Dim 宣告變數，As 指定型態。& 串接文字"
     },
     {
      "title": "微任務 3：讀寫儲存格",
      "time": "⏱ 2 分鐘",
      "code": [
       "Sub ReadWriteCells()",
       "' 寫入",
       "Range(\"A1\").Value = \"Hello\"",
       "Cells(2, 1).Value = \"World\"",
       "' 讀取",
       "Dim val As String",
       "val = Range(\"A1\").Value",
       "MsgBox \"A1 = \" & val",
       "End Sub"
      ],
      "tip": "💡 Range(\"A1\") 或 Cells(列,欄) 兩種方式操作儲存格"
     },
     {
      "title": "微任務 4：For 迴圈",
      "time": "⏱ 3 分鐘",
      "code": [
       "Sub ForLoop()",
       "Dim i As Integer",
       "' 在 A1:A10 填入 1~10 並上色",
       "For i = 1 To 10",
       "Cells(i, 1).Value = i",
       "Cells(i, 1).Interior.Color = _",
       "RGB(200, 230, 255)",
       "Next i",
       "MsgBox \"完成 A1:A10\"",
       "End Sub"
      ],
      "tip": "💡 For i = 起始 To 結束 ... Next i 是最常用迴圈"
     },
     {
      "title": "微任務 5：For Each 遍歷範圍",
      "time": "⏱ 2 分鐘",
      "code": [
       "Sub ForEachDemo()",
       "Dim cell As Range",
       "For Each cell In Range(\"A1:A10\")",
       "If cell.Value > 5 Then",
       "cell.Font.Bold = True",
       "cell.Font.Color = RGB(255, 0, 0)",
       "End If",
       "Next cell",
       "End Sub"
      ],
      "tip": "💡 For Each 逐一處理範圍內的每個儲存格"
     },
     {
      "title": "微任務 6：If 條件判斷",
      "time": "⏱ 2 分鐘",
      "code": [
       "Sub IfDemo()",
       "Dim score As Integer",
       "score = Range(\"B1\").Value",
       "If score >= 90 Then",
       "MsgBox \"優秀！\"",
       "ElseIf score >= 60 Then",
       "MsgBox \"及格\"",
       "Else",
       "MsgBox \"需要加油\"",
       "End If",
       "End Sub"
      ],
      "tip": "💡 If / ElseIf / Else / End If 完整條件架構"
     },
     {
      "title": "微任務 7：Do While 迴圈",
      "time": "⏱ 3 分鐘",
      "code": [
       "Sub DoWhileDemo()",
       "Dim r As Long",
       "r = 1",
       "' 讀取 A 欄直到空白為止",
       "Do While Cells(r, 1).Value <> \"\"",
       "Cells(r, 2).Value = Len(Cells(r, 1).Value)",
       "r = r + 1",
       "Loop",
       "MsgBox \"處理了 \" & (r - 1) & \" 列\"",
       "End Sub"
      ],
      "tip": "💡 Do While 條件成立就繼續。適合不知道幾筆資料時"
     },
     {
      "title": "微任務 8：Sub 與 Function",
      "time": "⏱ 3 分鐘",
      "code": [
       "Function CalcBonus(salary As Double, score As Integer) As Double",
       "If score >= 90 Then",
       "CalcBonus = salary * 3",
       "ElseIf score >= 75 Then",
       "CalcBonus = salary * 2",
       "Else",
       "CalcBonus = salary * 1",
       "End If",
       "End Function",
       "Sub TestBonus()",
       "MsgBox \"獎金 = \" & CalcBonus(50000, 92)",
       "End Sub"
      ],
      "tip": "💡 Function 可以回傳值，也能在儲存格當公式用！"
     },
     {
      "title": "微任務 9：格式化表格（實用模板）",
      "time": "⏱ 3 分鐘",
      "code": [
       "Sub FormatTable()",
       "With ActiveSheet.UsedRange",
       ".Borders.LineStyle = xlContinuous",
       ".Font.Name = \"Arial\"",
       ".Font.Size = 10",
       "End With",
       "With Rows(1)",
       ".Font.Bold = True",
       ".Font.Color = RGB(255, 255, 255)",
       ".Interior.Color = RGB(47, 84, 150)",
       "End With",
       "Cells.EntireColumn.AutoFit",
       "End Sub"
      ],
      "tip": "💡 With ... End With 減少重複寫物件名稱"
     },
     {
      "title": "微任務 10：刪除空白列（資料清洗）",
      "time": "⏱ 3 分鐘",
      "code": [
       "Sub DeleteEmptyRows()",
       "Dim i As Long",
       "For i = ActiveSheet.UsedRange.Rows.Count To 1 Step -1",
       "If WorksheetFunction.CountA(Rows(i)) = 0 Then",
       "Rows(i).Delete",
       "End If",
       "Next i",
       "MsgBox \"空白列已清除\"",
       "End Sub",
       "🎯 動手練習（做完打 ✓，難度漸進！）"
      ],
      "tip": "💡 從最後一列往上刪，避免索引錯亂"
     }
    ],
    "practiceTasks": [
     {
      "num": 1,
      "difficulty": "🟢 暖身",
      "desc": "開啟 VBE（工具→巨集→Visual Basic 編輯器）→ 插入模組 → 輸入 Sub Hello() / MsgBox \"Hello\" / End Sub → F5 執行"
     },
     {
      "num": 2,
      "difficulty": "🟢 暖身",
      "desc": "宣告變數：Dim name As String → name = Range(\"A1\").Value → MsgBox name。觀察 A1 的值是否正確顯示"
     },
     {
      "num": 3,
      "difficulty": "🔵 標準",
      "desc": "寫一個 For 迴圈：在 A1:A10 填入 1~10，並用 Interior.Color 幫偶數列加底色"
     },
     {
      "num": 4,
      "difficulty": "🔵 標準",
      "desc": "寫一個 If 判斷：讀取 B1 的分數，>=90 顯示「優秀」，>=60「及格」，其他「加油」"
     },
     {
      "num": 5,
      "difficulty": "🟡 變化",
      "desc": "用 For Each 遍歷 A1:A20，把所有大於 50000 的儲存格字體改為紅色粗體"
     },
     {
      "num": 6,
      "difficulty": "🟡 變化",
      "desc": "寫一個 Function：CalcTax(income)，收入 > 50000 稅率 20%，否則 10%，回傳稅額。在儲存格用 =CalcTax(A1) 測試"
     },
     {
      "num": 7,
      "difficulty": "🔴 挑戰",
      "desc": "寫一個完整的 Sub：讀取整個 UsedRange 到陣列 → 迴圈處理 → 寫回另一欄。比較跟逐格操作的速度差異"
     },
     {
      "num": 8,
      "difficulty": "🔴 挑戰",
      "desc": "加上錯誤處理：On Error GoTo ErrHandler → 主邏輯 → CleanUp 段。故意製造一個錯誤測試是否正確跳轉"
     }
    ]
   },
   "meta": {
    "phase": "🔧 Phase 5：VBA 自動化大師    第 11~14 週  |  目標：能寫程式自動化，能教別人",
    "stage": "第17階段",
    "topics": "變數 / 迴圈 / 條件 / 儲存格操作",
    "difficulty": "🟡 Lv.3",
    "taskCount": "10 題",
    "time": "30 分鐘",
    "xp": "300 XP"
   }
  },
  "P5-02": {
   "vba": {
    "title": "🔧 VBA 進階  ─  陣列 / 字典 / 錯誤處理 / 事件",
    "subtitle": "難度：🟠 Lv.4  |  這些技巧讓你從「會用」升級到「很強」",
    "warning": "⚠️ macOS 注意：VBA 編輯器從「工具→巨集→Visual Basic 編輯器」開啟（沒有 Alt+F11）。Mac 不支援 UserForm，需用 InputBox 或 MsgBox 替代。",
    "quickOps": [],
    "microTasks": [
     {
      "title": "微任務 1：陣列批次處理（效能大提升）",
      "time": "⏱ 3 分鐘",
      "code": [
       "Sub ArrayDemo()",
       "Dim arr() As Variant",
       "Dim i As Long",
       "' 一次讀入整個範圍",
       "arr = Range(\"A1:A1000\").Value",
       "For i = 1 To UBound(arr)",
       "arr(i, 1) = arr(i, 1) * 1.1  ' 加薪 10%",
       "Next i",
       "' 一次寫回",
       "Range(\"B1:B1000\").Value = arr",
       "End Sub"
      ],
      "tip": "💡 先讀到陣列 → 在記憶體處理 → 再寫回。比逐格操作快 100 倍"
     },
     {
      "title": "微任務 2：Dictionary 字典（去重/計數神器）",
      "time": "⏱ 3 分鐘",
      "code": [
       "Sub DictDemo()",
       "Dim dict As Object",
       "Set dict = CreateObject(\"Scripting.Dictionary\")",
       "Dim cell As Range",
       "For Each cell In Range(\"A1:A100\")",
       "If cell.Value <> \"\" Then",
       "If dict.Exists(cell.Value) Then",
       "dict(cell.Value) = dict(cell.Value) + 1",
       "Else",
       "dict.Add cell.Value, 1",
       "End If",
       "End If",
       "Next cell",
       "' 輸出結果",
       "Dim i As Long: i = 1",
       "Dim key As Variant",
       "For Each key In dict.Keys",
       "Cells(i, 5).Value = key",
       "Cells(i, 6).Value = dict(key)",
       "i = i + 1",
       "Next key",
       "End Sub"
      ],
      "tip": "💡 Dictionary 比 COUNTIF 更靈活，適合大量資料去重計數"
     },
     {
      "title": "微任務 3：錯誤處理 On Error",
      "time": "⏱ 2 分鐘",
      "code": [
       "Sub ErrorHandling()",
       "On Error GoTo ErrHandler",
       "Dim result As Double",
       "result = 10 / 0  ' 這會出錯",
       "Exit Sub",
       "ErrHandler:",
       "MsgBox \"發生錯誤: \" & Err.Description",
       "' 或 On Error Resume Next（忽略錯誤繼續）",
       "End Sub"
      ],
      "tip": "💡 On Error GoTo 標籤 = 跳到錯誤處理區。專業程式必備"
     },
     {
      "title": "微任務 4：With 語句 + 格式化",
      "time": "⏱ 2 分鐘",
      "code": [
       "Sub StyleDemo()",
       "With Range(\"A1:F1\")",
       ".Font.Bold = True",
       ".Font.Size = 12",
       ".Font.Color = RGB(255, 255, 255)",
       ".Interior.Color = RGB(47, 84, 150)",
       ".HorizontalAlignment = xlCenter",
       ".Borders.LineStyle = xlContinuous",
       "End With",
       "End Sub"
      ],
      "tip": "💡 With 減少重複輸入物件名，程式碼更乾淨"
     },
     {
      "title": "微任務 5：Worksheet 事件（自動觸發）",
      "time": "⏱ 3 分鐘",
      "code": [
       "' ⚠️ 這段放在「工作表」模組，不是一般模組",
       "Private Sub Worksheet_Change(ByVal Target As Range)",
       "' 當 A 欄被修改時，自動在 B 欄記錄時間",
       "If Target.Column = 1 Then",
       "Target.Offset(0, 1).Value = Now()",
       "End If",
       "End Sub"
      ],
      "tip": "💡 事件程式碼放在工作表模組（右鍵工作表→檢視程式碼）"
     },
     {
      "title": "微任務 6：Select Case（比 If 更乾淨）",
      "time": "⏱ 2 分鐘",
      "code": [
       "Sub SelectCaseDemo()",
       "Dim grade As String",
       "grade = Range(\"A1\").Value",
       "Select Case grade",
       "Case \"A\": MsgBox \"優秀\"",
       "Case \"B\": MsgBox \"良好\"",
       "Case \"C\": MsgBox \"及格\"",
       "Case Else: MsgBox \"未知等級\"",
       "End Select",
       "End Sub"
      ],
      "tip": "💡 Select Case 比多層 If/ElseIf 更易讀"
     },
     {
      "title": "微任務 7：FileSystemObject 檔案操作",
      "time": "⏱ 3 分鐘",
      "code": [
       "Sub ListFiles()",
       "Dim fso As Object, folder As Object, file As Object",
       "Set fso = CreateObject(\"Scripting.FileSystemObject\")",
       "Set folder = fso.GetFolder(\"C:\\Users\\你的帳號\\Desktop\")",
       "Dim r As Long: r = 1",
       "For Each file In folder.Files",
       "Cells(r, 1).Value = file.Name",
       "Cells(r, 2).Value = file.Size",
       "r = r + 1",
       "Next file",
       "End Sub"
      ],
      "tip": "💡 FSO 可以讀寫檔案系統，自動化必備技能"
     },
     {
      "title": "微任務 8：定時執行巨集",
      "time": "⏱ 2 分鐘",
      "code": [
       "Sub StartTimer()",
       "Application.OnTime Now + TimeValue(\"00:01:00\"), \"MyTask\"",
       "End Sub",
       "Sub MyTask()",
       "' 每分鐘執行的工作",
       "Range(\"A1\").Value = Now()",
       "StartTimer  ' 再次設定下次執行",
       "End Sub",
       "Sub StopTimer()",
       "Application.OnTime Now + TimeValue(\"00:01:00\"), \"MyTask\", , False",
       "End Sub",
       "🎯 動手練習（做完打 ✓，難度漸進！）"
      ],
      "tip": "💡 Application.OnTime 可以排程定時執行。StopTimer 停止"
     }
    ],
    "practiceTasks": [
     {
      "num": 1,
      "difficulty": "🟢 暖身",
      "desc": "建立 Dictionary：用 CreateObject(\"Scripting.Dictionary\") → 把 A 欄的值當 Key 計數 → 結果寫到 D:E 欄"
     },
     {
      "num": 2,
      "difficulty": "🟢 暖身",
      "desc": "用陣列批次處理：arr = Range(\"A1:A1000\").Value → 迴圈修改 → Range(\"B1:B1000\").Value = arr"
     },
     {
      "num": 3,
      "difficulty": "🔵 標準",
      "desc": "寫 Worksheet_Change 事件：A 欄任何修改時，自動在 B 欄寫入 Now() 時間戳"
     },
     {
      "num": 4,
      "difficulty": "🔵 標準",
      "desc": "寫效能優化開關：ScreenUpdating=False + Calculation=Manual → 大量操作 → 恢復。計時比較差異"
     },
     {
      "num": 5,
      "difficulty": "🟡 變化",
      "desc": "用 Dictionary 做去重：讀取 A 欄 → 只保留不重複的值 → 寫到新工作表"
     },
     {
      "num": 6,
      "difficulty": "🟡 變化",
      "desc": "寫 Application.OnTime 排程：每 60 秒自動更新 A1 為現在時間，並寫一個 StopTimer 停止"
     },
     {
      "num": 7,
      "difficulty": "🔴 挑戰",
      "desc": "完整專案：建立「資料清洗」模組 → 刪空白列 + 去重複 + 格式統一 + 錯誤處理，一鍵執行"
     },
     {
      "num": 8,
      "difficulty": "🔴 挑戰",
      "desc": "寫 Class Module：建立一個 Employee 類別（Name, Salary, CalcBonus 方法），在 Sub 中實例化並使用"
     }
    ]
   },
   "meta": {
    "phase": "🔧 Phase 5：VBA 自動化大師    第 11~14 週  |  目標：能寫程式自動化，能教別人",
    "stage": "第18階段",
    "topics": "陣列 / 字典 / 錯誤處理 / 事件",
    "difficulty": "🟠 Lv.4",
    "taskCount": "8 題",
    "time": "30 分鐘",
    "xp": "300 XP"
   }
  },
  "P5-03": {
   "vba": {
    "title": "🏗️ VBA 實戰  ─  UserForm / 自動報表 / 實務範本",
    "subtitle": "難度：🔴 Lv.5  |  做完這裡，你就能教別人了",
    "warning": "⚠️ macOS 注意：VBA 編輯器從「工具→巨集→Visual Basic 編輯器」開啟（沒有 Alt+F11）。Mac 不支援 UserForm，需用 InputBox 或 MsgBox 替代。",
    "quickOps": [],
    "microTasks": [
     {
      "title": "實戰 1：自動產生月報表",
      "time": "⏱ 5 分鐘",
      "code": [
       "Sub GenerateMonthlyReport()",
       "Dim wsData As Worksheet, wsReport As Worksheet",
       "Set wsData = Sheets(\"資料\")",
       "' 新增報表工作表",
       "Set wsReport = Sheets.Add(After:=Sheets(Sheets.Count))",
       "wsReport.Name = \"月報_\" & Format(Date, \"YYYYMM\")",
       "' 標題",
       "wsReport.Range(\"A1\").Value = Format(Date, \"YYYY年MM月\") & \" 月報表\"",
       "wsReport.Range(\"A1\").Font.Size = 16",
       "wsReport.Range(\"A1\").Font.Bold = True",
       "' 匯總資料",
       "Dim lastRow As Long",
       "lastRow = wsData.Cells(Rows.Count, 1).End(xlUp).Row",
       "wsReport.Range(\"A3\").Value = \"總筆數\"",
       "wsReport.Range(\"B3\").Value = lastRow - 1",
       "wsReport.Range(\"A4\").Value = \"總金額\"",
       "wsReport.Range(\"B4\").Value = WorksheetFunction.Sum(wsData.Range(\"E2:E\" & lastRow))",
       "wsReport.Range(\"B4\").NumberFormat = \"$#,##0\"",
       "' 格式化",
       "wsReport.Columns(\"A:B\").AutoFit",
       "MsgBox \"月報表已產生！\"",
       "End Sub"
      ],
      "tip": "💡 實務中最常見的需求：自動匯總+產生報表"
     },
     {
      "title": "實戰 2：UserForm 資料輸入介面",
      "time": "⏱ 5 分鐘",
      "code": [
       "' ⚠️ 先在 VBE 插入 UserForm，加入：",
       "' TextBox1 (姓名), TextBox2 (金額), CommandButton1 (送出)",
       "' 在 UserForm 的程式碼中：",
       "Private Sub CommandButton1_Click()",
       "Dim nextRow As Long",
       "nextRow = Cells(Rows.Count, 1).End(xlUp).Row + 1",
       "Cells(nextRow, 1).Value = TextBox1.Value",
       "Cells(nextRow, 2).Value = CDbl(TextBox2.Value)",
       "Cells(nextRow, 3).Value = Now()",
       "MsgBox \"已新增！\"",
       "TextBox1.Value = \"\"",
       "TextBox2.Value = \"\"",
       "End Sub",
       "' 在一般模組中（開啟表單）：",
       "Sub ShowForm()",
       "UserForm1.Show",
       "End Sub"
      ],
      "tip": "💡 UserForm 讓非技術同事也能安全輸入資料"
     },
     {
      "title": "實戰 3：匯出所有工作表為 PDF",
      "time": "⏱ 3 分鐘",
      "code": [
       "Sub ExportAllToPDF()",
       "Dim ws As Worksheet",
       "Dim path As String",
       "path = ThisWorkbook.Path & \"\\\"",
       "For Each ws In ThisWorkbook.Worksheets",
       "ws.ExportAsFixedFormat _",
       "Type:=xlTypePDF, _",
       "Filename:=path & ws.Name & \".pdf\", _",
       "Quality:=xlQualityStandard",
       "Next ws",
       "MsgBox \"所有工作表已匯出為 PDF！\"",
       "End Sub"
      ],
      "tip": "💡 批次匯出 PDF 是辦公室最受歡迎的自動化之一"
     },
     {
      "title": "實戰 4：自動寄送 Email（透過 Outlook）",
      "time": "⏱ 3 分鐘",
      "code": [
       "Sub SendEmail()",
       "Dim olApp As Object, olMail As Object",
       "Set olApp = CreateObject(\"Outlook.Application\")",
       "Set olMail = olApp.CreateItem(0)",
       "With olMail",
       ".To = \"colleague@company.com\"",
       ".Subject = \"自動報表 \" & Format(Date, \"YYYY/MM/DD\")",
       ".Body = \"附件為今日報表，請查閱。\"",
       ".Attachments.Add ThisWorkbook.FullName",
       ".Display  ' 改成 .Send 直接寄出",
       "End With",
       "End Sub"
      ],
      "tip": "💡 透過 Outlook 物件模型自動建立郵件"
     },
     {
      "title": "實戰 5：建立工作表目錄（含超連結）",
      "time": "⏱ 3 分鐘",
      "code": [
       "Sub CreateIndex()",
       "Dim ws As Worksheet, idxWs As Worksheet",
       "Dim i As Integer",
       "On Error Resume Next",
       "Application.DisplayAlerts = False",
       "Sheets(\"目錄\").Delete",
       "Application.DisplayAlerts = True",
       "On Error GoTo 0",
       "Set idxWs = Sheets.Add(Before:=Sheets(1))",
       "idxWs.Name = \"目錄\"",
       "idxWs.Range(\"A1\").Value = \"📋 工作表目錄\"",
       "idxWs.Range(\"A1\").Font.Size = 14",
       "idxWs.Range(\"A1\").Font.Bold = True",
       "i = 3",
       "For Each ws In ThisWorkbook.Sheets",
       "If ws.Name <> \"目錄\" Then",
       "idxWs.Hyperlinks.Add _",
       "Anchor:=idxWs.Cells(i, 1), _",
       "Address:=\"\", _",
       "SubAddress:=\"'\" & ws.Name & \"'!A1\", _",
       "TextToDisplay:=ws.Name",
       "i = i + 1",
       "End If",
       "Next ws",
       "idxWs.Columns(\"A\").AutoFit",
       "End Sub"
      ],
      "tip": "💡 自動目錄 + 超連結，方便在多工作表間導航"
     },
     {
      "title": "實戰 6：批次條件標色（取代手動格式化）",
      "time": "⏱ 3 分鐘",
      "code": [
       "Sub ConditionalColor()",
       "Dim cell As Range",
       "For Each cell In Range(\"E2:E\" & Cells(Rows.Count, 5).End(xlUp).Row)",
       "Select Case True",
       "Case cell.Value >= 50000",
       "cell.Interior.Color = RGB(198, 239, 206)  ' 綠",
       "Case cell.Value >= 30000",
       "cell.Interior.Color = RGB(255, 235, 156)  ' 黃",
       "Case cell.Value > 0",
       "cell.Interior.Color = RGB(255, 199, 206)  ' 紅",
       "End Select",
       "Next cell",
       "End Sub"
      ],
      "tip": "💡 Select Case True 是一個巧妙的多條件判斷寫法"
     },
     {
      "title": "實戰 7：效能優化技巧",
      "time": "⏱ 2 分鐘",
      "code": [
       "Sub OptimizedMacro()",
       "' ── 關閉更新提升效能 ──",
       "Application.ScreenUpdating = False",
       "Application.Calculation = xlCalculationManual",
       "Application.EnableEvents = False",
       "' ... 你的主要程式碼 ...",
       "' ── 恢復設定 ──",
       "Application.ScreenUpdating = True",
       "Application.Calculation = xlCalculationAutomatic",
       "Application.EnableEvents = True",
       "End Sub"
      ],
      "tip": "💡 大量操作時關閉畫面更新和自動計算，速度差 10 倍以上"
     },
     {
      "title": "實戰 8：完整錯誤處理框架",
      "time": "⏱ 3 分鐘",
      "code": [
       "Sub ProfessionalSub()",
       "On Error GoTo ErrHandler",
       "Application.ScreenUpdating = False",
       "' === 主要邏輯 ===",
       "' ... 你的程式碼 ...",
       "CleanUp:",
       "Application.ScreenUpdating = True",
       "Application.Calculation = xlCalculationAutomatic",
       "Exit Sub",
       "ErrHandler:",
       "MsgBox \"錯誤 \" & Err.Number & \": \" & Err.Description, _",
       "vbCritical, \"發生錯誤\"",
       "Resume CleanUp",
       "End Sub",
       "🎯 動手練習（做完打 ✓，難度漸進！）"
      ],
      "tip": "💡 專業模板：錯誤處理 + 效能優化 + 清理，養成好習慣"
     }
    ],
    "practiceTasks": [
     {
      "num": 1,
      "difficulty": "🟢 暖身",
      "desc": "複製「自動產生月報表」範本 → 修改為你的資料結構 → 執行看結果是否正確"
     },
     {
      "num": 2,
      "difficulty": "🟢 暖身",
      "desc": "複製「建立工作表目錄」範本 → 執行 → 觀察超連結是否能正確跳轉"
     },
     {
      "num": 3,
      "difficulty": "🔵 標準",
      "desc": "寫一個「批次匯出 PDF」巨集：每個工作表匯出為獨立 PDF 到指定資料夾"
     },
     {
      "num": 4,
      "difficulty": "🔵 標準",
      "desc": "寫一個「資料驗證報告」：掃描所有儲存格找出空白、異常值、格式不一致，輸出報告"
     },
     {
      "num": 5,
      "difficulty": "🟡 變化",
      "desc": "建立一個自動化流程：讀取 CSV → 清洗 → 產生樞紐摘要 → 匯出 PDF，一鍵完成"
     },
     {
      "num": 6,
      "difficulty": "🟡 變化",
      "desc": "寫一個 Ribbon 自訂按鈕（⚠️ Mac 限制：用工具→巨集→指定鍵盤快速鍵替代）"
     },
     {
      "num": 7,
      "difficulty": "🔴 挑戰",
      "desc": "完整專案：模擬你的「核對表自動填入」邏輯，用 VBA 重寫（不用 Python）→ 比較兩種方式"
     },
     {
      "num": 8,
      "difficulty": "🔴 挑戰",
      "desc": "建立「模板引擎」：讀取範本工作表 → 批次替換佔位符 → 產生多份客製報表"
     }
    ]
   },
   "meta": {
    "phase": "🔧 Phase 5：VBA 自動化大師    第 11~14 週  |  目標：能寫程式自動化，能教別人",
    "stage": "第19階段",
    "topics": "自動報表 / 串接 / 實務範本",
    "difficulty": "🔴 Lv.5",
    "taskCount": "8 題",
    "time": "40 分鐘",
    "xp": "400 XP"
   }
  },
  "P5-04": {
   "pro": {
    "title": "🏆 綜合挑戰  ─  融合所有技能的員工薪資分析",
    "subtitle": "難度：🔴 Lv.5  |  微任務數：5 題  |  建議時間：每題 2~3 分鐘",
    "dataHeader": [
     "員工編號",
     "姓名",
     "部門",
     "入職日期",
     "月薪",
     "績效",
     ""
    ],
    "dataRows": [
     [
      "E001",
      "王小明",
      "業務",
      "2019-03-15",
      55000,
      88,
      ""
     ],
     [
      "E002",
      "林美華",
      "財務",
      "2021-07-01",
      48000,
      72,
      ""
     ],
     [
      "E003",
      "張志偉",
      "業務",
      "2016-11-20",
      68000,
      95,
      ""
     ],
     [
      "E004",
      "陳雅婷",
      "人事",
      "2023-01-10",
      42000,
      65,
      ""
     ],
     [
      "E005",
      "李建宏",
      "資訊",
      "2018-05-28",
      75000,
      91,
      ""
     ],
     [
      "E006",
      "劉淑芬",
      "財務",
      "2020-09-01",
      52000,
      79,
      ""
     ],
     [
      "E007",
      "黃志強",
      "業務",
      "2015-04-12",
      82000,
      98,
      ""
     ],
     [
      "E008",
      "吳雅琪",
      "資訊",
      "2022-08-15",
      63000,
      84,
      ""
     ]
    ],
    "tasks": [
     {
      "numLabel": "🟢1",
      "desc": "計算每人年資（到 2026/3/31）",
      "time": "2m",
      "hint": "=DATEDIF(D6,DATE(2026,3,31),\"Y\")",
      "answer": "=DATEDIF(D6,DATE(2026,3,31),\"Y\")",
      "explain": "DATEDIF + DATE 組合"
     },
     {
      "numLabel": "🔵2",
      "desc": "薪資等級：>=70K高級/>=50K中級/其他初級",
      "time": "3m",
      "hint": "=IFS(E6>=70000,\"高級\",E6>=50000,\"中級\",TRUE,\"初級\")",
      "answer": "=IFS(E6>=70000,\"高級\",E6>=50000,\"中級\",TRUE,\"初級\")",
      "explain": "IFS 多層判斷"
     },
     {
      "numLabel": "🟡3",
      "desc": "年終獎金：績效>=90→×3/>=75→×2/其他→×1",
      "time": "3m",
      "hint": "=IFS(F6>=90,E6*3,F6>=75,E6*2,TRUE,E6)",
      "answer": "=IFS(F6>=90,E6*3,F6>=75,E6*2,TRUE,E6)",
      "explain": "IFS 結合計算"
     },
     {
      "numLabel": "🟡4",
      "desc": "SUMIFS：業務部年終總額",
      "time": "3m",
      "hint": "先完成獎金欄，再 =SUMIFS(獎金欄,部門欄,\"業務\")",
      "answer": "=SUMIFS(I6:I13,C6:C13,\"業務\")",
      "explain": "假設獎金在 I 欄"
     },
     {
      "numLabel": "🔴5",
      "desc": "INDEX+MATCH：查 E005 的姓名",
      "time": "2m",
      "hint": "=INDEX(B6:B13,MATCH(\"E005\",A6:A13,0))",
      "answer": "=INDEX(B6:B13,MATCH(\"E005\",A6:A13,0))",
      "explain": "反向查找"
     }
    ]
   },
   "meta": {
    "phase": "🔧 Phase 5：VBA 自動化大師    第 11~14 週  |  目標：能寫程式自動化，能教別人",
    "stage": "第20階段",
    "topics": "融合所有技能的實戰任務",
    "difficulty": "🔴 Lv.5",
    "taskCount": "5 題",
    "time": "30 分鐘",
    "xp": "500 XP"
   }
  }
 },
 "dashboard": {
  "⌨️ 快捷鍵大全": {
   "phase": "🏃 Phase 1：操作效率基礎    第 1~2 週  |  目標：養成鍵盤操作習慣，複習基本功",
   "stage": "第 1 階段",
   "topics": "導航/選取/編輯/KeyTips 全鍵盤",
   "difficulty": "🟢 Lv.1",
   "taskCount": "10 題",
   "time": "20 分鐘",
   "xp": "150 XP"
  },
  "📊 基礎函式": {
   "phase": "🏃 Phase 1：操作效率基礎    第 1~2 週  |  目標：養成鍵盤操作習慣，複習基本功",
   "stage": "第 2 階段",
   "topics": "SUM / AVERAGE / COUNT / MAX / MIN",
   "difficulty": "🟢 Lv.1",
   "taskCount": "8 題",
   "time": "15 分鐘",
   "xp": "100 XP"
  },
  "📐 條件判斷": {
   "phase": "🏃 Phase 1：操作效率基礎    第 1~2 週  |  目標：養成鍵盤操作習慣，複習基本功",
   "stage": "第 3 階段",
   "topics": "IF / IFS / 巢狀 IF / AND / OR",
   "difficulty": "🟢 Lv.1",
   "taskCount": "8 題",
   "time": "15 分鐘",
   "xp": "100 XP"
  },
  "📈 條件統計": {
   "phase": "💼 Phase 2：工作即戰力    第 3~4 週  |  目標：老闆交辦的任務都能搞定",
   "stage": "第 4 階段",
   "topics": "COUNTIFS / SUMIFS / AVERAGEIFS",
   "difficulty": "🔵 Lv.2",
   "taskCount": "8 題",
   "time": "20 分鐘",
   "xp": "150 XP"
  },
  "🔍 查找比對": {
   "phase": "💼 Phase 2：工作即戰力    第 3~4 週  |  目標：老闆交辦的任務都能搞定",
   "stage": "第 5 階段",
   "topics": "VLOOKUP / XLOOKUP / INDEX+MATCH",
   "difficulty": "🔵 Lv.2",
   "taskCount": "10 題",
   "time": "25 分鐘",
   "xp": "200 XP"
  },
  "📊 樞紐分析表": {
   "phase": "💼 Phase 2：工作即戰力    第 3~4 週  |  目標：老闆交辦的任務都能搞定",
   "stage": "第 6 階段",
   "topics": "建立/四大區域/日期群組/計算欄位",
   "difficulty": "🔵 Lv.2",
   "taskCount": "8 題",
   "time": "25 分鐘",
   "xp": "200 XP"
  },
  "🎨 條件式格式化": {
   "phase": "💼 Phase 2：工作即戰力    第 3~4 週  |  目標：老闆交辦的任務都能搞定",
   "stage": "第 7 階段",
   "topics": "資料橫條/色階/自訂公式規則",
   "difficulty": "🔵 Lv.2",
   "taskCount": "8 題",
   "time": "20 分鐘",
   "xp": "150 XP"
  },
  "✅ 資料驗證": {
   "phase": "📐 Phase 3：專業打磨    第 5~6 週  |  目標：做出讓人覺得「很專業」的報表",
   "stage": "第 8 階段",
   "topics": "下拉選單/連動選單/自訂驗證",
   "difficulty": "🔵 Lv.2",
   "taskCount": "8 題",
   "time": "15 分鐘",
   "xp": "150 XP"
  },
  "📝 文字日期": {
   "phase": "📐 Phase 3：專業打磨    第 5~6 週  |  目標：做出讓人覺得「很專業」的報表",
   "stage": "第 9 階段",
   "topics": "TEXT / MID / TRIM / DATEDIF / WORKDAY",
   "difficulty": "🟡 Lv.3",
   "taskCount": "10 題",
   "time": "20 分鐘",
   "xp": "150 XP"
  },
  "📈 圖表設計": {
   "phase": "📐 Phase 3：專業打磨    第 5~6 週  |  目標：做出讓人覺得「很專業」的報表",
   "stage": "第10階段",
   "topics": "圖表選擇/設計原則/走勢圖",
   "difficulty": "🟡 Lv.3",
   "taskCount": "8 題",
   "time": "25 分鐘",
   "xp": "200 XP"
  },
  "🔗 命名範圍&表格": {
   "phase": "📐 Phase 3：專業打磨    第 5~6 週  |  目標：做出讓人覺得「很專業」的報表",
   "stage": "第11階段",
   "topics": "命名/表格/結構化參照",
   "difficulty": "🟡 Lv.3",
   "taskCount": "8 題",
   "time": "20 分鐘",
   "xp": "150 XP"
  },
  "🛡️ 保護與安全": {
   "phase": "📐 Phase 3：專業打磨    第 5~6 週  |  目標：做出讓人覺得「很專業」的報表",
   "stage": "第12階段",
   "topics": "鎖定/保護/密碼/隱藏公式",
   "difficulty": "🔵 Lv.2",
   "taskCount": "6 題",
   "time": "15 分鐘",
   "xp": "100 XP"
  },
  "📊 陣列動態": {
   "phase": "⚡ Phase 4：進階自動化    第 7~10 週  |  目標：處理大量資料、自動化流程",
   "stage": "第13階段",
   "topics": "FILTER / SORT / UNIQUE / SEQUENCE",
   "difficulty": "🟡 Lv.3",
   "taskCount": "8 題",
   "time": "20 分鐘",
   "xp": "200 XP"
  },
  "🧮 進階函式": {
   "phase": "⚡ Phase 4：進階自動化    第 7~10 週  |  目標：處理大量資料、自動化流程",
   "stage": "第14階段",
   "topics": "INDIRECT / OFFSET / LET / AGGREGATE",
   "difficulty": "🟠 Lv.4",
   "taskCount": "8 題",
   "time": "25 分鐘",
   "xp": "250 XP"
  },
  "⚡ Power Query": {
   "phase": "⚡ Phase 4：進階自動化    第 7~10 週  |  目標：處理大量資料、自動化流程",
   "stage": "第15階段",
   "topics": "匯入/清洗/合併/從資料夾匯入",
   "difficulty": "🟠 Lv.4",
   "taskCount": "8 題",
   "time": "30 分鐘",
   "xp": "300 XP"
  },
  "🧩 假設分析": {
   "phase": "⚡ Phase 4：進階自動化    第 7~10 週  |  目標：處理大量資料、自動化流程",
   "stage": "第16階段",
   "topics": "目標搜尋/資料表/分析藍本",
   "difficulty": "🟠 Lv.4",
   "taskCount": "8 題",
   "time": "20 分鐘",
   "xp": "200 XP"
  },
  "⚙️ VBA基礎": {
   "phase": "🔧 Phase 5：VBA 自動化大師    第 11~14 週  |  目標：能寫程式自動化，能教別人",
   "stage": "第17階段",
   "topics": "變數 / 迴圈 / 條件 / 儲存格操作",
   "difficulty": "🟡 Lv.3",
   "taskCount": "10 題",
   "time": "30 分鐘",
   "xp": "300 XP"
  },
  "🔧 VBA進階": {
   "phase": "🔧 Phase 5：VBA 自動化大師    第 11~14 週  |  目標：能寫程式自動化，能教別人",
   "stage": "第18階段",
   "topics": "陣列 / 字典 / 錯誤處理 / 事件",
   "difficulty": "🟠 Lv.4",
   "taskCount": "8 題",
   "time": "30 分鐘",
   "xp": "300 XP"
  },
  "🏗️ VBA實戰": {
   "phase": "🔧 Phase 5：VBA 自動化大師    第 11~14 週  |  目標：能寫程式自動化，能教別人",
   "stage": "第19階段",
   "topics": "自動報表 / 串接 / 實務範本",
   "difficulty": "🔴 Lv.5",
   "taskCount": "8 題",
   "time": "40 分鐘",
   "xp": "400 XP"
  },
  "🏆 綜合挑戰": {
   "phase": "🔧 Phase 5：VBA 自動化大師    第 11~14 週  |  目標：能寫程式自動化，能教別人",
   "stage": "第20階段",
   "topics": "融合所有技能的實戰任務",
   "difficulty": "🔴 Lv.5",
   "taskCount": "5 題",
   "time": "30 分鐘",
   "xp": "500 XP"
  }
 },
 "badges": [
  {
   "name": "⌨️ 鍵盤戰士",
   "condition": "完成 Phase 1（第 1~3 階段）",
   "xp": "350 XP"
  },
  {
   "name": "💼 即戰力新星",
   "condition": "完成 Phase 2（第 4~7 階段）",
   "xp": "1,050 XP"
  },
  {
   "name": "📐 專業報表手",
   "condition": "完成 Phase 3（第 8~12 階段）",
   "xp": "1,800 XP"
  },
  {
   "name": "⚡ 資料工程師",
   "condition": "完成 Phase 4（第 13~16 階段）",
   "xp": "2,750 XP"
  },
  {
   "name": "🔧 自動化大師",
   "condition": "完成 VBA 基礎+進階（第 17~18 階段）",
   "xp": "3,350 XP"
  },
  {
   "name": "👑 全能導師",
   "condition": "完成全部 20 階段",
   "xp": "3,850 XP"
  }
 ],
 "strategies": [
  {
   "title": "⏱ 番茄鐘節奏",
   "desc": "15 分鐘專注 → 5 分鐘休息 → 再做 15 分鐘。一天最多 2~3 個番茄鐘就夠了"
  },
  {
   "title": "🎯 每天只做一件事",
   "desc": "今天練 3 個快捷鍵 OR 做 4 題函式 OR 看 1 個 VBA 範本。不要混著學！"
  },
  {
   "title": "🚫 卡住 3 分鐘就看答案",
   "desc": "看完答案後理解原理 → 關掉答案 → 自己重做一次。不要死磕"
  },
  {
   "title": "✅ 做完就打勾",
   "desc": "回到儀表板打 ✓ 是你的多巴胺獎勵。每個小 ✓ 都是真的進步"
  },
  {
   "title": "📌 答案欄先遮住",
   "desc": "用紙、便利貼、或把 G~H 欄暫時隱藏（選取 → 右鍵 → 隱藏）"
  },
  {
   "title": "🔄 間隔複習",
   "desc": "做完新的一階段後，花 5 分鐘回去「快速掃」前一階段的題目"
  },
  {
   "title": "💡 跟工作連結",
   "desc": "你的「核對表自動填入」專案已經用到了 XLOOKUP，學新東西時想想能不能用在工作上"
  }
 ],
 "startPoint": [
  {
   "stage": "已會的",
   "desc": "基本操作、基礎函式、XLOOKUP（核對表專案已實戰使用）"
  },
  {
   "stage": "Phase 1",
   "desc": "快捷鍵需要從頭練。基礎函式和條件判斷可快速掃過，預計 1 週搞定"
  },
  {
   "stage": "Phase 2",
   "desc": "條件統計是新的，查找比對有基礎可加速。樞紐和格式化要認真學。預計 2 週"
  },
  {
   "stage": "Phase 3",
   "desc": "全部是新技能但都不難，穩穩學 2 週"
  },
  {
   "stage": "Phase 4",
   "desc": "陣列函式和 Power Query 較吃力，預留 3~4 週"
  },
  {
   "stage": "Phase 5",
   "desc": "VBA 完全新手，但你有 Python 經驗（核對表），概念可遷移。預留 3~4 週"
  }
 ]
}
