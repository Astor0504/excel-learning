// 由 xlsx-content.js + lesson-demos-data.js 拆分(2026-07-14);本檔是 P1-01 的唯一資料來源
export const CONTENT = {
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
      "key": "⌃ + D",
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
      "key": "F2 或 fn+F2",
      "desc": "進入編輯模式",
      "note": "官方主路線是 F2；MacBook 常要搭配 fn"
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
      "key": "⌃ + T / ⌃ + L（選取資料時）",
      "desc": "轉換為表格",
      "note": "⚠️ 注意：⌘+T 在公式中=切換參照；建立表格請用 ⌃+T 或 ⌃+L"
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
      "note": "⚠️ Mac 不走 Alt+F11；請用 ⌥+F11 或從工具選單開啟"
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
      "note": "輸入/刪除/複製貼上、⌃+D 向下填滿、⌘+R 向右填滿"
     },
     {
      "key": "公式建立",
      "desc": "",
      "note": "所有公式輸入、⌘+T 切換絕對參照、Tab 自動完成函式名"
     },
     {
      "key": "表格操作",
      "desc": "",
      "note": "⌃+T 建立表格、⌘+⇧+L 開關篩選、排序（KeyTips 觸發）"
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
      "note": "⌃+T 表格 + ⌘+⇧+T 自動SUM + ⌘+1 格式化 + ⌘+⇧+L 篩選"
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
      "key": "⌃+T / ⌃+L 建立表格",
      "desc": "⌃+T 或 ⌃+L",
      "note": "WPS 快捷鍵可能依版本不同，建議以選單操作為主"
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
 },
 "knowledge": {
  "title": "⌨️ 新手熱身路線",
  "subtitle": "難度：🟢 Lv.1  |  先把最常用的 6 組操作練成反射  |  不求快，先求穩",
  "sections": [
   {
    "title": "📌 先練這 6 組就夠了",
    "items": [
     [
      "1. 存檔 / 復原",
      "⌘+S、⌘+Z",
      "先養成隨時存檔與敢試錯的節奏"
     ],
     [
      "2. 跳邊界",
      "⌘+方向鍵",
      "在有資料的表格裡快速移動，比滑鼠快很多"
     ],
     [
      "3. 選整欄 / 整列",
      "⌃+Space、⇧+Space",
      "整理資料與套格式時超常用"
     ],
     [
      "4. 編輯儲存格",
      "F2 或 fn+F2",
      "進入編輯模式，不用再雙擊儲存格"
     ],
     [
      "5. 向下填滿",
      "⌃+D",
      "把公式或內容快速複製到下方"
     ],
     [
      "6. 搜尋",
      "⌘+F",
      "表格大了以後，先找再動手"
     ]
    ]
   },
   {
    "title": "🎯 10 分鐘案例：整理一張名單",
    "items": [
     [
      "情境",
      "你拿到一張員工名單，要快速檢查資料、插入新列、搜尋特定姓名"
     ],
     [
      "起手式",
      "先按 ⌘+方向鍵確認資料邊界，再用 ⇧ 搭配方向鍵做選取"
     ],
     [
      "中段操作",
      "需要補欄位時用 ⌃+⇧+= 插入列或欄；要改內容先按 F2"
     ],
     [
      "結尾",
      "最後用 ⌘+F 搜尋關鍵字，確認你剛剛補的資料有沒有到位"
     ]
    ]
   },
   {
    "title": "⚠️ 新手最容易卡住的地方",
    "items": [
     [
      "按了沒反應",
      "先檢查是不是在中文輸入法組字中，或 F 鍵沒有開標準功能鍵"
     ],
     [
      "方向鍵變成捲動畫面",
      "可能是 Scroll Lock 狀態異常，先切回正常工作表焦點"
     ],
     [
      "記不起來太多鍵",
      "不要一次背 20 個。先背 3 個，每天強迫自己用一次"
     ]
    ]
   }
  ],
  "handsTasks": [
   {
    "num": 1,
    "difficulty": "🟢 暖身",
    "desc": "開一張任意工作表，連續做 3 次：⌘+S 存檔 → ⌘+Z 復原 → 再輸入一個值"
   },
   {
    "num": 2,
    "difficulty": "🟢 暖身",
    "desc": "在有資料的區塊中，練習用 ⌘+→ / ⌘+↓ 跳到資料邊界，再用 ⌘+← / ⌘+↑ 回來"
   },
   {
    "num": 3,
    "difficulty": "🔵 標準",
    "desc": "選一整欄後插入新欄，再用 F2 編輯其中一格內容，不用滑鼠完成"
   },
   {
    "num": 4,
    "difficulty": "🔵 標準",
    "desc": "在 A1 輸入公式或文字後，用 ⌃+D 向下填滿到 A10，感受和拖曳填滿的差異"
   },
   {
    "num": 5,
    "difficulty": "🟡 變化",
    "desc": "做一次完整小流程：搜尋姓名 → 編輯內容 → 插入一列 → 存檔"
   },
   {
    "num": 6,
    "difficulty": "🟡 變化",
    "desc": "把今天最常用的 3 個快捷鍵寫進筆記區，明天只強迫自己用這 3 個"
   }
  ]
 },
 "meta": {
  "phase": "🏃 Phase 1：操作效率基礎    第 1~2 週  |  目標：養成鍵盤操作習慣，複習基本功",
  "stage": "第 1 階段",
  "topics": "導航 / 選取 / 編輯 / 搜尋 / 表格基本操作",
  "difficulty": "🟢 Lv.1",
  "taskCount": "6 題",
  "time": "15 分鐘",
  "xp": "150 XP"
 }
};
export const DEMOS = [
 {
  "id": "mac-shortcut-navigation",
  "kind": "workflow",
  "badge": "操作節奏",
  "title": "快捷鍵不是背表：從定位、選取到填滿的 20 秒流程",
  "subtitle": "先練一條最常用的鍵盤路徑，手感會比單背快捷鍵更快建立起來。",
  "stages": [
   {
    "text": "定位資料",
    "tone": "c1"
   },
   {
    "text": "選取區域",
    "tone": "c2"
   },
   {
    "text": "插入新列",
    "tone": "c3"
   },
   {
    "text": "填滿公式",
    "tone": "c4"
   }
  ],
  "outcome": "結果：不用離開鍵盤，也能完成跳邊界、選資料、插入列、填滿公式。",
  "steps": [
   {
    "stage": 1,
    "label": "1. 用 ⌘ + 方向鍵跳到資料邊界",
    "tone": "c1",
    "text": "從資料區任一格出發，⌘ + ↓ 會跳到目前連續資料區的底部，⌘ + → 會跳到右側邊界。",
    "result": "游標快速抵達資料邊界，不用滾動找最後一列。",
    "caption": "新版 Excel for Mac 通常用 ⌘ + 方向鍵；舊版可試 ⌃ + 方向鍵。",
    "panels": [
     {
      "title": "鍵盤節奏",
      "lines": [
       "起點：A2",
       "按 ⌘ + ↓：跳到最後一筆資料",
       "按 ⌘ + →：跳到資料區右側"
      ]
     }
    ]
   },
   {
    "stage": 2,
    "label": "2. 用 Shift 延伸選取範圍",
    "tone": "c2",
    "text": "搭配 Shift，可以從目前位置延伸選取到資料邊界，快速圈出需要格式化或複製的整段資料。",
    "result": "資料區被一次選起來，後續操作不必用滑鼠拖拉。",
    "caption": "如果要選整欄，用 ⌃ Space；選整列，用 ⇧ Space。",
    "panels": [
     {
      "title": "選取策略",
      "columns": [
       "目的",
       "快捷鍵",
       "結果"
      ],
      "rows": [
       [
        "選整欄",
        "⌃ Space",
        "目前欄整欄選取"
       ],
       [
        "選整列",
        "⇧ Space",
        "目前列整列選取"
       ],
       [
        "延伸到邊界",
        "⇧ + ⌘ + 方向鍵",
        "選到資料邊界"
       ]
      ]
     }
    ]
   },
   {
    "stage": 3,
    "label": "3. 插入或刪除列欄",
    "tone": "c3",
    "text": "選好整列或整欄後，用 ⌃ ⇧ + 插入，用 ⌘ - 刪除。這是清資料時最常用的一組動作。",
    "result": "新增列欄不需要右鍵選單，節奏更連續。",
    "caption": "WPS 的插入快捷鍵可能不同，這課以 Excel for Mac 為主。",
    "panels": [
     {
      "title": "列欄動作",
      "lines": [
       "先選整列或整欄",
       "⌃ ⇧ +：插入",
       "⌘ -：刪除",
       "⌘ Z：復原"
      ]
     }
    ]
   },
   {
    "stage": 4,
    "label": "4. 用 ⌃ D 填滿向下",
    "tone": "c4",
    "text": "公式寫好一格後，選取下方需要填滿的範圍，按 ⌃ D 就能把上方公式向下填滿。",
    "result": "一整欄公式快速填滿，並保留相對參照邏輯。",
    "caption": "填滿前先檢查第一格公式是否正確，錯一格會一路錯到底。",
    "panels": [
     {
      "title": "填滿前/後",
      "columns": [
       "位置",
       "填滿前",
       "填滿後"
      ],
      "rows": [
       [
        "D2",
        "=B2*C2",
        "=B2*C2"
       ],
       [
        "D3",
        "空白",
        "=B3*C3"
       ],
       [
        "D4",
        "空白",
        "=B4*C4"
       ]
      ]
     }
    ]
   }
  ]
 }
];
