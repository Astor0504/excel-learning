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
  },
  "P1-02": {
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
  },
  "P1-03": {
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
  },
  "P2-02": {
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
        "選取資料範圍（建議先轉為表格 ⌃+T 或 ⌃+L）"
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
      "title": "🎯 典型案例：月業績彙總",
      "items": [
       [
        "原始資料",
        "每一列是一筆交易，欄位有：部門、月份、產品、金額"
       ],
       [
        "你想回答的問題",
        "哪個部門在每個月份的總業績最高？"
       ],
       [
        "樞紐配置",
        "列放部門、欄放月份、值放金額加總、篩選可放地區或付款狀態"
       ],
       [
        "做完後的結果",
        "原本幾百筆流水帳，會變成一張 2~3 秒就看懂的交叉表"
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
      "desc": "在「📈 條件統計」工作表的訂單資料上，按 ⌃+T 或 ⌃+L 轉為表格，再插入樞紐分析表"
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
      "title": "🎯 典型案例：主管一眼看出誰要追蹤",
      "items": [
       [
        "原始情境",
        "一張業務月報有姓名、部門、業績、成長率、客訴數，但全部都是黑白數字",
        "主管每次都得自己慢慢找出低於目標的人"
       ],
       [
        "第一層視覺",
        "業績欄加資料橫條、成長率加色階",
        "先讓高低差異能在 3 秒內看出來"
       ],
       [
        "第二層提醒",
        "成長率 < 0 標紅、客訴數 > 3 標黃、業績前 3 名加圖示集",
        "把「該追蹤、該表揚」直接浮出來"
       ],
       [
        "專業做法",
        "先決定規則代表什麼商業訊號，再套格式",
        "條件式格式化不是上色，而是把判斷邏輯視覺化"
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
     },
     {
      "title": "✅ 交付檢查",
      "items": [
       [
        "先測命中列",
        "新增一筆應該被標示的資料，確認格式會自動套用",
        "不要只看既有資料剛好有沒有亮"
       ],
       [
        "再測排除列",
        "新增一筆不該被標示的資料，確認規則不會誤傷",
        "尤其是自訂公式規則"
       ],
       [
        "寫下規則意義",
        "在報表旁註明紅色、黃色、圖示各代表什麼商業訊號",
        "讓接手者知道顏色不是裝飾"
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
      "title": "🎯 典型案例：建立不會填錯的報名表",
      "items": [
       [
        "原始情境",
        "人資要收新人資料，常有人把部門打錯、日期格式亂掉、手機少一碼",
        "後面整理名單很痛苦"
       ],
       [
        "第一步",
        "部門欄改成下拉選單，日期欄限制只能輸入今天以後",
        "先把最常見的格式錯誤擋掉"
       ],
       [
        "第二步",
        "手機欄限制 10 碼，班別欄做連動下拉",
        "讓使用者只能在合法範圍內輸入"
       ],
       [
        "做完效果",
        "資料進來時就乾淨，後面不用再花時間人工清理",
        "資料驗證的價值在於把錯誤擋在入口，不是事後補救"
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
     },
     {
      "title": "✅ 正式表單檢查",
      "items": [
       [
        "欄位角色",
        "先分出必填、選填、公式產生欄",
        "公式欄不應該讓使用者直接輸入"
       ],
       [
        "維護清單",
        "部門、區域、狀態集中放在清單區或表格",
        "新增選項時不必逐格修改規則"
       ],
       [
        "搭配保護",
        "長期給多人填的表單，要搭配工作表保護與條件式格式化",
        "資料驗證只擋輸入錯，不等於權限控管"
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
      "title": "🎯 典型案例：把月報從表格變成會說話的圖",
      "items": [
       [
        "原始情境",
        "你有 12 個月營收、3 位業務員業績、各產品占比，但全部都放在表格裡",
        "主管一眼看不出重點"
       ],
       [
        "怎麼分配",
        "看趨勢用折線圖、比大小用直條圖、看占比用圓餅或甜甜圈圖",
        "先決定問題，再決定圖型"
       ],
       [
        "常見錯誤",
        "把月份占比拿去畫折線，或把時間序列拿去畫圓餅",
        "圖表選錯，資料再漂亮也會誤導"
       ],
       [
        "專業成果",
        "一張月報最多保留 1 個主圖 + 1 個輔助圖",
        "讓閱讀順序清楚，比把所有圖都塞上去更重要"
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
        "資料轉表格 (⌃+T / ⌃+L) → 圖表會自動更新範圍"
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
     },
     {
      "title": "✅ 圖表交付檢查",
      "items": [
       [
        "先有結論",
        "圖表標題要能說出觀察，不只是欄位名稱",
        "例如把「月營收」改成「營收連續三月成長放緩」"
       ],
       [
        "控制標籤",
        "只標最後值、最高低點或討論中的資料點",
        "每個點都貼數字會讓圖更難讀"
       ],
       [
        "來源可更新",
        "正式月報的圖表來源優先轉成 Table",
        "新增資料後範圍才不會漏掉"
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
      "title": "📊 表格功能（⌃+T / ⌃+L）",
      "items": [
       [
        "建立",
        "選取資料 → ⌃+T 或 ⌃+L → 確認範圍",
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
      "title": "🎯 典型案例：讓每月新增資料不用重改公式",
      "items": [
       [
        "原始情境",
        "每個月都要把新訂單貼到報表底下，SUMIFS、圖表、樞紐都得重新拉範圍",
        "一忙就容易漏掉"
       ],
       [
        "做法",
        "先把資料轉成表格，再把關鍵區域命名",
        "表格負責自動擴展，命名範圍負責讓公式可讀"
       ],
       [
        "結果",
        "新資料貼到最下方後，表格自動吃進去，後面報表同步更新",
        "你不再是每個月重拉範圍，而是維護一個會自己長大的系統"
       ],
       [
        "判斷原則",
        "需要自動擴張時優先用表格；需要跨表引用時再搭配命名範圍",
        "兩者不是二選一，而是一起用效果最好"
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
       ],
       [
        "交付檢查",
        "新增一列測試資料，確認公式、樞紐與圖表都會吃到新列"
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
      "desc": "選取任意資料範圍 → ⌃+T 或 ⌃+L 轉為表格 → 觀察公式欄自動變成結構化參照"
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
       ]
      ]
     },
     {
      "title": "🔒 檔案安全",
      "items": [
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
      "title": "🎯 典型案例：交給同事填，但不能動到公式",
      "items": [
       [
        "原始情境",
        "你做了一份報價單模板，要交給同事或客戶填資料",
        "最怕對方把公式蓋掉、刪掉工作表、或把敏感公式複製走"
       ],
       [
        "基本作法",
        "輸入格解鎖、公式格保留鎖定，必要時把公式設成隱藏",
        "再開啟工作表保護"
       ],
       [
        "進一步",
        "若還要防止新增刪除頁籤，就再保護活頁簿結構",
        "如果真的涉及機密，再考慮開啟密碼而不是只靠工作表保護"
       ],
       [
        "專業觀念",
        "保護的目標是防誤改，不是做真正的資訊安全",
        "先分清楚你是在防手滑，還是在防外流"
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
       ],
       [
        "使用者測試",
        "保護後要用填表者視角測輸入、貼上、篩選與列印，不要只測密碼"
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
    "taskCount": "8 題",
    "time": "15 分鐘",
    "xp": "100 XP"
   }
  },
  "P4-01": {
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
  },
  "P4-02": {
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
       ]
      ]
     },
     {
      "title": "🧮 規劃求解 (Solver)",
      "items": [
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
     },
     {
      "title": "🎯 典型案例：定價調 5%，利潤會差多少？",
      "items": [
       [
        "原始情境",
        "老闆問：售價從 950、990、1050 到 1100，哪個價格最划算？",
        "如果每次都手改一輪公式，很慢也容易抄錯"
       ],
       [
        "怎麼拆",
        "只改一個參數就用目標搜尋或單變數資料表",
        "同時看售價 × 銷量兩個變數就用雙變數資料表"
       ],
       [
        "延伸",
        "樂觀、中性、悲觀三套假設值可用分析藍本收起來",
        "需要在限制條件下找最佳解時，再換 Solver"
       ],
       [
        "專業做法",
        "先問清楚自己是要反推答案、比較情境，還是找最佳解",
        "工具選對了，假設分析才會快"
       ],
       [
        "交付檢查",
        "把輸入假設、限制條件與最後輸出放在同一區，讓別人能追溯答案來源"
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
      "title": "🎯 典型案例：每月合併各分店 CSV",
      "items": [
       [
        "原始狀況",
        "每個分店每月都丟一份 CSV，欄位名稱差不多，但格式不乾淨"
       ],
       [
        "Power Query 要做的事",
        "從資料夾匯入 → 合併所有檔案 → 去掉空列 → 統一日期與數值格式 → 載入表格"
       ],
       [
        "你最後得到什麼",
        "下個月只要把新檔案放進資料夾，按重新整理就能重跑整套流程"
       ],
       [
        "專業分工",
        "Power Query 負責清資料；樞紐或圖表負責展示；VBA 只在需要一鍵交付時補上"
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
  "P4-05": {
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
  },
  "P5-01": {
   "vba": {
    "title": "⚙️ VBA 基礎  ─  變數 / 迴圈 / 條件 / 儲存格操作",
    "subtitle": "難度：🟡 Lv.3  |  複製程式碼到 VBE 執行（Mac 可用 ⌥+F11 或從工具選單開啟）",
    "warning": "⚠️ macOS 注意：不要照搬 Windows 的 Alt+F11 心智模型；請用「工具→巨集→Visual Basic 編輯器」或 ⌥+F11 開啟。Mac 不支援 UserForm，需用 InputBox 或 MsgBox 替代。",
    "quickOps": [
     {
      "op": "開啟 VBE",
      "how": "⌥ + F11 或 工具 → 巨集 → Visual Basic 編輯器"
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
    "taskCount": "8 題",
    "time": "30 分鐘",
    "xp": "300 XP"
   }
  },
  "P5-02": {
   "vba": {
    "title": "🔧 VBA 進階  ─  陣列 / 字典 / 錯誤處理 / 事件",
    "subtitle": "難度：🟠 Lv.4  |  這些技巧讓你從「會用」升級到「很強」",
    "warning": "⚠️ macOS 注意：不要照搬 Windows 的 Alt+F11 心智模型；請用「工具→巨集→Visual Basic 編輯器」或 ⌥+F11 開啟。Mac 不支援 UserForm、ActiveX 與多數 Windows COM 自動化；本課主線放在工作簿內資料處理、事件、錯誤處理與可交付流程。",
    "quickOps": [
     {
      "op": "先想 Power Query",
      "how": "如果需求是重複清資料、多檔合併，優先用 Power Query；VBA 負責刷新、檢查與輸出。"
     },
     {
      "op": "先寫 CleanUp",
      "how": "任何會關閉畫面更新、事件或自動計算的巨集，都要有統一復原區。"
     }
    ],
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
      "title": "微任務 3：錯誤處理 + CleanUp 安全收尾",
      "time": "⏱ 2 分鐘",
      "code": [
       "Sub SafeRun()",
       "On Error GoTo ErrHandler",
       "Application.ScreenUpdating = False",
       "Application.EnableEvents = False",
       "' 你的主要流程",
       "",
       "CleanUp:",
       "Application.EnableEvents = True",
       "Application.ScreenUpdating = True",
       "If Err.Number = 0 Then MsgBox \"完成\"",
       "Exit Sub",
       "",
       "ErrHandler:",
       "MsgBox \"發生錯誤: \" & Err.Description",
       "Resume CleanUp",
       "End Sub"
      ],
      "tip": "💡 專業 VBA 不是不出錯，而是出錯時也能把 Excel 設定復原"
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
       "Application.EnableEvents = False",
       "Target.Offset(0, 1).Value = Now()",
       "Application.EnableEvents = True",
       "End If",
       "End Sub"
      ],
      "tip": "💡 事件程式碼放在工作表模組；事件裡改格子時要暫停 EnableEvents，避免自己觸發自己"
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
      "title": "微任務 7：Mac 友善檔案清單",
      "time": "⏱ 3 分鐘",
      "code": [
       "Sub ListFiles()",
       "Dim folderPath As String",
       "folderPath = ThisWorkbook.Path & Application.PathSeparator",
       "Dim fileName As String",
       "fileName = Dir(folderPath & \"*.csv\")",
       "Dim r As Long: r = 1",
       "Do While fileName <> \"\"",
       "Cells(r, 1).Value = fileName",
       "Cells(r, 2).Value = folderPath & fileName",
       "r = r + 1",
       "fileName = Dir()",
       "Loop",
       "End Sub"
      ],
      "tip": "💡 用 ThisWorkbook.Path + Dir 可以避開 Windows 路徑範例，Mac/Windows 都更容易理解"
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
       "End Sub"
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
    "title": "🏗️ VBA 實戰  ─  RefreshAll / PDF 匯出 / 錯誤處理 / 交付流程",
    "subtitle": "難度：🔴 Lv.5  |  做完這裡，你就能教別人了",
    "warning": "⚠️ macOS 注意：不要照搬 Windows 的 Alt+F11 心智模型；請用「工具→巨集→Visual Basic 編輯器」或 ⌥+F11 開啟。Mac 不支援 UserForm，也不要把 Windows 的 Outlook COM 自動化當成主路線；請優先用 InputBox、工作表表單、PDF 匯出與檔案流程。",
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
      "title": "實戰 2：InputBox 快速輸入流程",
      "time": "⏱ 5 分鐘",
      "code": [
       "Sub QuickEntry()",
       "Dim userName As String, amountText As String",
       "Dim nextRow As Long",
       "userName = InputBox(\"請輸入姓名\")",
       "If userName = \"\" Then Exit Sub",
       "amountText = InputBox(\"請輸入金額\")",
       "If amountText = \"\" Then Exit Sub",
       "nextRow = Cells(Rows.Count, 1).End(xlUp).Row + 1",
       "Cells(nextRow, 1).Value = userName",
       "Cells(nextRow, 2).Value = CDbl(amountText)",
       "Cells(nextRow, 3).Value = Now()",
       "MsgBox \"已新增！\"",
       "End Sub"
      ],
      "tip": "💡 這是 Mac 可走的安全做法：不用 UserForm，也能快速建立半自動輸入流程"
     },
     {
      "title": "實戰 3：匯出所有工作表為 PDF",
      "time": "⏱ 3 分鐘",
      "code": [
       "Sub ExportAllToPDF()",
       "Dim ws As Worksheet",
       "Dim path As String",
       "path = ThisWorkbook.Path & Application.PathSeparator",
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
      "title": "實戰 4：建立交付資料夾與輸出摘要",
      "time": "⏱ 3 分鐘",
      "code": [
       "Sub PrepareDelivery()",
       "Dim folderPath As String",
       "folderPath = ThisWorkbook.Path & _",
       "Application.PathSeparator & \"exports\"",
       "If Dir(folderPath, vbDirectory) = \"\" Then MkDir folderPath",
       "Range(\"H1\").Value = \"交付摘要\"",
       "Range(\"H2\").Value = \"輸出路徑：\" & folderPath",
       "Range(\"H3\").Value = \"產出日期：\" & Format(Now(), \"yyyy/mm/dd hh:mm\")",
       "MsgBox \"交付資料夾已準備好：\" & vbCrLf & folderPath",
       "End Sub"
      ],
      "tip": "💡 先把輸出路徑、檔名規則與交付摘要整理穩，比直接碰 Windows 專屬郵件自動化更實用"
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
      "desc": "替常用巨集建立啟動入口：工作表按鈕、快速存取工具列，或指定鍵盤快速鍵（Mac 不走 Windows Ribbon 自訂主路線）"
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
    "topics": "RefreshAll / PDF 匯出 / 錯誤處理 / 交付流程",
    "difficulty": "🔴 Lv.5",
    "taskCount": "8 題",
    "time": "40 分鐘",
    "xp": "400 XP"
   }
  },
  "P5-04": {
   "pro": {
    "title": "🏆 綜合挑戰  ─  先選工具，再整合交付",
    "subtitle": "難度：🔴 Lv.5  |  微任務數：5 題  |  每題先判斷工具，再寫公式或流程",
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
      "desc": "公式題：計算每人年資（到 2026/3/31）",
      "time": "2m",
      "hint": "=DATEDIF(D6,DATE(2026,3,31),\"Y\")",
      "answer": "=DATEDIF(D6,DATE(2026,3,31),\"Y\")",
      "explain": "先判斷這是每列日期計算，所以用公式而不是樞紐或 VBA"
     },
     {
      "numLabel": "🔵2",
      "desc": "公式題：薪資等級 >=70K 高級 / >=50K 中級 / 其他初級",
      "time": "3m",
      "hint": "=IFS(E6>=70000,\"高級\",E6>=50000,\"中級\",TRUE,\"初級\")",
      "answer": "=IFS(E6>=70000,\"高級\",E6>=50000,\"中級\",TRUE,\"初級\")",
      "explain": "這是多門檻判斷，用 IFS 比巢狀 IF 更清楚"
     },
     {
      "numLabel": "🟡3",
      "desc": "公式題：年終獎金，績效 >=90 → ×3 / >=75 → ×2 / 其他 → ×1",
      "time": "3m",
      "hint": "=IFS(F6>=90,E6*3,F6>=75,E6*2,TRUE,E6)",
      "answer": "=IFS(F6>=90,E6*3,F6>=75,E6*2,TRUE,E6)",
      "explain": "仍是每列商業規則，用公式欄保留可追溯邏輯"
     },
     {
      "numLabel": "🟡4",
      "desc": "條件統計題：用 SUMIFS 算業務部年終總額",
      "time": "3m",
      "hint": "先完成獎金欄，再 =SUMIFS(獎金欄,部門欄,\"業務\")",
      "answer": "=SUMIFS(I6:I13,C6:C13,\"業務\")",
      "explain": "只問某部門合計，不需要先做樞紐；若要多部門交叉分析才改用樞紐"
     },
     {
      "numLabel": "🔴5",
      "desc": "查找題：用 INDEX+MATCH 查 E005 的姓名",
      "time": "2m",
      "hint": "=INDEX(B6:B13,MATCH(\"E005\",A6:A13,0))",
      "answer": "=INDEX(B6:B13,MATCH(\"E005\",A6:A13,0))",
      "explain": "查找值與回傳欄分開思考；新版環境也可改用 XLOOKUP"
     }
    ]
   },
   "meta": {
    "phase": "🔧 Phase 5：VBA 自動化大師    第 11~14 週  |  目標：能寫程式自動化，能教別人",
    "stage": "第20階段",
    "topics": "工具選擇 / 公式 / 條件統計 / 查找 / 綜合交付",
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
    "taskCount": "8 題",
   "time": "20 分鐘",
   "xp": "150 XP"
  },
	  "📊 基礎函式": {
	   "phase": "🏃 Phase 1：操作效率基礎    第 1~2 週  |  目標：養成鍵盤操作習慣，複習基本功",
	   "stage": "第 2 階段",
	   "topics": "SUM / AVERAGE / COUNT / COUNTA / MAX / MIN / MEDIAN / LARGE / SMALL",
	   "difficulty": "🟢 Lv.1",
	   "taskCount": "8 題",
	   "time": "15 分鐘",
   "xp": "100 XP"
  },
  "📐 條件判斷": {
   "phase": "🏃 Phase 1：操作效率基礎    第 1~2 週  |  目標：養成鍵盤操作習慣，複習基本功",
   "stage": "第 3 階段",
   "topics": "IF / IFS / AND / OR / SWITCH",
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
   "topics": "XLOOKUP / VLOOKUP / INDEX+MATCH / IFERROR",
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
   "topics": "TRIM / CLEAN / MID / DATE / TEXT / DATEDIF",
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
   "topics": "FILTER / SORT / UNIQUE / SEQUENCE / Spill",
   "difficulty": "🟡 Lv.3",
   "taskCount": "8 題",
   "time": "20 分鐘",
   "xp": "200 XP"
  },
  "🧮 進階函式": {
   "phase": "⚡ Phase 4：進階自動化    第 7~10 週  |  目標：處理大量資料、自動化流程",
   "stage": "第14階段",
   "topics": "LET / LAMBDA / AGGREGATE / SWITCH / INDIRECT",
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
   "topics": "RefreshAll / PDF 匯出 / 錯誤處理 / 交付流程",
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
   "desc": "15 分鐘學習 → 5 分鐘休息 → 再做 15 分鐘。一天最多 2~3 個番茄鐘就夠了"
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
