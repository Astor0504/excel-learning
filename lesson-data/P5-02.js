// 由 xlsx-content.js + lesson-demos-data.js 拆分(2026-07-14);本檔是 P5-02 的唯一資料來源
export const CONTENT = {
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
};
export const DEMOS = [
 {
  "id": "vba-array-dictionary-safety",
  "kind": "workflow",
  "badge": "流程演示",
  "title": "進階 VBA：一次讀入、記憶體處理、安全收尾",
  "subtitle": "速度不是靠更快點儲存格，而是少跟工作表來回溝通。",
  "stages": [
   {
    "text": "一次讀入",
    "tone": "c1"
   },
   {
    "text": "記憶體處理",
    "tone": "c2"
   },
   {
    "text": "一次寫回",
    "tone": "c3"
   },
   {
    "text": "錯誤出口",
    "tone": "c4"
   }
  ],
  "outcome": "結果：大量資料用陣列和 Dictionary 處理，速度更穩，也能在錯誤時恢復狀態。",
  "steps": [
   {
    "stage": 1,
    "label": "1. 把資料一次讀進陣列",
    "tone": "c1",
    "text": "把 A2:C10000 一次讀入 Variant 陣列，避免每一列都跨 Excel 物件模型讀取。",
    "result": "讀取動作從 10,000 次降成 1 次。",
    "caption": "慢的通常不是迴圈本身，而是一直 Cells(i,j) 讀寫工作表。",
    "panels": [
     {
      "title": "讀入策略",
      "lines": [
       "Dim arr As Variant",
       "arr = Range(\"A2:C10000\").Value",
       "後續迴圈都處理 arr，不直接碰儲存格"
      ]
     }
    ]
   },
   {
    "stage": 2,
    "label": "2. 用 Dictionary 做彙總或去重",
    "tone": "c2",
    "text": "在記憶體裡用 Dictionary 依部門加總、計數或去重，最後再整理成輸出陣列。",
    "result": "電子部兩筆資料在記憶體中累計成 208000。",
    "caption": "Dictionary 適合 key-value 類問題，例如每個部門累計多少金額。",
    "panels": [
     {
      "title": "處理邏輯",
      "columns": [
       "資料列",
       "部門",
       "金額",
       "Dictionary"
      ],
      "rows": [
       [
        "2",
        "電子",
        "120000",
        "電子=120000"
       ],
       [
        "3",
        "業務",
        "98000",
        "業務=98000"
       ],
       [
        "4",
        "電子",
        "88000",
        "電子=208000"
       ]
      ]
     }
    ]
   },
   {
    "stage": 3,
    "label": "3. 結果一次寫回",
    "tone": "c3",
    "text": "把輸出整理成二維陣列後，一次寫回結果區。這比在迴圈裡逐格寫入快很多。",
    "result": "輸出區一次取得部門、總金額與筆數。",
    "caption": "讀一次、算一次、寫一次，是大量資料巨集的基本節奏。",
    "panels": [
     {
      "title": "輸出結果",
      "columns": [
       "部門",
       "總金額",
       "筆數"
      ],
      "rows": [
       [
        "電子",
        "208000",
        "2"
       ],
       [
        "業務",
        "98000",
        "1"
       ],
       [
        "行政",
        "76000",
        "1"
       ]
      ]
     }
    ]
   },
   {
    "stage": 4,
    "label": "4. 加上錯誤處理與狀態恢復",
    "tone": "c4",
    "text": "關閉畫面更新或自動計算後，一定要在錯誤出口恢復設定，避免使用者的 Excel 狀態被巨集留壞。",
    "result": "就算中途失敗，也會回到安全的 Excel 狀態。",
    "caption": "進階巨集的安全感，常常來自完整的 CleanUp 區塊。",
    "panels": [
     {
      "title": "安全骨架",
      "lines": [
       "On Error GoTo ErrHandler",
       "Application.ScreenUpdating = False",
       "Application.EnableEvents = False",
       "主要流程...",
       "CleanUp: 復原 ScreenUpdating / EnableEvents",
       "ErrHandler: 顯示錯誤並跳回 CleanUp"
      ]
     }
    ]
   }
  ]
 }
];
