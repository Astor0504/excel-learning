// 由 xlsx-content.js + lesson-demos-data.js 拆分(2026-07-14);本檔是 P5-03 的唯一資料來源
export const CONTENT = {
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
};
export const DEMOS = [
 {
  "id": "vba-report-pipeline",
  "kind": "workflow",
  "badge": "專案演示",
  "title": "一鍵月報：刷新、檢查、輸出、紀錄",
  "subtitle": "實戰不是把程式寫長，而是讓每一段責任清楚、失敗時可追蹤。",
  "stages": [
   {
    "text": "刷新資料",
    "tone": "c1"
   },
   {
    "text": "檢查結果",
    "tone": "c2"
   },
   {
    "text": "輸出 PDF",
    "tone": "c3"
   },
   {
    "text": "留下紀錄",
    "tone": "c4"
   }
  ],
  "outcome": "結果：VBA 負責觸發和交付，Power Query 與樞紐負責資料處理。",
  "steps": [
   {
    "stage": 1,
    "label": "1. RefreshAll 觸發可重跑流程",
    "tone": "c1",
    "text": "清資料邏輯先放在 Power Query，VBA 只呼叫 ThisWorkbook.RefreshAll，避免把清理規則散落在巨集裡。",
    "result": "資料更新責任留給 Power Query，VBA 只負責啟動。",
    "caption": "這樣下個月資料來了，改 Query 比改 VBA 更好維護。",
    "panels": [
     {
      "title": "責任切分",
      "columns": [
       "階段",
       "負責工具",
       "輸出"
      ],
      "rows": [
       [
        "清資料",
        "Power Query",
        "乾淨資料表"
       ],
       [
        "彙總",
        "樞紐/公式",
        "月報工作表"
       ],
       [
        "觸發",
        "VBA",
        "一鍵更新"
       ]
      ]
     }
    ]
   },
   {
    "stage": 2,
    "label": "2. 檢查必要工作表與輸出區",
    "tone": "c2",
    "text": "輸出前先檢查月報工作表是否存在、關鍵儲存格是否有錯誤值、資料筆數是否合理。",
    "result": "錯誤在輸出 PDF 前被攔下，而不是寄出後才發現。",
    "caption": "自動化流程最怕安靜失敗，所以檢查步驟不能省。",
    "panels": [
     {
      "title": "檢查表",
      "columns": [
       "檢查項目",
       "期待",
       "失敗處理"
      ],
      "rows": [
       [
        "月報工作表",
        "存在",
        "提示缺工作表"
       ],
       [
        "錯誤值",
        "0 個",
        "停止輸出"
       ],
       [
        "資料筆數",
        "> 0",
        "提示來源異常"
       ]
      ]
     }
    ]
   },
   {
    "stage": 3,
    "label": "3. 輸出 PDF 到固定路徑",
    "tone": "c3",
    "text": "用 ExportAsFixedFormat 把月報工作表輸出成 PDF，檔名帶上年月，讓交付物可追蹤。",
    "result": "交付檔名固定成 月報_yyyymm.pdf。",
    "caption": "路徑與檔名規則固定，使用者才知道輸出去哪裡。",
    "panels": [
     {
      "title": "輸出前/後",
      "columns": [
       "項目",
       "執行前",
       "執行後"
      ],
      "rows": [
       [
        "PDF",
        "尚未產生",
        "月報_202605.pdf"
       ],
       [
        "路徑",
        "工作簿資料夾",
        "同資料夾輸出"
       ],
       [
        "提示",
        "無",
        "顯示輸出位置"
       ]
      ]
     }
    ]
   },
   {
    "stage": 4,
    "label": "4. CleanUp 復原設定並留紀錄",
    "tone": "c4",
    "text": "不管成功或失敗，都要復原 ScreenUpdating / EnableEvents，並留下執行時間、狀態、錯誤訊息。",
    "result": "流程可追蹤，也不會把使用者 Excel 狀態留壞。",
    "caption": "可交付自動化的重點，是下一次出錯時找得到線索。",
    "panels": [
     {
      "title": "執行紀錄",
      "columns": [
       "時間",
       "狀態",
       "訊息"
      ],
      "rows": [
       [
        "2026/05/03 12:00",
        "成功",
        "PDF 已輸出"
       ],
       [
        "2026/05/03 12:05",
        "失敗",
        "找不到月報工作表"
       ]
      ]
     }
    ]
   }
  ]
 }
];
