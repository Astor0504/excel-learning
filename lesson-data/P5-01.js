// 由 xlsx-content.js + lesson-demos-data.js 拆分(2026-07-14);本檔是 P5-01 的唯一資料來源
export const CONTENT = {
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
};
export const DEMOS = [
 {
  "id": "vba-safe-first-macro",
  "kind": "workflow",
  "badge": "流程演示",
  "title": "第一個安全巨集：從手動動作到可重跑流程",
  "subtitle": "VBA 入門先追求可理解、可復原，不急著把整份報表自動化。",
  "stages": [
   {
    "text": "手動跑順",
    "tone": "c1"
   },
   {
    "text": "寫入模組",
    "tone": "c2"
   },
   {
    "text": "測試輸出",
    "tone": "c3"
   },
   {
    "text": "安全保存",
    "tone": "c4"
   }
  ],
  "outcome": "結果：巨集只改指定區域，執行前知道會做什麼，執行後能檢查結果。",
  "steps": [
   {
    "stage": 1,
    "label": "1. 先把手動流程說清楚",
    "tone": "c1",
    "text": "先不要寫程式，先列出這段工作每次都固定做什麼：清空輸出區、寫標題、填入公式、提示完成。",
    "result": "巨集範圍先收斂成一段固定、可描述的流程。",
    "caption": "說不清楚的手動流程，寫成巨集後只會更難查錯。",
    "panels": [
     {
      "title": "手動流程",
      "lines": [
       "清空 A1:C10",
       "A1 寫入報表標題",
       "A2 寫入本月業績",
       "A3 寫入成長率公式"
      ]
     }
    ]
   },
   {
    "stage": 2,
    "label": "2. 放進標準模組",
    "tone": "c2",
    "text": "在 VBE 插入模組，把每個動作用清楚的 Range 寫出來。初學先避免 Select/Activate，減少焦點錯誤。",
    "result": "每個動作都指定明確範圍，降低誤改風險。",
    "caption": "Range(\"A1\").Value 比先選 A1 再輸入更穩。",
    "panels": [
     {
      "title": "巨集片段",
      "lines": [
       "Sub 建立月報()",
       "Range(\"A1:C10\").ClearContents",
       "Range(\"A1\").Value = \"本月業績摘要\"",
       "Range(\"A3\").Formula = \"=A2*1.05\"",
       "End Sub"
      ]
     }
    ]
   },
   {
    "stage": 3,
    "label": "3. 在測試檔執行",
    "tone": "c3",
    "text": "第一次執行請用測試檔或備份檔，確認它只改預期範圍，沒有清掉原始資料。",
    "result": "執行前後可對照，確認只有指定區域被改。",
    "caption": "VBA 的 Undo 通常不可靠，測試檔是基本保護。",
    "panels": [
     {
      "title": "執行前/後",
      "columns": [
       "區域",
       "執行前",
       "執行後"
      ],
      "rows": [
       [
        "A1",
        "空白",
        "本月業績摘要"
       ],
       [
        "A2",
        "100000",
        "100000"
       ],
       [
        "A3",
        "空白",
        "=A2*1.05"
       ]
      ]
     }
    ]
   },
   {
    "stage": 4,
    "label": "4. 存成 .xlsm 並註明用途",
    "tone": "c4",
    "text": "巨集檔要存成 .xlsm，並在工作表或模組註解中寫清楚用途、會修改的範圍與執行前注意事項。",
    "result": "可交付版本完成：副檔名、用途與安全提醒都清楚。",
    "caption": "可交付的巨集不是只有能跑，還要讓下一個人敢按。",
    "panels": [
     {
      "title": "安全提醒",
      "lines": [
       "保留原始資料備份",
       "巨集只處理指定工作表",
       "交付前關閉不必要的外部連結",
       "檔案副檔名使用 .xlsm"
      ]
     }
    ]
   }
  ]
 }
];
