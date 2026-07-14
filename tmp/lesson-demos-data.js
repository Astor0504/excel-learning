export const LESSON_DEMOS = {
  "P1-01": [
    {
      id: "mac-shortcut-navigation",
      kind: "workflow",
      badge: "操作節奏",
      title: "快捷鍵不是背表：從定位、選取到填滿的 20 秒流程",
      subtitle: "先練一條最常用的鍵盤路徑，手感會比單背快捷鍵更快建立起來。",
      stages: [
        { text: "定位資料", tone: "c1" },
        { text: "選取區域", tone: "c2" },
        { text: "插入新列", tone: "c3" },
        { text: "填滿公式", tone: "c4" }
      ],
      outcome: "結果：不用離開鍵盤，也能完成跳邊界、選資料、插入列、填滿公式。",
      steps: [
        {
          stage: 1,
          label: "1. 用 ⌘ + 方向鍵跳到資料邊界",
          tone: "c1",
          text: "從資料區任一格出發，⌘ + ↓ 會跳到目前連續資料區的底部，⌘ + → 會跳到右側邊界。",
          result: "游標快速抵達資料邊界，不用滾動找最後一列。",
          caption: "新版 Excel for Mac 通常用 ⌘ + 方向鍵；舊版可試 ⌃ + 方向鍵。",
          panels: [
            { title: "鍵盤節奏", lines: ["起點：A2", "按 ⌘ + ↓：跳到最後一筆資料", "按 ⌘ + →：跳到資料區右側"] }
          ]
        },
        {
          stage: 2,
          label: "2. 用 Shift 延伸選取範圍",
          tone: "c2",
          text: "搭配 Shift，可以從目前位置延伸選取到資料邊界，快速圈出需要格式化或複製的整段資料。",
          result: "資料區被一次選起來，後續操作不必用滑鼠拖拉。",
          caption: "如果要選整欄，用 ⌃ Space；選整列，用 ⇧ Space。",
          panels: [
            { title: "選取策略", columns: ["目的", "快捷鍵", "結果"], rows: [["選整欄", "⌃ Space", "目前欄整欄選取"], ["選整列", "⇧ Space", "目前列整列選取"], ["延伸到邊界", "⇧ + ⌘ + 方向鍵", "選到資料邊界"]] }
          ]
        },
        {
          stage: 3,
          label: "3. 插入或刪除列欄",
          tone: "c3",
          text: "選好整列或整欄後，用 ⌃ ⇧ + 插入，用 ⌘ - 刪除。這是清資料時最常用的一組動作。",
          result: "新增列欄不需要右鍵選單，節奏更連續。",
          caption: "WPS 的插入快捷鍵可能不同，這課以 Excel for Mac 為主。",
          panels: [
            { title: "列欄動作", lines: ["先選整列或整欄", "⌃ ⇧ +：插入", "⌘ -：刪除", "⌘ Z：復原"] }
          ]
        },
        {
          stage: 4,
          label: "4. 用 ⌃ D 填滿向下",
          tone: "c4",
          text: "公式寫好一格後，選取下方需要填滿的範圍，按 ⌃ D 就能把上方公式向下填滿。",
          result: "一整欄公式快速填滿，並保留相對參照邏輯。",
          caption: "填滿前先檢查第一格公式是否正確，錯一格會一路錯到底。",
          panels: [
            { title: "填滿前/後", columns: ["位置", "填滿前", "填滿後"], rows: [["D2", "=B2*C2", "=B2*C2"], ["D3", "空白", "=B3*C3"], ["D4", "空白", "=B4*C4"]] }
          ]
        }
      ]
    }
  ],
  "P1-02": [
    {
      id: "average-zero-blank",
      kind: "formula",
      badge: "概念拆解",
      title: "AVERAGE：0 會算進去，空白和文字會略過",
      subtitle: "這是本課最容易誤會的地方。平均值看起來怪怪的，先不要怪公式，先檢查資料內容。",
      outcome: "結果：(120,000 + 0 + 95,000) ÷ 3 = 71,667",
      formulaParts: [
        { text: "=AVERAGE(" },
        { text: "B2:B6", step: 1, tone: "c1" },
        { text: ")" }
      ],
      board: {
        title: "同一欄裡混了數字、0、空白、文字",
        columns: ["業績欄內容", "COUNT", "COUNTA", "AVERAGE"],
        rows: [
          [
            { text: "120,000", marks: { 1: "focus", 2: "match", 3: "match" } },
            { text: "算", marks: { 2: "result" } },
            { text: "算", marks: { 2: "result" } },
            { text: "算", marks: { 3: "result" } }
          ],
          [
            { text: "0", marks: { 1: "focus", 2: "match", 3: "match" } },
            { text: "算", marks: { 2: "result" } },
            { text: "算", marks: { 2: "result" } },
            { text: "算", marks: { 3: "result" } }
          ],
          [
            { text: "（空白）", marks: { 1: "focus" } },
            { text: "不算", marks: { 2: "result" } },
            { text: "不算", marks: { 2: "result" } },
            { text: "不算", marks: { 3: "result" } }
          ],
          [
            { text: "95,000", marks: { 1: "focus", 2: "match", 3: "match" } },
            { text: "算", marks: { 2: "result" } },
            { text: "算", marks: { 2: "result" } },
            { text: "算", marks: { 3: "result" } }
          ],
          [
            { text: "待補", marks: { 1: "focus" } },
            { text: "不算", marks: { 2: "result" } },
            { text: "算", marks: { 2: "result" } },
            { text: "不算", marks: { 3: "result" } }
          ]
        ]
      },
      steps: [
        {
          label: "1. 先看整段範圍",
          tone: "c1",
          text: "AVERAGE 第一件事只是先圈出要檢查的範圍。Excel 會在這整段裡，自己判斷哪些值要參與平均。",
          caption: "先記住：範圍裡的每格不一定都會真的被算進平均。"
        },
        {
          label: "2. COUNT / COUNTA 幫你辨認資料型態",
          tone: "c2",
          text: "COUNT 只數數字；COUNTA 數所有非空白。這一步最重要的觀念是：0 是數字，文字「待補」不是，空白更不是。",
          caption: "很多報表其實不是公式錯，而是你以為 0 跟空白是一樣的。"
        },
        {
          label: "3. AVERAGE 只拿數字做平均，但 0 不會被排除",
          tone: "c3",
          text: "AVERAGE 會忽略空白和文字，但 0 會當成真正的數字算進去。所以這組資料是用 120,000、0、95,000 三個值去平均。",
          caption: "只要格子裡是 0，它就會把平均拉低；想排除 0，要改用 AVERAGEIF。"
        }
      ]
    },
    {
      id: "large-ranking",
      kind: "formula",
      badge: "公式拆解",
      title: "LARGE：當你要的是名次，不是單純最大值",
      subtitle: "MAX 只會給你冠軍；要第 2 名、第 3 名，才輪到 LARGE 上場。",
      outcome: "結果：150,000 > 120,000 > 95,000，所以第 3 大是 95,000",
      formulaParts: [
        { text: "=LARGE(" },
        { text: "B2:B6", step: 1, tone: "c1" },
        { text: ", " },
        { text: "3", step: 2, tone: "c2" },
        { text: ")" }
      ],
      board: {
        title: "業績表",
        columns: ["業務員", "業績"],
        rows: [
          [
            { text: "王小明" },
            { text: "120,000", marks: { 1: "focus", 3: "match" } }
          ],
          [
            { text: "林美玲" },
            { text: "85,000", marks: { 1: "focus" } }
          ],
          [
            { text: "張志偉" },
            { text: "95,000", marks: { 1: "focus", 2: "result", 3: "result" } }
          ],
          [
            { text: "陳雅婷" },
            { text: "150,000", marks: { 1: "focus", 3: "match" } }
          ],
          [
            { text: "李大華" },
            { text: "72,000", marks: { 1: "focus" } }
          ]
        ]
      },
      steps: [
        {
          label: "1. 範圍：先決定在哪一群數字裡排名",
          tone: "c1",
          text: "B2:B6 是參與排名的資料範圍。LARGE 只會在這群數字裡找名次，不會跨到別欄去想。",
          caption: "非數字內容會被忽略，但你的範圍還是要抓對。"
        },
        {
          label: "2. 名次：3 代表第 3 大，不是第 3 筆",
          tone: "c2",
          text: "第二個參數的意思是『第幾名』。填 1 = 最大值，填 2 = 第 2 大，填 3 = 第 3 大。",
          caption: "問冠軍時用 MAX 就夠了；問前 3 名時才改用 LARGE。"
        },
        {
          label: "3. 讀結果：第 3 大是排序後的第 3 個值",
          tone: "c3",
          text: "這組資料排序後依序是 150,000、120,000、95,000、85,000、72,000，所以公式回傳 95,000。",
          caption: "如果你要看倒數第 2 名、倒數第 3 名，把 LARGE 換成 SMALL 即可。"
        }
      ]
    }
  ],
  "P1-03": [
    {
      id: "ifs-grade-flow",
      kind: "formula",
      badge: "公式拆解",
      title: "IFS：多個門檻要由高到低檢查",
      subtitle: "IFS 會從左到右找第一個 TRUE。一旦命中，就回傳結果，不會再看後面的條件。",
      outcome: "結果：85 >= 90 是 FALSE，85 >= 80 是 TRUE，所以回傳 B",
      formulaParts: [
        { text: "=IFS(" },
        { text: "B2>=90,\"A\"", step: 1, tone: "c1" },
        { text: ", " },
        { text: "B2>=80,\"B\"", step: 2, tone: "c2" },
        { text: ", " },
        { text: "B2>=70,\"C\"", step: 3, tone: "c3" },
        { text: ", " },
        { text: "TRUE,\"F\"", step: 4, tone: "c4" },
        { text: ")" }
      ],
      board: {
        title: "85 分的判斷流程",
        columns: ["分數", ">=90", ">=80", ">=70", "結果"],
        rows: [
          [
            { text: "85", marks: { 1: "focus", 2: "focus", 3: "focus", 4: "focus" } },
            { text: "FALSE", marks: { 1: "result" } },
            { text: "TRUE", marks: { 2: "match" } },
            { text: "跳過", marks: { 3: "result" } },
            { text: "B", marks: { 2: "result", 3: "result", 4: "result" } }
          ]
        ]
      },
      steps: [
        {
          label: "1. 先檢查最高門檻",
          tone: "c1",
          text: "85 沒有大於等於 90，所以第一組條件不成立。Excel 會繼續往右看下一組條件。",
          caption: "門檻判斷通常要從高到低寫，避免高分先被低門檻吃掉。"
        },
        {
          label: "2. 第二個條件命中",
          tone: "c2",
          text: "85 大於等於 80，這一組條件成立，所以 IFS 立即回傳 B。",
          caption: "IFS 的重點是一命中就結束，不會把後面所有條件都跑完。"
        },
        {
          label: "3. 後面的條件不再執行",
          tone: "c3",
          text: "雖然 85 也大於等於 70，但前面已經命中 B，這一段不會再影響結果。",
          caption: "這就是為什麼條件順序會直接改變答案。"
        },
        {
          label: "4. TRUE 是最後兜底",
          tone: "c4",
          text: "最後的 TRUE 等於 else。只有前面全部沒命中時，才會回傳這裡的 F。",
          caption: "沒有兜底條件時，IFS 找不到符合條件會出錯。"
        }
      ]
    },
    {
      id: "switch-department-map",
      kind: "formula",
      badge: "概念拆解",
      title: "SWITCH：固定值對應，比一串 IF 更清楚",
      subtitle: "當你不是判斷範圍，而是把一個固定值翻成另一個固定結果，SWITCH 會比 IFS 更乾淨。",
      outcome: "結果：B2 是「財務」，命中第二組對應，所以回傳 Finance",
      formulaParts: [
        { text: "=SWITCH(" },
        { text: "B2", step: 1, tone: "c1" },
        { text: ", " },
        { text: "\"業務\",\"Sales\"", step: 2, tone: "c2" },
        { text: ", " },
        { text: "\"財務\",\"Finance\"", step: 3, tone: "c3" },
        { text: ", " },
        { text: "\"資訊\",\"IT\"", step: 4, tone: "c4" },
        { text: ", \"其他\")" }
      ],
      board: {
        title: "部門名稱對應",
        columns: ["B2", "業務?", "財務?", "後續", "結果"],
        rows: [
          [
            { text: "財務", marks: { 1: "focus", 2: "focus", 3: "focus", 4: "focus" } },
            { text: "否", marks: { 2: "result" } },
            { text: "是", marks: { 3: "match" } },
            { text: "跳過", marks: { 4: "result" } },
            { text: "Finance", marks: { 3: "result", 4: "result" } }
          ]
        ]
      },
      steps: [
        {
          label: "1. 先拿 B2 當比對值",
          tone: "c1",
          text: "SWITCH 的第一個參數是要被拿來比對的值。這裡 B2 的內容是「財務」。",
          caption: "和 IFS 不同，SWITCH 不是每組都寫一個完整條件。"
        },
        {
          label: "2. 第一組對應沒有命中",
          tone: "c2",
          text: "Excel 先看 B2 是否等於「業務」。因為不是，所以不回傳 Sales，繼續看下一組。",
          caption: "每一組都是「要比對的值、命中後的結果」。"
        },
        {
          label: "3. 第二組對應命中",
          tone: "c3",
          text: "B2 等於「財務」，所以公式回傳 Finance。",
          caption: "這類部門、狀態、代碼對照就是 SWITCH 的舒適區。"
        },
        {
          label: "4. 後面對應與預設值不再影響結果",
          tone: "c4",
          text: "因為已經命中財務，後面的資訊與其他都不會再改變結果。",
          caption: "如果 B2 是完全沒列出的值，最後的「其他」才會出場。"
        }
      ]
    }
  ],
  "P2-01": [
    {
      id: "conditional-stats-decision-path",
      kind: "formula",
      role: "primary",
      badge: "主線示範",
      title: "條件統計決策：先選公式，再檢查命中列",
      subtitle: "不用一次背六個函數。先判斷要算什麼、條件幾個，再把命中的列加總或計數。",
      outcome: "結果：這題是多條件加總，所以用 SUMIFS；命中 4 筆已付款電子訂單，合計 86,500。",
      formulaParts: [
        { text: "問題：電子且已付的金額合計" },
        { text: " → 算合計", step: 1, tone: "c1" },
        { text: " → 2 個條件", step: 2, tone: "c2" },
        { text: " → SUMIFS", step: 3, tone: "c3" },
        { text: " → 驗算命中列", step: 4, tone: "c4" }
      ],
      board: {
        title: "從問題到命中列",
        columns: ["類別", "付款", "金額", "判斷"],
        rows: [
          [{ text: "電子", marks: { 2: "candidate", 3: "candidate" } }, { text: "已付", marks: { 2: "candidate", 3: "match" } }, { text: "15,000", marks: { 1: "focus", 4: "result" } }, { text: "命中", marks: { 4: "result" } }],
          [{ text: "服飾", marks: { 2: "exclude", 3: "exclude" } }, { text: "未付", marks: { 2: "focus", 3: "exclude" } }, { text: "8,500", marks: { 1: "focus" } }, { text: "排除", marks: { 4: "exclude" } }],
          [{ text: "電子", marks: { 2: "candidate", 3: "candidate" } }, { text: "已付", marks: { 2: "candidate", 3: "match" } }, { text: "22,000", marks: { 1: "focus", 4: "result" } }, { text: "命中", marks: { 4: "result" } }],
          [{ text: "服飾", marks: { 2: "exclude", 3: "exclude" } }, { text: "已付", marks: { 2: "focus", 3: "match" } }, { text: "12,000", marks: { 1: "focus" } }, { text: "排除", marks: { 4: "exclude" } }],
          [{ text: "電子", marks: { 2: "candidate", 3: "candidate" } }, { text: "未付", marks: { 2: "focus", 3: "exclude" } }, { text: "9,800", marks: { 1: "focus" } }, { text: "排除", marks: { 4: "exclude" } }],
          [{ text: "電子", marks: { 2: "candidate", 3: "candidate" } }, { text: "已付", marks: { 2: "candidate", 3: "match" } }, { text: "31,000", marks: { 1: "focus", 4: "result" } }, { text: "命中", marks: { 4: "result" } }],
          [{ text: "電子", marks: { 2: "candidate", 3: "candidate" } }, { text: "已付", marks: { 2: "candidate", 3: "match" } }, { text: "18,500", marks: { 1: "focus", 4: "result" } }, { text: "命中", marks: { 4: "result" } }],
          [{ text: "服飾", marks: { 2: "exclude", 3: "exclude" } }, { text: "已付", marks: { 2: "focus", 3: "match" } }, { text: "7,200", marks: { 1: "focus" } }, { text: "排除", marks: { 4: "exclude" } }]
        ]
      },
      steps: [
        {
          label: "1. 先問：要算什麼",
          tone: "c1",
          text: "題目問的是「金額合計」，所以不是 COUNT，也不是 AVERAGE，而是 SUM 家族。",
          caption: "幾筆用 COUNT；合計用 SUM；平均用 AVERAGE。"
        },
        {
          label: "2. 再問：條件幾個",
          tone: "c2",
          text: "這題同時要求「電子」和「已付」，所以是多條件。多條件版本要用結尾有 S 的函數。",
          caption: "COUNTIFS / SUMIFS / AVERAGEIFS 的條件預設都是 AND。"
        },
        {
          label: "3. 選公式並確認順序",
          tone: "c3",
          text: "多條件加總用 SUMIFS，而且加總範圍要放第一個，後面才是一組一組條件範圍與條件值。",
          caption: "SUMIF 是條件範圍在前；SUMIFS 是加總範圍在前。"
        },
        {
          label: "4. 驗算命中列",
          tone: "c4",
          text: "最後回到資料表，只把同時符合電子與已付的金額加起來。四筆命中列合計 86,500。",
          caption: "正式報表交付前，至少手動篩一次命中列確認公式沒有抓錯範圍。"
        }
      ]
    },
    {
      id: "countif-region",
      kind: "formula",
      role: "supporting",
      badge: "計數",
      title: "COUNTIF：符合條件的筆數怎麼數出來",
      subtitle: "問「有幾筆？」時，先用 COUNTIF 數條件欄裡命中的儲存格。",
      outcome: "結果：北區命中 4 格，所以回傳 4",
      formulaParts: [
        { text: "=COUNTIF(" },
        { text: "C2:C9", step: 1, tone: "c1" },
        { text: ", " },
        { text: "\"北區\"", step: 2, tone: "c2" },
        { text: ")" }
      ],
      board: {
        title: "訂單資料",
        columns: ["業務員", "區域", "類別", "金額"],
        rows: [
          [{ text: "王小明" }, { text: "北區", marks: { 1: "focus", 2: "match", 3: "result" } }, { text: "電子" }, { text: "15,000" }],
          [{ text: "林美華" }, { text: "南區", marks: { 1: "focus", 2: "exclude" } }, { text: "服飾" }, { text: "8,500" }],
          [{ text: "張志偉" }, { text: "北區", marks: { 1: "focus", 2: "match", 3: "result" } }, { text: "電子" }, { text: "22,000" }],
          [{ text: "王小明" }, { text: "北區", marks: { 1: "focus", 2: "match", 3: "result" } }, { text: "服飾" }, { text: "12,000" }],
          [{ text: "陳雅婷" }, { text: "東區", marks: { 1: "focus", 2: "exclude" } }, { text: "電子" }, { text: "9,800" }],
          [{ text: "林美華" }, { text: "南區", marks: { 1: "focus", 2: "exclude" } }, { text: "電子" }, { text: "31,000" }],
          [{ text: "張志偉" }, { text: "北區", marks: { 1: "focus", 2: "match", 3: "result" } }, { text: "電子" }, { text: "18,500" }],
          [{ text: "陳雅婷" }, { text: "東區", marks: { 1: "focus", 2: "exclude" } }, { text: "服飾" }, { text: "7,200" }]
        ]
      },
      steps: [
        { label: "1. 範圍：檢查區域欄", tone: "c1", text: "第一個參數是要檢查的範圍。這裡只看區域欄 C2:C9，不需要把金額欄圈進來。", caption: "COUNTIF 是數格數，不是加總。" },
        { label: "2. 條件：北區", tone: "c2", text: "條件填「北區」後，Excel 逐格比對。等於北區的留下，不等於北區的排除。", caption: "文字條件要放在引號裡。" },
        { label: "3. 結果：命中幾格就回傳幾", tone: "c3", text: "北區一共命中 4 格，所以公式回傳 4。", caption: "問幾筆、幾人、幾次時，先想到 COUNTIF / COUNTIFS。" }
      ]
    },
    {
      id: "sumif-category",
      kind: "formula",
      role: "supporting",
      badge: "加總",
      title: "SUMIF：單一條件加總，條件範圍在前",
      subtitle: "SUMIF 先找符合條件的列，再去加總同列的數字。",
      outcome: "結果：15,000 + 22,000 + 9,800 + 31,000 + 18,500 = 96,300",
      formulaParts: [
        { text: "=SUMIF(" },
        { text: "D2:D9", step: 1, tone: "c1" },
        { text: ", " },
        { text: "\"電子\"", step: 2, tone: "c2" },
        { text: ", " },
        { text: "E2:E9", step: 3, tone: "c3" },
        { text: ")" }
      ],
      board: {
        title: "依產品類別加總",
        columns: ["產品類別", "金額"],
        rows: [
          [{ text: "電子", marks: { 1: "focus", 2: "match" } }, { text: "15,000", marks: { 3: "result", 4: "result" } }],
          [{ text: "服飾", marks: { 1: "focus", 2: "exclude" } }, { text: "8,500", marks: { 3: "focus" } }],
          [{ text: "電子", marks: { 1: "focus", 2: "match" } }, { text: "22,000", marks: { 3: "result", 4: "result" } }],
          [{ text: "服飾", marks: { 1: "focus", 2: "exclude" } }, { text: "12,000", marks: { 3: "focus" } }],
          [{ text: "電子", marks: { 1: "focus", 2: "match" } }, { text: "9,800", marks: { 3: "result", 4: "result" } }],
          [{ text: "電子", marks: { 1: "focus", 2: "match" } }, { text: "31,000", marks: { 3: "result", 4: "result" } }],
          [{ text: "電子", marks: { 1: "focus", 2: "match" } }, { text: "18,500", marks: { 3: "result", 4: "result" } }],
          [{ text: "服飾", marks: { 1: "focus", 2: "exclude" } }, { text: "7,200", marks: { 3: "focus" } }]
        ]
      },
      steps: [
        { label: "1. 條件範圍", tone: "c1", text: "SUMIF 第一個參數是條件範圍。這裡先看產品類別欄。", caption: "單條件 SUMIF 是先條件、後加總。" },
        { label: "2. 條件值", tone: "c2", text: "條件填「電子」後，只有產品類別等於電子的列會留下。", caption: "其他類別不參與最後加總。" },
        { label: "3. 加總範圍", tone: "c3", text: "第三個參數才是要加總的金額欄。Excel 只會加總剛剛命中的同列金額。", caption: "這是 SUMIF 和 COUNTIF 最大的不同。" },
        { label: "4. 結果", tone: "c4", text: "把電子類別命中的金額加起來，得到 96,300。", caption: "問合計多少時，用 SUMIF / SUMIFS。" }
      ]
    },
    {
      id: "sumifs-category-status",
      kind: "formula",
      role: "supporting",
      badge: "多條件加總",
      title: "SUMIFS：加總範圍放第一個",
      subtitle: "多條件加總先說要加總哪一欄，再一組一組寫條件。",
      outcome: "結果：15,000 + 22,000 + 31,000 + 18,500 = 86,500",
      formulaParts: [
        { text: "=SUMIFS(" },
        { text: "E2:E9", step: 1, tone: "c1" },
        { text: ", " },
        { text: "D2:D9", step: 2, tone: "c2" },
        { text: ", " },
        { text: "\"電子\"", step: 3, tone: "c3" },
        { text: ", " },
        { text: "F2:F9", step: 4, tone: "c4" },
        { text: ", " },
        { text: "\"已付\"", step: 5, tone: "c5" },
        { text: ")" }
      ],
      board: {
        title: "電子且已付款才加總",
        columns: ["類別", "金額", "付款"],
        rows: [
          [{ text: "電子", marks: { 2: "focus", 3: "candidate" } }, { text: "15,000", marks: { 1: "focus", 5: "result" } }, { text: "已付", marks: { 4: "focus", 5: "match" } }],
          [{ text: "服飾", marks: { 2: "focus", 3: "exclude" } }, { text: "8,500", marks: { 1: "focus" } }, { text: "未付", marks: { 4: "focus", 5: "exclude" } }],
          [{ text: "電子", marks: { 2: "focus", 3: "candidate" } }, { text: "22,000", marks: { 1: "focus", 5: "result" } }, { text: "已付", marks: { 4: "focus", 5: "match" } }],
          [{ text: "服飾", marks: { 2: "focus", 3: "exclude" } }, { text: "12,000", marks: { 1: "focus" } }, { text: "已付", marks: { 4: "focus", 5: "match" } }],
          [{ text: "電子", marks: { 2: "focus", 3: "candidate" } }, { text: "9,800", marks: { 1: "focus", 5: "exclude" } }, { text: "未付", marks: { 4: "focus", 5: "exclude" } }],
          [{ text: "電子", marks: { 2: "focus", 3: "candidate" } }, { text: "31,000", marks: { 1: "focus", 5: "result" } }, { text: "已付", marks: { 4: "focus", 5: "match" } }],
          [{ text: "電子", marks: { 2: "focus", 3: "candidate" } }, { text: "18,500", marks: { 1: "focus", 5: "result" } }, { text: "已付", marks: { 4: "focus", 5: "match" } }],
          [{ text: "服飾", marks: { 2: "focus", 3: "exclude" } }, { text: "7,200", marks: { 1: "focus" } }, { text: "已付", marks: { 4: "focus", 5: "match" } }]
        ]
      },
      steps: [
        { label: "1. 加總範圍", tone: "c1", text: "SUMIFS 的第一個參數永遠是要加總的數字欄。這點和 SUMIF 順序不同。", caption: "先說最後要加哪一欄，再寫條件。" },
        { label: "2. 第一個條件範圍", tone: "c2", text: "接著指定第一個條件要檢查的欄位：產品類別。", caption: "條件範圍和條件值要成對出現。" },
        { label: "3. 第一個條件：電子", tone: "c3", text: "產品類別等於電子的列先成為候選資料，服飾列排除。", caption: "這一步只是第一輪篩選，還沒完成。" },
        { label: "4. 第二個條件範圍", tone: "c4", text: "第二組條件改看付款狀態欄。", caption: "多條件可以一直往後接，但每個範圍大小要一致。" },
        { label: "5. 第二個條件：已付", tone: "c5", text: "同時是電子且已付的列才會被加總；電子但未付也會被排除。", caption: "SUMIFS 的多條件預設是 AND，全部成立才算。" }
      ]
    },
    {
      id: "countifs-region-amount",
      kind: "formula",
      role: "supporting",
      badge: "多條件計數",
      title: "COUNTIFS：多條件是 AND 關係",
      subtitle: "COUNTIFS 不加總金額，只數同時滿足所有條件的列數。",
      outcome: "結果：北區且金額 >= 15000 的訂單有 3 筆",
      formulaParts: [
        { text: "=COUNTIFS(" },
        { text: "C2:C9", step: 1, tone: "c1" },
        { text: ", \"北區\", " },
        { text: "E2:E9", step: 2, tone: "c2" },
        { text: ", \">=15000\")", step: 3, tone: "c3" }
      ],
      board: {
        title: "北區且達標",
        columns: ["區域", "金額"],
        rows: [
          [{ text: "北區", marks: { 1: "candidate" } }, { text: "15,000", marks: { 2: "focus", 3: "result" } }],
          [{ text: "南區", marks: { 1: "exclude" } }, { text: "8,500", marks: { 2: "focus", 3: "exclude" } }],
          [{ text: "北區", marks: { 1: "candidate" } }, { text: "22,000", marks: { 2: "focus", 3: "result" } }],
          [{ text: "北區", marks: { 1: "candidate" } }, { text: "12,000", marks: { 2: "focus", 3: "exclude" } }],
          [{ text: "東區", marks: { 1: "exclude" } }, { text: "9,800", marks: { 2: "focus", 3: "exclude" } }],
          [{ text: "南區", marks: { 1: "exclude" } }, { text: "31,000", marks: { 2: "focus", 3: "exclude" } }],
          [{ text: "北區", marks: { 1: "candidate" } }, { text: "18,500", marks: { 2: "focus", 3: "result" } }],
          [{ text: "東區", marks: { 1: "exclude" } }, { text: "7,200", marks: { 2: "focus", 3: "exclude" } }]
        ]
      },
      steps: [
        { label: "1. 條件一：北區", tone: "c1", text: "第一組條件先留下北區列。這些只是候選，還要通過下一個條件。", caption: "COUNTIFS 每一組都是條件範圍、條件值。" },
        { label: "2. 條件二：看金額欄", tone: "c2", text: "第二組條件檢查金額欄，準備判斷是否達到 15,000。", caption: "比較條件通常寫成引號裡的字串，例如 \">=15000\"。" },
        { label: "3. 同時成立才計數", tone: "c3", text: "只有北區且金額 >= 15,000 的列被計入，所以結果是 3。", caption: "如果你想做 OR，通常要把多個 COUNTIFS 相加。" }
      ]
    },
    {
      id: "averageifs-region-category",
      kind: "formula",
      role: "supporting",
      badge: "多條件平均",
      title: "AVERAGEIFS：只平均命中的數字",
      subtitle: "AVERAGEIFS 的語法和 SUMIFS 很像，只是第一個參數換成要取平均的範圍。",
      outcome: "結果：(15,000 + 22,000 + 18,500) / 3 = 18,500",
      formulaParts: [
        { text: "=AVERAGEIFS(" },
        { text: "E2:E9", step: 1, tone: "c1" },
        { text: ", " },
        { text: "C2:C9", step: 2, tone: "c2" },
        { text: ", \"北區\", " },
        { text: "D2:D9", step: 3, tone: "c3" },
        { text: ", \"電子\")", step: 4, tone: "c4" }
      ],
      board: {
        title: "北區電子類平均金額",
        columns: ["區域", "類別", "金額"],
        rows: [
          [{ text: "北區", marks: { 2: "candidate" } }, { text: "電子", marks: { 3: "candidate" } }, { text: "15,000", marks: { 1: "focus", 4: "result" } }],
          [{ text: "南區", marks: { 2: "exclude" } }, { text: "服飾", marks: { 3: "exclude" } }, { text: "8,500", marks: { 1: "focus" } }],
          [{ text: "北區", marks: { 2: "candidate" } }, { text: "電子", marks: { 3: "candidate" } }, { text: "22,000", marks: { 1: "focus", 4: "result" } }],
          [{ text: "北區", marks: { 2: "candidate" } }, { text: "服飾", marks: { 3: "exclude" } }, { text: "12,000", marks: { 1: "focus", 4: "exclude" } }],
          [{ text: "東區", marks: { 2: "exclude" } }, { text: "電子", marks: { 3: "candidate" } }, { text: "9,800", marks: { 1: "focus" } }],
          [{ text: "南區", marks: { 2: "exclude" } }, { text: "電子", marks: { 3: "candidate" } }, { text: "31,000", marks: { 1: "focus" } }],
          [{ text: "北區", marks: { 2: "candidate" } }, { text: "電子", marks: { 3: "candidate" } }, { text: "18,500", marks: { 1: "focus", 4: "result" } }],
          [{ text: "東區", marks: { 2: "exclude" } }, { text: "服飾", marks: { 3: "exclude" } }, { text: "7,200", marks: { 1: "focus" } }]
        ]
      },
      steps: [
        { label: "1. 平均範圍", tone: "c1", text: "AVERAGEIFS 第一個參數是要取平均的數字欄。這和 SUMIFS 的位置邏輯一致。", caption: "不是先寫條件欄，而是先寫最後要計算的欄。" },
        { label: "2. 條件一：北區", tone: "c2", text: "第一組條件留下北區資料。", caption: "非北區不會進入最後平均。" },
        { label: "3. 條件二：電子", tone: "c3", text: "第二組條件再留下電子類別。北區但不是電子，也會被排除。", caption: "多條件平均也是 AND 關係。" },
        { label: "4. 結果：只平均命中的金額", tone: "c4", text: "最後只拿北區電子類的三筆金額平均，得到 18,500。", caption: "問符合條件的平均時，用 AVERAGEIF / AVERAGEIFS。" }
      ]
    },
    {
      id: "wildcards-comparisons",
      kind: "formula",
      role: "supporting",
      badge: "條件寫法",
      title: "萬用字元與比較條件：條件不只等於",
      subtitle: "條件可以是文字片段、不等於，也可以是大於等於。",
      outcome: "結果：姓王、非已付、金額達標都能用條件統計快速回答",
      formulaParts: [
        { text: "=COUNTIF(B2:B9," },
        { text: "\"王*\"", step: 1, tone: "c1" },
        { text: ")  " },
        { text: "=COUNTIF(F2:F9," },
        { text: "\"<>已付\"", step: 2, tone: "c2" },
        { text: ")  " },
        { text: "=SUMIFS(E2:E9,C2:C9,\"北區\",E2:E9," },
        { text: "\">=15000\"", step: 3, tone: "c3" },
        { text: ")" }
      ],
      board: {
        title: "常用條件寫法",
        columns: ["業務員", "區域", "金額", "付款"],
        rows: [
          [{ text: "王小明", marks: { 1: "match" } }, { text: "北區", marks: { 3: "candidate" } }, { text: "15,000", marks: { 3: "result" } }, { text: "已付", marks: { 2: "exclude" } }],
          [{ text: "林美華", marks: { 1: "exclude" } }, { text: "南區" }, { text: "8,500" }, { text: "未付", marks: { 2: "match" } }],
          [{ text: "張志偉", marks: { 1: "exclude" } }, { text: "北區", marks: { 3: "candidate" } }, { text: "22,000", marks: { 3: "result" } }, { text: "已付", marks: { 2: "exclude" } }],
          [{ text: "王小明", marks: { 1: "match" } }, { text: "北區", marks: { 3: "candidate" } }, { text: "12,000", marks: { 3: "exclude" } }, { text: "已付", marks: { 2: "exclude" } }],
          [{ text: "陳雅婷", marks: { 1: "exclude" } }, { text: "東區" }, { text: "9,800" }, { text: "未付", marks: { 2: "match" } }],
          [{ text: "林美華", marks: { 1: "exclude" } }, { text: "南區" }, { text: "31,000" }, { text: "已付", marks: { 2: "exclude" } }],
          [{ text: "張志偉", marks: { 1: "exclude" } }, { text: "北區", marks: { 3: "candidate" } }, { text: "18,500", marks: { 3: "result" } }, { text: "已付", marks: { 2: "exclude" } }],
          [{ text: "陳雅婷", marks: { 1: "exclude" } }, { text: "東區" }, { text: "7,200" }, { text: "已付", marks: { 2: "exclude" } }]
        ]
      },
      steps: [
        { label: "1. 王*：開頭是王", tone: "c1", text: "星號代表任意長度文字。王* 會命中王小明、王大華這類以王開頭的文字。", caption: "用在姓名、產品名、公司名都很常見。" },
        { label: "2. <>：不等於", tone: "c2", text: "\"<>已付\" 代表付款狀態不是已付。它可以拿來找未付、取消、待確認等例外資料。", caption: "不等於很適合做追蹤清單。" },
        { label: "3. >=：達到門檻", tone: "c3", text: "比較條件通常放在引號裡，例如 \">=15000\"。如果門檻在 H2，寫成 \">=\"&H2。", caption: "條件統計常搭配 KPI 門檻使用。" }
      ]
    }
  ],
  "P2-02": [
    {
      id: "xlookup-direct-return",
      kind: "formula",
      badge: "公式拆解",
      title: "XLOOKUP：查找欄和回傳欄分開寫",
      subtitle: "新版 Excel 的主路線。你不用手算第幾欄，只要明確說出去哪裡找、回傳哪裡。",
      outcome: "結果：P003 在產品代碼欄命中，所以回傳同列的 27吋螢幕",
      formulaParts: [
        { text: "=XLOOKUP(" },
        { text: "\"P003\"", step: 1, tone: "c1" },
        { text: ", " },
        { text: "A2:A6", step: 2, tone: "c2" },
        { text: ", " },
        { text: "B2:B6", step: 3, tone: "c3" },
        { text: ", " },
        { text: "\"查無此商品\"", step: 4, tone: "c4" },
        { text: ")" }
      ],
      board: {
        title: "商品表",
        columns: ["產品代碼", "產品名稱", "單價"],
        rows: [
          [
            { text: "P001", marks: { 2: "focus" } },
            { text: "無線滑鼠", marks: { 3: "focus" } },
            { text: "590" }
          ],
          [
            { text: "P002", marks: { 2: "focus" } },
            { text: "機械鍵盤", marks: { 3: "focus" } },
            { text: "2,200" }
          ],
          [
            { text: "P003", marks: { 1: "focus", 2: "match" } },
            { text: "27吋螢幕", marks: { 3: "result" } },
            { text: "8,900" }
          ],
          [
            { text: "P004", marks: { 2: "focus" } },
            { text: "USB集線器", marks: { 3: "focus" } },
            { text: "380" }
          ],
          [
            { text: "P005", marks: { 2: "focus" } },
            { text: "網路攝影機", marks: { 3: "focus" } },
            { text: "1,500" }
          ]
        ]
      },
      steps: [
        {
          label: "1. 查找值：你手上拿什麼去找",
          tone: "c1",
          text: "這裡拿 P003 去查商品表。查找值可以直接寫在公式裡，也可以引用訂單表上的產品代碼儲存格。",
          caption: "先確認查找值乾淨一致，像是 P003 不要混到多餘空白。"
        },
        {
          label: "2. 查找範圍：去哪一欄找",
          tone: "c2",
          text: "A2:A6 是產品代碼欄。XLOOKUP 會在這一欄尋找 P003。",
          caption: "查找範圍只需要放查找欄，不必把整張表都圈進去。"
        },
        {
          label: "3. 回傳範圍：命中後拿哪一欄回來",
          tone: "c3",
          text: "B2:B6 是產品名稱欄。P003 命中第 3 列，所以回傳同一列的 27吋螢幕。",
          caption: "這就是 XLOOKUP 比 VLOOKUP 清楚的地方：回傳欄直接指定，不用數第幾欄。"
        },
        {
          label: "4. 找不到時的結果",
          tone: "c4",
          text: "如果找不到 P003，公式會顯示「查無此商品」，而不是讓 #N/A 直接出現在報表上。",
          caption: "正式交付報表時，找不到資料也要有可讀訊息。"
        }
      ]
    },
    {
      id: "index-match-reverse",
      kind: "formula",
      badge: "反向查找",
      title: "INDEX + MATCH：舊版也能往左查",
      subtitle: "MATCH 先找位置，INDEX 再依位置取值。這個拆法讓你不受 VLOOKUP 只能往右查的限制。",
      outcome: "結果：機械鍵盤在名稱欄第 2 筆，所以 INDEX 從代碼欄取回 P002",
      formulaParts: [
        { text: "=INDEX(" },
        { text: "A2:A6", step: 1, tone: "c1" },
        { text: ", MATCH(" },
        { text: "\"機械鍵盤\"", step: 2, tone: "c2" },
        { text: ", " },
        { text: "B2:B6", step: 3, tone: "c3" },
        { text: ", " },
        { text: "0", step: 4, tone: "c4" },
        { text: "))" }
      ],
      board: {
        title: "已知產品名稱，反查產品代碼",
        columns: ["產品代碼", "產品名稱"],
        rows: [
          [
            { text: "P001", marks: { 1: "focus" } },
            { text: "無線滑鼠", marks: { 3: "focus" } }
          ],
          [
            { text: "P002", marks: { 1: "focus", 4: "result" } },
            { text: "機械鍵盤", marks: { 2: "focus", 3: "match" } }
          ],
          [
            { text: "P003", marks: { 1: "focus" } },
            { text: "27吋螢幕", marks: { 3: "focus" } }
          ],
          [
            { text: "P004", marks: { 1: "focus" } },
            { text: "USB集線器", marks: { 3: "focus" } }
          ],
          [
            { text: "P005", marks: { 1: "focus" } },
            { text: "網路攝影機", marks: { 3: "focus" } }
          ]
        ]
      },
      steps: [
        {
          label: "1. INDEX 先指定要拿結果的欄",
          tone: "c1",
          text: "A2:A6 是產品代碼欄，也就是最後要回傳結果的範圍。",
          caption: "和 VLOOKUP 不同，回傳欄可以在查找欄左邊。"
        },
        {
          label: "2. MATCH 指定你要找的值",
          tone: "c2",
          text: "這裡要在產品名稱欄尋找「機械鍵盤」。",
          caption: "MATCH 的工作不是回傳名稱，而是找出它在範圍中的第幾個。"
        },
        {
          label: "3. MATCH 指定去哪一欄找",
          tone: "c3",
          text: "B2:B6 是產品名稱欄。機械鍵盤在這個範圍裡的第 2 個位置。",
          caption: "MATCH 回傳的是相對位置，不是 Excel 的實際列號。"
        },
        {
          label: "4. 0 代表精確比對，再交給 INDEX 取值",
          tone: "c4",
          text: "MATCH 回傳 2，INDEX 就從 A2:A6 的第 2 個位置取回 P002。",
          caption: "這就是 INDEX + MATCH 的核心：先定位，再取值。"
        }
      ]
    }
  ],
  "P2-03": [
    {
      id: "pivot-from-raw",
      kind: "media",
      badge: "操作演示",
      title: "樞紐分析：從流水帳變成可讀報表",
      subtitle: "不是重做一張新表，而是把原始資料的欄位拖進四大區域。",
      media: {
        kicker: "慢速循環短片",
        title: "先看一次完整操作，再直接去練習區跟做",
        src: "/media/p2-03-pivot-demo.webp",
        video: "/media/p2-03-pivot-demo.mp4",
        poster: "/media/p2-03-pivot-demo-poster.jpg",
        alt: "樞紐分析操作短片：從原始資料、插入樞紐分析表、把部門拖到列、月份拖到欄、業績拖到值，最後形成彙總報表。",
        note: "這支短片刻意放慢節奏，讓你看清楚每一次拖曳、停頓和報表長出的順序。這一課不再重複放第二套 Pivot 拆步，避免同頁內容重複。"
      },
      outcome: "結果：原始 5 筆交易，逐步變成部門 × 月份彙總表",
      observations: [
        { title: "先看欄位拖曳", text: "部門進列、月份進欄、業績進值，報表角度就是這樣被組出來。" },
        { title: "再看彙總形成", text: "同部門同月份的多筆交易會被自動加總，不是手動重打一張表。" },
        { title: "最後看可重整性", text: "來源資料更新後，樞紐只要重新整理就能吃到新資料。" }
      ],
      followup: "看完一次後，直接切到練習區照做就好；需要回想時，再回來重播這支短片。"
    }
  ],
  "P2-04": [
    {
      id: "conditional-formatting-story",
      kind: "media",
      badge: "操作演示",
      title: "條件式格式化：讓主管一眼看出誰該追",
      subtitle: "先比較，再提醒，最後才考慮美化。這樣的報表才真的有判讀效率。",
      media: {
        kicker: "慢速循環短片",
        title: "從黑白表格，變成會自己說重點的月報",
        src: "/media/p2-04-conditional-demo.webp",
        video: "/media/p2-04-conditional-demo.mp4",
        poster: "/media/p2-04-conditional-demo-poster.jpg",
        alt: "條件式格式化操作短片：選取業績欄位、建立規則、套用資料橫條與醒目提醒，最後形成可快速判讀的月報。",
        note: "這支短片聚焦在三個實務動作：業績橫條、負成長標紅、客訴偏高提醒。看懂之後再去練習區自己重做，會比先讀一堆規則更容易進狀況。"
      },
      outcome: "結果：同一張月報裡，高績效、負成長和高客訴會自己浮出來",
      observations: [
        { title: "先選正確範圍", text: "規則是套在選取範圍上，選錯範圍就會只亮局部或亮錯列。" },
        { title: "每層規則只回答一件事", text: "資料橫條看大小，紅色提醒負成長，高客訴再用另一層規則。" },
        { title: "最後檢查規則順序", text: "多個規則同時命中時，管理規則裡的順序會影響最後畫面。" }
      ],
      followup: "練習時先只做一層規則，再慢慢疊第二層和第三層，才不會一開始就把表格弄亂。"
    }
  ],
  "P3-01": [
    {
      id: "data-validation-gate",
      kind: "media",
      badge: "操作演示",
      title: "資料驗證：把錯誤擋在入口，而不是事後補救",
      subtitle: "對使用者來說只是多一個下拉選單，對你來說卻是整份資料品質的起點。",
      media: {
        kicker: "慢速循環短片",
        title: "從下拉選單到錯誤提醒，示範一次完整表單防呆",
        src: "/media/p3-01-validation-demo.webp",
        video: "/media/p3-01-validation-demo.mp4",
        poster: "/media/p3-01-validation-demo-poster.jpg",
        alt: "資料驗證操作短片：設定部門清單、開啟下拉選單、輸入錯誤手機時跳出提醒，最後形成乾淨可交付的表單。",
        note: "這支短片故意把節奏放慢，讓你看清楚規則設定、使用者實際填表，以及錯誤被阻止時的差別。資料驗證真正有價值的地方，就是把清理成本往前移。"
      },
      outcome: "結果：部門、日期、手機這類常見錯誤，在輸入當下就被攔下來",
      observations: [
        { title: "先看允許類型", text: "清單、日期、數字、自訂公式分別擋不同錯誤，不要一律只做下拉。" },
        { title: "再看使用者輸入", text: "合法值會通過，錯誤格式會立刻被攔住，清資料成本被往前移。" },
        { title: "最後看錯誤提醒", text: "錯誤訊息要講人話，讓填表者知道該怎麼修正。" }
      ],
      followup: "練習時先從最簡單的『清單』和『文字長度』開始，等手感穩了再做連動下拉與自訂公式。"
    }
  ],
  "P3-02": [
    {
      id: "text-cleaning-stack",
      kind: "formula",
      badge: "清理順序",
      title: "TRIM + CLEAN + SUBSTITUTE：先處理看不見的髒資料",
      subtitle: "外部系統匯出的空白不一定是一般空白。先換掉 CHAR(160)，再清不可見字元，最後 TRIM。",
      outcome: "結果：看起來正常但混亂的姓名，變成可比對、可查找的乾淨文字",
      formulaParts: [
        { text: "=TRIM(" },
        { text: "CLEAN(", step: 2, tone: "c2" },
        { text: "SUBSTITUTE(" },
        { text: "A2", step: 1, tone: "c1" },
        { text: ", CHAR(160), \"\")", step: 1, tone: "c1" },
        { text: ")", step: 2, tone: "c2" },
        { text: ")", step: 3, tone: "c3" },
        { text: ")" }
      ],
      board: {
        title: "貼上來的姓名欄",
        columns: ["原始值", "問題", "清理後"],
        rows: [
          [
            { text: "王小明", marks: { 1: "focus" } },
            { text: "前後空白", marks: { 3: "focus" } },
            { text: "王小明", marks: { 3: "result" } }
          ],
          [
            { text: "林 美華", marks: { 1: "focus" } },
            { text: "多個空格", marks: { 3: "focus" } },
            { text: "林 美華", marks: { 3: "result" } }
          ],
          [
            { text: "張志偉", marks: { 1: "focus" } },
            { text: "CHAR(160)", marks: { 1: "match" } },
            { text: "張志偉", marks: { 3: "result" } }
          ],
          [
            { text: "陳雅琪", marks: { 2: "focus" } },
            { text: "不可見字元", marks: { 2: "match" } },
            { text: "陳雅琪", marks: { 3: "result" } }
          ]
        ]
      },
      steps: [
        {
          label: "1. 先處理非斷行空白 CHAR(160)",
          tone: "c1",
          text: "網頁或系統貼上來的空白常常不是一般空白，而是 CHAR(160)。TRIM 不一定能處理它，所以先用 SUBSTITUTE 換掉。",
          caption: "如果 VLOOKUP 明明看起來一樣卻查不到，常常就是這種隱藏字元。"
        },
        {
          label: "2. CLEAN 移除不可見控制字元",
          tone: "c2",
          text: "CLEAN 會處理換行、控制字元等肉眼看不出的髒資料，讓文字更穩定。",
          caption: "它不會修正所有問題，但常是清資料流程裡很值得加的一層。"
        },
        {
          label: "3. 最後才用 TRIM 整理空白",
          tone: "c3",
          text: "TRIM 會移除前後空白，並把連續空白壓成單一空白。放在最後，能把前面清完後剩下的空白整理乾淨。",
          caption: "清資料要有順序：先換掉怪空白，再清不可見字元，最後整理空白。"
        }
      ]
    },
    {
      id: "order-id-to-date",
      kind: "formula",
      badge: "日期轉換",
      title: "DATE + MID：從訂單號拆出真正日期",
      subtitle: "不要只用 TEXT 做顯示格式。若後面還要排序或計算天數，要把文字轉成真正日期。",
      outcome: "結果：A2026031101 會被轉成 2026/03/11，可排序、分月與計算天數",
      formulaParts: [
        { text: "=DATE(" },
        { text: "MID(A2,2,4)", step: 1, tone: "c1" },
        { text: ", " },
        { text: "MID(A2,6,2)", step: 2, tone: "c2" },
        { text: ", " },
        { text: "MID(A2,8,2)", step: 3, tone: "c3" },
        { text: ")" }
      ],
      board: {
        title: "訂單號碼 A2026031101",
        columns: ["片段", "公式", "結果"],
        rows: [
          [
            { text: "年份", marks: { 1: "focus" } },
            { text: "MID(A2,2,4)", marks: { 1: "match" } },
            { text: "2026", marks: { 1: "result" } }
          ],
          [
            { text: "月份", marks: { 2: "focus" } },
            { text: "MID(A2,6,2)", marks: { 2: "match" } },
            { text: "03", marks: { 2: "result" } }
          ],
          [
            { text: "日期", marks: { 3: "focus" } },
            { text: "MID(A2,8,2)", marks: { 3: "match" } },
            { text: "11", marks: { 3: "result" } }
          ],
          [
            { text: "真日期", marks: { 4: "focus" } },
            { text: "DATE(2026,03,11)", marks: { 4: "match" } },
            { text: "2026/03/11", marks: { 4: "result" } }
          ]
        ]
      },
      steps: [
        {
          label: "1. 從第 2 碼開始取 4 碼作為年份",
          tone: "c1",
          text: "訂單號第一碼是字母 A，所以年份從第 2 碼開始。MID(A2,2,4) 會取出 2026。",
          caption: "MID 的位置從 1 開始算，不是從 0。"
        },
        {
          label: "2. 再取第 6 到第 7 碼作為月份",
          tone: "c2",
          text: "MID(A2,6,2) 會取出 03。DATE 可以接受這種文字型數字並轉成日期的一部分。",
          caption: "月份如果是 03，不要先用 TEXT 格式化，先轉真日期。"
        },
        {
          label: "3. 最後取日期並交給 DATE",
          tone: "c3",
          text: "MID(A2,8,2) 取出 11。DATE(年,月,日) 會把三段合成真正日期。",
          caption: "真正日期才能排序、計算天數、接樞紐日期群組。"
        },
        {
          label: "4. TEXT 是顯示格式，不是日期本身",
          tone: "c4",
          text: "如果只是要顯示成 2026-03-11，可以再用 TEXT；但請保留原本的真日期欄位供後續計算。",
          caption: "這是很多報表日期軸怪掉的原因。"
        }
      ]
    }
  ],
  "P3-03": [
    {
      id: "chart-design-story",
      kind: "media",
      badge: "操作演示",
      title: "圖表設計：從原始圖到有觀點的專業圖",
      subtitle: "真正的差別不是插入圖表，而是知道哪些元素該留下、哪些該刪掉。",
      media: {
        kicker: "慢速循環短片",
        title: "示範一次：選資料、插組合圖、去雜訊、補觀點",
        src: "/media/p3-03-chart-demo.webp",
        video: "/media/p3-03-chart-demo.mp4",
        poster: "/media/p3-03-chart-demo-poster.jpg",
        alt: "圖表設計操作短片：選取月份、營收和成長率資料，插入組合圖，刪除雜訊並加上有觀點的標題，最後形成更專業的圖表。",
        note: "這裡刻意不是教你所有圖表類型，而是示範一條最常見的職場流程：把原始自動圖修成主管願意直接貼進簡報的版本。"
      },
      outcome: "結果：同一份資料，從『只是有圖』變成『真的有在講結論』",
      observations: [
        { title: "先看圖表選型", text: "金額和成長率單位不同，所以用組合圖比硬塞單一圖表清楚。" },
        { title: "再看去雜訊", text: "拿掉粗網格線、外框和多餘標籤，視線才會回到資料本身。" },
        { title: "最後看觀點標題", text: "標題不是寫圖表名稱，而是直接講這張圖想傳達的結論。" }
      ],
      followup: "自己重做時先練三件事：圖型選對、格線刪掉、標題改成有觀點。這三件就已經能拉開很大差距。"
    }
  ],
  "P3-04": [
    {
      id: "table-structured-reference",
      kind: "formula",
      badge: "公式拆解",
      title: "普通範圍 vs 表格：結構化參照讓公式更穩",
      subtitle: "同樣是 SUMIFS，用欄位名稱寫出來，維護時比較不會漏列或選錯欄。",
      outcome: "結果：新增資料列後，業績表[金額] 會自動延伸，不必手動把 C2:C500 改成 C2:C501",
      formulaParts: [
        { text: "=SUMIFS(" },
        { text: "業績表[金額]", step: 1, tone: "c1" },
        { text: ", " },
        { text: "業績表[部門]", step: 2, tone: "c2" },
        { text: ", " },
        { text: "\"電子\"", step: 3, tone: "c3" },
        { text: ")" }
      ],
      board: {
        title: "業績表",
        columns: ["日期", "部門", "業務", "金額", "公式狀態"],
        rows: [
          [
            { text: "2026/05/01", marks: { 1: "focus" } },
            { text: "電子", marks: { 2: "match", 3: "match", 4: "result" } },
            { text: "王小明", marks: { 4: "result" } },
            { text: "120,000", marks: { 1: "match", 4: "result" } },
            { text: "已納入", marks: { 4: "result" } }
          ],
          [
            { text: "2026/05/02", marks: { 1: "focus" } },
            { text: "業務", marks: { 2: "focus", 3: "exclude" } },
            { text: "林美華" },
            { text: "98,000", marks: { 1: "match" } },
            { text: "部門不符", marks: { 4: "exclude" } }
          ],
          [
            { text: "2026/05/03", marks: { 1: "focus" } },
            { text: "電子", marks: { 2: "match", 3: "match", 4: "result" } },
            { text: "張志偉", marks: { 4: "result" } },
            { text: "146,000", marks: { 1: "match", 4: "result" } },
            { text: "已納入", marks: { 4: "result" } }
          ],
          [
            { text: "新增列", marks: { 1: "candidate", 4: "candidate" } },
            { text: "電子", marks: { 2: "candidate", 3: "match", 4: "result" } },
            { text: "新資料", marks: { 4: "result" } },
            { text: "88,000", marks: { 1: "candidate", 4: "result" } },
            { text: "自動延伸", marks: { 4: "result" } }
          ]
        ]
      },
      steps: [
        {
          label: "1. 先指定要加總的欄位",
          tone: "c1",
          text: "業績表[金額] 直接指向表格裡的金額欄，比 C2:C500 更容易讀，也會隨著表格新增列自動延伸。",
          result: "目前先鎖定要加總的欄位：金額。",
          caption: "表格名稱 + 欄位名稱，就是結構化參照。"
        },
        {
          label: "2. 條件範圍也用欄位名稱",
          tone: "c2",
          text: "業績表[部門] 告訴公式要在部門欄判斷條件。公式不再依賴你記住 A 欄或 B 欄是哪個欄位。",
          result: "條件欄位已清楚指向：部門。",
          caption: "欄位調整位置時，結構化參照比裸範圍更耐改。"
        },
        {
          label: "3. 條件值仍照一般 SUMIFS 寫法",
          tone: "c3",
          text: "條件值是 \"電子\"，所以只有部門等於電子的列會被加總。",
          result: "命中部門 = 電子的資料列，排除其他部門。",
          caption: "語法邏輯沒變，只是把範圍改成更可讀的表格欄位。"
        },
        {
          label: "4. 新增列會自動納入",
          tone: "c4",
          text: "當你在表格底部新增一筆電子部資料，業績表[金額] 和業績表[部門] 會自動包含它。",
          result: "新資料列自動進入公式範圍，結果會跟著更新。",
          caption: "這就是表格比普通範圍適合長期報表的原因。"
        }
      ]
    }
  ],
  "P3-05": [
    {
      id: "protect-input-flow",
      kind: "workflow",
      badge: "流程演示",
      title: "保護工作表：先解鎖輸入格，再鎖住公式區",
      subtitle: "保護不是最後按密碼而已，真正關鍵是先定義使用者可以改哪些地方。",
      stages: [
        { text: "標出輸入區", tone: "c1" },
        { text: "取消鎖定", tone: "c2" },
        { text: "啟用保護", tone: "c3" },
        { text: "使用者測試", tone: "c4" }
      ],
      outcome: "結果：使用者能填黃色輸入格，公式、標題和彙總區都被保護住。",
      steps: [
        {
          stage: 1,
          label: "1. 先分出輸入格與公式格",
          tone: "c1",
          text: "把使用者需要填的欄位標成輸入區，例如日期、部門、金額；公式欄、彙總欄和標題列不要開放編輯。",
          result: "先定義可編輯區，不急著按保護。",
          caption: "保護工作表前，先從使用者流程反推哪些格子真的需要改。",
          panels: [
            { title: "表單規劃", columns: ["欄位", "用途", "狀態"], rows: [["日期", "使用者填寫", "可編輯"], ["部門", "下拉選單", "可編輯"], ["金額", "使用者填寫", "可編輯"], ["稅額公式", "=金額*5%", "不可編輯"]] }
          ]
        },
        {
          stage: 2,
          label: "2. 對輸入區取消鎖定",
          tone: "c2",
          text: "選取可填欄位，開啟儲存格格式，把「鎖定」取消。其他儲存格維持預設鎖定。",
          result: "只有輸入欄位被解鎖，公式區維持鎖定。",
          caption: "Excel 預設所有格子都鎖定，但要等工作表保護啟用後才會生效。",
          panels: [
            { title: "儲存格保護設定", lines: ["可填欄位：取消鎖定", "公式欄位：維持鎖定", "標題與說明：維持鎖定"] }
          ]
        },
        {
          stage: 3,
          label: "3. 啟用保護工作表",
          tone: "c3",
          text: "到校閱 → 保護工作表，保留選取未鎖定儲存格，視需求允許排序、篩選或插入列。",
          result: "保護啟用後，鎖定格才真正不能被修改。",
          caption: "密碼是防誤改，不是資安控管；機密檔案要另外做檔案權限。",
          panels: [
            { title: "建議勾選", lines: ["允許選取未鎖定的儲存格", "需要報表操作時才允許排序/篩選", "不要開放格式化儲存格，避免版面被改亂"] }
          ]
        },
        {
          stage: 4,
          label: "4. 用使用者視角測試",
          tone: "c4",
          text: "試著填輸入格、改公式格、貼上資料、操作篩選。真正的驗收是流程能不能順利完成。",
          result: "使用者能填該填的，不能碰不該碰的。",
          caption: "交付前請自己扮演使用者走一次，很多鎖太死的問題會立刻浮出來。",
          panels: [
            { title: "測試結果", columns: ["動作", "期待結果", "狀態"], rows: [["填金額", "可輸入", "通過"], ["改稅額公式", "被阻擋", "通過"], ["篩選部門", "可操作", "通過"]] }
          ]
        }
      ]
    }
  ],
  "P4-01": [
    {
      id: "filter-multi-condition",
      kind: "formula",
      badge: "公式拆解",
      title: "FILTER：用 TRUE/FALSE 陣列挑出整列資料",
      subtitle: "FILTER 的重點不是只會篩選，而是理解每一列都會先被條件判斷成 TRUE 或 FALSE。",
      outcome: "結果：只有業務部且績效 >= 90 的列會溢出成新清單",
      formulaParts: [
        { text: "=FILTER(" },
        { text: "A2:E7", step: 1, tone: "c1" },
        { text: ", " },
        { text: "(B2:B7=\"業務\")", step: 2, tone: "c2" },
        { text: "*" },
        { text: "(E2:E7>=90)", step: 3, tone: "c3" },
        { text: ", " },
        { text: "\"查無資料\"", step: 4, tone: "c4" },
        { text: ")" }
      ],
      board: {
        title: "員工績效表",
        columns: ["員工", "部門", "月薪", "績效"],
        rows: [
          [
            { text: "王小明", marks: { 1: "focus" } },
            { text: "業務", marks: { 2: "match" } },
            { text: "55,000", marks: { 1: "focus" } },
            { text: "88", marks: { 3: "focus" } }
          ],
          [
            { text: "張志偉", marks: { 1: "focus", 4: "result" } },
            { text: "業務", marks: { 2: "match", 4: "result" } },
            { text: "68,000", marks: { 1: "focus", 4: "result" } },
            { text: "95", marks: { 3: "match", 4: "result" } }
          ],
          [
            { text: "李建宏", marks: { 1: "focus" } },
            { text: "資訊", marks: { 2: "focus" } },
            { text: "75,000", marks: { 1: "focus" } },
            { text: "91", marks: { 3: "match" } }
          ],
          [
            { text: "黃志強", marks: { 1: "focus", 4: "result" } },
            { text: "業務", marks: { 2: "match", 4: "result" } },
            { text: "82,000", marks: { 1: "focus", 4: "result" } },
            { text: "98", marks: { 3: "match", 4: "result" } }
          ]
        ]
      },
      steps: [
        {
          label: "1. 先圈出要回傳的整張資料",
          tone: "c1",
          text: "A2:E7 是 FILTER 最後要吐出的資料範圍。只要某一列通過條件，這一整列就會出現在結果裡。",
          caption: "FILTER 不是只回傳單一值，而是能回傳整片資料。"
        },
        {
          label: "2. 第一個條件：部門必須是業務",
          tone: "c2",
          text: "B2:B7=\"業務\" 會產生一串 TRUE/FALSE。業務部是 TRUE，其他部門是 FALSE。",
          caption: "條件範圍的列數必須和資料範圍一致。"
        },
        {
          label: "3. 第二個條件：績效必須 >= 90",
          tone: "c3",
          text: "E2:E7>=90 也會產生一串 TRUE/FALSE。兩個條件中間用 *，代表兩邊都要成立。",
          caption: "多條件 AND 用 *；OR 情境通常用 +。"
        },
        {
          label: "4. 找不到資料時要有可讀訊息",
          tone: "c4",
          text: "第 3 參數填「查無資料」後，如果沒有任何列符合條件，就不會直接丟出 #CALC!。",
          caption: "交付報表時，錯誤訊息要變成使用者看得懂的文字。"
        }
      ]
    },
    {
      id: "unique-sort-spill",
      kind: "formula",
      badge: "動態清單",
      title: "SORT + UNIQUE：做一份會自動更新的下拉來源",
      subtitle: "先去重，再排序，最後用 D2# 引用整個溢出結果。",
      outcome: "結果：重複的部門只留下 1 次，並排序成可引用的動態清單",
      formulaParts: [
        { text: "=SORT(", step: 3, tone: "c3" },
        { text: "UNIQUE(", step: 2, tone: "c2" },
        { text: "B2:B9", step: 1, tone: "c1" },
        { text: "))" }
      ],
      board: {
        title: "部門欄位會重複出現",
        columns: ["原始部門", "UNIQUE 後", "SORT 後"],
        rows: [
          [
            { text: "業務", marks: { 1: "focus", 2: "match" } },
            { text: "業務", marks: { 2: "result" } },
            { text: "人事", marks: { 3: "result", 4: "result" } }
          ],
          [
            { text: "財務", marks: { 1: "focus", 2: "match" } },
            { text: "財務", marks: { 2: "result" } },
            { text: "業務", marks: { 3: "result", 4: "result" } }
          ],
          [
            { text: "業務", marks: { 1: "focus" } },
            { text: "人事", marks: { 2: "result" } },
            { text: "財務", marks: { 3: "result", 4: "result" } }
          ],
          [
            { text: "資訊", marks: { 1: "focus", 2: "match" } },
            { text: "資訊", marks: { 2: "result" } },
            { text: "資訊", marks: { 3: "result", 4: "result" } }
          ]
        ]
      },
      steps: [
        {
          label: "1. 來源範圍：先抓會重複的欄",
          tone: "c1",
          text: "B2:B9 是原始部門欄。因為資料會新增或重複，所以不適合手動維護清單。",
          caption: "實務上可改用 Table 欄位，例如 員工表[部門]。"
        },
        {
          label: "2. UNIQUE 先去掉重複值",
          tone: "c2",
          text: "UNIQUE 會保留每個部門第一次出現的結果，做出一份不重複清單。",
          caption: "這一步常拿來做資料驗證的來源。"
        },
        {
          label: "3. SORT 再把清單排整齊",
          tone: "c3",
          text: "SORT 包在外層，讓去重後的部門清單按照順序排列。",
          caption: "巢狀公式從內往外讀：先 UNIQUE，再 SORT。"
        },
        {
          label: "4. 用 D2# 引用整片溢出結果",
          tone: "c4",
          text: "公式只寫在 D2，但結果會向下溢出。其他公式或下拉選單可以用 D2# 引用整個動態清單。",
          caption: "# 是動態陣列最重要的引用符號。"
        }
      ]
    }
  ],
  "P4-02": [
    {
      id: "let-quarterly-bonus",
      kind: "formula",
      badge: "公式重構",
      title: "LET：把季度獎金公式拆成可讀段落",
      subtitle: "先命名全年業績 total，再命名獎金 bonus，最後回傳 total + bonus。",
      outcome: "結果：810,000 達標，獎金 40,500，最後回傳 850,500",
      formulaParts: [
        { text: "=LET(" },
        { text: "total", step: 1, tone: "c1" },
        { text: ", " },
        { text: "SUM(B2:E2)", step: 2, tone: "c2" },
        { text: ", " },
        { text: "bonus", step: 3, tone: "c3" },
        { text: ", IF(total>=500000,total*5%,0), " },
        { text: "total+bonus", step: 4, tone: "c4" },
        { text: ")" }
      ],
      board: {
        title: "張志偉季度業績",
        columns: ["Q1", "Q2", "Q3", "Q4", "公式結果"],
        rows: [
          [
            { text: "210,000", marks: { 2: "focus" } },
            { text: "185,000", marks: { 2: "focus" } },
            { text: "220,000", marks: { 2: "focus" } },
            { text: "195,000", marks: { 2: "focus" } },
            { text: "850,500", marks: { 4: "result" } }
          ],
          [
            { text: "total", marks: { 1: "focus" } },
            { text: "810,000", marks: { 2: "result" } },
            { text: "bonus", marks: { 3: "focus" } },
            { text: "40,500", marks: { 3: "result" } },
            { text: "total + bonus", marks: { 4: "match" } }
          ]
        ]
      },
      steps: [
        {
          label: "1. 先替全年業績取名 total",
          tone: "c1",
          text: "LET 的變數名稱可以自己命名。total 讓後面的人一眼知道這段代表全年業績。",
          caption: "命名不是為了省字，而是為了讓公式自我解釋。"
        },
        {
          label: "2. total 的值只計算一次",
          tone: "c2",
          text: "SUM(B2:E2) 把 Q1 到 Q4 加總成 810,000。之後公式裡每次寫 total 都引用這個結果。",
          caption: "長公式中重複出現的計算，很適合先用 LET 存起來。"
        },
        {
          label: "3. 再命名 bonus，把判斷邏輯收起來",
          tone: "c3",
          text: "IF(total>=500000,total*5%,0) 代表達標才給 5% 獎金。把這段叫 bonus，比直接塞在最後更好讀。",
          caption: "LET 可以定義多組名稱和值。"
        },
        {
          label: "4. 最後一段就是回傳值",
          tone: "c4",
          text: "LET 的最後一個引數是最終答案。這裡回傳 total+bonus，也就是 850,500。",
          caption: "讀 LET 時可以先找最後一段，知道這條公式最後要吐什麼。"
        }
      ]
    },
    {
      id: "lambda-reusable-tax",
      kind: "formula",
      badge: "封裝邏輯",
      title: "LAMBDA：把公式變成可重用的自訂函式",
      subtitle: "先在儲存格立即執行測試，確認沒錯後再放進名稱管理員。",
      outcome: "結果：120,000 加上 5% 後回傳 126,000",
      formulaParts: [
        { text: "=LAMBDA(" },
        { text: "amount, rate", step: 1, tone: "c1" },
        { text: ", " },
        { text: "amount*(1+rate)", step: 2, tone: "c2" },
        { text: ")(" },
        { text: "120000, 5%", step: 3, tone: "c3" },
        { text: ")" }
      ],
      board: {
        title: "從測試公式到自訂函式",
        columns: ["階段", "內容", "結果"],
        rows: [
          [
            { text: "參數", marks: { 1: "focus" } },
            { text: "amount, rate", marks: { 1: "result" } },
            { text: "等待輸入" }
          ],
          [
            { text: "計算式", marks: { 2: "focus" } },
            { text: "amount*(1+rate)", marks: { 2: "result" } },
            { text: "加上比例" }
          ],
          [
            { text: "測試值", marks: { 3: "focus" } },
            { text: "120000, 5%", marks: { 3: "match" } },
            { text: "126,000", marks: { 3: "result", 4: "result" } }
          ],
          [
            { text: "命名後", marks: { 4: "focus" } },
            { text: "=含稅價(120000,5%)", marks: { 4: "match" } },
            { text: "126,000", marks: { 4: "result" } }
          ]
        ]
      },
      steps: [
        {
          label: "1. 先列出這個邏輯需要哪些參數",
          tone: "c1",
          text: "amount 和 rate 是這個自訂函式需要的兩個輸入。參數名稱要清楚，之後才看得懂。",
          caption: "LAMBDA 的第一段是參數清單。"
        },
        {
          label: "2. 寫出要重用的計算邏輯",
          tone: "c2",
          text: "amount*(1+rate) 是實際要重複使用的公式邏輯。這裡示範把金額加上一個比例。",
          caption: "這段可以很短，也可以包進 LET 讓它更清楚。"
        },
        {
          label: "3. 先立即執行測試",
          tone: "c3",
          text: "最後括號裡的 120000, 5% 是測試輸入。立即執行能確認公式真的回傳 126,000。",
          caption: "不要把沒測過的 LAMBDA 直接丟進名稱管理員。"
        },
        {
          label: "4. 測通後再放進名稱管理員",
          tone: "c4",
          text: "命名成含稅價後，就能像內建函數一樣使用：=含稅價(120000,5%)。",
          caption: "LAMBDA 的價值是讓常用邏輯可重用，而不是把公式藏起來。"
        }
      ]
    }
  ],
  "P4-03": [
    {
      id: "goal-seek-price",
      kind: "workflow",
      badge: "流程演示",
      title: "目標搜尋：用目標利潤反推售價",
      subtitle: "先把利潤公式建好，再讓 Excel 反覆試算單一變數。",
      stages: [
        { text: "建立模型", tone: "c1" },
        { text: "設定目標", tone: "c2" },
        { text: "變更售價", tone: "c3" },
        { text: "檢查結果", tone: "c4" }
      ],
      outcome: "結果：在成本 90、銷量 16,000 的假設下，目標利潤 960,000 需要售價 150。",
      steps: [
        {
          stage: 1,
          label: "1. 先有可反推的公式",
          tone: "c1",
          text: "目標搜尋不能憑空找答案，必須先有一格公式能根據售價算出利潤。",
          result: "起始狀態：售價 120 時，利潤只有 480,000。",
          caption: "這裡的利潤公式是 (售價 - 單位成本) * 銷量。",
          panels: [
            { title: "目前模型", columns: ["項目", "儲存格", "值/公式"], rows: [["售價", "B2", "120"], ["單位成本", "B3", "90"], ["銷量", "B4", "16,000"], ["利潤", "B5", "=(B2-B3)*B4 = 480,000"]] }
          ]
        },
        {
          stage: 2,
          label: "2. 指定結果格與目標值",
          tone: "c2",
          text: "資料 → 模擬分析 → 目標搜尋，把「設定儲存格」指定為利潤 B5，目標值輸入 960000。",
          result: "目標已設定：讓 B5 利潤達到 960,000。",
          caption: "設定儲存格一定要是公式格，不能是手打數字。",
          panels: [
            { title: "目標搜尋設定", lines: ["設定儲存格：B5 利潤", "目標值：960000", "變更儲存格：B2 售價"] }
          ]
        },
        {
          stage: 3,
          label: "3. 讓 Excel 改變售價",
          tone: "c3",
          text: "Excel 會調整 B2 售價，反覆重算 B5，直到利潤接近 960,000。",
          result: "Excel 找到售價 150，可讓利潤達標。",
          caption: "Goal Seek 只適合反推單一變數；多變數就要改用 Solver。",
          panels: [
            { title: "試算過程", columns: ["售價", "成本", "銷量", "利潤"], rows: [["120", "90", "16,000", "480,000"], ["140", "90", "16,000", "800,000"], ["150", "90", "16,000", "960,000"]] }
          ]
        },
        {
          stage: 4,
          label: "4. 檢查答案是否合理",
          tone: "c4",
          text: "得到售價 150 後，不要直接交付；回頭檢查成本、銷量、折扣和稅率假設是否完整。",
          result: "最終答案：售價 150，但必須附上假設限制。",
          caption: "假設分析算得再準，也只跟你的模型一樣可靠。",
          panels: [
            { title: "結果檢查", columns: ["項目", "結果", "檢查"], rows: [["售價", "150", "可達成目標"], ["利潤", "960,000", "符合目標"], ["模型假設", "成本/銷量固定", "需註明限制"]] }
          ]
        }
      ]
    }
  ],
  "P4-04": [
    {
      id: "power-query-refresh",
      kind: "media",
      badge: "流程演示",
      title: "Power Query：把每月手工整理變成可重跑流程",
      subtitle: "第一次是在設流程，第二次開始才是真的省時間。",
      media: {
        kicker: "慢速循環短片",
        title: "從髒 CSV 到月報自動更新，先看一次完整路徑",
        src: "/media/p4-04-power-query-demo.webp",
        video: "/media/p4-04-power-query-demo.mp4",
        poster: "/media/p4-04-power-query-demo-poster.jpg",
        alt: "Power Query 操作短片：連到原始 CSV、記住清理步驟、附加每月資料，最後在新增四月檔後按重新整理更新報表。",
        note: "這支短片保留了最有感的四個瞬間：接來源、記步驟、附加月份、重新整理。先建立整體節奏感，再去課文和練習裡補細節，會比較容易懂為什麼 Power Query 值得學。"
      },
      outcome: "結果：下個月資料一到，只要按重新整理，不用整套重做",
      observations: [
        { title: "先看資料來源", text: "Power Query 先連到檔案或資料夾，後續才有可重跑的基礎。" },
        { title: "再看轉換步驟", text: "拆欄、篩選、改型態都會被記成步驟，不是一次性的手動整理。" },
        { title: "最後看重新整理", text: "新增月份資料後按重新整理，舊流程會重新跑一次。" }
      ],
      followup: "自己實作時先讓第一條流程能穩定重跑，等這件事成功後，再去追求更多轉換步驟。"
    }
  ],
  "P4-05": [
    {
      id: "data-model-relationship",
      kind: "media",
      badge: "概念演示",
      title: "資料模型：把多張表串成同一套分析系統",
      subtitle: "Power Query 清資料，資料模型管關聯，樞紐只負責把答案展示出來。",
      media: {
        kicker: "慢速循環短片",
        title: "從多表、關聯、量值，到最後交給樞紐展示",
        src: "/media/p4-05-data-model-demo.webp",
        video: "/media/p4-05-data-model-demo.mp4",
        poster: "/media/p4-05-data-model-demo-poster.jpg",
        alt: "資料模型操作短片：先分開訂單表與產品表，建立產品ID關聯，寫量值公式，最後把結果交給樞紐分析表展示。",
        note: "這一課本來就偏觀念，所以短片重點不是花俏拖曳，而是把『多表 → 關聯 → 量值 → 樞紐』這條路徑建立起來。對 Mac 使用者尤其重要，因為你更需要判斷何時該交給模型，而不是硬把所有邏輯塞進樞紐。"
      },
      outcome: "結果：量值只定義一次，之後很多張報表都能共用同一套邏輯",
      observations: [
        { title: "先分清楚表角色", text: "訂單表是事實表，產品表、日期表是維度表，角色清楚關聯才穩。" },
        { title: "再看關聯線", text: "用產品 ID 或日期欄建立關聯，樞紐才能跨表分析。" },
        { title: "最後看量值", text: "總業績這類指標定義成量值後，多張報表可以共用同一套算法。" }
      ],
      followup: "學這課時先把表的角色和關聯想清楚，比急著背 DAX 函數更重要。"
    }
  ],
  "P5-01": [
    {
      id: "vba-safe-first-macro",
      kind: "workflow",
      badge: "流程演示",
      title: "第一個安全巨集：從手動動作到可重跑流程",
      subtitle: "VBA 入門先追求可理解、可復原，不急著把整份報表自動化。",
      stages: [
        { text: "手動跑順", tone: "c1" },
        { text: "寫入模組", tone: "c2" },
        { text: "測試輸出", tone: "c3" },
        { text: "安全保存", tone: "c4" }
      ],
      outcome: "結果：巨集只改指定區域，執行前知道會做什麼，執行後能檢查結果。",
      steps: [
        {
          stage: 1,
          label: "1. 先把手動流程說清楚",
          tone: "c1",
          text: "先不要寫程式，先列出這段工作每次都固定做什麼：清空輸出區、寫標題、填入公式、提示完成。",
          result: "巨集範圍先收斂成一段固定、可描述的流程。",
          caption: "說不清楚的手動流程，寫成巨集後只會更難查錯。",
          panels: [
            { title: "手動流程", lines: ["清空 A1:C10", "A1 寫入報表標題", "A2 寫入本月業績", "A3 寫入成長率公式"] }
          ]
        },
        {
          stage: 2,
          label: "2. 放進標準模組",
          tone: "c2",
          text: "在 VBE 插入模組，把每個動作用清楚的 Range 寫出來。初學先避免 Select/Activate，減少焦點錯誤。",
          result: "每個動作都指定明確範圍，降低誤改風險。",
          caption: "Range(\"A1\").Value 比先選 A1 再輸入更穩。",
          panels: [
            { title: "巨集片段", lines: ["Sub 建立月報()", "Range(\"A1:C10\").ClearContents", "Range(\"A1\").Value = \"本月業績摘要\"", "Range(\"A3\").Formula = \"=A2*1.05\"", "End Sub"] }
          ]
        },
        {
          stage: 3,
          label: "3. 在測試檔執行",
          tone: "c3",
          text: "第一次執行請用測試檔或備份檔，確認它只改預期範圍，沒有清掉原始資料。",
          result: "執行前後可對照，確認只有指定區域被改。",
          caption: "VBA 的 Undo 通常不可靠，測試檔是基本保護。",
          panels: [
            { title: "執行前/後", columns: ["區域", "執行前", "執行後"], rows: [["A1", "空白", "本月業績摘要"], ["A2", "100000", "100000"], ["A3", "空白", "=A2*1.05"]] }
          ]
        },
        {
          stage: 4,
          label: "4. 存成 .xlsm 並註明用途",
          tone: "c4",
          text: "巨集檔要存成 .xlsm，並在工作表或模組註解中寫清楚用途、會修改的範圍與執行前注意事項。",
          result: "可交付版本完成：副檔名、用途與安全提醒都清楚。",
          caption: "可交付的巨集不是只有能跑，還要讓下一個人敢按。",
          panels: [
            { title: "安全提醒", lines: ["保留原始資料備份", "巨集只處理指定工作表", "交付前關閉不必要的外部連結", "檔案副檔名使用 .xlsm"] }
          ]
        }
      ]
    }
  ],
  "P5-02": [
    {
      id: "vba-array-dictionary-safety",
      kind: "workflow",
      badge: "流程演示",
      title: "進階 VBA：一次讀入、記憶體處理、安全收尾",
      subtitle: "速度不是靠更快點儲存格，而是少跟工作表來回溝通。",
      stages: [
        { text: "一次讀入", tone: "c1" },
        { text: "記憶體處理", tone: "c2" },
        { text: "一次寫回", tone: "c3" },
        { text: "錯誤出口", tone: "c4" }
      ],
      outcome: "結果：大量資料用陣列和 Dictionary 處理，速度更穩，也能在錯誤時恢復狀態。",
      steps: [
        {
          stage: 1,
          label: "1. 把資料一次讀進陣列",
          tone: "c1",
          text: "把 A2:C10000 一次讀入 Variant 陣列，避免每一列都跨 Excel 物件模型讀取。",
          result: "讀取動作從 10,000 次降成 1 次。",
          caption: "慢的通常不是迴圈本身，而是一直 Cells(i,j) 讀寫工作表。",
          panels: [
            { title: "讀入策略", lines: ["Dim arr As Variant", "arr = Range(\"A2:C10000\").Value", "後續迴圈都處理 arr，不直接碰儲存格"] }
          ]
        },
        {
          stage: 2,
          label: "2. 用 Dictionary 做彙總或去重",
          tone: "c2",
          text: "在記憶體裡用 Dictionary 依部門加總、計數或去重，最後再整理成輸出陣列。",
          result: "電子部兩筆資料在記憶體中累計成 208000。",
          caption: "Dictionary 適合 key-value 類問題，例如每個部門累計多少金額。",
          panels: [
            { title: "處理邏輯", columns: ["資料列", "部門", "金額", "Dictionary"], rows: [["2", "電子", "120000", "電子=120000"], ["3", "業務", "98000", "業務=98000"], ["4", "電子", "88000", "電子=208000"]] }
          ]
        },
        {
          stage: 3,
          label: "3. 結果一次寫回",
          tone: "c3",
          text: "把輸出整理成二維陣列後，一次寫回結果區。這比在迴圈裡逐格寫入快很多。",
          result: "輸出區一次取得部門、總金額與筆數。",
          caption: "讀一次、算一次、寫一次，是大量資料巨集的基本節奏。",
          panels: [
            { title: "輸出結果", columns: ["部門", "總金額", "筆數"], rows: [["電子", "208000", "2"], ["業務", "98000", "1"], ["行政", "76000", "1"]] }
          ]
        },
        {
          stage: 4,
          label: "4. 加上錯誤處理與狀態恢復",
          tone: "c4",
          text: "關閉畫面更新或自動計算後，一定要在錯誤出口恢復設定，避免使用者的 Excel 狀態被巨集留壞。",
          result: "就算中途失敗，也會回到安全的 Excel 狀態。",
          caption: "進階巨集的安全感，常常來自完整的 CleanUp 區塊。",
          panels: [
            { title: "安全骨架", lines: ["On Error GoTo ErrHandler", "Application.ScreenUpdating = False", "Application.EnableEvents = False", "主要流程...", "CleanUp: 復原 ScreenUpdating / EnableEvents", "ErrHandler: 顯示錯誤並跳回 CleanUp"] }
          ]
        }
      ]
    }
  ],
  "P5-03": [
    {
      id: "vba-report-pipeline",
      kind: "workflow",
      badge: "專案演示",
      title: "一鍵月報：刷新、檢查、輸出、紀錄",
      subtitle: "實戰不是把程式寫長，而是讓每一段責任清楚、失敗時可追蹤。",
      stages: [
        { text: "刷新資料", tone: "c1" },
        { text: "檢查結果", tone: "c2" },
        { text: "輸出 PDF", tone: "c3" },
        { text: "留下紀錄", tone: "c4" }
      ],
      outcome: "結果：VBA 負責觸發和交付，Power Query 與樞紐負責資料處理。",
      steps: [
        {
          stage: 1,
          label: "1. RefreshAll 觸發可重跑流程",
          tone: "c1",
          text: "清資料邏輯先放在 Power Query，VBA 只呼叫 ThisWorkbook.RefreshAll，避免把清理規則散落在巨集裡。",
          result: "資料更新責任留給 Power Query，VBA 只負責啟動。",
          caption: "這樣下個月資料來了，改 Query 比改 VBA 更好維護。",
          panels: [
            { title: "責任切分", columns: ["階段", "負責工具", "輸出"], rows: [["清資料", "Power Query", "乾淨資料表"], ["彙總", "樞紐/公式", "月報工作表"], ["觸發", "VBA", "一鍵更新"]] }
          ]
        },
        {
          stage: 2,
          label: "2. 檢查必要工作表與輸出區",
          tone: "c2",
          text: "輸出前先檢查月報工作表是否存在、關鍵儲存格是否有錯誤值、資料筆數是否合理。",
          result: "錯誤在輸出 PDF 前被攔下，而不是寄出後才發現。",
          caption: "自動化流程最怕安靜失敗，所以檢查步驟不能省。",
          panels: [
            { title: "檢查表", columns: ["檢查項目", "期待", "失敗處理"], rows: [["月報工作表", "存在", "提示缺工作表"], ["錯誤值", "0 個", "停止輸出"], ["資料筆數", "> 0", "提示來源異常"]] }
          ]
        },
        {
          stage: 3,
          label: "3. 輸出 PDF 到固定路徑",
          tone: "c3",
          text: "用 ExportAsFixedFormat 把月報工作表輸出成 PDF，檔名帶上年月，讓交付物可追蹤。",
          result: "交付檔名固定成 月報_yyyymm.pdf。",
          caption: "路徑與檔名規則固定，使用者才知道輸出去哪裡。",
          panels: [
            { title: "輸出前/後", columns: ["項目", "執行前", "執行後"], rows: [["PDF", "尚未產生", "月報_202605.pdf"], ["路徑", "工作簿資料夾", "同資料夾輸出"], ["提示", "無", "顯示輸出位置"]] }
          ]
        },
        {
          stage: 4,
          label: "4. CleanUp 復原設定並留紀錄",
          tone: "c4",
          text: "不管成功或失敗，都要復原 ScreenUpdating / EnableEvents，並留下執行時間、狀態、錯誤訊息。",
          result: "流程可追蹤，也不會把使用者 Excel 狀態留壞。",
          caption: "可交付自動化的重點，是下一次出錯時找得到線索。",
          panels: [
            { title: "執行紀錄", columns: ["時間", "狀態", "訊息"], rows: [["2026/05/03 12:00", "成功", "PDF 已輸出"], ["2026/05/03 12:05", "失敗", "找不到月報工作表"]] }
          ]
        }
      ]
    }
  ],
  "P5-04": [
    {
      id: "capstone-tool-choice",
      kind: "workflow",
      badge: "整合演示",
      title: "綜合挑戰：先選工具，再動手做",
      subtitle: "最後一課驗收的不是記憶力，而是面對任務時能不能選對 Excel 工具。",
      stages: [
        { text: "辨識任務", tone: "c1" },
        { text: "選工具", tone: "c2" },
        { text: "串流程", tone: "c3" },
        { text: "驗收交付", tone: "c4" }
      ],
      outcome: "結果：公式、樞紐、Power Query、VBA 各自負責適合的段落，不互相硬扛。",
      steps: [
        {
          stage: 1,
          label: "1. 先把題目拆成資料問題",
          tone: "c1",
          text: "不要看到題目就開始寫公式。先拆：是清資料、查找、彙總、視覺化，還是交付流程？",
          result: "任務被拆成多個小問題，而不是一坨模糊需求。",
          caption: "能拆問題，才有機會選對工具。",
          panels: [
            { title: "任務拆解", columns: ["需求", "資料問題", "候選工具"], rows: [["年資獎金", "日期 + 條件判斷", "公式"], ["跨表訂單", "查找 + 彙總", "XLOOKUP + 樞紐"], ["月報合併", "多檔清理", "Power Query"]] }
          ]
        },
        {
          stage: 2,
          label: "2. 依工作型態選工具",
          tone: "c2",
          text: "公式適合列邏輯與查找，樞紐適合快速彙總，Power Query 適合可重跑清理，VBA 適合觸發與輸出。",
          result: "每種工具只做自己最擅長的事。",
          caption: "工具選錯時，通常不是做不到，而是後續很難維護。",
          panels: [
            { title: "工具配對", columns: ["工作", "優先工具", "理由"], rows: [["單列計算", "公式", "即時、可檢查"], ["分類彙總", "樞紐", "快速切片"], ["多表指標", "Data Model", "關聯與量值集中治理"], ["多檔清理", "Power Query", "可重跑"], ["輸出 PDF", "VBA", "控制工作簿操作"]] }
          ]
        },
        {
          stage: 3,
          label: "3. 串成可重跑流程",
          tone: "c3",
          text: "先讓資料進 Power Query，再讓樞紐接乾淨表，最後用 VBA 觸發刷新與輸出。",
          result: "流程從資料來源到交付物有固定順序。",
          caption: "真正可交付的報表，要能下個月重跑。",
          panels: [
            { title: "端到端路徑", lines: ["CSV 資料夾 → Power Query", "乾淨資料表 → 樞紐 / 公式", "月報頁 → VBA 輸出 PDF", "紀錄表 → 留下執行結果"] }
          ]
        },
        {
          stage: 4,
          label: "4. 用新增資料驗收",
          tone: "c4",
          text: "最後不要只看當下答案對不對，要新增一筆資料或替換下個月檔案，確認整條流程仍能更新。",
          result: "新增資料後，報表、圖表、輸出檔都跟著更新。",
          caption: "能重跑，才算真的會做。"
          ,
          panels: [
            { title: "驗收清單", columns: ["檢查", "期待結果", "狀態"], rows: [["新增資料", "公式/樞紐更新", "必測"], ["換下月檔案", "Power Query 重跑", "必測"], ["輸出 PDF", "檔名與路徑正確", "必測"]] }
          ]
        }
      ]
    }
  ]
};
