export const LESSON_DEMOS = {
  "P2-01": [
    {
      id: "sumifs-month-dept",
      kind: "formula",
      badge: "實作動畫",
      title: "SUMIFS：算出一月業務部業績",
      subtitle: "先鎖定金額欄，再用兩個條件把資料縮到真正該加總的列。",
      outcome: "結果：32,000 + 28,000 = 60,000",
      formulaParts: [
        { text: "=SUMIFS(" },
        { text: "C2:C7", step: 1, tone: "c1" },
        { text: ", " },
        { text: "A2:A7", step: 2, tone: "c2" },
        { text: ", " },
        { text: "\"一月\"", step: 3, tone: "c3" },
        { text: ", " },
        { text: "B2:B7", step: 4, tone: "c4" },
        { text: ", " },
        { text: "\"業務\"", step: 5, tone: "c5" },
        { text: ")" }
      ],
      board: {
        title: "原始資料",
        columns: ["月份", "部門", "金額"],
        rows: [
          [
            { text: "一月", marks: { 2: "focus", 3: "match" } },
            { text: "業務", marks: { 4: "focus", 5: "match" } },
            { text: "32,000", marks: { 1: "result", 5: "result" } }
          ],
          [
            { text: "一月", marks: { 2: "focus", 3: "match" } },
            { text: "行政", marks: { 4: "focus" } },
            { text: "18,000", marks: { 1: "result" } }
          ],
          [
            { text: "二月", marks: { 2: "focus" } },
            { text: "業務", marks: { 4: "focus" } },
            { text: "25,000", marks: { 1: "result" } }
          ],
          [
            { text: "一月", marks: { 2: "focus", 3: "match" } },
            { text: "業務", marks: { 4: "focus", 5: "match" } },
            { text: "28,000", marks: { 1: "result", 5: "result" } }
          ],
          [
            { text: "二月", marks: { 2: "focus" } },
            { text: "財務", marks: { 4: "focus" } },
            { text: "22,000", marks: { 1: "result" } }
          ],
          [
            { text: "一月", marks: { 2: "focus", 3: "match" } },
            { text: "客服", marks: { 4: "focus" } },
            { text: "12,000", marks: { 1: "result" } }
          ]
        ]
      },
      steps: [
        {
          label: "1. 加總範圍",
          tone: "c1",
          text: "先指定 Excel 要加總哪一欄。SUMIFS 的第一個參數永遠是最終要加總的數字範圍。",
          caption: "綠色欄位是最後會被加總的金額欄。"
        },
        {
          label: "2. 第一個條件範圍",
          tone: "c2",
          text: "第二個參數不是條件值，而是『用來比對條件的欄位』。這裡先看月份欄。",
          caption: "藍色欄位是第一個條件要檢查的月份欄。"
        },
        {
          label: "3. 條件值 = 一月",
          tone: "c3",
          text: "當條件值填入「一月」後，Excel 先把所有一月的列挑出來，其他月份先排除。",
          caption: "橘色列是第一輪篩選後留下來的候選資料。"
        },
        {
          label: "4. 第二個條件範圍",
          tone: "c4",
          text: "接著再指定第二個條件要看的欄位。這裡要檢查的，是部門欄。",
          caption: "紫色欄位是第二個條件要檢查的部門欄。"
        },
        {
          label: "5. 條件值 = 業務",
          tone: "c5",
          text: "只有同時滿足『一月』和『業務』的列，才會被加進結果。這就是 SUMIFS 的核心思路。",
          caption: "最後亮綠色的兩筆，才是被真正加總的資料。"
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
        src: "/media/p2-03-pivot-demo.gif",
        webp: "/media/p2-03-pivot-demo.webp",
        alt: "樞紐分析操作短片：從原始資料、插入樞紐分析表、把部門拖到列、月份拖到欄、業績拖到值，最後形成彙總報表。",
        note: "這支短片刻意放慢節奏，讓你看清楚每一次拖曳、停頓和報表長出的順序。這一課不再重複放第二套 Pivot 拆步，避免同頁內容重複。"
      },
      outcome: "結果：原始 5 筆交易，逐步變成部門 × 月份彙總表",
      followup: "看完一次後，直接切到練習區照做就好；需要回想時，再回來重播這支短片。"
    }
  ],
  "P4-04": [
    {
      id: "power-query-refresh",
      kind: "workflow",
      badge: "流程動畫",
      title: "Power Query：把每月手工整理變成可重跑流程",
      subtitle: "第一次是在設流程，第二次開始才是真的省時間。",
      outcome: "結果：下個月資料一到，只要按重新整理",
      stages: [
        { text: "連到來源", tone: "c1" },
        { text: "清欄位", tone: "c2" },
        { text: "附加月份", tone: "c3" },
        { text: "重新整理", tone: "c4" }
      ],
      steps: [
        {
          stage: 1,
          label: "1. 先連到原始來源",
          tone: "c1",
          text: "Power Query 的起點不是公式，而是資料來源。你可以連到 CSV、資料夾或其他表格，先把原始檔讀進來。",
          panels: [
            {
              type: "sheet",
              title: "March.csv 原始資料",
              sheetName: "March.csv",
              nameBox: "A2",
              formula: "先讀進原始來源，不急著在這一步整理乾淨",
              columns: ["A", "B", "C", "D"],
              rows: [
                [
                  { text: "訂單日期", tone: "header selection" },
                  { text: "部門", tone: "header selection" },
                  { text: "金額", tone: "header selection" },
                  { text: "備註", tone: "header selection" }
                ],
                [
                  { text: "2026/03/01", tone: "selection" },
                  { text: "電子", tone: "selection" },
                  { text: "32,000", tone: "selection result" },
                  { text: "北區", tone: "selection" }
                ],
                [
                  { text: "2026/03/04", tone: "selection" },
                  { text: "家電", tone: "selection" },
                  { text: "18,500", tone: "selection" },
                  { text: "南區", tone: "selection" }
                ],
                [
                  { text: "2026/03/05", tone: "selection" },
                  { text: "電子", tone: "selection" },
                  { text: "29,000", tone: "selection result" },
                  { text: "北區", tone: "selection" }
                ]
              ],
              activeCell: "A2",
              sidebar: {
                title: "來源觀念",
                items: [
                  "先接受資料很醜是正常的",
                  "Power Query 最擅長接住 CSV / 資料夾 / 匯出報表",
                  "第一次先把來源接進來"
                ]
              },
              footer: "這一步先建立連線，不要急著在工作表手修。"
            }
          ],
          caption: "這時候先接受資料很醜，Power Query 的工作就是接住這種資料。"
        },
        {
          stage: 2,
          label: "2. 把欄位整理乾淨",
          tone: "c2",
          text: "真正有價值的是把每一步都記下來，例如改日期格式、移除備註欄、把金額轉成數值欄位。",
          panels: [
            {
              type: "sheet",
              title: "Query Editor 整理後預覽",
              sheetName: "Query1",
              context: "Power Query Editor",
              nameBox: "A2",
              formula: "已記住：變更型別 / 移除欄位 / 重新命名",
              columns: ["A", "B", "C"],
              rows: [
                [
                  { text: "訂單日期", tone: "header" },
                  { text: "部門", tone: "header" },
                  { text: "金額", tone: "header" }
                ],
                [
                  { text: "2026-03-01", tone: "selection" },
                  { text: "電子", tone: "selection" },
                  { text: "32,000", tone: "result strong" }
                ],
                [
                  { text: "2026-03-04", tone: "selection" },
                  { text: "家電", tone: "selection" },
                  { text: "18,500", tone: "selection" }
                ],
                [
                  { text: "2026-03-05", tone: "selection" },
                  { text: "電子", tone: "selection" },
                  { text: "29,000", tone: "result strong" }
                ]
              ],
              activeCell: "C2",
              sidebar: {
                title: "已記住的步驟",
                items: [
                  "變更型別：訂單日期 → 日期",
                  "變更型別：金額 → 數值",
                  "移除欄位：備註",
                  "欄位重新命名：訂單日期 / 部門 / 金額"
                ]
              },
              footer: "關鍵不是這次清乾淨，而是下次同樣的髒資料也能自動套這些步驟。"
            }
          ],
          caption: "重點不是這次清乾淨，而是下次同樣的髒資料也能自動被清乾淨。"
        },
        {
          stage: 3,
          label: "3. 把每月資料附加成同一張表",
          tone: "c3",
          text: "當三個月份結構一樣時，不要手貼。用附加查詢把它們接成同一張長表，後面樞紐與圖表才會穩。",
          panels: [
            {
              type: "sheet",
              title: "Append 後的正式長表",
              sheetName: "Fact_Sales",
              nameBox: "A2",
              formula: "把 1 月、2 月、3 月資料附加成同一張分析表",
              columns: ["A", "B", "C"],
              rows: [
                [
                  { text: "月份", tone: "header" },
                  { text: "部門", tone: "header" },
                  { text: "金額", tone: "header" }
                ],
                [
                  { text: "1月", tone: "selection" },
                  { text: "電子", tone: "selection" },
                  { text: "120,000", tone: "result strong" }
                ],
                [
                  { text: "1月", tone: "selection" },
                  { text: "家電", tone: "selection" },
                  { text: "80,000", tone: "selection" }
                ],
                [
                  { text: "2月", tone: "selection" },
                  { text: "電子", tone: "selection" },
                  { text: "95,000", tone: "selection" }
                ],
                [
                  { text: "2月", tone: "selection" },
                  { text: "家電", tone: "selection" },
                  { text: "110,000", tone: "selection" }
                ],
                [
                  { text: "3月", tone: "selection" },
                  { text: "電子", tone: "selection" },
                  { text: "132,000", tone: "result strong" }
                ]
              ],
              activeCell: "A2",
              sidebar: {
                title: "現在這張表可拿去",
                items: [
                  "餵樞紐分析表",
                  "餵圖表和儀表板",
                  "餵資料模型與 Power Pivot"
                ]
              },
              footer: "真正專業的做法，是把月資料做成同一張長表，而不是每月各一張表。"
            }
          ],
          caption: "這時你得到的是可以餵給樞紐、圖表與資料模型的正式長表。"
        },
        {
          stage: 4,
          label: "4. 下個月只按重新整理",
          tone: "c4",
          text: "真正的時間節省發生在第二次之後。四月檔案放進資料夾後，你不再重做步驟，只是重跑同一條流程。",
          panels: [
            {
              type: "sheet",
              title: "重新整理後的月營收報表",
              sheetName: "Monthly_Report",
              nameBox: "A2",
              formula: "新增 2026-04.csv 後按重新整理，報表自動更新",
              columns: ["A", "B", "C", "D"],
              rows: [
                [
                  { text: "月份", tone: "header" },
                  { text: "電子", tone: "header" },
                  { text: "家電", tone: "header" },
                  { text: "總計", tone: "header" }
                ],
                [
                  { text: "1月", tone: "selection" },
                  { text: "120,000", tone: "selection" },
                  { text: "80,000", tone: "selection" },
                  { text: "200,000", tone: "result strong" }
                ],
                [
                  { text: "2月", tone: "selection" },
                  { text: "95,000", tone: "selection" },
                  { text: "110,000", tone: "selection" },
                  { text: "205,000", tone: "result strong" }
                ],
                [
                  { text: "3月", tone: "selection" },
                  { text: "132,000", tone: "selection" },
                  { text: "98,000", tone: "selection" },
                  { text: "230,000", tone: "result strong" }
                ],
                [
                  { text: "4月", tone: "focus" },
                  { text: "141,000", tone: "focus result strong" },
                  { text: "105,000", tone: "focus" },
                  { text: "246,000", tone: "focus result strong" }
                ]
              ],
              activeCell: "A5",
              sidebar: {
                title: "四月進來後",
                items: [
                  "新增來源：2026-04.csv",
                  "動作：重新整理查詢",
                  "新資料自動套用既有清理步驟",
                  "樞紐與圖表同步更新"
                ]
              },
              footer: "真正省時間的不是第一次，而是之後每個月都不用再重做。"
            }
          ],
          caption: "這就是 Power Query 比手動整理更專業的地方。"
        }
      ]
    }
  ],
  "P4-05": [
    {
      id: "data-model-relationship",
      kind: "workflow",
      badge: "概念動畫",
      title: "資料模型：把多張表串成同一套分析系統",
      subtitle: "Power Query 清資料，資料模型管關聯，樞紐只負責把答案展示出來。",
      outcome: "結果：一個量值，可以供多張樞紐共用",
      stages: [
        { text: "匯入多表", tone: "c1" },
        { text: "建立關聯", tone: "c2" },
        { text: "寫量值", tone: "c3" },
        { text: "交給樞紐", tone: "c4" }
      ],
      steps: [
        {
          stage: 1,
          label: "1. 先把事實表和維度表分開",
          tone: "c1",
          text: "資料模型不是把所有欄位攤成一張超大表，而是把『交易』和『主檔』各自保留，再用關聯串起來。",
          panels: [
            {
              title: "訂單表（事實表）",
              columns: ["訂單ID", "產品ID", "金額"],
              rows: [
                ["SO-001", "P001", "32,000"],
                ["SO-002", "P003", "18,500"],
                ["SO-003", "P001", "29,000"]
              ]
            },
            {
              title: "產品表（維度表）",
              columns: ["產品ID", "產品名稱", "類別"],
              rows: [
                ["P001", "無線滑鼠", "周邊"],
                ["P003", "27 吋螢幕", "顯示"],
                ["P005", "網路攝影機", "周邊"]
              ]
            }
          ],
          caption: "關鍵不是表多，而是每張表有自己的角色。"
        },
        {
          stage: 2,
          label: "2. 用共同鍵值建立關聯",
          tone: "c2",
          text: "資料模型的核心不是公式，而是關聯。只要產品 ID 能對上，訂單表就能借用產品表的類別與名稱。",
          panels: [
            {
              title: "關聯邏輯",
              lines: [
                "訂單[產品ID] → 產品[產品ID]",
                "一個產品，可對應多筆訂單",
                "這等於在模型層先把 JOIN 定義好"
              ]
            }
          ],
          caption: "之後的樞紐不需要再寫 XLOOKUP，把關聯交給模型就好。"
        },
        {
          stage: 3,
          label: "3. 在模型上寫量值",
          tone: "c3",
          text: "量值（Measure）和一般公式的差別在於：它是在模型層定義的計算規則，可以被很多報表共用。",
          panels: [
            {
              title: "量值示意",
              lines: [
                "總業績 := SUM(訂單[金額])",
                "周邊業績 := CALCULATE([總業績], 產品[類別] = \"周邊\")",
                "重點：指標定義只寫一次"
              ]
            }
          ],
          caption: "這就是資料模型比傳統樞紐更像 BI 工具的地方。"
        },
        {
          stage: 4,
          label: "4. 最後才交給樞紐展示",
          tone: "c4",
          text: "當關聯和量值都在模型裡定好了，樞紐分析表只是切片器與展示層，不再承擔計算邏輯本身。",
          panels: [
            {
              title: "樞紐可直接使用",
              columns: ["類別", "總業績"],
              rows: [
                ["周邊", "61,000"],
                ["顯示", "18,500"]
              ]
            }
          ],
          caption: "樞紐只是出口，真正的商業邏輯放在模型裡。"
        }
      ]
    }
  ]
};
