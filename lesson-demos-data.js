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
        src: "/media/p2-03-pivot-demo.webp",
        alt: "樞紐分析操作短片：從原始資料、插入樞紐分析表、把部門拖到列、月份拖到欄、業績拖到值，最後形成彙總報表。",
        note: "這支短片刻意放慢節奏，讓你看清楚每一次拖曳、停頓和報表長出的順序。這一課不再重複放第二套 Pivot 拆步，避免同頁內容重複。"
      },
      outcome: "結果：原始 5 筆交易，逐步變成部門 × 月份彙總表",
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
        alt: "條件式格式化操作短片：選取業績欄位、建立規則、套用資料橫條與醒目提醒，最後形成可快速判讀的月報。",
        note: "這支短片聚焦在三個實務動作：業績橫條、負成長標紅、客訴偏高提醒。看懂之後再去練習區自己重做，會比先讀一堆規則更容易進狀況。"
      },
      outcome: "結果：同一張月報裡，高績效、負成長和高客訴會自己浮出來",
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
        alt: "資料驗證操作短片：設定部門清單、開啟下拉選單、輸入錯誤手機時跳出提醒，最後形成乾淨可交付的表單。",
        note: "這支短片故意把節奏放慢，讓你看清楚規則設定、使用者實際填表，以及錯誤被阻止時的差別。資料驗證真正有價值的地方，就是把清理成本往前移。"
      },
      outcome: "結果：部門、日期、手機這類常見錯誤，在輸入當下就被攔下來",
      followup: "練習時先從最簡單的『清單』和『文字長度』開始，等手感穩了再做連動下拉與自訂公式。"
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
        alt: "圖表設計操作短片：選取月份、營收和成長率資料，插入組合圖，刪除雜訊並加上有觀點的標題，最後形成更專業的圖表。",
        note: "這裡刻意不是教你所有圖表類型，而是示範一條最常見的職場流程：把原始自動圖修成主管願意直接貼進簡報的版本。"
      },
      outcome: "結果：同一份資料，從『只是有圖』變成『真的有在講結論』",
      followup: "自己重做時先練三件事：圖型選對、格線刪掉、標題改成有觀點。這三件就已經能拉開很大差距。"
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
        alt: "Power Query 操作短片：連到原始 CSV、記住清理步驟、附加每月資料，最後在新增四月檔後按重新整理更新報表。",
        note: "這支短片保留了最有感的四個瞬間：接來源、記步驟、附加月份、重新整理。先建立整體節奏感，再去課文和練習裡補細節，會比較容易懂為什麼 Power Query 值得學。"
      },
      outcome: "結果：下個月資料一到，只要按重新整理，不用整套重做",
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
        alt: "資料模型操作短片：先分開訂單表與產品表，建立產品ID關聯，寫量值公式，最後把結果交給樞紐分析表展示。",
        note: "這一課本來就偏觀念，所以短片重點不是花俏拖曳，而是把『多表 → 關聯 → 量值 → 樞紐』這條路徑建立起來。對 Mac 使用者尤其重要，因為你更需要判斷何時該交給模型，而不是硬把所有邏輯塞進樞紐。"
      },
      outcome: "結果：量值只定義一次，之後很多張報表都能共用同一套邏輯",
      followup: "學這課時先把表的角色和關聯想清楚，比急著背 DAX 函數更重要。"
    }
  ]
};
