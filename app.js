
import { SEARCH_INDEX } from './search-index.js';
import { QUIZ_CARDS } from './quiz-data.js';

function pageKey(){
  const parts = location.pathname.split("/").filter(Boolean);
  return parts.slice(-2).join("/");
}
const PK = pageKey();
const DEPTH = document.body?.dataset.depth || "";
const idx = SEARCH_INDEX;
const LESSON_SEQUENCE = [
  "P1-01","P1-02","P1-03",
  "P2-01","P2-02","P2-03","P2-04",
  "P3-01","P3-02","P3-03","P3-04","P3-05",
  "P4-01","P4-02","P4-03","P4-04","P4-05",
  "P5-01","P5-02","P5-03","P5-04",
];
const LESSON_WORKBOOK_MAP = {
  "P1-01": "P1-01 ⌨️ 鍵盤操作",
  "P1-02": "P1-02 📊 基礎函式",
  "P1-03": "P1-03 📐 條件判斷",
  "P2-01": "P2-01 📈 條件統計",
  "P2-02": "P2-02 🔍 查找比對",
  "P2-03": "P2-03 📊 樞紐分析表",
  "P2-04": "P2-04 🎨 條件式格式化",
  "P3-01": "P3-01 ✅ 資料驗證",
  "P3-02": "P3-02 ✂️ 文字與日期",
  "P3-03": "P3-03 📈 圖表設計",
  "P3-04": "P3-04 🔗 命名範圍",
  "P3-05": "P3-05 🛡️ 保護與安全",
  "P4-01": "P4-01 🔢 陣列與動態",
  "P4-02": "P4-02 🧮 進階函式",
  "P4-03": "P4-03 🧩 假設分析",
  "P4-04": "P4-04 ⚡ Power Query",
  "P4-05": "P4-05 🔗 資料模型",
  "P5-01": "P5-01 ⚙️ VBA 基礎",
  "P5-02": "P5-02 🔧 VBA 進階",
  "P5-03": "P5-03 🏗️ VBA 實戰",
  "P5-04": "P5-04 🏆 綜合挑戰",
};
const LESSON_GUIDE = {
  "P1-01": {
    focus: "先把手從滑鼠搬回鍵盤，建立最基本的 Excel 操作節奏。",
    unlock: "做資料整理、寫公式、切換工作表時都不會一直被滑鼠打斷。",
    nextWhy: "下一課會開始碰公式，快捷鍵和輸入節奏先熟，後面學函數會輕鬆很多。",
  },
  "P1-02": {
    focus: "把 SUM、AVERAGE、COUNT 這些最常用統計函數變成反射動作。",
    unlock: "看到一欄數字時，會自然知道要怎麼加總、取平均、算筆數和找排名。",
    nextWhy: "有了基本函數後，下一課才能進到 IF / IFS 這種會做判斷的公式。",
  },
  "P1-03": {
    focus: "讓 Excel 不只會算數，還會根據條件做出不同結果。",
    unlock: "你會開始理解公式的邏輯結構，為 Phase 2 的條件統計打底。",
    nextWhy: "下一課的 SUMIF / COUNTIF 本質上就是把條件判斷搬進統計公式裡。",
  },
  "P2-01": {
    focus: "把條件邏輯和加總、計數結合，開始做真正的報表分析。",
    unlock: "老闆問某部門、某月份、某類型資料時，你可以直接用公式答出來。",
    nextWhy: "做完條件統計後，下一課會進到跨表查找，補齊職場最常見的第二塊能力。",
  },
  "P2-02": {
    focus: "學會在不同工作表之間拉資料，解決職場最常見的查表任務。",
    unlock: "你能處理代碼對照、名單補欄位、跨檔查價格這類高頻需求。",
    nextWhy: "下一課的樞紐分析會把查回來、整理好的資料做成真正可讀的報表。",
  },
  "P2-03": {
    focus: "把大量原始資料快速變成可以閱讀與回答問題的報表。",
    unlock: "你會知道什麼時候該用公式，什麼時候該直接上樞紐比較快。",
    nextWhy: "有了樞紐之後，下一課會教你把重點自動凸顯，讓報表更容易被看懂。",
  },
  "P2-04": {
    focus: "把報表從『算得出來』升級成『一眼看得懂』。",
    unlock: "你能用顏色、資料橫條和公式規則，讓異常值和重點自己跳出來。",
    nextWhy: "下一個 phase 會進到表單與資料品質，從『看懂資料』走向『避免髒資料』。",
  },
  "P3-01": {
    focus: "先把輸入端管好，減少後面清資料和修資料的時間。",
    unlock: "你會知道怎麼限制輸入格式、做下拉選單、避免表單被亂填。",
    nextWhy: "資料品質穩了之後，下一課就能更有效處理文字與日期這些常見髒資料。",
  },
  "P3-02": {
    focus: "把最常見的文字清理和日期處理手法一次補齊。",
    unlock: "你能拆字串、補格式、算期間，處理從外部貼進來的雜亂資料。",
    nextWhy: "資料整理乾淨後，下一課才適合做圖表，否則視覺化只會放大混亂。",
  },
  "P3-03": {
    focus: "學會選對圖表，不讓圖表只是漂亮而已。",
    unlock: "你能根據資料類型選柱狀圖、折線圖、圓餅圖或散佈圖。",
    nextWhy: "下一課會把資料來源本身變穩，用表格和命名範圍讓整套報表更耐用。",
  },
  "P3-04": {
    focus: "把公式和資料來源變得可讀、可維護、可自動擴張。",
    unlock: "你會開始寫出半年後回來也看得懂的公式，而不是只有當下能用。",
    nextWhy: "報表結構穩了之後，下一課再加上保護與安全，才算能安心交付別人使用。",
  },
  "P3-05": {
    focus: "替報表補上最後一道防線，避免公式和結構被誤改。",
    unlock: "你會知道什麼該鎖、什麼該留給使用者填，交付檔案時更安心。",
    nextWhy: "接下來會進到進階自動化，從保護既有流程走向減少手工流程本身。",
  },
  "P4-01": {
    focus: "學會用動態陣列一次吐出整片結果，減少輔助欄和重複公式。",
    unlock: "你能開始用 FILTER、SORT、UNIQUE 這些新世代函數處理大量清單。",
    nextWhy: "下一課會進一步用 LET、LAMBDA 等把公式模組化，處理更複雜的場景。",
  },
  "P4-02": {
    focus: "把公式從『能算』再往前推到『能封裝、能重用』。",
    unlock: "你會知道何時該用 LET 提升可讀性，何時該用 LAMBDA 做可重複邏輯。",
    nextWhy: "公式能力再往上後，下一課會換一種思考方式，開始做假設分析與反推。",
  },
  "P4-03": {
    focus: "從單向計算切到反推決策，讓 Excel 幫你找出達標條件。",
    unlock: "你能用目標搜尋、規劃求解和資料表做更接近商業決策的分析。",
    nextWhy: "下一課會把焦點轉到 Power Query，解決大量重複清資料的流程問題。",
  },
  "P4-04": {
    focus: "把一再重複的清資料流程變成可以重新整理的一鍵流程。",
    unlock: "你會理解 Extract / Transform / Load 的流程，讓 Excel 更像資料工具。",
    nextWhy: "學會 Query 後，下一課就能把多表關聯和量值計算交給 Power Pivot。",
  },
  "P4-05": {
    focus: "把 Excel 從單表思維提升到資料模型思維。",
    unlock: "你能處理跨表關聯、建立量值，開始接近 BI 工具的分析方式。",
    nextWhy: "最後一個 phase 會把這些能力接到 VBA，自動化整個處理流程與交付流程。",
  },
  "P5-01": {
    focus: "第一次讓 Excel 替你執行一連串步驟，而不是自己手動點完。",
    unlock: "你會知道錄巨集、開 VBE、寫基本變數與流程控制的實戰起點。",
    nextWhy: "下一課會把 VBA 從入門腳本推到更像真正程式：陣列、字典、錯誤處理。",
  },
  "P5-02": {
    focus: "把 VBA 從會跑提升到更穩、更快、更像工程化腳本。",
    unlock: "你能處理大量資料、避免卡頓，也知道程式壞掉時怎麼保底。",
    nextWhy: "下一課會把這些能力放進完整專案，做真正的自動化交付流程。",
  },
  "P5-03": {
    focus: "把分散技巧組成完整專案，從讀資料到產報表都交給 VBA。",
    unlock: "你會看到宏觀流程怎麼設計，而不是只會寫零散的小巨集。",
    nextWhy: "最後一課會用綜合挑戰驗收整張學習地圖，看看哪些能力還需要補強。",
  },
  "P5-04": {
    focus: "把前面 20 課的能力串起來，確認自己已經能獨立解題。",
    unlock: "你會更清楚自己在哪一段已經穩了、哪一段還值得回頭補強。",
    nextWhy: "做完後最有價值的下一步不是再往下衝，而是回頭重做弱點課與建立自己的模板。",
  },
};
const LESSON_PRO_NOTES = {
  "P1-01": {
    title: "Mac 鍵位節奏是你的操作地基",
    eyebrow: "Professional Default",
    items: [
      "這站以 macOS Excel 為主路線。你在這裡學到的快捷鍵，預設就是 Mac 版 Excel 的實際按法，不是從 Windows 翻譯過來的。",
      "先把 ⌘ S / ⌘ Z / ⌘+方向鍵 / F2 這幾個練成反射，後面學公式時才不會一直被操作介面卡住。",
      "WPS Office 的說明是補充資訊，不是主路線。如果你主要用 Excel for Mac，照 Excel 欄位的做法走即可。",
    ],
  },
  "P1-02": {
    title: "先把基礎函數練穩，再往複雜走",
    eyebrow: "Professional Default",
    items: [
      "SUM / AVERAGE / COUNT / COUNTA 看起來簡單，但很多錯誤報表都是因為混用 COUNT 和 COUNTA、或誤解 AVERAGE 對空白與 0 的處理而出現的。",
      "這課的函數在 Excel for Mac 與 Windows 版行為完全相同，不需要注意平台差異。",
      "這課是後面所有進階課的地基：先把這幾個練到不用想，SUMIFS / XLOOKUP 才會學得順。",
    ],
  },
  "P1-03": {
    title: "知道 IF 的邊界，才知道什麼時候該出場",
    eyebrow: "Professional Default",
    items: [
      "IFS 比巢狀 IF 好讀，且不容易出現括號配對錯誤。有 IFS 可以用的環境，沒有理由繼續寫 IF(A, X, IF(B, Y, ...))。",
      "IF 適合處理欄位層級的條件邏輯。如果你需要『依條件加總或計數』，那是 SUMIFS / COUNTIFS 的場景，不要用 IF 再配 SUM 硬撐。",
      "這課函數的語法在 macOS Excel 和 Windows 版相同，IFS 在 Excel for Mac 2019 以後可用。",
    ],
  },
  "P2-01": {
    title: "條件統計是報表工作的基本盤",
    eyebrow: "Professional Default",
    items: [
      "SUMIFS、COUNTIFS 這類條件統計公式，通常比複雜巢狀 IF 更適合做正式報表。",
      "如果需求只是依條件加總與計數，先用條件統計公式；不要太早把簡單問題做成很重的模型。",
      "真正專業的差別，在於你知道什麼時候該停在公式層，什麼時候才該進樞紐或 Power Query。",
    ],
  },
  "P2-02": {
    title: "專業現場怎麼選查找工具",
    eyebrow: "Professional Default",
    items: [
      "如果是 Microsoft 365 或 Excel 2021 以上，預設優先用 XLOOKUP；它不用手算欄位，也能向左查找。",
      "如果你在維護舊公司檔案或 2016 / 2019 環境，請把 INDEX + MATCH 視為真正的通用備案，不要只會 VLOOKUP。",
      "職場上真正專業的差別不是背函數名，而是先判斷同事的版本、檔案相容性和後續維護成本。",
    ],
  },
  "P2-03": {
    title: "樞紐不是取代公式，而是接手彙總",
    eyebrow: "Professional Default",
    items: [
      "樞紐分析最適合接手大量原始資料的彙總、切片與觀察，不適合拿來取代所有細部商業規則。",
      "比較穩的工作流通常是：原始資料先整理成 Table，再用樞紐做彙總，必要時才補公式欄位。",
      "你越早建立『先整理資料，再彙總報表』的習慣，後面接 Power Query 會越順。",
    ],
  },
  "P2-04": {
    title: "格式化要幫助決策，不只是上色",
    eyebrow: "Professional Default",
    items: [
      "條件式格式化最有價值的地方，是幫人快速看到異常、門檻與趨勢，而不是把整張表染滿顏色。",
      "正式報表應先定義哪些數值真的需要被看見，再設規則，不要讓格式化搶走資料本身的可讀性。",
      "如果每個欄位都在亮，等於沒有任何欄位真正重要。",
    ],
  },
  "P3-01": {
    title: "專業報表先從防呆開始",
    eyebrow: "Professional Default",
    items: [
      "資料驗證的價值不是限制使用者，而是把髒資料擋在輸入端，減少之後清理與追錯成本。",
      "如果一份工作簿會交給別人長期輸入，資料驗證幾乎應該視為預設，而不是可選配件。",
      "真正穩定的流程，永遠比事後補救更省時間。",
    ],
  },
  "P3-02": {
    title: "清資料能力決定你能不能接住真實世界資料",
    eyebrow: "Professional Default",
    items: [
      "文字與日期函數看起來零碎，但它們處理的是最常見的外部貼上資料、系統匯出資料與格式混亂問題。",
      "正式工作流中，先把欄位定義清楚，再清理文字與日期，會比邊清邊猜更穩。",
      "當清理步驟開始重複出現時，就是往 Power Query 過渡的訊號。",
    ],
  },
  "P3-03": {
    title: "圖表真正的專業感來自選對，不是堆特效",
    eyebrow: "Professional Default",
    items: [
      "專業圖表最重要的是讓讀者更快理解，不是讓圖表本身更花。",
      "先決定你要比較類別、看趨勢、看占比，圖表類型自然就會縮小到少數幾種合理選項。",
      "如果資料本身還沒整理乾淨，再漂亮的圖表也只是在放大混亂。",
    ],
  },
  "P3-04": {
    title: "專業工作簿的預設習慣",
    eyebrow: "Professional Default",
    items: [
      "正式報表或要長期維護的資料，應先轉成 Table，再開始寫公式與做樞紐。",
      "結構化參照的價值不只是好看，而是新增列、欄位調整、交接給同事時都更穩。",
      "專業 Excel 使用者通常不是公式最炫的人，而是最能讓檔案在三個月後還維持可讀性的人。",
    ],
  },
  "P3-05": {
    title: "保護的目的，是讓檔案能安心交付",
    eyebrow: "Professional Default",
    items: [
      "工作表保護不是為了把檔案藏起來，而是區分哪些欄位可以填、哪些公式不該被誤改。",
      "真正專業的交付，會把可輸入區、計算區、輸出區責任分清楚，再決定保護策略。",
      "如果沒有先整理好結構，光上鎖只會讓後續維護更痛苦。",
    ],
  },
  "P4-01": {
    title: "動態陣列在實務上的邊界",
    eyebrow: "Professional Default",
    items: [
      "動態陣列適合拿來輸出分析結果與中間結果，但溢出公式本身不應該寫在 Table 裡。",
      "比較穩的做法是：原始資料放 Table，動態陣列公式寫在表格外，用結構化參照去吃資料。",
      "如果你要交付給舊版 Excel 使用者，請提前規劃降級方案，否則動態函數會變成相容性風險。",
    ],
  },
  "P4-02": {
    title: "進階函數的重點是可讀與可重用",
    eyebrow: "Professional Default",
    items: [
      "LET 和 LAMBDA 的價值，不是把公式寫得更炫，而是把重複邏輯封裝得更清楚。",
      "當一條公式長到同事看不懂時，先想可讀性與拆解，而不是再往裡面硬塞一層巢狀。",
      "如果需求會被反覆複製到很多欄位，才真正值得把邏輯做成可重用的公式模組。",
    ],
  },
  "P4-03": {
    title: "假設分析是決策工具，不是計算工具",
    eyebrow: "Professional Default",
    items: [
      "目標搜尋適合「只改一個變數、反推輸入」的場景；Solver 才適合多變數加約束條件的最佳化——不要一開始就用 Solver，先看目標搜尋夠不夠。",
      "規劃求解（Solver）在 macOS 上需要先到「工具 → Excel 增益集」啟用，不是預設開啟的。",
      "這課工具的價值在商業決策推演，不是取代公式計算。真正的報表通常是公式算結果，假設分析是反問「要達到這個結果，條件得是什麼」。",
    ],
  },
  "P4-04": {
    title: "Power Query 為什麼是專業分水嶺",
    eyebrow: "Professional Default",
    items: [
      "真正專業的差別，通常不在於手動清資料多快，而在於能不能把清理流程變成可重跑的步驟。",
      "Power Query 的核心思維是 Connect → Transform → Combine → Load；一旦建立好，之後應優先按重新整理，而不是重做。",
      "如果流程每週、每月都要重複一次，通常就該先想 Power Query，而不是再多寫幾欄公式硬撐。",
    ],
  },
  "P4-05": {
    title: "什麼時候該進資料模型",
    eyebrow: "Professional Default",
    items: [
      "只要開始遇到多張表關聯、數十萬列以上資料，或需要穩定量值計算，就應該考慮 Data Model / Power Pivot。",
      "這一塊最像真正的分析建模，不再只是單一工作表上的公式技巧。",
      "如果團隊常做跨表分析與經營報表，這課其實是從 Excel 進入 BI 思維的重要門檻。",
    ],
  },
  "P5-01": {
    title: "專業自動化的起點不是炫技",
    eyebrow: "Professional Default",
    items: [
      "VBA 最適合拿來解決重複、規則清楚、人工又容易出錯的流程。",
      "專業做法不是一開始就寫很大的巨集，而是先把手動流程拆清楚，再自動化最穩定的一段。",
      "如果資料清理是可重跑流程，先想 Power Query；如果是 Excel 內部操作與報表交付流程，再考慮 VBA。",
    ],
  },
  "P5-02": {
    title: "VBA 進階真正要學的是穩定性",
    eyebrow: "Professional Default",
    items: [
      "專業 VBA 不只是跑得動，而是大量資料也跑得穩、出錯時知道怎麼停在安全的位置。",
      "陣列、Dictionary、錯誤處理這些看起來比較工程化，但它們真正解決的是交付品質與維護成本。",
      "如果你的流程邏輯還沒固定，先不要急著把所有步驟都寫成巨集；先穩定流程，再談優化速度。",
    ],
  },
  "P5-04": {
    title: "能選對工具，才算真的學完了",
    eyebrow: "Professional Default",
    items: [
      "綜合挑戰的重點不是把題目做完，而是每題開始前先判斷：這是公式題、樞紐題、Power Query 題，還是 VBA 題？",
      "能選對工具再動手，比堆功能完成題目更接近真正的職場能力。",
      "這課所有挑戰題在 macOS Excel 都可以完整執行，不依賴任何 Windows-only 功能。",
    ],
  },
  "P5-03": {
    title: "把巨集升級成可交付的工作系統",
    eyebrow: "Professional Default",
    items: [
      "實戰專案最重要的不是功能越多越好，而是輸入、處理、輸出三段責任要清楚。",
      "真正專業的自動化專案會先想檔名規則、錯誤處理、執行前檢查與結果驗證，而不是只想主流程。",
      "當流程會反覆跑在每月報表或批次檔案上時，文件化和可重複執行比一時寫得快更重要。",
    ],
  },
};
const LESSON_BADGES = {
  "P1-01": [
    { kind: "mac", label: "macOS 主路線", note: "這站以 Mac 鍵位與 Excel for Mac 操作為主，不混用 Windows 快捷鍵。" },
  ],
  "P1-02": [
    { kind: "pro", label: "專業基礎", note: "先學會穩定算數，再碰更複雜的邏輯與查找。" },
  ],
  "P1-03": [
    { kind: "workflow", label: "邏輯基礎", note: "這課是從算數跨到條件判斷的分水嶺，後面 SUMIF / COUNTIF 都靠這個思維。" },
    { kind: "pro", label: "專業基礎", note: "先把條件邏輯學穩，後面做報表規則才不會一直繞路。" },
  ],
  "P2-01": [
    { kind: "workflow", label: "報表核心", note: "條件統計是多數日常報表與 KPI 彙總的核心工具。" },
    { kind: "pro", label: "專業預設", note: "能用 SUMIFS / COUNTIFS 解決的彙總，通常比複雜巢狀公式更穩。" },
  ],
  "P2-02": [
    { kind: "mac", label: "macOS 可用", note: "XLOOKUP 在 Excel for Microsoft 365 for Mac 與 Excel 2021 for Mac 可用。" },
    { kind: "version", label: "M365 / 2021+", note: "新版環境優先用 XLOOKUP。" },
    { kind: "fallback", label: "舊版備案", note: "2016 / 2019 更該熟 INDEX + MATCH。" },
    { kind: "pro", label: "職場核心", note: "跨表查找幾乎是所有報表工作的基本功。" },
  ],
  "P2-03": [
    { kind: "mac", label: "macOS 可用", note: "PivotTable、Slicer、PivotChart 在新版 Excel for Mac 都可用。" },
    { kind: "pro", label: "專業預設", note: "正式資料先轉 Table，再做樞紐會更穩。" },
    { kind: "workflow", label: "報表入口", note: "這課開始從算公式走向真正整理報表。" },
  ],
  "P2-04": [
    { kind: "workflow", label: "可視化判讀", note: "這課重點是讓異常值與重點自己浮出來，提升報表可讀性。" },
    { kind: "pro", label: "專業預設", note: "格式化應幫助判讀，不應該把整張表染成裝飾品。" },
  ],
  "P3-01": [
    { kind: "workflow", label: "防呆輸入", note: "這課是從分析端回頭補輸入端品質，直接影響後面清資料成本。" },
    { kind: "pro", label: "專業預設", note: "長期會給別人填的表單，資料驗證應該是預設配備。" },
  ],
  "P3-02": [
    { kind: "workflow", label: "清資料核心", note: "真實工作裡最常見的不是新函數，而是貼上後亂掉的文字與日期。" },
    { kind: "pro", label: "專業預設", note: "當清理動作開始重複，應該意識到這是往 Power Query 過渡的前兆。" },
  ],
  "P3-03": [
    { kind: "workflow", label: "溝通輸出", note: "這課是把分析結果轉成別人看得懂的輸出，不只是畫圖。" },
    { kind: "pro", label: "專業預設", note: "先想比較、趨勢、占比哪一種問題，再決定圖表類型。" },
  ],
  "P3-04": [
    { kind: "pro", label: "專業預設", note: "長期維護的檔案，Table 和結構化參照幾乎是必備。" },
    { kind: "workflow", label: "維護性", note: "這課直接影響交接、擴充和長期穩定度。" },
  ],
  "P3-05": [
    { kind: "workflow", label: "交付安全", note: "這課重點是讓可填區與公式區分清楚，避免交付後被誤改。" },
    { kind: "pro", label: "專業預設", note: "保護應建立在清楚的結構分層上，不是先鎖再說。" },
  ],
  "P4-01": [
    { kind: "mac", label: "macOS 視版本", note: "動態陣列在較新的 Mac 版 Excel 可用，交付舊版前要先確認環境。" },
    { kind: "version", label: "M365 / 2021+", note: "動態陣列屬於新世代 Excel 能力。" },
    { kind: "compat", label: "相容性注意", note: "交付舊版使用者前，要先規劃降級方案。" },
  ],
  "P4-02": [
    { kind: "mac", label: "macOS 視版本", note: "LET、LAMBDA 等進階函數需要較新的 Excel 版本，交付前請先確認環境。" },
    { kind: "version", label: "M365 / 2021+", note: "這課屬於新世代 Excel 公式能力。" },
    { kind: "pro", label: "專業預設", note: "進階函數的重點是提升可讀性與重用性，不是把公式寫得更長。" },
  ],
  "P4-03": [
    { kind: "mac", label: "macOS 可用", note: "目標搜尋可直接用；若要用規劃求解，先到 Tools > Excel Add-ins 啟用 Solver。" },
    { kind: "workflow", label: "決策分析", note: "這課是從公式計算走向商業決策推演的入口。" },
  ],
  "P4-04": [
    { kind: "mac", label: "macOS 可用", note: "Power Query 已支援 Excel for Microsoft 365 for Mac，但部分資料來源與功能深度仍要看版本。" },
    { kind: "workflow", label: "可重跑流程", note: "重複清資料時，應優先想 Power Query。" },
    { kind: "platform", label: "平台差異", note: "Mac 版可匯入與整形多種資料，但像部分資料來源仍可能比 Windows 更受限。" },
  ],
  "P4-05": [
    { kind: "mac", label: "macOS 非主路線", note: "Power Pivot 不屬於 Office for Mac；這課在 Mac 上更適合理解資料模型概念，而不是當主要操作路線。" },
    { kind: "pro", label: "分析建模", note: "這課是從 Excel 技巧邁向 BI 思維的重要門檻。" },
    { kind: "workflow", label: "多表關聯", note: "遇到關聯、多表、量值時，就該開始思考 Data Model。" },
  ],
  "P5-01": [
    { kind: "mac", label: "macOS 可用", note: "Excel for Mac 支援 Developer tab、錄製 macro，以及用 VBE 撰寫 VBA。" },
    { kind: "workflow", label: "工具分工", note: "流程可重跑先想 Power Query，Excel 操作自動化再想 VBA。" },
    { kind: "pro", label: "不是炫技", note: "真正專業是自動化穩定流程，不是寫最長的巨集。" },
  ],
  "P5-02": [
    { kind: "mac", label: "macOS 可用", note: "陣列、錯誤處理、事件與大部分 VBA 核心語法都可在 Excel for Mac 使用。" },
    { kind: "compat", label: "相容性注意", note: "部分 Windows 專屬 COM / ActiveX 自動化做法不能直接搬到 Mac。" },
    { kind: "workflow", label: "穩定交付", note: "這課的重點是讓巨集跑得穩、改得動，而不是只追求功能變多。" },
  ],
  "P5-03": [
    { kind: "mac", label: "macOS 可用", note: "用 VBA 串接工作簿內流程與批次處理，在 Mac 上仍是可行主路線。" },
    { kind: "workflow", label: "專案化交付", note: "這課開始把前面零散技巧組成完整輸入 → 處理 → 輸出的系統。" },
    { kind: "pro", label: "專業預設", note: "先定輸入規則與錯誤處理，再做月報自動化，專案才會穩。" },
  ],
  "P5-04": [
    { kind: "mac", label: "macOS 主路線", note: "這課重點是整合 Excel 能力，而不是依賴某個單一 Windows-only 功能。" },
    { kind: "workflow", label: "能力驗收", note: "建議把挑戰題當成真實工作流練習，而不是只求做完。" },
  ],
};
const LESSON_BADGE_KIND_ORDER = [
  "mac",
  "version",
  "fallback",
  "pro",
  "workflow",
  "platform",
  "compat",
];
const HOME_CAPABILITY_MAP = {
  title: "從新手走到專業的能力地圖",
  intro: "真正專業不是會很多功能，而是知道什麼時候該用哪一種方法、怎麼做出可維護、可交接、可重跑的工作流程。",
  defaults: [
    "本站以 macOS Excel 為主；遇到 Windows-only 或 Mac 受限功能，會直接標出來。",
    "原始資料優先轉成 Table，再開始做公式、樞紐與分析。",
    "新版 Excel 能用 XLOOKUP 就不要把 VLOOKUP 當唯一答案。",
    "重複清資料流程先想 Power Query，不要只靠手動欄位硬撐。",
    "多表關聯與量值分析要開始用 Data Model / Power Pivot 思維。",
  ],
  stages: [
    {
      eyebrow: "Stage 1",
      title: "操作與公式基本功",
      summary: "先把輸入速度、基礎函數與條件邏輯練成反射動作。",
      skills: ["快捷鍵節奏", "基礎函數", "IF / IFS 判斷"],
    },
    {
      eyebrow: "Stage 2",
      title: "職場報表核心",
      summary: "學會查找、條件統計、樞紐與格式化，開始能獨立做分析報表。",
      skills: ["XLOOKUP / INDEX+MATCH", "SUMIFS", "PivotTable", "條件式格式化"],
    },
    {
      eyebrow: "Stage 3",
      title: "資料結構與可維護性",
      summary: "從會做結果，升級成能把檔案做得穩、做得久、做得能交接。",
      skills: ["資料驗證", "文字日期清理", "圖表設計", "Table / Structured References"],
    },
    {
      eyebrow: "Stage 4",
      title: "可重跑的資料流程",
      summary: "把分析從單次操作，推進到能重新整理、能處理多來源資料的流程。",
      skills: ["動態陣列", "進階函數", "Power Query", "Data Model"],
    },
    {
      eyebrow: "Stage 5",
      title: "自動化與系統化交付",
      summary: "最後才是巨集與專案化，把前面能力組成真正能落地的工作系統。",
      skills: ["VBA 基礎", "錯誤處理", "專案化自動化", "綜合挑戰"],
    },
  ],
};
const HOME_LEARNING_MODES = [
  {
    key: "all",
    label: "全部",
    phases: [],
    summary: "顯示完整 5 個階段，從基礎操作一路看到自動化。",
    mindset: "最適合你，如果你想先看完整路線，再決定優先補哪一段。",
    anchorPhase: 1,
    roadmapTitle: "從新手走到專業的能力地圖",
    roadmapIntro: "完整模式會保留整張學習地圖，讓你先看懂能力升級順序，再決定從哪段開始切入。",
    featuredSkills: ["完整路線", "先看全圖", "再選起點"],
    featuredLessons: [
      { slug: "P1-02", badge: "先學這課", reason: "先把函數基本功練成反射動作，後面的查找與分析才會順。" },
      { slug: "P2-02", badge: "高頻核心", reason: "跨表查找是職場最常遇到的 Excel 任務之一。" },
      { slug: "P4-04", badge: "專業分水嶺", reason: "開始理解可重跑流程，你的 Excel 會從手工操作升級成系統化工作流。" },
    ],
  },
  {
    key: "starter",
    label: "新手入門",
    phases: [1, 2],
    summary: "聚焦 Phase 1-2，先把公式、查找與報表核心打穩。",
    mindset: "最適合你，如果你還不熟 Excel 基本操作，或想先補職場最常用的核心能力。",
    anchorPhase: 1,
    roadmapTitle: "先建立公式與報表核心",
    roadmapIntro: "這個模式把重心放在 Phase 1-2：先把輸入節奏、基礎函數、條件邏輯與查找報表能力打穩，再往更進階的資料流程前進。",
    featuredSkills: ["快捷鍵節奏", "基礎函數", "XLOOKUP / SUMIFS"],
    featuredLessons: [
      { slug: "P1-01", badge: "先暖身", reason: "先把 Mac 鍵位節奏建立起來，後面學公式時會更順手。" },
      { slug: "P1-02", badge: "主推薦", reason: "這是公式基本功的起點，後面幾乎每課都會用到。" },
      { slug: "P2-02", badge: "下一個里程碑", reason: "查找比對是從會寫公式走向會做職場報表的關鍵一步。" },
    ],
  },
  {
    key: "work",
    label: "職場即用",
    phases: [2, 3],
    summary: "聚焦 Phase 2-3，適合想先補職場報表與資料整理實力。",
    mindset: "最適合你，如果你已經會基本函數，現在想更快補齊工作上最常用的報表能力。",
    anchorPhase: 2,
    roadmapTitle: "先補職場報表與資料整理實力",
    roadmapIntro: "這個模式把重心放在 Phase 2-3：先學會查找、條件統計與樞紐，再把資料結構、清理與可維護性補齊。",
    featuredSkills: ["SUMIFS", "XLOOKUP", "PivotTable", "Table"],
    featuredLessons: [
      { slug: "P2-01", badge: "先學這課", reason: "條件統計是報表工作最常出現的核心技術。" },
      { slug: "P2-03", badge: "快速見效", reason: "樞紐分析能讓你很快做出真正能交付的報表。" },
      { slug: "P3-04", badge: "補維護性", reason: "學會 Table 和結構化參照，檔案才真的能交接、能長期維護。" },
    ],
  },
  {
    key: "auto",
    label: "自動化進階",
    phases: [4, 5],
    summary: "聚焦 Phase 4-5，先看 Power Query、資料模型與 VBA。",
    mindset: "最適合你，如果你已經會基本報表，現在想把重複流程變成可重跑、可交付的系統。",
    anchorPhase: 4,
    roadmapTitle: "從重複操作走向可重跑流程與自動化",
    roadmapIntro: "這個模式把重心放在 Phase 4-5：先理解動態陣列與 Power Query（Mac 主路線），再了解資料模型概念（Mac 上以觀念為主），最後把 VBA 接進完整自動化流程。",
    featuredSkills: ["動態陣列", "Power Query", "Data Model", "VBA"],
    featuredLessons: [
      { slug: "P4-01", badge: "先補新工具", reason: "先理解新世代公式能力，後面做資料流程會更順。" },
      { slug: "P4-04", badge: "主推薦", reason: "Power Query 是把手工清資料升級成可重跑流程的關鍵。" },
      { slug: "P5-01", badge: "自動化入口", reason: "當流程已經夠穩，再把 VBA 接進來，會更像真正的工作系統。" },
    ],
  },
];
const FOCUS_MODE_KEY = "lesson-focus-mode";
if (document.body?.dataset.lessonSlug) {
  try {
    document.body.dataset.focusMode = localStorage.getItem(FOCUS_MODE_KEY) === "1" ? "on" : "off";
  } catch (e) {
    document.body.dataset.focusMode = "off";
  }
}

// Shared HTML escape — used by AI chat and search
function escHtml(s){ return (s||"").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c])); }
function usesReadingLessonFlow(){
  const variant = document.body?.dataset.lessonVariant || "";
  return variant !== "command" && variant !== "studio";
}
function isLessonPage(){
  return !!document.body?.dataset.lessonSlug;
}
function isFocusModeOn(){
  return document.body?.dataset.focusMode === "on";
}
function getLessonFocusState(){
  const boxes = [...document.querySelectorAll(".checklist input[type=checkbox]")];
  const done = boxes.filter(function(box){ return box.checked; }).length;
  const firstOpen = boxes.find(function(box){ return !box.checked; }) || null;
  const firstText = firstOpen?.closest("label")?.querySelector(".txt")?.textContent?.trim() || "";
  const practiceBtn = document.querySelector('.lt-tab[data-tab="practice"]');
  const activeTab = document.querySelector('.lt-tab.is-active')?.getAttribute("data-tab") || "";
  const nextLink = document.querySelector(".lesson-nav .btn.primary[href]");
  return {
    total: boxes.length,
    done: done,
    firstOpen: firstOpen,
    firstText: firstText,
    practiceBtn: practiceBtn,
    activeTab: activeTab,
    nextLink: nextLink,
  };
}
function syncFocusModeUI(){
  const on = isFocusModeOn();
  const navBtn = document.getElementById("focusModeBtn");
  if (navBtn) {
    navBtn.classList.toggle("is-active", on);
    navBtn.setAttribute("aria-pressed", on ? "true" : "false");
    const icon = navBtn.querySelector(".tool-btn-icon");
    const label = navBtn.querySelector(".tool-btn-label");
    if (icon) icon.textContent = on ? "🎯" : "🧠";
    if (label) label.textContent = on ? "專注中" : "專注";
    navBtn.title = on ? "結束專注模式" : "開啟專注模式";
    navBtn.setAttribute("aria-label", navBtn.title);
  }
  const stripToggle = document.querySelector(".lesson-focus-toggle");
  if (stripToggle) {
    stripToggle.textContent = on ? "結束專注模式" : "開啟專注模式";
    stripToggle.classList.toggle("is-active", on);
    stripToggle.setAttribute("aria-pressed", on ? "true" : "false");
  }
  const state = document.querySelector(".lesson-focus-state strong");
  if (state) state.textContent = on ? "已開啟" : "未開啟";
}
function setFocusMode(on){
  if (!isLessonPage()) return;
  document.body.dataset.focusMode = on ? "on" : "off";
  try {
    localStorage.setItem(FOCUS_MODE_KEY, on ? "1" : "0");
  } catch (e) {}
  if (on) {
    document.getElementById("aiPanel")?.classList.remove("open");
    document.getElementById("aiFab")?.classList.remove("hidden");
    document.getElementById("ttsPanel")?.classList.remove("open");
    document.getElementById("cheatDrawer")?.classList.remove("open");
    document.getElementById("cheatBackdrop")?.classList.remove("show");
  }
  syncFocusModeUI();
  updateLessonFocusStrip();
}
function updateLessonFocusStrip(){
  const strip = document.querySelector(".lesson-focus-strip");
  if (!strip) return;
  const state = getLessonFocusState();
  const progress = strip.querySelector(".lesson-focus-progress");
  const step = strip.querySelector(".lesson-focus-step");
  const task = strip.querySelector("[data-focus-task]");
  const caption = strip.querySelector("[data-focus-caption]");
  const primary = strip.querySelector(".lesson-focus-primary");

  if (progress) {
    progress.textContent = state.total
      ? `${state.done} / ${state.total} 已完成`
      : "先讀重點，再做一小題";
  }
  if (step) step.textContent = state.firstOpen ? String(state.done + 1).padStart(2, "0") : "✓";

  let action = "study";
  let taskText = state.firstText || "先把這頁的核心案例看完";
  let captionText = "一次只做一小步，不用現在就把整課想完。";
  let primaryLabel = "去下一步";

  if (state.practiceBtn && state.activeTab !== "practice") {
    action = "practice";
    taskText = state.firstText || "先切去練習區做第一題";
    captionText = "先做第一題就好，不需要一次把整個練習區清完。";
    primaryLabel = "切到練習";
  } else if (state.firstOpen) {
    action = "checklist";
    taskText = state.firstText;
    captionText = "看完這段後，回到任務清單把這一項打勾就好。";
    primaryLabel = "前往任務";
  } else if (state.nextLink) {
    action = "next";
    taskText = "這課已完成，可以往下一課前進";
    captionText = state.nextLink.textContent.trim();
    primaryLabel = "前往下一課";
  }

  if (task) task.textContent = taskText;
  if (caption) caption.textContent = captionText;
  if (primary) {
    primary.textContent = primaryLabel;
    primary.dataset.action = action;
  }
  syncFocusModeUI();
}
function runLessonFocusPrimaryAction(){
  const state = getLessonFocusState();
  const primary = document.querySelector(".lesson-focus-primary");
  const action = primary?.dataset.action || "study";
  if (action === "practice" && state.practiceBtn) {
    state.practiceBtn.click();
    setTimeout(function(){
      document.querySelector('.lt-panel.is-active')?.scrollIntoView({ behavior: "smooth", block: "start" });
    }, 90);
    return;
  }
  if (action === "checklist") {
    const checklist = document.querySelector(".checklist");
    checklist?.scrollIntoView({ behavior: "smooth", block: "start" });
    state.firstOpen?.focus({ preventScroll: true });
    return;
  }
  if (action === "next" && state.nextLink) {
    location.href = state.nextLink.href;
    return;
  }
  document.querySelector(".lesson .md-body")?.scrollIntoView({ behavior: "smooth", block: "start" });
}

// Toolbar labels for lesson pages
(function(){
  if (!document.querySelector(".lesson")) return;
  var labels = {
    printBtn: "列印",
    markWeakBtn: "複習",
    searchBtn: "搜尋",
    themeBtn: document.documentElement.dataset.theme === "dark" ? "亮色" : "深色",
  };
  Object.keys(labels).forEach(function(id){
    var btn = document.getElementById(id);
    if (!btn) return;
    btn.classList.add("tool-btn");
    var icon = (btn.textContent || "").trim();
    btn.innerHTML = '<span class="tool-btn-icon">' + escHtml(icon) + '</span><span class="tool-btn-label">' + labels[id] + '</span>';
    btn.setAttribute("aria-label", labels[id]);
    btn.title = labels[id];
  });
})();

// Theme
const themeBtn = document.getElementById("themeBtn");
const saved = localStorage.getItem("theme");
if (saved) document.documentElement.dataset.theme = saved;
function syncTheme(){
  if (!themeBtn) return;
  const isDark = document.documentElement.dataset.theme === "dark";
  const icon = isDark ? "☀️" : "🌙";
  const label = isDark ? "亮色" : "深色";
  const iconEl = themeBtn.querySelector(".tool-btn-icon");
  const labelEl = themeBtn.querySelector(".tool-btn-label");
  if (iconEl && labelEl) {
    iconEl.textContent = icon;
    labelEl.textContent = label;
  } else {
    themeBtn.textContent = icon;
  }
  themeBtn.setAttribute("aria-label", "切換到" + label + "模式");
  themeBtn.title = "切換到" + label + "模式";
}
themeBtn?.addEventListener("click", () => {
  const cur = document.documentElement.dataset.theme === "dark" ? "" : "dark";
  document.documentElement.dataset.theme = cur;
  localStorage.setItem("theme", cur);
  syncTheme();
});
syncTheme();

// Focus mode toggle
(function(){
  if (!document.querySelector(".lesson")) return;
  const tools = document.querySelector(".nav .tools");
  if (!tools || document.getElementById("focusModeBtn")) return;
  const btn = document.createElement("button");
  btn.type = "button";
  btn.id = "focusModeBtn";
  btn.className = "btn tool-btn";
  btn.innerHTML = '<span class="tool-btn-icon">🧠</span><span class="tool-btn-label">專注</span>';
  btn.addEventListener("click", function(){
    setFocusMode(!isFocusModeOn());
  });
  tools.insertBefore(btn, document.getElementById("searchBtn") || tools.firstChild);
  syncFocusModeUI();
})();

// 記錄上次學習時間
function touchPage(){ localStorage.setItem("seen:" + PK, Date.now().toString()); }
if (document.querySelector(".lesson")) touchPage();
try {
  if (document.querySelector(".lesson")) sessionStorage.setItem("last-lesson", PK);
} catch(e){}

// Checklist
function initChecklist(){
  document.querySelectorAll(".checklist input[type=checkbox], .core-list input[type=checkbox]").forEach(cb => {
    const key = "check:" + PK + ":" + cb.dataset.key;
    if (localStorage.getItem(key) === "1") cb.checked = true;
    cb.addEventListener("change", () => {
      localStorage.setItem(key, cb.checked ? "1" : "0");
      updateLessonProgress();
      markPageDone();
      updateLessonFocusStrip();
    });
  });
  updateLessonProgress();
  markPageDone();
  updateLessonFocusStrip();
}
function updateLessonProgress(){
  const boxes = document.querySelectorAll(".checklist input[type=checkbox]");
  if (!boxes.length) return;
  const done = [...boxes].filter(b => b.checked).length;
  const bar = document.querySelector(".lesson .progress > span");
  const lbl = document.querySelector(".lesson .progress-label");
  if (bar) bar.style.width = (done/boxes.length*100) + "%";
  if (lbl) lbl.textContent = `本課進度 ${done} / ${boxes.length}`;
}
function markPageDone(){
  const boxes = document.querySelectorAll(".checklist input[type=checkbox]");
  if (!boxes.length) return;
  const allDone = [...boxes].every(b => b.checked);
  const key = "done:" + PK;
  if (allDone) {
    if (!localStorage.getItem(key)) {
      localStorage.setItem(key, new Date().toISOString());
      const today = new Date().toDateString();
      const tk = "today:" + today;
      localStorage.setItem(tk, String((parseInt(localStorage.getItem(tk)||"0"))+1));
    }
  } else localStorage.removeItem(key);
}
initChecklist();

// Lesson kickoff
(function(){
  const lesson = document.querySelector(".lesson");
  const checklist = lesson?.querySelector(".checklist");
  if (!lesson || !checklist || lesson.querySelector(".lesson-kickoff")) return;
  const isReadingLayout = usesReadingLessonFlow();
  const slug = document.body?.dataset.lessonSlug || "";
  const workbookSheet = LESSON_WORKBOOK_MAP[slug] || "";
  const workbookNote = workbookSheet
    ? `這課有對應的練習簿工作表：${workbookSheet}`
    : "這課目前沒有對應的練習簿工作表，先跟著本頁案例與任務清單做就好。";

  const tasks = [...checklist.querySelectorAll("label .txt")]
    .map(el => (el.textContent || "").trim())
    .filter(Boolean)
    .slice(0, 3);
  if (!tasks.length) return;

  const box = document.createElement("section");
  box.className = "lesson-kickoff";
  box.innerHTML = `
    <div class="lesson-kickoff-head">
      <div>
        <div class="lesson-kickoff-eyebrow">開始前先看這裡</div>
        <h2 class="lesson-kickoff-title">先完成這 ${tasks.length} 件事</h2>
      </div>
      <div class="lesson-kickoff-actions">
        <button type="button" class="btn primary lesson-kickoff-btn">去做練習</button>
        ${workbookSheet ? `<a class="btn lesson-kickoff-btn secondary" href="${DEPTH}practice.xlsx" download>打開練習簿</a>` : ""}
      </div>
    </div>
    <div class="lesson-kickoff-guidance">建議節奏：先看 TL;DR 與這課案例，再做第一題，最後回來把任務打勾。</div>
    <div class="lesson-kickoff-list">
      ${tasks.map((task, i) => `
        <div class="lesson-kickoff-item">
          <span class="lesson-kickoff-num">0${i + 1}</span>
          <span class="lesson-kickoff-text">${escHtml(task)}</span>
        </div>
      `).join("")}
    </div>
    <div class="lesson-kickoff-workbook">
      <strong>練習方式</strong>
      <span>${escHtml(workbookNote)}</span>
    </div>
  `;

  box.querySelector(".lesson-kickoff-btn")?.addEventListener("click", () => {
    const practiceTab = document.querySelector('.lt-tab[data-tab="practice"]');
    if (practiceTab) {
      practiceTab.click();
      setTimeout(() => {
        const activePanel = document.querySelector('.lt-panel.is-active');
        activePanel?.scrollIntoView({ behavior: "smooth", block: "start" });
      }, 80);
      return;
    }
    checklist.scrollIntoView({ behavior: "smooth", block: "start" });
  });

  const anchor = isReadingLayout
    ? lesson.querySelector(".md-body")
    : (lesson.querySelector(".tldr") || lesson.querySelector(".progress-label"));
  if (anchor) anchor.insertAdjacentElement("afterend", box);
})();

// Lesson badge rail
(function(){
  const lesson = document.querySelector(".lesson");
  if (!lesson || lesson.querySelector(".lesson-badge-rail")) return;
  const slug = document.body?.dataset.lessonSlug || "";
  const badges = LESSON_BADGES[slug];
  if (!badges || !badges.length) return;
  const orderedBadges = badges.slice().sort(function(a, b){
    const ai = LESSON_BADGE_KIND_ORDER.indexOf(a.kind || "");
    const bi = LESSON_BADGE_KIND_ORDER.indexOf(b.kind || "");
    return (ai === -1 ? 999 : ai) - (bi === -1 ? 999 : bi);
  });

  const rail = document.createElement("section");
  rail.className = "lesson-badge-rail";
  rail.innerHTML = orderedBadges.map(function(badge){
    return `
      <div class="lesson-badge lesson-badge-${escHtml(badge.kind || "note")}">
        <span class="lesson-badge-label">${escHtml(badge.label)}</span>
        <span class="lesson-badge-note">${escHtml(badge.note)}</span>
      </div>
    `;
  }).join("");

  const anchor = lesson.querySelector(".tldr") || lesson.querySelector(".progress-label");
  if (anchor) anchor.insertAdjacentElement("afterend", rail);
})();

// Lesson focus strip
(function(){
  const lesson = document.querySelector(".lesson");
  if (!lesson || lesson.querySelector(".lesson-focus-strip")) return;

  const strip = document.createElement("section");
  strip.className = "lesson-focus-strip";
  strip.innerHTML = `
    <div class="lesson-focus-head">
      <div>
        <div class="lesson-focus-eyebrow">Focus Friendly</div>
        <h2 class="lesson-focus-title">現在先做 1 件事</h2>
      </div>
      <div class="lesson-focus-state">專注模式 <strong>未開啟</strong></div>
    </div>
    <div class="lesson-focus-next">
      <span class="lesson-focus-step">01</span>
      <div class="lesson-focus-copy">
        <strong data-focus-task>先把這頁的核心案例看完</strong>
        <span data-focus-caption>一次只做一小步，不用現在就把整課想完。</span>
      </div>
      <div class="lesson-focus-progress">0 / 0 已完成</div>
    </div>
    <div class="lesson-focus-actions">
      <button type="button" class="btn primary lesson-focus-primary" data-action="study">去下一步</button>
      <button type="button" class="btn lesson-focus-toggle" aria-pressed="false">開啟專注模式</button>
    </div>
  `;
  strip.querySelector(".lesson-focus-primary")?.addEventListener("click", runLessonFocusPrimaryAction);
  strip.querySelector(".lesson-focus-toggle")?.addEventListener("click", function(){
    setFocusMode(!isFocusModeOn());
  });

  const anchor = lesson.querySelector(".lesson-badge-rail")
    || lesson.querySelector(".tldr")
    || lesson.querySelector(".progress-label");
  if (anchor) anchor.insertAdjacentElement("afterend", strip);
  else lesson.insertAdjacentElement("afterbegin", strip);

  document.addEventListener("click", function(ev){
    if (ev.target.closest(".lt-tab")) {
      setTimeout(updateLessonFocusStrip, 80);
    }
  });
  document.addEventListener("xlsx-integrator:ready", function(){
    setTimeout(updateLessonFocusStrip, 360);
  }, { once: true });
  setTimeout(updateLessonFocusStrip, 500);
  syncFocusModeUI();
})();

// Lesson compass
(function(){
  const lesson = document.querySelector(".lesson");
  if (!lesson || lesson.querySelector(".lesson-compass")) return;
  const isReadingLayout = usesReadingLessonFlow();

  const slug = document.body?.dataset.lessonSlug || "";
  if (!slug || !LESSON_GUIDE[slug]) return;

  const currentIndex = LESSON_SEQUENCE.indexOf(slug);
  if (currentIndex === -1) return;

  const current = idx.find(e => e.u.endsWith(slug + ".html"));
  const prevSlug = LESSON_SEQUENCE[currentIndex - 1];
  const nextSlug = LESSON_SEQUENCE[currentIndex + 1];
  const prev = prevSlug ? idx.find(e => e.u.endsWith(prevSlug + ".html")) : null;
  const next = nextSlug ? idx.find(e => e.u.endsWith(nextSlug + ".html")) : null;
  const phaseLabel = current?.b || "";
  const phaseMatch = phaseLabel.match(/Phase\s+\d+/);
  const phaseName = phaseMatch ? phaseMatch[0] : "";
  const guide = LESSON_GUIDE[slug];
  const box = document.createElement("section");
  box.className = "lesson-compass";

  const previousHtml = prev ? `
    <a class="lesson-compass-link" href="${DEPTH + prev.u}">
      <span class="lesson-compass-kicker">上一站</span>
      <strong>${escHtml(prev.t)}</strong>
      <span>${escHtml(prev.b || "")}</span>
    </a>
  ` : `
    <div class="lesson-compass-link is-static">
      <span class="lesson-compass-kicker">起點</span>
      <strong>這是整張學習地圖的第一課</strong>
      <span>先建立操作手感，再一路往公式、分析與自動化推進。</span>
    </div>
  `;

  const nextHtml = next ? `
    <a class="lesson-compass-link" href="${DEPTH + next.u}">
      <span class="lesson-compass-kicker">下一站</span>
      <strong>${escHtml(next.t)}</strong>
      <span>${escHtml(guide.nextWhy)}</span>
    </a>
  ` : `
    <div class="lesson-compass-link is-static">
      <span class="lesson-compass-kicker">下一步</span>
      <strong>回頭補強弱點課，開始做自己的模板</strong>
      <span>${escHtml(guide.nextWhy)}</span>
    </div>
  `;

  box.innerHTML = `
    <div class="lesson-compass-head">
      <div>
        <div class="lesson-compass-eyebrow">學習指南針</div>
        <h2 class="lesson-compass-title">${escHtml(phaseName ? phaseName + " · " : "")}你現在在這裡</h2>
      </div>
      <div class="lesson-compass-pill">${currentIndex + 1} / ${LESSON_SEQUENCE.length}</div>
    </div>
    <div class="lesson-compass-summary">
      <div class="lesson-compass-block">
        <span class="lesson-compass-label">這課在做什麼</span>
        <p>${escHtml(guide.focus)}</p>
      </div>
      <div class="lesson-compass-block">
        <span class="lesson-compass-label">學完會解鎖</span>
        <p>${escHtml(guide.unlock)}</p>
      </div>
    </div>
    <div class="lesson-compass-rail">
      ${previousHtml}
      ${nextHtml}
    </div>
  `;

  const anchor = isReadingLayout
    ? (lesson.querySelector(".lesson-kickoff") || lesson.querySelector(".md-body"))
    : (lesson.querySelector(".lesson-kickoff") || lesson.querySelector(".tldr"));
  if (anchor) anchor.insertAdjacentElement("afterend", box);
})();

// Professional note
(function(){
  const lesson = document.querySelector(".lesson");
  if (!lesson || lesson.querySelector(".lesson-pro-note")) return;
  const isReadingLayout = usesReadingLessonFlow();
  const slug = document.body?.dataset.lessonSlug || "";
  const note = LESSON_PRO_NOTES[slug];
  if (!note) return;

  const box = document.createElement("section");
  box.className = "lesson-pro-note";
  box.innerHTML = `
    <div class="lesson-pro-note-head">
      <div class="lesson-pro-note-eyebrow">${escHtml(note.eyebrow || "Professional Note")}</div>
      <h2 class="lesson-pro-note-title">${escHtml(note.title)}</h2>
    </div>
    <div class="lesson-pro-note-body">
      ${note.items.map(item => `
        <div class="lesson-pro-note-item">
          <span class="lesson-pro-note-dot"></span>
          <p>${escHtml(item)}</p>
        </div>
      `).join("")}
    </div>
  `;

  const anchor = isReadingLayout
    ? (lesson.querySelector(".lesson-compass") || lesson.querySelector(".lesson-kickoff") || lesson.querySelector(".md-body"))
    : (lesson.querySelector(".lesson-compass") || lesson.querySelector(".lesson-kickoff"));
  if (anchor) anchor.insertAdjacentElement("afterend", box);
})();

// 標記為「不熟」(spaced repetition)
document.getElementById("markWeakBtn")?.addEventListener("click", () => {
  const k = "weak:" + PK;
  if (localStorage.getItem(k)) { localStorage.removeItem(k); alert("已取消標記"); }
  else { localStorage.setItem(k, Date.now().toString()); alert("已標記為「不熟」，3 天後會在首頁提醒複習"); }
});

// Notes
(function(){
  const ta = document.querySelector("textarea[data-note]");
  if (!ta) return;
  const key = "note:" + PK;
  ta.value = localStorage.getItem(key) || "";
  let t;
  ta.addEventListener("input", () => { clearTimeout(t); t = setTimeout(() => localStorage.setItem(key, ta.value), 300); });
})();

// Streak
(function(){
  if (!document.querySelector(".lesson")) return;
  const today = new Date().toDateString();
  const last = localStorage.getItem("streak:lastDay");
  let count = parseInt(localStorage.getItem("streak:count")||"0");
  if (last !== today){
    const yest = new Date(Date.now()-86400000).toDateString();
    if (last === yest) count++; else count = 1;
    localStorage.setItem("streak:lastDay", today);
    localStorage.setItem("streak:count", String(count));
  }
})();

// Pomodoro
let timer=null, remain=25*60, running=false;
const timeEl = document.querySelector(".pomo .time");
const playBtn = document.querySelector(".pomo .play");
const resetBtn = document.querySelector(".pomo .reset");
const fmt = s => String(Math.floor(s/60)).padStart(2,"0") + ":" + String(s%60).padStart(2,"0");
const renderTime = () => { if (timeEl) timeEl.textContent = fmt(remain); };
playBtn?.addEventListener("click", () => {
  running = !running;
  playBtn.textContent = running ? "⏸" : "▶";
  if (running) {
    timer = setInterval(() => {
      remain--;
      if (remain <= 0) { remain=25*60; running=false; playBtn.textContent="▶"; clearInterval(timer); alert("專注完成！休息一下吧 ☕"); }
      renderTime();
    }, 1000);
  } else clearInterval(timer);
});
resetBtn?.addEventListener("click", () => { clearInterval(timer); remain=25*60; running=false; if(playBtn) playBtn.textContent="▶"; renderTime(); });
renderTime();

// TOC
(function(){
  const tocNav = document.getElementById("tocNav");
  if (!tocNav) return;
  const heads = document.querySelectorAll(".lesson .md-body h2, .lesson .md-body h3");
  if (!heads.length) { tocNav.parentElement.style.display = "none"; return; }
  const ul = document.createElement("ul");
  heads.forEach((h,i) => {
    if (!h.id) h.id = "h-" + i;
    const li = document.createElement("li");
    li.className = h.tagName.toLowerCase();
    const a = document.createElement("a");
    a.href = "#" + h.id;
    a.textContent = h.textContent;
    li.appendChild(a);
    ul.appendChild(li);
  });
  tocNav.appendChild(ul);
  const links = tocNav.querySelectorAll("a");
  const obs = new IntersectionObserver(entries => {
    entries.forEach(e => {
      if (e.isIntersecting) {
        links.forEach(l => l.classList.toggle("active", l.getAttribute("href") === "#" + e.target.id));
      }
    });
  }, { rootMargin: "0px 0px -70% 0px" });
  heads.forEach(h => obs.observe(h));
})();

// 列印按鈕
document.getElementById("printBtn")?.addEventListener("click", () => window.print());

// Progress / cards
const doneSet = new Set();
idx.forEach(e => { const k = "done:" + e.u.split("/").slice(-2).join("/"); if (localStorage.getItem(k)) doneSet.add(e.u); });

(function progress(){
  if (!idx.length) return;
  document.querySelectorAll('a[data-lesson]').forEach(link => {
    const u = link.dataset.lesson;
    const card = link.querySelector(".card");
    if (!u || !card) return;
    const k = "done:" + u.split("/").slice(-2).join("/");
    if (localStorage.getItem(k)) card.classList.add("done");
    // 上次學習時間
    const seen = localStorage.getItem("seen:" + u.split("/").slice(-2).join("/"));
    if (seen) {
      const days = Math.floor((Date.now() - parseInt(seen)) / 86400000);
      const ago = days === 0 ? "今天剛看過" : days === 1 ? "昨天看過" : `${days} 天前看過`;
      const lbl = card.querySelector(".last-seen");
      if (lbl) lbl.textContent = "🕐 " + ago;
    }
  });
  const overall = document.getElementById("overallProgress");
  if (overall) {
    const total = idx.length;
    const done = doneSet.size;
    const pct = total ? Math.round(done/total*100) : 0;
    overall.textContent = `整體進度 ${done} / ${total}（${pct}%）`;
    const bar = document.querySelector(".hero .progress > span");
    if (bar) bar.style.width = pct + "%";
  }
})();

// 倒數計時
(function(){
  const el = document.getElementById("countdown");
  if (!el || !el.dataset.target) return;
  const target = new Date(el.dataset.target);
  const ms = target - new Date();
  const days = Math.ceil(ms / 86400000);
  if (days < 0) { el.textContent = "🎉 考試已過"; return; }
  el.innerHTML = `📅 距離考試還有 <strong>${days}</strong> 天`;
  if (days <= 14) el.classList.add("urgent");
})();

// Streak / today / 推薦 / 匯出
(function(){
  const streakNum = document.getElementById("streakNum");
  if (streakNum) {
    const today = new Date().toDateString();
    const last = localStorage.getItem("streak:lastDay");
    let count = parseInt(localStorage.getItem("streak:count")||"0");
    if (last && last !== today) {
      const yest = new Date(Date.now()-86400000).toDateString();
      if (last !== yest) count = 0;
    }
    streakNum.textContent = count;
  }
  const td = document.getElementById("todayDone");
  if (td) {
    const today = new Date().toDateString();
    td.textContent = localStorage.getItem("today:" + today) || "0";
  }

  const helper = document.createElement("div");
  helper.className = "hero-helper";
  const helperTarget = document.querySelector(".hero .utility-actions") || document.querySelector(".hero .hero-actions:last-of-type");
  if (helperTarget) helperTarget.insertAdjacentElement("afterend", helper);

  const quickNav = document.createElement("div");
  quickNav.className = "hero-quicknav";
  if (helperTarget) helperTarget.insertAdjacentElement("afterend", quickNav);

  function updateHeroHelper(){
    if (!helper) return;
    const weak = idx.filter(e => {
      const k = "weak:" + e.u.split("/").slice(-2).join("/");
      const t = localStorage.getItem(k);
      if (!t) return false;
      const days = (Date.now() - parseInt(t)) / 86400000;
      return days >= 3;
    });
    let pick = null;
    let prefix = "";
    let cta = "開始這課";
    if (weak.length) {
      pick = weak[0];
      prefix = "建議先複習";
      cta = "開始複習";
    } else {
      const undone = idx.filter(e => !doneSet.has(e.u));
      if (undone.length) {
        pick = undone[0];
        prefix = "建議下一課";
      }
    }
    if (!pick) {
      helper.innerHTML = `
        <div class="hero-helper-copy">
          <span class="hero-helper-label">今天先做這一件</span>
          <strong>先做一題快速測驗或回顧筆記</strong>
          <span>目前進度很完整，現在最值得做的是保持熟悉感，而不是再塞新內容。</span>
        </div>
        <button type="button" class="btn primary hero-helper-btn">開始測驗</button>
      `;
      helper.querySelector(".hero-helper-btn")?.addEventListener("click", function(){
        document.getElementById("quizBox")?.scrollIntoView({ behavior: "smooth", block: "start" });
      });
      return;
    }
    helper.innerHTML = `
      <div class="hero-helper-copy">
        <span class="hero-helper-label">${escHtml(prefix)}</span>
        <strong>${escHtml(pick.t)}</strong>
        <span>${escHtml(pick.b || "")}</span>
      </div>
      <button type="button" class="btn primary hero-helper-btn">${escHtml(cta)}</button>
    `;
    helper.querySelector(".hero-helper-btn")?.addEventListener("click", function(){
      location.href = DEPTH + pick.u;
    });
  }
  updateHeroHelper();

  (function buildQuickNav(){
    if (!quickNav) return;
    const lastLesson = sessionStorage.getItem("last-lesson");
    const groups = HOME_LEARNING_MODES;
    const state = { active: "all" };
    let hasHydratedMode = false;

    const wrap = document.createElement("div");
    wrap.className = "hero-mode-switch";
    const title = document.createElement("div");
    title.className = "hero-mode-title";
    title.textContent = "選一條現在最適合你的路線";
    wrap.appendChild(title);

    const chips = document.createElement("div");
    chips.className = "hero-mode-grid";
    wrap.appendChild(chips);

    const meta = document.createElement("div");
    meta.className = "hero-mode-meta";
    wrap.appendChild(meta);

    const continueWrap = document.createElement("div");
    continueWrap.className = "hero-continue";
    const featureWrap = document.createElement("section");
    featureWrap.className = "hero-featured-lessons";
    const roadmapWrap = document.createElement("details");
    roadmapWrap.className = "hero-capability-map";
    roadmapWrap.open = false;
    roadmapWrap.innerHTML = `
      <summary class="hero-capability-summary">
        <div class="hero-capability-head">
          <div>
            <div class="hero-capability-eyebrow">Professional Path</div>
            <h2 class="hero-capability-title"></h2>
          </div>
          <p class="hero-capability-intro"></p>
        </div>
        <span class="hero-capability-toggle">展開完整路線</span>
      </summary>
      <div class="hero-capability-body">
        <div class="hero-capability-defaults">
          ${HOME_CAPABILITY_MAP.defaults.map(function(item){
            return `<div class="hero-capability-default"><span>◆</span><span>${escHtml(item)}</span></div>`;
          }).join("")}
        </div>
        <div class="hero-capability-timeline">
          ${HOME_CAPABILITY_MAP.stages.map(function(stage, index){
            return `
              <div class="hero-capability-node" data-phase="${index + 1}">
                <div class="hero-capability-node-num">${String(index + 1).padStart(2, "0")}</div>
                <div class="hero-capability-node-copy">
                  <span>${escHtml(stage.eyebrow)}</span>
                  <strong>${escHtml(stage.title)}</strong>
                </div>
              </div>
            `;
          }).join("")}
        </div>
        <div class="hero-capability-grid">
          ${HOME_CAPABILITY_MAP.stages.map(function(stage, index){
            return `
              <article class="hero-capability-card" data-phase="${index + 1}">
                <div class="hero-capability-card-top">
                  <div class="hero-capability-card-eyebrow">${escHtml(stage.eyebrow)}</div>
                  <div class="hero-capability-card-num">${String(index + 1).padStart(2, "0")}</div>
                </div>
                <h3>${escHtml(stage.title)}</h3>
                <p>${escHtml(stage.summary)}</p>
                <div class="hero-capability-skills">
                  ${stage.skills.map(function(skill){ return `<span>${escHtml(skill)}</span>`; }).join("")}
                </div>
              </article>
            `;
          }).join("")}
        </div>
      </div>
    `;
    roadmapWrap.addEventListener("toggle", function(){
      const toggle = roadmapWrap.querySelector(".hero-capability-toggle");
      if (toggle) toggle.textContent = roadmapWrap.open ? "收起完整路線" : "展開完整路線";
    });

    function applyMode(modeKey){
      state.active = modeKey;
      chips.querySelectorAll("button").forEach(function(btn){
        btn.classList.toggle("is-active", btn.dataset.mode === modeKey);
      });
      const selected = groups.find(g => g.key === modeKey) || groups[0];
      const allowed = new Set(selected.phases);
      let visibleLessons = 0;
      let firstVisibleSection = null;
      document.querySelectorAll(".phase-section").forEach(function(section){
        const phaseMatch = section.className.match(/phase-(\d)/);
        const phaseNo = phaseMatch ? parseInt(phaseMatch[1], 10) : 0;
        const visible = !allowed.size || allowed.has(phaseNo);
        section.style.display = visible ? "" : "none";
        if (visible && !firstVisibleSection) firstVisibleSection = section;
        var next = section.nextElementSibling;
        while (next && !next.classList.contains("phase-section") && !next.classList.contains("phase-group")) {
          next.style.display = visible ? "" : "none";
          if (visible && next.matches('a[data-lesson]')) visibleLessons++;
          next = next.nextElementSibling;
        }
        const group = section.nextElementSibling;
        if (group && group.classList.contains("phase-group")) {
          group.style.display = visible ? "" : "none";
        }
      });
      const phaseText = selected.phases.length ? `Phase ${selected.phases.join("、")}` : "全部 5 個 Phase";
      meta.innerHTML = `
        <strong>${escHtml(selected.label)}</strong>
        <span>${escHtml(selected.summary)}</span>
        <em>${escHtml(selected.mindset)} · 目前顯示 ${phaseText} · ${visibleLessons} 堂課</em>
      `;
      roadmapWrap.dataset.mode = selected.key;
      roadmapWrap.querySelector(".hero-capability-title").textContent = selected.roadmapTitle;
      roadmapWrap.querySelector(".hero-capability-intro").textContent = selected.roadmapIntro;
      roadmapWrap.querySelectorAll("[data-phase]").forEach(function(node){
        var phaseNo = parseInt(node.getAttribute("data-phase"), 10);
        var isActive = !allowed.size || allowed.has(phaseNo);
        node.classList.toggle("is-mode-active", isActive);
        node.classList.toggle("is-mode-muted", !isActive);
      });
      featureWrap.innerHTML = `
        <div class="hero-featured-head">
          <div>
            <div class="hero-featured-eyebrow">Start Here</div>
            <h3 class="hero-featured-title">${escHtml(selected.label)} 先學哪 3 課</h3>
          </div>
          <div class="hero-featured-skills">
            ${selected.featuredSkills.map(function(skill){ return `<span>${escHtml(skill)}</span>`; }).join("")}
          </div>
        </div>
        <div class="hero-featured-grid">
          ${selected.featuredLessons.map(function(item, index){
            const lesson = idx.find(function(entry){ return entry.u.endsWith(item.slug + ".html"); });
            if (!lesson) return "";
            return `
              <a class="hero-featured-card ${index === 0 ? "is-primary" : ""}" href="${DEPTH + lesson.u}">
                <span class="hero-featured-badge">${escHtml(item.badge)}</span>
                <strong>${escHtml(lesson.t)}</strong>
                <p>${escHtml(item.reason)}</p>
                <em>${escHtml(lesson.b || "")}</em>
              </a>
            `;
          }).join("")}
        </div>
      `;
      featureWrap.dataset.mode = selected.key;
      if (firstVisibleSection) {
        if (hasHydratedMode) {
          var y = window.pageYOffset + firstVisibleSection.getBoundingClientRect().top - 120;
          window.scrollTo({ top: Math.max(0, y), behavior: "smooth" });
        }
      }
      hasHydratedMode = true;
    }

    groups.forEach(function(group){
      const starter = group.featuredLessons[0];
      const starterLesson = starter ? idx.find(function(entry){ return entry.u.endsWith(starter.slug + ".html"); }) : null;
      const phaseLabel = group.phases.length ? `Phase ${group.phases.join(" · ")}` : "Phase 1-5";
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "hero-mode-card";
      btn.dataset.mode = group.key;
      btn.innerHTML = `
        <span class="hero-mode-card-head">
          <span class="hero-mode-card-title">${escHtml(group.label)}</span>
          <span class="hero-mode-card-phase">${escHtml(phaseLabel)}</span>
        </span>
        <span class="hero-mode-card-summary">${escHtml(group.summary)}</span>
        <span class="hero-mode-card-start">${escHtml(starterLesson ? "從 " + starterLesson.t + " 開始" : "先看完整路線")}</span>
      `;
      btn.addEventListener("click", function(){ applyMode(group.key); });
      chips.appendChild(btn);
    });

    if (lastLesson) {
      const match = idx.find(e => e.u.endsWith(lastLesson));
      if (match) {
        continueWrap.innerHTML = `
          <button type="button" class="btn hero-continue-btn">
            <span>▶</span>
            <span>繼續上次看到：${escHtml(match.t)}</span>
          </button>
        `;
        continueWrap.querySelector("button")?.addEventListener("click", function(){
          location.href = DEPTH + match.u;
        });
      }
    }

    quickNav.appendChild(wrap);
    if (continueWrap.children.length) quickNav.appendChild(continueWrap);
    applyMode("all");

    (function composeHero(){
      const hero = document.querySelector(".hero");
      if (!hero || hero.querySelector(".hero-shell")) return;
      const titleNode = hero.querySelector("h1");
      const copyNode = hero.querySelector("p");
      const streakNode = hero.querySelector(".streak-row");
      const progressNode = hero.querySelector(".progress");
      const progressLabelNode = hero.querySelector(".progress-label");
      const primaryActionsNode = hero.querySelector(".primary-actions");
      const utilityActionsNode = hero.querySelector(".utility-actions");
      const helperNode = hero.querySelector(".hero-helper");
      const quickNavNode = hero.querySelector(".hero-quicknav");
      if (!titleNode || !copyNode || !quickNavNode) return;

      const shell = document.createElement("div");
      shell.className = "hero-shell";
      const mainCol = document.createElement("div");
      mainCol.className = "hero-main";
      const sideCol = document.createElement("div");
      sideCol.className = "hero-side";
      const copyGroup = document.createElement("div");
      copyGroup.className = "hero-copy-group";
      const eyebrow = document.createElement("div");
      eyebrow.className = "hero-eyebrow";
      eyebrow.textContent = "Modern Academy · 從 Excel 新手一路走到專業";
      copyGroup.appendChild(eyebrow);
      copyGroup.appendChild(titleNode);
      copyGroup.appendChild(copyNode);

      const progressStack = document.createElement("div");
      progressStack.className = "hero-progress-stack";
      [streakNode, progressNode, progressLabelNode].forEach(function(node){
        if (node) progressStack.appendChild(node);
      });

      const actionStack = document.createElement("div");
      actionStack.className = "hero-action-stack";
      [primaryActionsNode, utilityActionsNode].forEach(function(node){
        if (node) actionStack.appendChild(node);
      });

      mainCol.appendChild(copyGroup);
      if (progressStack.children.length) mainCol.appendChild(progressStack);
      if (actionStack.children.length) mainCol.appendChild(actionStack);

      if (helperNode) sideCol.appendChild(helperNode);
      sideCol.appendChild(quickNavNode);
      shell.appendChild(mainCol);
      shell.appendChild(sideCol);
      hero.appendChild(shell);

      const academyStack = document.createElement("section");
      academyStack.className = "home-academy-stack";
      academyStack.appendChild(featureWrap);
      academyStack.appendChild(roadmapWrap);
      requestAnimationFrame(function(){
        const dashboard = document.querySelector(".idx-dashboard");
        if (dashboard) dashboard.insertAdjacentElement("afterend", academyStack);
        else hero.insertAdjacentElement("afterend", academyStack);
      });
    })();
  })();

  // 今天學一課（優先選 weak、再選未完成）
  document.getElementById("todayBtn")?.addEventListener("click", () => {
    const weak = idx.filter(e => {
      const k = "weak:" + e.u.split("/").slice(-2).join("/");
      const t = localStorage.getItem(k);
      if (!t) return false;
      const days = (Date.now() - parseInt(t)) / 86400000;
      return days >= 3;
    });
    let pick;
    if (weak.length) {
      pick = weak[Math.floor(Math.random()*weak.length)];
      if (!confirm(`📌 該複習了：\n\n${pick.t}\n${pick.b||""}\n\n（你 3 天前標記為「不熟」）\n\n要開始嗎？`)) return;
    } else {
      const undone = idx.filter(e => !doneSet.has(e.u));
      if (!undone.length) { alert("太強了，全部學完了 🎉"); return; }
      pick = undone[Math.floor(Math.random()*undone.length)];
      if (!confirm(`今天推薦你學：\n\n📘 ${pick.t}\n${pick.b||""}\n\n要開始嗎？`)) return;
    }
    location.href = DEPTH + pick.u;
  });

  // 匯出
  document.getElementById("exportBtn")?.addEventListener("click", () => {
    const lines = ["# 我的學習筆記", "", "匯出時間：" + new Date().toLocaleString("zh-TW"), ""];
    lines.push(`## 📊 學習統計`);
    lines.push(`- 已完成：**${doneSet.size} / ${idx.length}** 課`);
    lines.push(`- 連續學習：**${localStorage.getItem("streak:count")||0}** 天`);
    lines.push("");
    let any = false;
    idx.forEach(e => {
      const pk = e.u.split("/").slice(-2).join("/");
      const done = doneSet.has(e.u);
      const note = localStorage.getItem("note:" + pk) || "";
      if (!done && !note) return;
      any = true;
      lines.push(`### ${done ? "✅" : "📝"} ${e.t}`);
      if (e.b) lines.push(`*${e.b}*`);
      lines.push("");
      if (e.s) { lines.push("> " + e.s); lines.push(""); }
      if (note) { lines.push("**我的筆記：**"); lines.push(""); lines.push(note); lines.push(""); }
    });
    if (!any) lines.push("_還沒有完成的課程或筆記_");
    const blob = new Blob([lines.join("\n")], { type:"text/markdown;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `學習筆記_${new Date().toISOString().slice(0,10)}.md`;
    a.click();
  });

  // 匯出 Anki
  document.getElementById("ankiBtn")?.addEventListener("click", () => {
    const cards = window.ANKI_CARDS || [];
    if (!cards.length) { alert("這個網站沒有 Anki 卡片資料"); return; }
    const tsv = cards.map(c => `${c.q.replace(/\t/g," ")}\t${c.a.replace(/\t/g," ")}`).join("\n");
    const blob = new Blob([tsv], { type:"text/tab-separated-values;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `anki_cards_${new Date().toISOString().slice(0,10)}.tsv`;
    a.click();
  });
})();

// Quiz 模式
(function(){
  const quizBox = document.getElementById("quizBox");
  if (!quizBox) return;
  const cards = QUIZ_CARDS;
  if (!cards.length) { quizBox.innerHTML = "<p>沒有測驗資料</p>"; return; }
  let i = 0;
  function render() {
    const c = cards[i];
    const opts = [...c.options];
    quizBox.innerHTML = `
      <div class="quiz-card">
        <div style="color:var(--muted);font-size:13px">第 ${i+1} / ${cards.length} 題</div>
        <div class="quiz-q">${c.q}</div>
        <div class="quiz-options">
          ${opts.map((o,j) => `<div class="quiz-option" data-i="${j}">${o}</div>`).join("")}
        </div>
        <div class="quiz-feedback" style="display:none"></div>
        <div style="margin-top:14px;display:flex;gap:10px">
          <button class="btn" id="quizPrev">← 上一題</button>
          <button class="btn primary" id="quizNext">下一題 →</button>
        </div>
      </div>`;
    quizBox.querySelectorAll(".quiz-option").forEach(el => {
      el.addEventListener("click", () => {
        const idx = parseInt(el.dataset.i);
        const fb = quizBox.querySelector(".quiz-feedback");
        quizBox.querySelectorAll(".quiz-option").forEach((x,j) => {
          if (j === c.answer) x.classList.add("correct");
          else if (j === idx) x.classList.add("wrong");
          x.style.pointerEvents = "none";
        });
        fb.style.display = "block";
        fb.innerHTML = `${idx === c.answer ? "✅ 答對了！" : "❌ 答錯了"}<br>${c.explain || ""}`;
      });
    });
    quizBox.querySelector("#quizPrev").onclick = () => { if (i>0) { i--; render(); } };
    quizBox.querySelector("#quizNext").onclick = () => { if (i<cards.length-1) { i++; render(); } else alert("做完啦 🎉"); };
  }
  render();
})();

// ========= AI 助教浮動聊天 =========
(function(){
  if (!document.getElementById("aiFab")) return;
  const __apiMeta = document.querySelector('meta[name="api-base"]');
  const CHAT_API_BASE = (__apiMeta && __apiMeta.content) ? __apiMeta.content.replace(/\/$/,'') : '';
  const fab = document.getElementById("aiFab");
  const panel = document.getElementById("aiPanel");
  const closeBtn = document.getElementById("aiClose");
  const log = document.getElementById("aiLog");
  const input = document.getElementById("aiInput");
  const sendBtn = document.getElementById("aiSend");
  const copyBtn = document.getElementById("aiCopy");

  // 收集當前頁面內容當 context
  const lessonTitle = document.querySelector(".lesson h1")?.textContent || document.title;
  const breadcrumb = document.querySelector(".lesson > main > div")?.textContent?.replace("← 回到目錄 ·","").trim() || "";
  const bodyText = document.querySelector(".md-body")?.innerText?.slice(0, 2500) || "";
  const SITE_NAME = document.querySelector(".brand")?.textContent?.trim() || "學習網站";
  const sysPrompt = `你是一位友善、簡潔的學習教練，使用繁體中文回答。學生正在閱讀「${SITE_NAME}」中的單元：「${lessonTitle}」（${breadcrumb}）。\n\n本課內容摘要：\n${bodyText}\n\n回答原則：\n- 用最白話的方式解釋\n- 優先用條列、表格或範例\n- 如果學生問題和本課無關，也可以回答\n- 保持簡短，重點優先`;

  let messages = [];
  function open(){ panel.classList.add("open"); fab.classList.add("hidden"); setTimeout(()=>input.focus(),200); }
  function close(){ panel.classList.remove("open"); fab.classList.remove("hidden"); }
  fab.addEventListener("click", open);
  closeBtn.addEventListener("click", close);

  function add(role, text){
    const div = document.createElement("div");
    div.className = "ai-msg ai-" + role;
    div.innerHTML = role === "user" ? esc(text) : renderMd(text);
    log.appendChild(div);
    log.scrollTop = log.scrollHeight;
  }
  const esc = escHtml;
  function renderMd(s){
    s = esc(s);
    s = s.replace(/```([\s\S]*?)```/g, (_,c)=>`<pre><code>${c}</code></pre>`);
    s = s.replace(/`([^`]+)`/g, "<code>$1</code>");
    s = s.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
    s = s.replace(/^### (.+)$/gm, "<h4>$1</h4>");
    s = s.replace(/^## (.+)$/gm, "<h3>$1</h3>");
    s = s.replace(/^(\|.+\|\n)((?:\|[-: ]+)+\|\n)((?:\|.+\|\n?)+)/gm, (_, hdr, _sep, body) => {
      const parse = r => r.trim().replace(/^\||\|$/g, '').split('|').map(c => c.trim());
      const ths = parse(hdr).map(h => `<th>${h}</th>`).join('');
      const trs = body.trim().split('\n').map(r => `<tr>${parse(r).map(c=>`<td>${c}</td>`).join('')}</tr>`).join('');
      return `<table><thead><tr>${ths}</tr></thead><tbody>${trs}</tbody></table>`;
    });
    s = s.replace(/^- (.+)$/gm, "<li>$1</li>");
    s = s.replace(/(<li>.*<\/li>)/s, "<ul>$1</ul>");
    s = s.replace(/\n\n/g, "<br><br>");
    return s;
  }

  async function send(){
    const text = input.value.trim();
    if (!text) return;
    add("user", text);
    messages.push({role:"user", content:text});
    input.value = "";
    sendBtn.disabled = true;
    const thinking = document.createElement("div");
    thinking.className = "ai-msg ai-assistant ai-thinking";
    thinking.textContent = "🤔 思考中…";
    log.appendChild(thinking);
    log.scrollTop = log.scrollHeight;
    const ctrl = new AbortController();
    const timeout = setTimeout(() => ctrl.abort(), 30_000);
    try {
      const resp = await fetch(CHAT_API_BASE + '/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        signal: ctrl.signal,
        body: JSON.stringify({
          model: 'claude-haiku-4-5',
          max_tokens: 1024,
          system: sysPrompt,
          messages: messages
        })
      });
      const data = await resp.json();
      thinking.remove();
      if (data.content && data.content[0]) {
        const reply = data.content[0].text;
        messages.push({role:"assistant", content:reply});
        add("assistant", reply);
      } else {
        add("assistant", "❌ " + (data.error?.message || data.error || "API 回應異常，可改用「複製問題」貼到 Claude/ChatGPT 網頁版"));
      }
    } catch(err) {
      thinking.remove();
      const msg = err.name === 'AbortError' ? '請求逾時（30 秒）' : err.message;
      add("assistant", "🌐 連線失敗：" + msg + "\n\n你可以按「📋 複製到剪貼簿」貼到 Claude/ChatGPT 網頁版繼續問。");
    } finally {
      clearTimeout(timeout);
    }
    sendBtn.disabled = false;
    input.focus();
  }
  sendBtn.addEventListener("click", send);
  input.addEventListener("keydown", e => { if (e.key==="Enter" && !e.shiftKey && !e.isComposing) { e.preventDefault(); send(); } });

  copyBtn.addEventListener("click", () => {
    const q = input.value.trim() || "請幫我解釋這個單元的重點";
    const full = `我正在學「${lessonTitle}」（${breadcrumb}）。\n\n本課內容：\n${bodyText}\n\n我的問題：${q}`;
    navigator.clipboard.writeText(full).then(() => {
      copyBtn.textContent = "✅ 已複製！貼到 Claude/ChatGPT";
      setTimeout(()=>copyBtn.textContent = "📋 複製問題到剪貼簿", 2500);
    });
  });

  // 預設打招呼
  add("assistant", `嗨！我是這課的 AI 學習教練 👋\n\n你正在學「**${lessonTitle}**」。卡住或想更深入的話，直接問我吧～`);
})();

// Search
(function(){
  const btn = document.getElementById("searchBtn");
  const modal = document.getElementById("searchModal");
  const input = document.getElementById("searchInput");
  const results = document.getElementById("searchResults");
  if (!btn || !modal) return;
  let cursor = 0, current = [];
  function open(){ modal.classList.add("open"); input.value=""; render(""); setTimeout(()=>input.focus(),50); }
  function close(){ modal.classList.remove("open"); }
  function render(q){
    q = q.trim().toLowerCase();
    current = !q ? idx.slice(0, 30) : idx.filter(e => (e.t+" "+(e.b||"")+" "+(e.s||"")).toLowerCase().includes(q)).slice(0, 50);
    cursor = 0;
    if (!current.length) { results.innerHTML = '<div class="search-empty">沒有結果</div>'; return; }
    results.innerHTML = current.map((e,i) =>
      `<a class="search-item${i===0?' active':''}" href="${DEPTH}${e.u}">
         <div class="si-title">${esc(e.t)}</div>
         <div class="si-meta">${esc(e.b||"")}</div>
       </a>`).join("");
  }
  const esc = escHtml;
  btn.addEventListener("click", open);
  modal.addEventListener("click", e => { if (e.target === modal) close(); });
  input?.addEventListener("input", e => render(e.target.value));
  document.addEventListener("keydown", e => {
    if ((e.metaKey || e.ctrlKey) && e.key === "k") { e.preventDefault(); open(); return; }
    if (!modal.classList.contains("open")) return;
    if (e.key === "Escape") close();
    if (e.key === "ArrowDown") { e.preventDefault(); cursor=Math.min(cursor+1,current.length-1); upd(); }
    if (e.key === "ArrowUp")   { e.preventDefault(); cursor=Math.max(cursor-1,0); upd(); }
    if (e.key === "Enter") { const a = results.querySelectorAll(".search-item")[cursor]; if (a) location.href = a.href; }
  });
  function upd(){
    results.querySelectorAll(".search-item").forEach((el,i) => el.classList.toggle("active", i===cursor));
    const el = results.querySelectorAll(".search-item")[cursor];
    if (el) el.scrollIntoView({ block:"nearest" });
  }
})();

/* ───────── 🔊 TTS 朗讀（Web Speech API + Azure 自然語音） ───────── */
window.__TTS = (function(){
  if (!('speechSynthesis' in window)) return {};
  const synth = window.speechSynthesis;
  const LS = { rate:'tts.rate', voice:'tts.voice', mode:'tts.mode', azVoice:'tts.azVoice' };
  let voices = [], queue = [], qIdx = 0;
  let azVoices = [], audio = null, azAvailable = false;
  let epoch = 0;

  const __ttsMeta = document.querySelector('meta[name="api-base"]');
  const API_BASE = (__ttsMeta && __ttsMeta.content) ? __ttsMeta.content.replace(/\/$/,'')
                 : (location.protocol === 'file:' ? 'http://localhost:5173' : '');
  function detectBackend(){
    return fetch(API_BASE + '/api/health').then(r=>r.json()).then(d => {
      azAvailable = !!(d && d.azure);
      refreshVoiceList();
      if (azAvailable){
        fetch(API_BASE + '/api/voices').then(r=>r.json()).then(list => {
          azVoices = (list||[]).filter(v => /^zh-(TW|CN|HK)/.test(v.locale));
          refreshVoiceList();
        }).catch(()=>{});
      }
    }).catch(()=>{ azAvailable = false; refreshVoiceList(); });
  }
  detectBackend();

  function loadVoices(){
    voices = synth.getVoices().filter(v => /zh|cmn|yue/i.test(v.lang));
    if (!voices.length) voices = synth.getVoices();
  }
  loadVoices();
  speechSynthesis.onvoiceschanged = loadVoices;

  function pickVoice(){
    const saved = localStorage.getItem(LS.voice);
    if (saved){ const v = voices.find(x => x.name === saved); if (v) return v; }
    return voices.find(v => /zh-TW|zh_TW|zh-Hant|Taiwan/i.test(v.lang+v.name))
        || voices.find(v => /zh/i.test(v.lang)) || voices[0];
  }
  function getRate(){ return parseFloat(localStorage.getItem(LS.rate) || '1.05'); }
  function getMode(){ return localStorage.getItem(LS.mode) || (azAvailable?'azure':'browser'); }
  function getAzVoice(){ return localStorage.getItem(LS.azVoice) || 'zh-TW-HsiaoChenNeural'; }

  function stopAll(){
    epoch++;
    synth.cancel();
    if (audio){ audio.onended = audio.onerror = null; audio.pause(); audio.src=''; audio=null; }
  }
  function speakChunks(chunks, startIdx=0){
    stopAll(); queue = chunks; qIdx = startIdx; nextChunk();
  }
  async function nextChunk(){
    if (qIdx >= queue.length){ setStatus('idle'); return; }
    const myEpoch = epoch;
    const text = queue[qIdx];
    if (getMode() === 'azure' && azAvailable){
      try {
        setStatus('playing');
        const ratePct = Math.round((getRate()-1)*100);
        const r = (ratePct>=0?'+':'')+ratePct+'%';
        const resp = await fetch(API_BASE + '/api/tts', {
          method:'POST', headers:{'Content-Type':'application/json'},
          body: JSON.stringify({ text, voice:getAzVoice(), rate:r })
        });
        if (epoch !== myEpoch) return;
        if (!resp.ok){ throw new Error('tts '+resp.status); }
        const blob = await resp.blob();
        if (epoch !== myEpoch) return;
        audio = new Audio(URL.createObjectURL(blob));
        audio.onended = () => { if (epoch === myEpoch){ qIdx++; nextChunk(); } };
        audio.onerror = () => { if (epoch === myEpoch){ qIdx++; nextChunk(); } };
        audio.play();
      } catch(e){
        if (epoch !== myEpoch) return;
        console.warn('Azure TTS failed, fallback to browser', e);
        localStorage.setItem(LS.mode,'browser'); browserSpeak(text);
      }
      return;
    }
    browserSpeak(text);
  }
  function browserSpeak(text){
    const myEpoch = epoch;
    const u = new SpeechSynthesisUtterance(text);
    const v = pickVoice(); if (v) u.voice = v;
    u.lang = (v && v.lang) || 'zh-TW';
    u.rate = getRate(); u.pitch = 1;
    u.onend = () => { if (epoch === myEpoch){ qIdx++; nextChunk(); } };
    u.onerror = () => { if (epoch === myEpoch){ qIdx++; nextChunk(); } };
    setStatus('playing');
    synth.speak(u);
  }
  function splitText(text){
    return text.replace(/[\p{Emoji_Presentation}\p{Extended_Pictographic}]/gu, '').replace(/\s+/g,' ').split(/(?<=[。！？!?；;])/).map(s=>s.trim()).filter(s=>s.length);
  }
  function getSectionText(sec){
    const clone = sec.cloneNode(true);
    clone.querySelectorAll('button,.tts-btn,pre').forEach(n => n.remove());
    return splitText(clone.innerText || clone.textContent || '');
  }

  function injectButtons(){
    document.querySelectorAll('.lesson .md-body h2').forEach(h2 => {
      if (h2.querySelector('.tts-btn')) return;
      const b = document.createElement('button');
      b.className = 'tts-btn'; b.title = '朗讀本段'; b.textContent = '🔊';
      b.onclick = (e) => {
        e.stopPropagation();
        const section = h2.parentElement;
        speakChunks(getSectionText(section));
        panel.classList.add('open');
      };
      h2.appendChild(b);
    });
  }

  const panel = document.createElement('div');
  panel.id = 'ttsPanel';
  panel.innerHTML = `
    <button class="tts-ico" data-act="prev" title="上一句">⏮</button>
    <button class="tts-ico" data-act="toggle" title="點擊暫停/繼續，雙擊停止">⏸</button>
    <button class="tts-ico" data-act="next" title="下一句">⏭</button>
    <label class="tts-rate">速度<select data-act="rate">
      <option value="0.85">0.85x</option><option value="1">1x</option>
      <option value="1.05">1.05x</option><option value="1.2">1.2x</option>
      <option value="1.4">1.4x</option><option value="1.6">1.6x</option>
    </select></label>
    <select class="tts-voice" data-act="voice"></select>
    <label class="tts-rate"><input type="checkbox" data-act="mode"> Azure 自然</label>
    <label class="tts-rate"><input type="checkbox" data-act="autoai"> AI 朗讀</label>
    <button class="tts-ico" data-act="page" title="朗讀整頁">📖</button>
    <button class="tts-ico" data-act="close" title="收起">✕</button>
  `;
  document.body.appendChild(panel);

  function refreshVoiceList(){
    const sel = panel.querySelector('[data-act="voice"]');
    if (!sel) return;
    sel.innerHTML = '';
    const mode = getMode();
    if (mode === 'azure' && azVoices.length){
      azVoices.forEach(v => {
        const o = document.createElement('option');
        o.value = v.name;
        o.textContent = `${v.display||v.name} · ${v.locale} ${v.gender==='Female'?'♀':'♂'}`;
        sel.appendChild(o);
      });
      sel.value = getAzVoice();
    } else {
      voices.forEach(v => {
        const o = document.createElement('option');
        o.value = v.name; o.textContent = `${v.name} (${v.lang})`;
        sel.appendChild(o);
      });
      const saved = localStorage.getItem(LS.voice);
      if (saved && voices.find(v => v.name === saved)) sel.value = saved;
    }
    panel.querySelector('[data-act="rate"]').value = String(getRate());
    const m = panel.querySelector('[data-act="mode"]');
    if (m){
      m.checked = (mode === 'azure' && azAvailable);
      m.disabled = false;
      m.title = azAvailable ? '切換 Azure 自然語音' : '未偵測到後端，請確認 http://localhost:5173 已啟動';
    }
    const a = panel.querySelector('[data-act="autoai"]');
    if (a) a.checked = localStorage.getItem('tts.autoai') === '1';
  }
  setTimeout(refreshVoiceList, 300);
  speechSynthesis.addEventListener && speechSynthesis.addEventListener('voiceschanged', refreshVoiceList);

  function setStatus(s){
    const t = panel.querySelector('[data-act="toggle"]');
    if (t) t.textContent = (s === 'playing') ? '⏸' : '▶';
  }

  function readWholePage(){
    const all = [];
    const body = document.querySelector('.lesson .md-body') || document.querySelector('.md-body') || document.querySelector('main');
    if (body) all.push(...splitText(body.innerText || ''));
    if (all.length) speakChunks(all);
  }
  let _toggleTimer = null;
  panel.addEventListener('click', (e) => {
    const el = e.target.closest('[data-act]');
    if (!el) return;
    const act = el.dataset.act;
    if (el.tagName === 'INPUT' || el.tagName === 'SELECT') return;
    if (act === 'toggle'){
      if (_toggleTimer){ clearTimeout(_toggleTimer); _toggleTimer = null; stopAll(); setStatus('idle'); return; }
      _toggleTimer = setTimeout(() => {
        _toggleTimer = null;
        if (audio && !audio.paused){ audio.pause(); setStatus('paused'); return; }
        if (audio && audio.paused && audio.src){ audio.play(); setStatus('playing'); return; }
        if (synth.paused){ synth.resume(); setStatus('playing'); return; }
        if (synth.speaking){ synth.pause(); setStatus('paused'); return; }
        if (queue.length && qIdx < queue.length){ nextChunk(); return; }
        readWholePage();
      }, 250);
      return;
    }
    if (act === 'next'){ stopAll(); qIdx = Math.min(queue.length, qIdx+1); nextChunk(); }
    else if (act === 'prev'){ stopAll(); qIdx = Math.max(0, qIdx-1); nextChunk(); }
    else if (act === 'page'){ readWholePage(); }
    else if (act === 'close'){ stopAll(); panel.classList.remove('open'); }
  });
  function restartFromCurrent(){
    const playing = (audio && !audio.paused) || synth.speaking || synth.paused;
    if (!playing || !queue.length) return;
    stopAll(); nextChunk();
  }
  panel.addEventListener('change', (e) => {
    const act = e.target.dataset.act;
    if (act === 'rate'){
      localStorage.setItem(LS.rate, e.target.value);
      restartFromCurrent();
    } else if (act === 'voice'){
      if (getMode()==='azure') localStorage.setItem(LS.azVoice, e.target.value);
      else localStorage.setItem(LS.voice, e.target.value);
      restartFromCurrent();
    } else if (act === 'mode'){
      if (e.target.checked && !azAvailable){
        detectBackend().then(()=>{
          if (!azAvailable){
            alert('無法連到 Azure 後端\n\n請確認後端已啟動並在 index.html 設定 meta[name="api-base"]');
            e.target.checked = false;
            localStorage.setItem(LS.mode, 'browser');
          } else {
            localStorage.setItem(LS.mode, 'azure');
            refreshVoiceList();
            restartFromCurrent();
          }
        });
        return;
      }
      localStorage.setItem(LS.mode, e.target.checked ? 'azure' : 'browser');
      refreshVoiceList();
      restartFromCurrent();
    } else if (act === 'autoai'){
      localStorage.setItem('tts.autoai', e.target.checked ? '1' : '0');
    }
  });

  const fab = document.createElement('button');
  fab.id = 'ttsFab'; fab.title = '朗讀工具'; fab.setAttribute('aria-label', '打開朗讀工具'); fab.textContent = '🔊';
  fab.onclick = () => panel.classList.toggle('open');
  document.body.appendChild(fab);

  window.addEventListener('beforeunload', () => synth.cancel());

  if (document.readyState === 'loading')
    document.addEventListener('DOMContentLoaded', injectButtons);
  else injectButtons();

  return { speak: speakChunks, splitText, stop: stopAll };
})();
