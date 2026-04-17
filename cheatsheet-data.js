/* Excel 速查共用資料層
   被 cheatsheet.html 與 lesson 頁的速查抽屜 (drawer.js) 共用。
   修改這份資料 = 同步更新速查手冊和所有 lesson 頁的抽屜。
*/
// ── Formula Data ──
const formulas = [
  // Math/Stats (Lv.1)
  {name:"SUM",cat:"math",lv:1,syntax:'=SUM(範圍)',desc:"加總範圍內所有數字",example:'=SUM(A1:A10)  →  將 A1 到 A10 的數字全部加起來',detail:"最基本的函式。可以多範圍：=SUM(A1:A10, C1:C10)。也支援直接輸入數字：=SUM(1,2,3)",tip:"SUM 忽略文字和空白，只加數字",tags:["加總","合計","基礎"]},
  {name:"AVERAGE",cat:"math",lv:1,syntax:'=AVERAGE(範圍)',desc:"計算平均值",example:'=AVERAGE(B2:B10)  →  所有數字的平均',detail:"等於 SUM(範圍)/COUNT(範圍)。忽略空白儲存格。",tip:"如果資料有 0 但不想算入，用 AVERAGEIF 排除",tags:["平均","統計"]},
  {name:"COUNT",cat:"math",lv:1,syntax:'=COUNT(範圍)',desc:"計算範圍內有多少個「數字」",example:'=COUNT(A1:A10)  →  計算有幾格是數字',detail:"COUNT 只數數字。COUNTA 數所有非空白（含文字）。COUNTBLANK 數空白格。",tip:"面試常考：COUNT vs COUNTA 的差別",tags:["計數","統計"]},
  {name:"MAX / MIN",cat:"math",lv:1,syntax:'=MAX(範圍)  /  =MIN(範圍)',desc:"找出最大值 / 最小值",example:'=MAX(C2:C100)  →  找出 C 欄最大的數字',detail:"LARGE(範圍,k) 找第 k 大的值。SMALL(範圍,k) 找第 k 小的值。",tip:"想找前3名？用 LARGE(範圍,{1,2,3})",tags:["最大","最小","排名"]},
  {name:"MEDIAN",cat:"math",lv:2,syntax:'=MEDIAN(範圍)',desc:"中位數（排序後的中間值）",example:'=MEDIAN(A1:A10)  →  不受極端值影響的中間值',detail:"適合薪資、房價等有極端值的資料。比 AVERAGE 更能反映真實水準。",tip:"面試題常考：什麼時候用 MEDIAN 比 AVERAGE 好？",tags:["中位數","統計","分析"]},

  // Conditions (Lv.1~2)
  {name:"IF",cat:"condition",lv:1,syntax:'=IF(條件, 真的值, 假的值)',desc:"根據條件回傳不同結果",example:'=IF(A1>=60, "及格", "不及格")',detail:'可以巢狀：=IF(A1>=90,"A",IF(A1>=80,"B","C"))。但超過 3 層建議改用 IFS。',tip:"IF 是所有條件函式的基礎，務必熟練",tags:["條件","判斷","基礎"]},
  {name:"IFS",cat:"condition",lv:2,syntax:'=IFS(條件1,值1, 條件2,值2, ..., TRUE,預設值)',desc:"多條件判斷，比巢狀 IF 乾淨",example:'=IFS(A1>=90,"A", A1>=80,"B", TRUE,"C")',detail:"依序判斷，第一個成立就回傳。最後用 TRUE 當「其他」。僅 Excel 2019+ / 365。",tip:"TRUE 放最後等於 else，一定要加！",tags:["多條件","判斷"]},
  {name:"AND / OR",cat:"condition",lv:1,syntax:'=AND(條件1, 條件2)  /  =OR(條件1, 條件2)',desc:"多條件組合：全部成立 / 任一成立",example:'=IF(AND(A1>=60, B1>=60), "雙科及格", "未通過")',detail:"AND = 全部都要成立。OR = 任一成立即可。通常搭配 IF 使用。",tip:"NOT() 可以反轉條件：NOT(A1>60) 等於 A1<=60",tags:["條件","邏輯","組合"]},
  {name:"IFERROR",cat:"condition",lv:2,syntax:'=IFERROR(公式, 錯誤時的值)',desc:"當公式出錯時顯示替代值",example:'=IFERROR(VLOOKUP(...), "查無資料")',detail:"包住任何可能出錯的公式。常見錯誤：#N/A, #REF!, #DIV/0!, #VALUE!",tip:"專業做法：每個 VLOOKUP 外面都包 IFERROR",tags:["錯誤處理","防呆"]},
  {name:"SWITCH",cat:"condition",lv:3,syntax:'=SWITCH(值, 比對1,結果1, 比對2,結果2, 預設)',desc:"根據值選擇對應結果",example:'=SWITCH(A1, 1,"一月", 2,"二月", 3,"三月", "未知")',detail:"比 IF/IFS 更適合「值對應」的情境。類似程式語言的 switch/case。",tip:"最後一個參數（沒有配對的）就是預設值",tags:["對應","選擇"]},

  // Conditional Stats (Lv.2)
  {name:"COUNTIF",cat:"condition",lv:2,syntax:'=COUNTIF(範圍, 條件)',desc:"計算符合條件的儲存格個數",example:'=COUNTIF(B2:B100, "業務")  →  計算幾個是「業務」',detail:'條件可以是文字、數字、或含運算子的字串如 ">=60"。',tip:"COUNTIFS 可以多條件，注意 S 結尾",tags:["計數","條件","統計"]},
  {name:"COUNTIFS",cat:"condition",lv:2,syntax:'=COUNTIFS(範圍1,條件1, 範圍2,條件2, ...)',desc:"多條件計數",example:'=COUNTIFS(B:B,"業務", C:C,">=60")',detail:"所有條件之間是 AND 關係（全部都要成立）。範圍大小必須一致。",tip:"要做 OR 計數？用多個 COUNTIFS 相加",tags:["多條件","計數","統計"]},
  {name:"SUMIF",cat:"math",lv:2,syntax:'=SUMIF(條件範圍, 條件, 加總範圍)',desc:"條件加總",example:'=SUMIF(B:B, "電子", E:E)  →  電子類的金額合計',detail:"注意參數順序：條件範圍→條件→加總範圍。SUMIFS 順序不同！",tip:"SUMIF 和 SUMIFS 的參數順序不同，容易搞混",tags:["條件加總","統計"]},
  {name:"SUMIFS",cat:"math",lv:2,syntax:'=SUMIFS(加總範圍, 條件範圍1,條件1, ...)',desc:"多條件加總",example:'=SUMIFS(E:E, B:B,"電子", F:F,"已付")',detail:"加總範圍放「第一個」！這跟 SUMIF 不同。後面才接條件對。",tip:"記住：SUMIFS 的 S = 多條件，加總欄移到最前面",tags:["多條件","加總","統計"]},
  {name:"AVERAGEIFS",cat:"math",lv:2,syntax:'=AVERAGEIFS(平均範圍, 條件範圍1,條件1, ...)',desc:"多條件平均",example:'=AVERAGEIFS(E:E, B:B,"北區", C:C,"電子")',detail:"語法同 SUMIFS，加總範圍換成要計算平均的範圍。",tip:"整個 xIFS 家族語法一致：SUMIFS、COUNTIFS、AVERAGEIFS",tags:["條件平均","統計"]},

  // Lookup (Lv.2~3)
  {name:"VLOOKUP",cat:"lookup",lv:2,syntax:'=VLOOKUP(查找值, 表格範圍, 欄號, 0)',desc:"垂直查找：在最左欄找值，回傳同列指定欄",example:'=VLOOKUP("P003", A:D, 2, 0)  →  查產品名稱',detail:"第 4 參數一定填 0（精確查找）。只能往右查。查找欄必須在最左。",tip:"VLOOKUP 限制多，新版 Excel 建議改用 XLOOKUP",tags:["查找","比對","必學"]},
  {name:"INDEX + MATCH",cat:"lookup",lv:3,syntax:'=INDEX(回傳範圍, MATCH(查找值, 查找範圍, 0))',desc:"更靈活的查找組合，可往左查",example:'=INDEX(A:A, MATCH("螢幕", B:B, 0))  →  已知名稱找代碼',detail:"MATCH 找出位置，INDEX 根據位置取值。不受欄位方向限制。",tip:"這是面試必考組合，熟練後效率遠超 VLOOKUP",tags:["查找","靈活","面試"]},
  {name:"XLOOKUP",cat:"lookup",lv:2,syntax:'=XLOOKUP(查找值, 查找範圍, 回傳範圍)',desc:"最新的查找函式（365 專用），取代 VLOOKUP",example:'=XLOOKUP("P003", A:A, B:B)  →  直接查找',detail:"支援左右查找、多值回傳、找不到時的預設值。語法最簡潔。",tip:"如果用 365 版，XLOOKUP 是首選",tags:["查找","新版","推薦"]},
  {name:"MATCH",cat:"lookup",lv:2,syntax:'=MATCH(查找值, 範圍, 0)',desc:"回傳查找值在範圍中的「相對位置」",example:'=MATCH("張志偉", A:A, 0)  →  回傳 3（第 3 個）',detail:"第 3 參數：0=精確、1=小於等於（需升序）、-1=大於等於（需降序）。",tip:"MATCH 回傳的是相對位置，不是列號",tags:["位置","查找"]},

  // Text (Lv.2~3)
  {name:"TRIM",cat:"text",lv:1,syntax:'=TRIM(文字)',desc:"去除多餘空格（前後+連續空格變單一空格）",example:'=TRIM("  Hello   World  ")  →  "Hello World"',detail:"從外部系統貼進來的資料常有隱藏空格，TRIM 是清洗第一步。搭配 CLEAN 去除不可見字元。",tip:"資料清洗三連：TRIM(CLEAN(SUBSTITUTE(A1,CHAR(160),\"\")))",tags:["清洗","空格","文字"]},
  {name:"LEFT / MID / RIGHT",cat:"text",lv:2,syntax:'=LEFT(文字,字數)  =MID(文字,起始,字數)  =RIGHT(文字,字數)',desc:"擷取字串的左/中/右部分",example:'=MID("A20260311", 2, 4)  →  "2026"',detail:"LEFT 從左邊取。RIGHT 從右邊取。MID 指定起始位置和長度。常搭配 FIND/LEN 使用。",tip:"MID 的起始位置從 1 開始算",tags:["擷取","字串","文字"]},
  {name:"FIND / SEARCH",cat:"text",lv:2,syntax:'=FIND(要找的字, 文字, [起始位])  ',desc:"找出某個字在文字中的位置",example:'=FIND("@", "alice@co.com")  →  6',detail:"FIND 區分大小寫，SEARCH 不區分。找不到會出 #VALUE! 錯誤。",tip:"搭配 LEFT/MID 使用：=LEFT(A1, FIND(\"@\",A1)-1) 取 email 帳號",tags:["搜尋","位置","文字"]},
  {name:"TEXT",cat:"text",lv:3,syntax:'=TEXT(值, 格式碼)',desc:"將數字/日期格式化為指定文字",example:'=TEXT(1234567, "#,##0")  →  "1,234,567"',detail:'常用格式碼：#,##0（千分位）、0.00（兩位小數）、YYYY-MM-DD（日期）、0000（補零）。',tip:"TEXT 的結果是「文字」不是數字，無法直接拿來計算",tags:["格式化","數字","日期"]},
  {name:"TEXTJOIN",cat:"text",lv:3,syntax:'=TEXTJOIN(分隔符, 忽略空白, 範圍)',desc:"合併多格文字，可自動略過空白",example:'=TEXTJOIN(", ", TRUE, A1:A5)  →  "王,林,張"',detail:"比 CONCATENATE 和 & 更強大。第二個參數 TRUE = 略過空白儲存格。",tip:"取代老舊的 CONCATENATE，365 必備",tags:["合併","文字","串接"]},
  {name:"SUBSTITUTE",cat:"text",lv:2,syntax:'=SUBSTITUTE(文字, 舊字, 新字)',desc:"替換文字中的特定字串",example:'=SUBSTITUTE("2026/03/31", "/", "-")  →  "2026-03-31"',detail:"可以加第 4 參數指定替換第幾個。不加就全部替換。",tip:"SUBSTITUTE 是文字替換，REPLACE 是位置替換",tags:["替換","文字"]},

  // Date (Lv.2~3)
  {name:"DATEDIF",cat:"date",lv:2,syntax:'=DATEDIF(起始日, 結束日, "D"/"M"/"Y")',desc:"計算兩日期之間的差距",example:'=DATEDIF(A1, TODAY(), "Y")  →  年資幾年',detail:'"D"=天數、"M"=月數、"Y"=年數。這是隱藏函式，不會自動完成。',tip:"計算年資的標準公式，面試常考",tags:["日期差","年資","天數"]},
  {name:"WORKDAY",cat:"date",lv:2,syntax:'=WORKDAY(起始日, 工作天數)',desc:"計算 N 個工作天後的日期（跳過週末）",example:'=WORKDAY(TODAY(), 30)  →  30 個工作天後',detail:"WORKDAY.INTL 可以自訂哪幾天是休息日。可加第 3 參數指定假日清單。",tip:"專案管理必備函式",tags:["工作日","排程","日期"]},
  {name:"EOMONTH",cat:"date",lv:2,syntax:'=EOMONTH(日期, 月份偏移)',desc:"取得月底日期",example:'=EOMONTH(TODAY(), 0)  →  本月最後一天',detail:"偏移量 0=當月、1=下個月、-1=上個月。結果是日期序號，需要格式化。",tip:"=EOMONTH(A1,0)+1 = 下個月的第一天",tags:["月底","日期"]},
  {name:"TODAY / NOW",cat:"date",lv:1,syntax:'=TODAY()  /  =NOW()',desc:"取得今天日期 / 現在的日期時間",example:'=TODAY()  →  2026-03-31',detail:"TODAY 只有日期。NOW 含時分秒。每次開檔或計算時自動更新。",tip:"如果不想自動更新，改用 ⌃+; 輸入靜態日期",tags:["今天","現在","日期"]},
  {name:"YEAR / MONTH / DAY",cat:"date",lv:1,syntax:'=YEAR(日期)  =MONTH(日期)  =DAY(日期)',desc:"從日期中取出年/月/日",example:'=YEAR("2026-03-31")  →  2026',detail:"搭配 DATE 函式可以重組日期：=DATE(YEAR(A1), MONTH(A1)+1, 1) = 下個月一號。",tip:"DATE(年,月,日) 是反向操作，把數字組合成日期",tags:["年","月","日","拆解"]},

  // Array/Dynamic (Lv.3)
  {name:"FILTER",cat:"array",lv:3,syntax:'=FILTER(範圍, 條件)',desc:"動態篩選，結果自動展開",example:'=FILTER(A1:D10, B1:B10="業務")',detail:'多條件用 * 相乘（AND）或 + 相加（OR）。可加第 3 參數當沒有結果時的預設值。',tip:"取代輔助欄+進階篩選，365 最實用的新函式之一",tags:["篩選","動態","陣列","365"]},
  {name:"SORT",cat:"array",lv:3,syntax:'=SORT(範圍, 依哪欄, 升降序)',desc:"動態排序",example:'=SORT(A1:D10, 3, -1)  →  依第 3 欄降序',detail:"1=升序、-1=降序。可以巢狀：SORT(FILTER(...), ...) 先篩再排。",tip:"搭配 FILTER 使用效果最好",tags:["排序","動態","365"]},
  {name:"UNIQUE",cat:"array",lv:3,syntax:'=UNIQUE(範圍)',desc:"取出不重複的值",example:'=UNIQUE(B2:B100)  →  所有不重複的部門名稱',detail:"可加第 2 參數 TRUE = 按欄。第 3 參數 TRUE = 只出現一次的值。",tip:"做動態下拉選單的最佳搭檔",tags:["去重","不重複","365"]},
  {name:"SEQUENCE",cat:"array",lv:3,syntax:'=SEQUENCE(列數, 欄數, 起始, 間隔)',desc:"產生序號陣列",example:'=SEQUENCE(10)  →  產生 1~10',detail:"=SEQUENCE(5,3,1,2) 產生 5 列 3 欄，從 1 開始，間隔 2。",tip:"搭配 TEXT 可以快速產生日期序列",tags:["序號","陣列","365"]},

  // Advanced (Lv.4)
  {name:"INDIRECT",cat:"advanced",lv:4,syntax:'=INDIRECT(文字參照)',desc:"把文字轉成儲存格參照",example:'=INDIRECT("A"&B1)  →  動態參照 A 欄某列',detail:"可以動態組合工作表名稱：=INDIRECT(\"'\"&A1&\"'!B2\")。但效能較差。",tip:"強大但效能差，大量資料時避免使用",tags:["動態","參照","進階"]},
  {name:"OFFSET",cat:"advanced",lv:4,syntax:'=OFFSET(基準格, 下移, 右移, [高度], [寬度])',desc:"從基準格偏移指定距離取得範圍",example:'=SUM(OFFSET(A1, 0, 0, 10, 1))  →  A1 起算 10 列的 SUM',detail:"揮發性函式（volatile），每次計算都重算。搭配 COUNTA 可以做自動擴展範圍。",tip:"可以用 INDEX 取代 OFFSET 來提升效能",tags:["偏移","動態","範圍"]},
  {name:"LET",cat:"advanced",lv:4,syntax:'=LET(名稱1,值1, 名稱2,值2, ..., 計算)',desc:"在公式中定義命名變數",example:'=LET(x, AVERAGE(A:A), y, MAX(A:A), y-x)',detail:"讓複雜公式更易讀。變數只在這個公式內有效。可以避免重複計算。",tip:"專業公式必備，讓同事能看懂你的公式",tags:["變數","可讀性","365"]},
  {name:"LAMBDA",cat:"advanced",lv:5,syntax:'=LAMBDA(參數1, 參數2, ..., 計算)',desc:"建立自訂函式（365 專用）",example:'=LAMBDA(f, (f-32)*5/9)(100)  →  華氏轉攝氏',detail:"定義在「名稱管理員」中，就能像內建函式一樣使用。搭配 MAP/REDUCE 更強。",tip:"這是 Excel 最強的新功能，能創造可重用的函式",tags:["自訂函式","365","專家"]},
  {name:"AGGREGATE",cat:"advanced",lv:4,syntax:'=AGGREGATE(函式碼, 選項, 範圍)',desc:"超強統計函式，可忽略錯誤/隱藏列",example:'=AGGREGATE(1, 6, A:A)  →  忽略錯誤值的 AVERAGE',detail:"函式碼：1=AVERAGE, 4=MAX, 5=MIN, 9=SUM, 14=LARGE。選項 6=忽略錯誤值。",tip:"處理有 #N/A 的資料時特別好用",tags:["統計","忽略錯誤","進階"]},

  // ── Keyboard Shortcuts (macOS) ──
  {name:"🏦 投行級全鍵盤訓練法",cat:"shortcut",lv:3,syntax:'目標：80~90% 無滑鼠操作（macOS 可達成）',desc:"華爾街投行會把新人的滑鼠拿走訓練。Mac 上透過 KeyTips + 快捷鍵可達 80~90%",example:'訓練順序：導航鍵(⌘+方向) → 選取鍵(⇧) → 編輯鍵(⌘+T/D/R) → KeyTips(⌥) → Tab 切換對話框',detail:"【投行訓練背景】華爾街投行（Goldman Sachs、Morgan Stanley 等）會在新人訓練期直接拿走滑鼠，強制練習鍵盤操作。原因：鍵盤操作比滑鼠快 2~5 倍。\\n\\n【macOS 可達 80~90%】自從 2024 年 Mac Excel 新增 KeyTips（按 ⌥ 啟動 Ribbon 快捷鍵）後，macOS 與 Windows 的鍵盤操作差距大幅縮小。\\n\\n【仍需滑鼠的 10~20%】樞紐分析表拖拉欄位到四大區域、圖表細部微調（資料標籤位置/趨勢線）、條件式格式化的自訂規則對話框中的色彩選取、交叉分析篩選器的多選操作、Power Query 編輯器的部分進階 UI。\\n\\n【建議練法】每天挑 3 個快捷鍵，貼在螢幕旁邊，強迫自己用一天。兩週後就會變成肌肉記憶。不要一次背 50 個！",tip:"先掌握本頁前 5 張卡片的快捷鍵，就已經贏過 90% 的 Excel 使用者了",tags:["投行","無滑鼠","全鍵盤","訓練","macOS"]},
  {name:"⌘+方向鍵",cat:"shortcut",lv:1,syntax:'⌘ + ←/→/↑/↓',desc:"跳到資料邊界（連續資料的盡頭）",example:'在 A1 按 ⌘+↓ → 跳到 A 欄最後一筆資料',detail:"搭配 Shift 可以跳選：⌘+⇧+↓ 選取到底部。這是最基本也最重要的導航技。",tip:"面試時操作 Excel 被觀察的第一件事就是你會不會用這個",tags:["導航","效率","必學","macOS"]},
  {name:"⌘+T 建立表格",cat:"shortcut",lv:1,syntax:'⌘ + T',desc:"將資料轉為「表格」，自動篩選+結構化參照+自動擴展",example:'選取資料 → ⌘+T → 確認範圍',detail:"表格是 Excel 最被低估的功能。轉為表格後：公式自動填充、圖表自動更新範圍、可用結構化參照如 =SUM(表格[金額])。",tip:"專業人士的第一步：拿到資料先 ⌘+T",tags:["表格","效率","必學","macOS"]},
  {name:"⌘+T 切換參照",cat:"shortcut",lv:1,syntax:'⌘ + T（在公式編輯中按）',desc:"切換絕對/相對參照：A1 → $A$1 → A$1 → $A1",example:'輸入 =A1*B1 後，游標在 A1 上按 ⌘+T → 變成 $A$1',detail:"$A$1=完全鎖定、A$1=鎖列、$A1=鎖欄。複製公式時，沒鎖的部分會自動偏移。⚠️ macOS 上 F4 不一定有效，改用 ⌘+T（在公式編輯模式中）。若使用外接鍵盤可嘗試 Fn+F4。",tip:"這是初學者最常忽略但影響最大的操作",tags:["參照","公式","必學","macOS"]},
  {name:"⌘+E Flash Fill",cat:"shortcut",lv:2,syntax:'⌘ + E',desc:"快速填入：Excel 自動辨識模式並填充",example:'A1=王小明，B1 手動輸入「王」，B2 按 ⌘+E → 自動擷取所有姓氏',detail:"Flash Fill 能自動辨識你的模式：拆分、合併、格式轉換都行。只要給它一兩個範例。⚠️ 部分舊版 Mac Excel 可能不支援，建議更新至 365 最新版。",tip:"比寫 LEFT/MID/RIGHT 快 10 倍的神技",tags:["快速填入","模式","效率","macOS"]},
  {name:"⌘+⇧+T 自動 SUM",cat:"shortcut",lv:1,syntax:'⌘ + ⇧ + T',desc:"一鍵自動 SUM（macOS 版）",example:'選取 A1:A10 下方空格 → ⌘+⇧+T → 自動產生 =SUM(A1:A10)',detail:"Excel 會自動偵測連續數字範圍並產生 SUM 公式。也可以選取整個區域一次加總多欄。⚠️ Windows 版為 Alt+=，macOS 改為 ⌘+⇧+T。",tip:"不用手動打 SUM 了",tags:["加總","快速","效率","macOS"]},
  {name:"⌃+; 與 ⌘+;",cat:"shortcut",lv:1,syntax:'⌃+;（日期）  ⌘+⇧+;（時間）',desc:"插入靜態日期/時間（不會自動更新）",example:'⌃+; → 輸入 2026-04-01（固定值）',detail:"跟 TODAY()/NOW() 不同，這裡輸入的是固定值，不會隨時間變化。⚠️ macOS 上日期用 ⌃+;（Control+分號），時間用 ⌘+;。",tip:"記錄「完成日期」時用這個，不要用 TODAY()",tags:["日期","靜態","效率","macOS"]},
  {name:"⌃+` 顯示公式",cat:"shortcut",lv:2,syntax:'⌃ + `（反引號）',desc:"切換顯示/隱藏所有公式",example:'按一次：所有儲存格顯示公式文字；再按一次：恢復顯示計算結果',detail:"檢查報表公式、找出硬編碼值的最快方式。搭配 ⌃+[ 可以追蹤來源儲存格。",tip:"審計/檢查報表時的必備技巧",tags:["公式","檢查","審計","macOS"]},
  {name:"KeyTips 快速存取",cat:"shortcut",lv:2,syntax:'按住 ⌥（Option）啟動 KeyTips',desc:"macOS 版的 Ribbon 鍵盤導航（2024 新功能）",example:'按住 ⌥ → 出現字母提示 → 按 H 進入常用 → 按 B 粗體',detail:"2024 年新增功能，需在 Excel → 偏好設定 → 輔助使用 → 勾選「啟用 KeyTips」。啟用後按住 ⌥ 即可看到 Ribbon 各按鈕的快捷字母，類似 Windows 的 Alt 鍵功能。這大幅提升了 macOS 上的鍵盤操作效率。",tip:"這是 Mac 用戶實現「全鍵盤操作」的關鍵功能，務必開啟！",tags:["KeyTips","Ribbon","鍵盤導航","macOS","必學"]},
  {name:"其他重要快捷鍵（macOS）",cat:"shortcut",lv:2,syntax:'⌘+D / ⌘+R / ⌃+Space / ⇧+Space / ⌘+⇧+L',desc:"填滿/選取/篩選等常用組合",example:'⌘+D=向下填滿、⌘+R=向右填滿、⌘+⇧+L=開關篩選',detail:"⌃+Space=選整欄、⇧+Space=選整列、⌃+⇧++=插入列/欄、⌘+-=刪除列/欄、⌃+⌥+Return=格內換行、⌘+1=格式化對話框。⚠️ ⌃+Space 可能被 macOS Spotlight 搶走，需在系統設定 → 鍵盤 → 快捷鍵中調整。",tip:"每天記 3 個，兩週就全會了",tags:["快捷鍵","效率","綜合","macOS"]},

  // ── Pivot Table ──
  {name:"樞紐分析表",cat:"pivot",lv:2,syntax:'插入 → 樞紐分析表（macOS: ⌥ 啟動 KeyTips → N → V）',desc:"資料分析的核心武器，拖拉即可彙總",example:'選取銷售資料 → 插入樞紐 → 列放「部門」、值放「金額」→ 即時看到各部門合計',detail:"四大區域：列（分組）、欄（展開）、值（計算）、篩選（全域篩選）。建議先把資料轉表格(⌘+T)再建樞紐。",tip:"面試最常考的 Excel 功能，沒有之一",tags:["樞紐","分析","面試","必學"]},
  {name:"樞紐-日期群組",cat:"pivot",lv:3,syntax:'右鍵日期欄位 → 群組 → 選月/季/年',desc:"把逐日資料自動彙整為月報或季報",example:'日期欄自動群組為「月」→ 一秒看到月度彙總',detail:"群組化是樞紐最被忽略但最實用的功能。可以同時依「年+月」群組。",tip:"老闆問「幫我做月報」→ 就用這個",tags:["樞紐","日期","群組","月報"]},
  {name:"樞紐-計算欄位",cat:"pivot",lv:3,syntax:'分析 → 欄位/項目/集 → 計算欄位',desc:"在樞紐表內新增公式欄（如利潤率）",example:'新增計算欄位：利潤率 = 利潤/營收',detail:"不用在原始資料加欄位，直接在樞紐表內計算衍生指標。",tip:"適合計算比率、差異、成長率等",tags:["樞紐","計算","公式"]},
  {name:"交叉分析篩選器",cat:"pivot",lv:3,syntax:'插入 → 交叉分析篩選器',desc:"漂亮的按鈕式篩選器，適合做儀表板",example:'插入部門篩選器 → 點「業務」→ 樞紐表即時更新',detail:"多個樞紐表可以共用同一個篩選器。適合做互動式報表/儀表板。",tip:"做給老闆看的報表一定要加這個",tags:["樞紐","篩選器","儀表板"]},

  // ── Conditional Formatting ──
  {name:"條件式格式化-基本",cat:"format",lv:2,syntax:'常用 → 條件式格式化 → 醒目提示/頂端底端/資料橫條/色階/圖示集',desc:"讓數字自己說故事，一眼看出異常",example:'選取金額欄 → 條件式格式化 → 資料橫條 → 即時看到大小比較',detail:"五大類型：醒目提示規則（大於/小於）、頂端底端（前10名）、資料橫條、色階（熱力圖）、圖示集（紅黃綠燈號）。",tip:"報表必加：讓不懂數字的人也能一眼看出重點",tags:["格式化","視覺化","報表"]},
  {name:"條件式格式化-自訂公式",cat:"format",lv:3,syntax:'條件式格式化 → 新增規則 → 使用公式',desc:"用公式做更靈活的條件格式",example:'整列標色：=$E2>50000（注意 $ 鎖欄不鎖列）',detail:"常用公式：=MOD(ROW(),2)=0 隔列變色、=COUNTIF($A:$A,$A2)>1 標記重複、=($D2-TODAY())<7 到期提醒、=ISBLANK($C2) 空白警告。",tip:"$ 的位置是關鍵：$E2 表示「鎖住 E 欄但列號會變」",tags:["格式化","公式","進階"]},

  // ── Data Validation ──
  {name:"資料驗證-下拉選單",cat:"format",lv:2,syntax:'資料 → 資料驗證 → 清單 → 輸入選項（逗號分隔）',desc:"建立下拉選單，限制輸入選項",example:'資料驗證 → 清單 → 業務,財務,資訊,人事 → 儲存格出現下拉選單',detail:"也可以參照範圍：=A1:A10 或命名範圍。連動下拉用 INDIRECT：子選單來源 =INDIRECT(主選單儲存格)。",tip:"給同事填的表格一定要加下拉選單，減少錯誤",tags:["驗證","下拉","資料品質"]},
  {name:"資料驗證-自訂規則",cat:"format",lv:3,syntax:'資料 → 資料驗證 → 自訂 → 輸入公式',desc:"用公式自訂任意驗證條件",example:'=AND(LEN(A2)=10, LEFT(A2,1)="A") → 限制只能輸入 A 開頭的 10 碼',detail:"可以限制整數範圍、日期範圍、文字長度、自訂公式。錯誤提醒可選：停止（不允許）、警告（提醒但可繼續）。",tip:"搭配「圈選無效資料」功能可以找出不符合規則的現有資料",tags:["驗證","公式","進階"]},

  // ── Charts ──
  {name:"圖表選擇指南",cat:"chart",lv:2,syntax:'插入 → 圖表（macOS: Fn+⌥+F1 一鍵插入）',desc:"選對圖表類型：直條=比較、折線=趨勢、圓餅=比例",example:'月營收趨勢 → 折線圖；各部門比較 → 直條圖；市佔率 → 圓餅圖',detail:"直條/橫條=類別比較、折線=時間趨勢、圓餅=比例構成（限3~6類）、散佈=相關性、組合圖=不同量級的資料。",tip:"專業人士很少用圓餅圖，因為人眼不擅長比較角度",tags:["圖表","視覺化","選擇"]},
  {name:"圖表設計原則",cat:"chart",lv:3,syntax:'設計原則：刪多餘、標數據、統色彩、寫觀點',desc:"讓圖表從「能看」變成「專業」的四大原則",example:'標題從「月營收」改成「營收連續 3 月成長 15%」→ 有觀點的標題',detail:"1.刪除多餘元素（格線、框線、不必要的圖例）2.直接在數據點上標記數字 3.統一2~3個品牌色系 4.標題要有觀點不只是描述。",tip:"老闆不想看「數據」，老闆想看「結論」",tags:["圖表","設計","專業"]},
  {name:"走勢圖 Sparklines",cat:"chart",lv:2,syntax:'插入 → 走勢圖 → 折線/直條/盈虧',desc:"放在儲存格裡的迷你圖表",example:'每位業務旁邊加一個走勢圖 → 一眼看出趨勢',detail:"走勢圖不佔額外空間，嵌在儲存格裡。適合儀表板和摘要表格。有折線、直條、盈虧三種。",tip:"儀表板必備元素",tags:["走勢圖","迷你","儀表板"]},

  // ── Named Ranges & Tables ──
  {name:"命名範圍",cat:"format",lv:3,syntax:'選取範圍 → 名稱方塊輸入名稱（或 公式→定義名稱）',desc:"給範圍取名字，公式更易讀",example:'把 E2:E100 命名為「月薪」→ =SUM(月薪) 比 =SUM(E2:E100) 好讀',detail:"建立方式：1.名稱方塊直接輸入 2.公式→定義名稱 3.⌘+⇧+F3 依選取範圍自動建立。管理：公式→名稱管理員。",tip:"專業公式必備：讓別人看得懂你的公式",tags:["命名","範圍","可讀性"]},
  {name:"結構化參照",cat:"format",lv:3,syntax:'=SUM(表格名稱[欄名])',desc:"表格專用的公式寫法，自動擴展",example:'=SUMIFS(銷售表[金額], 銷售表[部門], "業務")',detail:"⌘+T 轉表格後自動啟用。@=目前列、[欄名]=整欄、[#All]=含標題、[#Headers]=標題列。",tip:"搭配表格使用，公式永遠不用手動調整範圍",tags:["表格","參照","動態"]},

  // ── Protection ──
  {name:"工作表保護",cat:"protect",lv:2,syntax:'校閱 → 保護工作表',desc:"鎖定公式、防止誤改，但開放輸入格",example:'取消輸入格的鎖定 → 保護工作表 → 公式被保護，輸入格可編輯',detail:"預設所有儲存格都是「鎖定」的，保護工作表後才生效。要讓特定格可編輯：右鍵→儲存格格式→保護→取消鎖定。",tip:"Excel 密碼保護很容易被破解，不要當成真正的安全措施",tags:["保護","安全","鎖定"]},
  {name:"隱藏公式",cat:"protect",lv:2,syntax:'儲存格格式 → 保護 tab → 勾「隱藏」→ 保護工作表',desc:"保護工作表後，公式欄不顯示公式內容",example:'公式格勾「隱藏」→ 保護工作表 → 選取該格時公式欄顯示空白',detail:"需要同時保護工作表才會生效。適合保護商業機密的計算邏輯。",tip:"搭配 VBA 密碼保護模組可以更安全",tags:["隱藏","公式","安全"]},
  {name:"活頁簿保護",cat:"protect",lv:2,syntax:'校閱 → 保護活頁簿結構 / 檔案→另存→工具→一般選項',desc:"保護結構（防新增刪除工作表）或設開啟密碼",example:'保護活頁簿結構 → 無法新增/刪除/重新命名工作表',detail:"活頁簿結構保護：防止修改工作表結構。開啟密碼：沒密碼打不開檔案。修改密碼：可以看但不能改。",tip:"重要檔案一定要設開啟密碼，結構保護只是防手滑",tags:["活頁簿","密碼","安全"]},

  // ── Power Query ──
  {name:"Power Query 入門",cat:"pquery",lv:3,syntax:'資料 → 取得資料 → 從 Excel/CSV/資料夾/Web',desc:"Excel 內建的 ETL 工具，不用寫程式就能清洗轉換資料",example:'匯入 CSV → 移除空列 → 分割欄位 → 載入到表格',detail:"核心概念：建立「查詢步驟」，資料更新時一鍵重跑所有步驟。不需要公式或 VBA。支援幾十種資料來源。",tip:"10 萬列以上的資料，用 Power Query 比公式快很多倍",tags:["Power Query","ETL","匯入","清洗"]},
  {name:"PQ 合併查詢",cat:"pquery",lv:4,syntax:'常用 → 合併查詢（= VLOOKUP 的升級版）',desc:"跨表合併資料，支援 Inner/Left/Full Join",example:'訂單表 + 客戶表 → 合併查詢 → 依客戶ID關聯',detail:"比 VLOOKUP 更強：支援多欄比對、模糊比對、多種 Join 方式。附加查詢=垂直合併（相同結構的表上下接在一起）。",tip:"需要跨表查找時，先考慮 Power Query 再考慮 VLOOKUP",tags:["Power Query","合併","Join","進階"]},
  {name:"PQ 從資料夾匯入",cat:"pquery",lv:4,syntax:'取得資料 → 從資料夾 → 選擇資料夾',desc:"自動合併資料夾內所有檔案（月報神技）",example:'選取「月報」資料夾 → 自動合併所有 Excel/CSV 檔案',detail:"每月的報表放同一個資料夾，Power Query 自動合併。新增檔案後只要按「重新整理」就自動包含。搭配參數化查詢更靈活。",tip:"辦公室最值得學的自動化技巧之一",tags:["Power Query","資料夾","合併","自動化"]},

  // ── What-If Analysis ──
  {name:"目標搜尋",cat:"advanced",lv:3,syntax:'資料 → 假設分析 → 目標搜尋',desc:"已知結果，反推輸入值",example:'利潤公式在 C1，目標搜尋：C1 要達到 100 萬，B1（售價）應為多少？',detail:"設定：目標儲存格=結果格、目標值=想達到的數字、變更儲存格=要調整的輸入格。",tip:"做損益兩平分析的標準工具",tags:["假設分析","反推","目標"]},
  {name:"資料表(模擬分析)",cat:"advanced",lv:4,syntax:'資料 → 假設分析 → 資料表',desc:"一次看多個假設情境的結果（敏感度分析）",example:'左欄放不同利率、上列放不同年期 → 交叉看月付款金額',detail:"單變數：一個輸入值變化看結果。雙變數：兩個輸入值同時變化，矩陣式呈現。適合定價策略和貸款試算。",tip:"做給決策者看的分析一定要用這個",tags:["假設分析","敏感度","矩陣"]},
];

// ── VBA Data ──
const vbaTemplates = [
  {title:"Hello World",cat:"basic",lv:1,code:`Sub HelloWorld()\n    MsgBox "Hello, Excel VBA!"\nEnd Sub`,explain:"MsgBox 彈出對話框，你的第一個 VBA"},
  {title:"變數與資料型態",cat:"basic",lv:1,code:`Sub Variables()\n    Dim name As String\n    Dim age As Integer\n    Dim salary As Double\n    Dim isActive As Boolean\n    \n    name = "Astor"\n    age = 30\n    salary = 55000.5\n    isActive = True\n    \n    MsgBox name & " / " & age\nEnd Sub`,explain:"Dim 宣告，As 指定型態，& 串接"},
  {title:"讀寫儲存格",cat:"basic",lv:1,code:`Sub ReadWrite()\n    Range("A1").Value = "Hello"\n    Cells(2, 1).Value = "World"\n    \n    Dim val As String\n    val = Range("A1").Value\n    MsgBox "A1 = " & val\nEnd Sub`,explain:"Range(\"A1\") 或 Cells(列,欄) 兩種方式"},
  {title:"For 迴圈",cat:"basic",lv:1,code:`Sub ForLoop()\n    Dim i As Integer\n    For i = 1 To 10\n        Cells(i, 1).Value = i\n        Cells(i, 1).Interior.Color = RGB(200, 230, 255)\n    Next i\nEnd Sub`,explain:"For i = 起始 To 結束 ... Next i"},
  {title:"For Each 遍歷",cat:"basic",lv:2,code:`Sub ForEachDemo()\n    Dim cell As Range\n    For Each cell In Range("A1:A10")\n        If cell.Value > 5 Then\n            cell.Font.Bold = True\n            cell.Font.Color = RGB(255, 0, 0)\n        End If\n    Next cell\nEnd Sub`,explain:"逐一處理範圍內的每個儲存格"},
  {title:"If 條件判斷",cat:"basic",lv:1,code:`Sub IfDemo()\n    Dim score As Integer\n    score = Range("B1").Value\n    \n    If score >= 90 Then\n        MsgBox "優秀"\n    ElseIf score >= 60 Then\n        MsgBox "及格"\n    Else\n        MsgBox "加油"\n    End If\nEnd Sub`,explain:"If / ElseIf / Else / End If"},
  {title:"Select Case",cat:"basic",lv:2,code:`Sub SelectCaseDemo()\n    Select Case Range("A1").Value\n        Case "A": MsgBox "優秀"\n        Case "B": MsgBox "良好"\n        Case "C": MsgBox "及格"\n        Case Else: MsgBox "未知"\n    End Select\nEnd Sub`,explain:"比多層 If 更乾淨的寫法"},
  {title:"Function 自訂函式",cat:"basic",lv:2,code:`Function CalcBonus(salary As Double, score As Integer) As Double\n    If score >= 90 Then\n        CalcBonus = salary * 3\n    ElseIf score >= 75 Then\n        CalcBonus = salary * 2\n    Else\n        CalcBonus = salary * 1\n    End If\nEnd Function`,explain:"Function 可當公式用：=CalcBonus(A1,B1)"},
  {title:"陣列批次處理",cat:"data",lv:3,code:`Sub ArrayBatch()\n    Dim arr() As Variant\n    arr = Range("A1:A1000").Value\n    Dim i As Long\n    For i = 1 To UBound(arr)\n        arr(i, 1) = arr(i, 1) * 1.1\n    Next i\n    Range("B1:B1000").Value = arr\nEnd Sub`,explain:"先讀入陣列、記憶體處理、再寫回，比逐格快 100 倍"},
  {title:"Dictionary 計數",cat:"data",lv:3,code:`Sub DictCount()\n    Dim d As Object\n    Set d = CreateObject("Scripting.Dictionary")\n    Dim c As Range\n    For Each c In Range("A1:A100")\n        If c.Value <> "" Then\n            d(c.Value) = d(c.Value) + 1\n        End If\n    Next c\n    Dim i As Long: i = 1\n    Dim k As Variant\n    For Each k In d.Keys\n        Cells(i, 5) = k\n        Cells(i, 6) = d(k)\n        i = i + 1\n    Next k\nEnd Sub`,explain:"Dictionary 去重+計數，比 COUNTIF 更靈活"},
  {title:"刪除空白列",cat:"data",lv:2,code:`Sub DeleteEmptyRows()\n    Dim i As Long\n    For i = ActiveSheet.UsedRange.Rows.Count To 1 Step -1\n        If WorksheetFunction.CountA(Rows(i)) = 0 Then\n            Rows(i).Delete\n        End If\n    Next i\nEnd Sub`,explain:"從最後一列往上刪，避免索引錯亂"},
  {title:"格式化表格",cat:"format",lv:2,code:`Sub FormatTable()\n    With ActiveSheet.UsedRange\n        .Borders.LineStyle = xlContinuous\n        .Font.Name = "Arial"\n        .Font.Size = 10\n    End With\n    With Rows(1)\n        .Font.Bold = True\n        .Font.Color = RGB(255, 255, 255)\n        .Interior.Color = RGB(47, 84, 150)\n    End With\n    Cells.EntireColumn.AutoFit\nEnd Sub`,explain:"With 減少重複，一鍵美化表格"},
  {title:"條件標色",cat:"format",lv:2,code:`Sub ConditionalColor()\n    Dim c As Range\n    For Each c In Range("E2:E" & Cells(Rows.Count, 5).End(xlUp).Row)\n        Select Case True\n            Case c.Value >= 50000: c.Interior.Color = RGB(198, 239, 206)\n            Case c.Value >= 30000: c.Interior.Color = RGB(255, 235, 156)\n            Case c.Value > 0: c.Interior.Color = RGB(255, 199, 206)\n        End Select\n    Next c\nEnd Sub`,explain:"Select Case True 是多條件標色的巧妙寫法"},
  {title:"匯出全部為 PDF",cat:"file",lv:3,code:`Sub ExportPDF()\n    Dim ws As Worksheet\n    Dim p As String: p = ThisWorkbook.Path & "\\"\n    For Each ws In ThisWorkbook.Worksheets\n        ws.ExportAsFixedFormat xlTypePDF, p & ws.Name & ".pdf"\n    Next ws\n    MsgBox "匯出完成"\nEnd Sub`,explain:"批次匯出 PDF，辦公室最受歡迎的自動化"},
  {title:"建立工作表目錄",cat:"file",lv:3,code:`Sub CreateIndex()\n    Dim ws As Worksheet, idx As Worksheet, i As Integer\n    On Error Resume Next\n    Application.DisplayAlerts = False\n    Sheets("目錄").Delete\n    Application.DisplayAlerts = True\n    On Error GoTo 0\n    Set idx = Sheets.Add(Before:=Sheets(1))\n    idx.Name = "目錄"\n    idx.Range("A1") = "工作表目錄"\n    idx.Range("A1").Font.Size = 14\n    i = 3\n    For Each ws In ThisWorkbook.Sheets\n        If ws.Name <> "目錄" Then\n            idx.Hyperlinks.Add idx.Cells(i, 1), "", "'" & ws.Name & "'!A1", , ws.Name\n            i = i + 1\n        End If\n    Next ws\nEnd Sub`,explain:"自動產生含超連結的目錄頁"},
  {title:"自動月報表",cat:"file",lv:4,code:`Sub MonthlyReport()\n    Dim wsData As Worksheet, wsRpt As Worksheet\n    Set wsData = Sheets("資料")\n    Set wsRpt = Sheets.Add(After:=Sheets(Sheets.Count))\n    wsRpt.Name = "月報_" & Format(Date, "YYYYMM")\n    wsRpt.Range("A1") = Format(Date, "YYYY年MM月") & " 月報表"\n    Dim lr As Long: lr = wsData.Cells(Rows.Count, 1).End(xlUp).Row\n    wsRpt.Range("A3") = "總筆數"\n    wsRpt.Range("B3") = lr - 1\n    wsRpt.Range("A4") = "總金額"\n    wsRpt.Range("B4") = WorksheetFunction.Sum(wsData.Range("E2:E" & lr))\n    wsRpt.Range("B4").NumberFormat = "$#,##0"\nEnd Sub`,explain:"自動匯總+產生新工作表的實務範本"},
  {title:"Outlook 自動寄信",cat:"file",lv:4,code:`Sub SendEmail()\n    Dim olApp As Object, olMail As Object\n    Set olApp = CreateObject("Outlook.Application")\n    Set olMail = olApp.CreateItem(0)\n    With olMail\n        .To = "colleague@company.com"\n        .Subject = "報表 " & Format(Date, "YYYY/MM/DD")\n        .Body = "附件為今日報表"\n        .Attachments.Add ThisWorkbook.FullName\n        .Display\n    End With\nEnd Sub`,explain:"透過 Outlook 自動建立郵件並附檔"},
  {title:"錯誤處理框架",cat:"pro",lv:3,code:`Sub ProfessionalTemplate()\n    On Error GoTo ErrHandler\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n    \n    ' === 你的主要邏輯 ===\n    \nCleanUp:\n    Application.ScreenUpdating = True\n    Application.Calculation = xlCalculationAutomatic\n    Exit Sub\nErrHandler:\n    MsgBox "錯誤 " & Err.Number & ": " & Err.Description\n    Resume CleanUp\nEnd Sub`,explain:"專業模板：錯誤處理 + 效能優化 + 清理"},
  {title:"效能優化開關",cat:"pro",lv:3,code:`Sub SpeedUp()\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n    Application.EnableEvents = False\n    \n    ' ... 大量操作 ...\n    \n    Application.ScreenUpdating = True\n    Application.Calculation = xlCalculationAutomatic\n    Application.EnableEvents = True\nEnd Sub`,explain:"大量操作時關閉更新，速度提升 10 倍以上"},
  {title:"Worksheet 事件",cat:"pro",lv:4,code:`' 放在「工作表」模組（右鍵工作表 → 檢視程式碼）\nPrivate Sub Worksheet_Change(ByVal Target As Range)\n    If Target.Column = 1 Then\n        Target.Offset(0, 1).Value = Now()\n    End If\nEnd Sub`,explain:"自動觸發：A 欄改動時在 B 欄記錄時間戳"},
  {title:"定時執行巨集",cat:"pro",lv:4,code:`Sub StartTimer()\n    Application.OnTime Now + TimeValue("00:01:00"), "MyTask"\nEnd Sub\n\nSub MyTask()\n    Range("A1").Value = Now()\n    StartTimer\nEnd Sub\n\nSub StopTimer()\n    Application.OnTime Now + TimeValue("00:01:00"), "MyTask", , False\nEnd Sub`,explain:"Application.OnTime 排程定時執行"},
];

// ── Lesson → 速查分類對應 ──
// pick: 本課關鍵卡片（inline 顯示在 lesson 頂部）
// highlights: practice.xlsx 瀏覽器要高亮的關鍵字
const LESSON_CHEAT_MAP = {
  "P1-01": {cats:["shortcut"], title:"快捷鍵速查",
    pick:["⌘+方向鍵","⌘+T 建立表格","⌘+T 切換參照","⌘+⇧+T 自動 SUM","KeyTips 快速存取"],
    highlights:["快捷鍵","⌘","⌃"]},
  "P1-02": {cats:["math","condition"], title:"基礎函式速查",
    pick:["SUM","AVERAGE","COUNT","MAX / MIN","MEDIAN"],
    highlights:["SUM","AVERAGE","COUNT","MAX","MIN","MEDIAN","LARGE"]},
  "P1-03": {cats:["condition"], title:"條件判斷速查",
    pick:["IF","IFS","AND / OR","IFERROR","SWITCH"],
    highlights:["IF","IFS","AND","OR","IFERROR","SWITCH"]},
  "P2-01": {cats:["math","condition"], title:"條件統計速查",
    pick:["COUNTIF","COUNTIFS","SUMIF","SUMIFS","AVERAGEIFS"],
    highlights:["COUNTIF","SUMIF","AVERAGEIFS"]},
  "P2-02": {cats:["lookup"], title:"查找函式速查",
    pick:["VLOOKUP","XLOOKUP","INDEX + MATCH","MATCH"],
    highlights:["VLOOKUP","XLOOKUP","INDEX","MATCH"]},
  "P2-03": {cats:["pivot"], title:"樞紐分析速查",
    pick:["樞紐分析表","樞紐-日期群組","樞紐-計算欄位","交叉分析篩選器"],
    highlights:["樞紐","pivot"]},
  "P2-04": {cats:["format"], title:"格式化速查",
    pick:["條件式格式化-基本","條件式格式化-自訂公式"],
    highlights:["條件式格式化","格式化"]},
  "P3-01": {cats:["format"], title:"資料驗證速查",
    pick:["資料驗證-下拉選單","資料驗證-自訂規則"],
    highlights:["資料驗證","下拉"]},
  "P3-02": {cats:["text","date"], title:"文字/日期函式速查",
    pick:["TRIM","LEFT / MID / RIGHT","TEXT","TEXTJOIN","DATEDIF"],
    highlights:["TRIM","LEFT","MID","RIGHT","TEXT","DATEDIF","WORKDAY"]},
  "P3-03": {cats:["chart"], title:"圖表速查",
    pick:["圖表選擇指南","圖表設計原則","走勢圖 Sparklines"],
    highlights:["圖表","折線","直條","圓餅"]},
  "P3-04": {cats:["format","shortcut"], title:"命名範圍與表格速查",
    pick:["⌘+T 建立表格","命名範圍","結構化參照"],
    highlights:["表格","命名範圍"]},
  "P3-05": {cats:["protect"], title:"保護與安全速查",
    pick:["工作表保護","隱藏公式","活頁簿保護"],
    highlights:["保護","密碼"]},
  "P4-01": {cats:["array"], title:"動態陣列速查",
    pick:["FILTER","SORT","UNIQUE","SEQUENCE"],
    highlights:["FILTER","SORT","UNIQUE","SEQUENCE"]},
  "P4-02": {cats:["advanced"], title:"進階函式速查",
    pick:["INDIRECT","OFFSET","LET","LAMBDA","AGGREGATE"],
    highlights:["INDIRECT","OFFSET","LET","LAMBDA","AGGREGATE"]},
  "P4-03": {cats:["advanced"], title:"假設分析速查",
    pick:["目標搜尋","資料表(模擬分析)"],
    highlights:["假設分析","目標搜尋"]},
  "P4-04": {cats:["pquery"], title:"Power Query 速查",
    pick:["Power Query 入門","PQ 合併查詢","PQ 從資料夾匯入"],
    highlights:["Power Query","匯入"]},
  "P4-05": {cats:["pquery","advanced"], title:"Power Pivot 速查",
    pick:["Power Query 入門","LET","LAMBDA"],
    highlights:["Power Pivot","DAX"]},
  "P5-01": {cats:[], kinds:["vba"], vbaCats:["basic"], title:"VBA 基礎範本",
    pick:[], pickVba:["Hello World","變數與資料型態","讀寫儲存格","For 迴圈","If 條件判斷"],
    highlights:["Sub","Dim"]},
  "P5-02": {cats:[], kinds:["vba"], vbaCats:["data","format"], title:"VBA 進階範本",
    pick:[], pickVba:["陣列批次處理","Dictionary 計數","格式化表格","條件標色"],
    highlights:["For Each","Dictionary"]},
  "P5-03": {cats:[], kinds:["vba"], vbaCats:["file","pro"], title:"VBA 實戰範本",
    pick:[], pickVba:["匯出全部為 PDF","建立工作表目錄","自動月報表","Outlook 自動寄信"],
    highlights:["Workbook","Worksheet"]},
  "P5-04": {cats:[], kinds:["vba"], title:"VBA 綜合範本",
    pick:[], pickVba:["錯誤處理框架","效能優化開關","Worksheet 事件"],
    highlights:["VBA"]},
};

export const CHEATSHEET_DATA = { formulas: formulas, vbaTemplates: vbaTemplates, lessonMap: LESSON_CHEAT_MAP };
window.CHEATSHEET_DATA = CHEATSHEET_DATA; // cheatsheet.html inline script 需要此 window global
