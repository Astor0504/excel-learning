# Excel 21 課的詳細教學內容（語法 / 範例 / 場景 / 常見錯誤）
EXCEL_LESSONS = {
"P1-01": {
  "tldr": "掌握 macOS 版 Excel 的核心快捷鍵，速度比滑鼠快 10 倍",
  "intro": "Excel 高手和新手最大的差別不是函數，而是手不離鍵盤。先把這 20 個最常用的快捷鍵練熟，未來操作時間會減半。",
  "blocks": [
    {"h":"必背快捷鍵 Top 10", "type":"table",
     "rows":[["動作","macOS","說明"],
             ["儲存","⌘ S","隨時存檔"],
             ["復原 / 重做","⌘ Z / ⌘ ⇧ Z",""],
             ["全選資料區","⌘ A 或 ⌃ A",""],
             ["選到資料邊界","⌘ + 方向鍵","跳到非空儲存格邊緣"],
             ["選取一整欄/列","⌃ Space / ⇧ Space",""],
             ["插入欄/列","⌃ ⇧ +",""],
             ["刪除欄/列","⌘ -","macOS 正確版，⌃ - 在新版已失效"],
             ["編輯儲存格","⌃ U 或 fn+F2","⌃ U 可能被中文輸入法攔截，MacBook 按 F2 要配 fn"],
             ["填滿向下","⌃ D","複製上方儲存格"],
             ["快速搜尋","⌘ F",""]]},
    {"h":"與 Windows 的差異", "type":"p",
     "text":"Mac 沒有 Alt 鍵 → 用 Option（⌥）；Mac 版「插入/刪除列欄」用 ⌃ ⇧ + / ⌘ -；編輯儲存格用 ⌃ U 或 fn+F2（MacBook 的 F 列要配 fn）；KeyTips 在 macOS 2024 後已支援（需在偏好設定啟用）。"},
    {"h":"常見錯誤", "type":"warn",
     "text":"⌘ + 方向鍵在新版才能跳邊界，舊版可能要用 ⌃ + 方向鍵。如果按了沒反應，先試另一個。"}
  ],
  "tasks":["背熟存檔/復原/全選 3 個鍵","練習用 ⌘+方向鍵跳到表格邊緣","練習 ⌃ ⇧ + 插入新列","用 ⌘ F 找一個關鍵字"],
},

"P1-02": {
  "tldr": "7 個最常用的統計函數，幾乎所有報表都從這裡開始",
  "intro": "這是 Excel 的「九九乘法表」——背熟這幾個，才有資格往後學進階函數。",
  "blocks": [
    {"h":"基礎函數速覽", "type":"table",
     "rows":[["函數","用途","範例"],
             ["SUM","加總","=SUM(A2:A100)"],
             ["AVERAGE","平均","=AVERAGE(B2:B100)"],
             ["COUNT","計算數字個數","=COUNT(C2:C100)"],
             ["COUNTA","計算非空個數（含文字）","=COUNTA(D2:D100)"],
             ["MAX / MIN","最大 / 最小","=MAX(E2:E100)"],
             ["MEDIAN","中位數","=MEDIAN(F2:F100)"],
             ["LARGE / SMALL","第 N 大 / 小","=LARGE(G2:G100, 3) → 第 3 大"]]},
    {"h":"實戰範例", "type":"code",
     "lang":"excel",
     "code":"# 業績統計（A欄=業務員，B欄=金額）\n=SUM(B:B)              # 全部業績\n=AVERAGE(B2:B100)      # 平均單筆\n=COUNT(B:B)            # 有幾筆業績\n=COUNTA(A:A) - 1       # 業務員人數（扣標題）\n=LARGE(B:B, 1)         # 最高一筆\n=LARGE(B:B, 3)         # 第 3 高（找前三名）"},
    {"h":"常見錯誤", "type":"warn",
     "text":"COUNT 只算「數字」，文字會被忽略！要算筆數請用 COUNTA。AVERAGE 會自動忽略空白和文字，但會把 0 算進去。"}
  ],
  "tasks":["分清楚 COUNT vs COUNTA","用 LARGE 找前 3 名","知道 AVERAGE 怎麼處理 0 跟空白"],
},

"P1-03": {
  "tldr": "IF 是 Excel 的「if 陳述式」，會了它才能寫條件邏輯",
  "intro": "從這課開始 Excel 變成「會思考」的工具。IF 加上 AND/OR 幾乎可以表達所有判斷。",
  "blocks": [
    {"h":"語法", "type":"code","lang":"excel",
     "code":"=IF(條件, 條件成立的值, 條件不成立的值)\n\n# 範例：成績及格判斷\n=IF(B2>=60, \"及格\", \"不及格\")\n\n# 多重條件：用 IFS（比巢狀 IF 好讀）\n=IFS(B2>=90, \"A\", B2>=80, \"B\", B2>=70, \"C\", TRUE, \"F\")\n\n# 複合條件：AND / OR\n=IF(AND(B2>=60, C2>=60), \"兩科都過\", \"補考\")\n=IF(OR(D2=\"VIP\", E2>10000), \"享折扣\", \"原價\")"},
    {"h":"巢狀 IF vs IFS", "type":"p",
     "text":"舊寫法 =IF(A>90,\"A\",IF(A>80,\"B\",IF(A>70,\"C\",\"F\"))) → 括號地獄。新版直接用 IFS 一行搞定，最後一個 TRUE 當 else。"},
    {"h":"小技巧", "type":"p",
     "text":"MOD(A2, 2)=0 可以判斷偶數。配合 IF 可以做「奇數列上色」之類的判斷。"}
  ],
  "tasks":["寫一個 IF 判斷及格","用 IFS 寫等第","用 AND 寫雙條件 IF","知道 IFS 的 TRUE 是什麼"],
},

"P2-01": {
  "tldr": "SUMIF 系列：在條件下才加總/計算的「會挑食」函數",
  "intro": "業績報表的核心函數。看懂這 6 個，老闆問你「電子部 3 月業績」可以 30 秒答出來。",
  "blocks": [
    {"h":"6 大條件統計函數", "type":"table",
     "rows":[["函數","用途"],
             ["COUNTIF","符合條件的「個數」"],
             ["COUNTIFS","多條件版"],
             ["SUMIF","符合條件的「加總」"],
             ["SUMIFS","多條件版"],
             ["AVERAGEIF","符合條件的「平均」"],
             ["AVERAGEIFS","多條件版"]]},
    {"h":"語法差異（關鍵）", "type":"code","lang":"excel",
     "code":"# 單條件：條件範圍在前\n=SUMIF(條件範圍, 條件, 加總範圍)\n=SUMIF(A:A, \"電子\", C:C)\n\n# 多條件：加總範圍在前！順序不同！\n=SUMIFS(加總範圍, 條件範圍1, 條件1, 條件範圍2, 條件2, ...)\n=SUMIFS(C:C, A:A, \"電子\", B:B, \">2024-01-01\")"},
    {"h":"萬用字元", "type":"code","lang":"excel",
     "code":"=COUNTIF(A:A, \"陳*\")        # 開頭是「陳」\n=COUNTIF(A:A, \"*股份*\")      # 包含「股份」\n=SUMIF(A:A, \"<>已取消\", B:B) # 不等於已取消"},
    {"h":"常見錯誤", "type":"warn",
     "text":"SUMIF 跟 SUMIFS 的參數順序不一樣！前者「條件範圍在前、加總範圍在後」，後者「加總範圍在最前面」——記反了結果會錯。"}
  ],
  "tasks":["寫一個 SUMIF","寫一個 SUMIFS（雙條件）","用萬用字元 *","知道兩者參數順序差異"],
},

"P2-02": {
  "tldr": "VLOOKUP / XLOOKUP：跨表查資料的命脈函數",
  "intro": "面試最常被問的就是這個。能解釋清楚 VLOOKUP 和 XLOOKUP 的差別，就贏一半。",
  "blocks": [
    {"h":"VLOOKUP（舊版王者）", "type":"code","lang":"excel",
     "code":"=VLOOKUP(查的值, 查表範圍, 第幾欄, [是否模糊比對])\n\n# 用員工編號查姓名\n=VLOOKUP(A2, 員工表!A:D, 2, FALSE)\n\n# ⚠️ 4 個限制：\n# 1. 只能往「右邊」查，查左邊要用 INDEX+MATCH\n# 2. 第四個參數必填 FALSE（精確比對），否則會出怪結果\n# 3. 中間欄位數要手動數，加減欄位就會錯位\n# 4. 找不到回傳 #N/A，要包 IFERROR"},
    {"h":"XLOOKUP（M365 新王）", "type":"code","lang":"excel",
     "code":"=XLOOKUP(查的值, 查找範圍, 回傳範圍, [找不到回傳], [比對模式], [搜尋方向])\n\n=XLOOKUP(A2, 員工表!A:A, 員工表!B:B, \"查無此人\")\n\n# ✅ 優點：\n# - 可以往任何方向查\n# - 內建找不到的處理\n# - 不用數欄位\n# - 支援由下往上找"},
    {"h":"INDEX + MATCH（萬能組合）", "type":"code","lang":"excel",
     "code":"=INDEX(回傳範圍, MATCH(查的值, 查找範圍, 0))\n=INDEX(B:B, MATCH(\"陳小明\", A:A, 0))\n# 適合舊版 Excel、需要查左邊、雙向查表"},
    {"h":"哪個該用？", "type":"p",
     "text":"M365 / Excel 2021 → 一律用 XLOOKUP；舊版 Office → INDEX+MATCH > VLOOKUP。VLOOKUP 只在維護舊檔案時才碰。"}
  ],
  "tasks":["寫一個 VLOOKUP","寫一個 XLOOKUP","寫一個 INDEX+MATCH","知道 VLOOKUP 4 個限制"],
},

"P2-03": {
  "tldr": "樞紐分析表：5 秒做出 30 分鐘的報表",
  "intro": "如果只能學一個 Excel 功能，請學樞紐。它能把幾萬筆原始資料瞬間變成可讀報表。",
  "blocks": [
    {"h":"建立步驟", "type":"p",
     "text":"① 點任一儲存格 → ② 插入 → 樞紐分析表 → ③ 拖欄位到 4 大區域。"},
    {"h":"4 大區域怎麼用", "type":"table",
     "rows":[["區域","用途","範例"],
             ["列（Rows）","變成左邊的「分組」","部門、產品類別"],
             ["欄（Columns）","變成上方的「分組」","月份、年度"],
             ["值（Values）","要計算的數字","業績、數量"],
             ["篩選（Filters）","頂部下拉篩選","地區、年度"]]},
    {"h":"超實用功能", "type":"p",
     "text":"右鍵 → 群組 → 可以把日期自動分成「年/季/月」。值欄位設定 → 改顯示為「百分比」、「累計」、「差異」等。插入交叉分析篩選器（Slicer）→ 視覺化篩選按鈕。"},
    {"h":"常見錯誤", "type":"warn",
     "text":"原始資料有合併儲存格、空白欄位、欄位名稱重複 → 樞紐會壞。建議資料先轉成「表格」（⌘+T）再做樞紐，新增資料會自動納入。"}
  ],
  "tasks":["建立一個樞紐","拖欄位到 4 個區域","用日期群組成月","插入一個交叉分析篩選器"],
},

"P2-04": {
  "tldr": "條件式格式化：讓重要的資料自己跳出來",
  "intro": "報表要給人看，不上色等於沒做。條件式格式可以自動把高/低值上色，看一眼就知道重點。",
  "blocks": [
    {"h":"3 種視覺化", "type":"table",
     "rows":[["類型","適合場景"],
             ["資料橫條","快速比較大小（像橫條圖）"],
             ["色階","熱力圖效果（紅高綠低）"],
             ["圖示集","加箭頭/燈號/星星"]]},
    {"h":"自訂公式規則（最強）", "type":"code","lang":"excel",
     "code":"# 整列上色：當 D 欄是「逾期」就把整列變紅\n# 選整列 → 條件式格式化 → 新增規則 → 使用公式\n=$D2=\"逾期\"\n# ⚠️ $D2 加 $ 是鎖欄不鎖列，這樣每一列都會用自己的 D 欄判斷\n\n# 高亮週末\n=WEEKDAY($A2, 2)>=6\n\n# 標記高於平均\n=B2>AVERAGE($B$2:$B$100)"},
    {"h":"常見錯誤", "type":"warn",
     "text":"公式裡 $ 用錯地方是最常犯的錯。記住：要鎖「欄」就 $X，要鎖「列」就 X$。整列上色一定要鎖欄不鎖列。"}
  ],
  "tasks":["套一個資料橫條","套一個色階","用公式規則整列上色","知道 $ 鎖欄 vs 鎖列"],
},

"P3-01": {
  "tldr": "資料驗證：讓使用者只能輸入合法的值",
  "intro": "不想再收到「2024/13/45」「電子部電子部電子部」這種髒資料？資料驗證就是答案。",
  "blocks": [
    {"h":"最常用：下拉選單", "type":"p",
     "text":"資料 → 資料驗證 → 允許「清單」→ 來源輸入「電子,業務,行政」（用半形逗號）或選一個範圍。儲存格右側出現下拉箭頭。"},
    {"h":"串聯下拉（進階）", "type":"code","lang":"excel",
     "code":"# 第一個下拉選「部門」\n# 第二個下拉根據第一個顯示對應「員工」\n=INDIRECT(A2)\n# A2 是「電子」就會顯示名為「電子」的範圍"},
    {"h":"其他驗證類型", "type":"table",
     "rows":[["類型","用途"],
             ["整數 / 小數","限定數值範圍"],
             ["日期","限定日期區間"],
             ["文字長度","限定字數"],
             ["自訂","用公式 =A2>0"]]}
  ],
  "tasks":["建立一個下拉選單","設一個日期範圍驗證","用 INDIRECT 做串聯下拉","知道「自訂」可以放公式"],
},

"P3-02": {
  "tldr": "文字 + 日期函數：清資料最常用的工具組",
  "intro": "從別處貼過來的資料總是亂七八糟？這課的函數就是專門用來「整理髒資料」的。",
  "blocks": [
    {"h":"文字函數", "type":"code","lang":"excel",
     "code":"=TRIM(A2)              # 去除前後空白\n=MID(A2, 3, 4)         # 從第 3 個字取 4 個\n=LEFT(A2, 2)           # 取最左 2 個字\n=RIGHT(A2, 2)          # 取最右 2 個字\n=FIND(\"@\", A2)         # 找符號位置\n=SUBSTITUTE(A2, \"-\", \"\")  # 取代\n=TEXT(A2, \"yyyy-mm-dd\")   # 數字/日期 → 格式化字串\n=A2 & \"-\" & B2         # 串接（也可用 &）"},
    {"h":"日期函數", "type":"code","lang":"excel",
     "code":"=TODAY()                       # 今天\n=NOW()                         # 現在（含時間）\n=DATEDIF(A2, TODAY(), \"Y\")    # 年資（年）\n=DATEDIF(A2, TODAY(), \"M\")    # 月差\n=DATEDIF(A2, TODAY(), \"D\")    # 天數\n=WORKDAY(TODAY(), 5)          # 5 個工作天後\n=EOMONTH(TODAY(), 0)          # 本月最後一天\n=EOMONTH(TODAY(), -1)+1       # 本月第一天"},
    {"h":"常見錯誤", "type":"warn",
     "text":"DATEDIF 是隱藏函數，輸入時不會自動提示，但能用。MID 從 1 開始算，不是 0。"}
  ],
  "tasks":["用 TRIM 清空白","用 LEFT/MID/RIGHT 取字","用 DATEDIF 算年資","用 EOMONTH 算月底"],
},

"P3-03": {
  "tldr": "圖表設計：選對類型比畫得漂亮重要",
  "intro": "Excel 提供 17 種圖表，但 80% 的場合只需要 4 種。選錯類型再美也是錯。",
  "blocks": [
    {"h":"4 種圖表配對什麼資料", "type":"table",
     "rows":[["圖表","適合"],
             ["柱狀 / 長條","比較類別大小（部門業績）"],
             ["折線","時間趨勢（每月收入）"],
             ["圓餅","佔比（市場份額，類別 ≤ 5）"],
             ["散佈","兩變數關聯（廣告 vs 銷售）"]]},
    {"h":"設計三原則", "type":"p",
     "text":"① 一張圖只講一件事 ② 拿掉所有不必要的東西（網格線、3D 效果、陰影）③ 重點顏色用深色，其他用灰色——眼睛會自動看到重點。"},
    {"h":"走勢圖（Sparklines）", "type":"p",
     "text":"插入 → 走勢圖 → 在單一儲存格內塞迷你折線圖。最適合做 KPI 表格旁邊的趨勢列。"},
    {"h":"常見錯誤", "type":"warn",
     "text":"圓餅圖類別超過 5 個就難讀，改用橫條圖。3D 圓餅圖永遠不要用。"}
  ],
  "tasks":["插入一個柱狀圖","插入一個折線圖","插入一個走勢圖","知道圓餅圖什麼時候不該用"],
},

"P3-04": {
  "tldr": "命名範圍 + 表格：讓公式變得「會說人話」",
  "intro": "=SUMIFS(C2:C500, A2:A500, \"電子\") vs =SUMIFS(業績[金額], 業績[部門], \"電子\")——後者半年後回來看也懂。",
  "blocks": [
    {"h":"命名範圍", "type":"p",
     "text":"選範圍 → 左上角「名稱方塊」輸入名字 → Enter。之後公式可以直接寫名字。例如把 B2:B100 取名為「業績」，公式變成 =SUM(業績)。"},
    {"h":"表格（最強功能之一）", "type":"code","lang":"excel",
     "code":"# 選資料 → ⌘+T → 套用表格樣式\n# 表格會自動：\n# 1. 加標題排序篩選按鈕\n# 2. 公式自動展開（新增一列就自動套用）\n# 3. 範圍動態擴展（新資料樞紐自動更新）\n\n# 結構化參照（公式裡用欄位名）\n=SUM(業績表[金額])\n=SUMIFS(業績表[金額], 業績表[部門], \"電子\")\n=業績表[@金額] * 0.05  # @ 代表本列"},
    {"h":"小技巧", "type":"p",
     "text":"養成把所有資料先轉成「表格」的習慣，後面所有函數和樞紐都會變得超穩定。"}
  ],
  "tasks":["建立一個命名範圍","用 ⌘+T 把資料轉成表格","用結構化參照寫一個 SUMIFS","知道 [@欄位] 是什麼"],
},

"P3-05": {
  "tldr": "保護工作表：避免別人亂改你的公式",
  "intro": "把報表寄出去之前，先鎖起來——避免使用者誤刪公式還怪你。",
  "blocks": [
    {"h":"3 層保護", "type":"table",
     "rows":[["層級","怎麼做"],
             ["儲存格鎖定","格式 → 儲存格 → 保護 → 鎖定（預設都鎖定）"],
             ["工作表保護","校閱 → 保護工作表 → 設密碼"],
             ["活頁簿保護","校閱 → 保護活頁簿（防止新增/刪除工作表）"],
             ["檔案密碼","檔案 → 密碼（開啟才需要密碼）"]]},
    {"h":"常見流程", "type":"p",
     "text":"① 先選「使用者可以填寫的儲存格」→ 取消鎖定 ② 再保護工作表 → 這樣使用者只能改沒鎖的格子。"},
    {"h":"⚠️ 重要", "type":"warn",
     "text":"Excel 的保護不是真的加密，只是「防呆」。要真正加密請用「檔案密碼」。"}
  ],
  "tasks":["把一個儲存格設為不鎖定","保護一張工作表","設一個檔案密碼","知道保護不等於加密"],
},

"P4-01": {
  "tldr": "動態陣列函數：M365 的革命，一個公式回傳整個範圍",
  "intro": "這是 Excel 近 30 年最大的功能升級。會了之後再也回不去傳統公式了。",
  "blocks": [
    {"h":"6 個必學動態函數", "type":"code","lang":"excel",
     "code":"=FILTER(A2:C100, B2:B100=\"電子\")\n# → 篩選出電子部所有列（不用樞紐）\n\n=SORT(A2:C100, 3, -1)\n# → 依第 3 欄降序排序\n\n=UNIQUE(A2:A100)\n# → 取出不重複值\n\n=SEQUENCE(10)\n# → 生成 1~10 的序號\n\n=SEQUENCE(12, 1, 1, 1)\n# → 12 個月\n\n=SORT(UNIQUE(FILTER(A2:A100, B2:B100>1000)))\n# → 業績破千的人去重後排序，一行搞定！"},
    {"h":"溢出範圍 #", "type":"p",
     "text":"動態函數的結果會「溢出」到下方儲存格。寫 D2# 表示「從 D2 開始的整個溢出範圍」。"},
    {"h":"SUMPRODUCT（老牌動態）", "type":"code","lang":"excel",
     "code":"# 沒有動態陣列也能玩\n=SUMPRODUCT((A:A=\"電子\")*(B:B>1000)*C:C)\n# → 電子部且金額>1000 的加總"},
    {"h":"常見錯誤", "type":"warn",
     "text":"FILTER/SORT/UNIQUE 只在 Microsoft 365 / Excel 2021 以後才有。給舊版用會 #NAME?。"}
  ],
  "tasks":["寫一個 FILTER","寫一個 UNIQUE","用 SORT 排序","知道溢出範圍 # 是什麼"],
},

"P4-02": {
  "tldr": "進階函數：解決傳統函數做不到的場景",
  "intro": "這些函數平常用不到，但碰到時你會慶幸自己學過。",
  "blocks": [
    {"h":"重點函數", "type":"code","lang":"excel",
     "code":"=LET(x, A2*0.05, y, B2+x, y*1.1)\n# 命名變數，避免重複計算同一個東西\n\n=INDIRECT(\"工作表\" & A2 & \"!B5\")\n# 用文字組出儲存格參照（動態跨表查詢）\n\n=OFFSET(A1, 5, 2, 1, 1)\n# 從 A1 偏移 5 列 2 欄 → C6\n\n=AGGREGATE(9, 7, A:A)\n# 強化版 SUM，可忽略隱藏列、錯誤值\n\n=SWITCH(A2, \"A\", 100, \"B\", 80, \"C\", 60, 0)\n# 比巢狀 IF 好讀的多分支\n\n=LAMBDA(x, y, x^2 + y^2)(3, 4)\n# 自訂函數（M365 only），可放在「名稱管理員」變成可重用函式"},
    {"h":"使用建議", "type":"p",
     "text":"LET 是現在最被推薦的「公式重構」工具——把 =A1*0.05 + (A1*0.05)*1.1 改成 =LET(t, A1*0.05, t + t*1.1) 容易讀又只算一次。"}
  ],
  "tasks":["寫一個 LET","用 SWITCH 取代巢狀 IF","知道 INDIRECT 的用法","知道 LAMBDA 是什麼"],
},

"P4-03": {
  "tldr": "假設分析：給結果，反推輸入",
  "intro": "「我想要利潤 100 萬，售價該訂多少？」這種問題就是假設分析的舞台。",
  "blocks": [
    {"h":"4 種工具", "type":"table",
     "rows":[["工具","適合"],
             ["目標搜尋（Goal Seek）","只有 1 個變數要求解"],
             ["模擬分析表（Data Table）","看 1~2 個變數變化的結果矩陣"],
             ["分析藍本（Scenario Manager）","比較幾個情境（樂觀/悲觀/正常）"],
             ["規劃求解（Solver）","多變數+條件約束的最佳化"]]},
    {"h":"目標搜尋範例", "type":"p",
     "text":"資料 → 模擬分析 → 目標搜尋。「將儲存格 B5（總利潤）設為 1000000，藉由變更儲存格 B2（售價）」→ Excel 會自動算出 B2 該是多少。"},
    {"h":"規劃求解（Solver）", "type":"p",
     "text":"macOS 要先到「工具 → Excel 增益集 → Solver」啟用。可以設目標、可變儲存格、限制式。適合做產能規劃、運送路線等最佳化問題。"}
  ],
  "tasks":["用目標搜尋解一個方程式","建立一個模擬分析表","知道分析藍本和規劃求解的差別","在 Mac 啟用 Solver"],
},

"P4-04": {
  "tldr": "Power Query：Excel 內建的 ETL 神器",
  "intro": "再也不用每月手動清資料！Power Query 把整理流程錄起來，下個月一鍵重跑。",
  "blocks": [
    {"h":"什麼是 ETL", "type":"p",
     "text":"Extract（取資料）→ Transform（清資料）→ Load（載入 Excel）。Power Query 把這三步全部視覺化，不用寫程式碼。"},
    {"h":"常見操作", "type":"p",
     "text":"資料 → 取得資料 → 從檔案/資料庫/網頁。匯入後進入 Power Query 編輯器：可以拆欄、合併欄、移除列、轉置、樞紐/反樞紐、合併查詢（類似 JOIN）、附加查詢（類似 UNION）。"},
    {"h":"M 語言", "type":"code","lang":"powerquery",
     "code":"let\n    來源 = Excel.CurrentWorkbook(){[Name=\"訂單\"]}[Content],\n    篩選 = Table.SelectRows(來源, each [金額] > 1000),\n    排序 = Table.Sort(篩選, {{\"金額\", Order.Descending}})\nin\n    排序"},
    {"h":"為什麼要用", "type":"p",
     "text":"傳統複製貼上的工作流程：每個月花 2 小時。用 Power Query 第一次設定 30 分鐘，之後每月「重新整理」按一下，10 秒搞定。"}
  ],
  "tasks":["從 CSV 匯入一份資料","練習移除欄、篩選列","知道 ETL 是什麼","知道 M 語言長什麼樣"],
},

"P4-05": {
  "tldr": "Power Pivot + DAX：把 Excel 變成迷你 BI 工具",
  "intro": "資料超過樞紐處理上限（約百萬列）？要跨多張表分析？這時 Power Pivot 就出場了。",
  "blocks": [
    {"h":"資料模型", "type":"p",
     "text":"傳統樞紐只能用「一張表」。資料模型可以匯入多張表 → 設定關聯（像資料庫的 JOIN）→ 建立樞紐時跨表分析。"},
    {"h":"DAX 基礎", "type":"code","lang":"dax",
     "code":"# 度量值（Measures）= 公式語言版的計算欄位\n總業績 := SUM(訂單[金額])\n\n電子部業績 :=\n    CALCULATE(\n        SUM(訂單[金額]),\n        部門[部門名] = \"電子\"\n    )\n\n# 同期比較\n去年同期 :=\n    CALCULATE([總業績], SAMEPERIODLASTYEAR('日期'[日期]))"},
    {"h":"和 Power Query 的差別", "type":"p",
     "text":"Power Query：清資料、整形資料；Power Pivot：建模型、寫 DAX、做分析。一個負責「清」，一個負責「算」，常常一起用。"},
    {"h":"⚠️ macOS 注意", "type":"warn",
     "text":"Power Pivot 的資料模型 macOS 版支援有限，部分功能要 Windows 版。Power Query 在 Mac 已經完整支援。"}
  ],
  "tasks":["知道資料模型是什麼","知道度量值（Measure）跟計算欄位的差","認識 CALCULATE 這個 DAX 核心","了解 Mac 版的限制"],
},

"P5-01": {
  "tldr": "VBA 基礎：用程式碼自動化重複工作",
  "intro": "VBA 是 Excel 的內建程式語言。會寫一點 VBA，等於有個 24 小時不抱怨的助理。",
  "blocks": [
    {"h":"⚠️ macOS 限制", "type":"warn",
     "text":"Mac 沒有 Alt+F11，要用「工具 → 巨集 → Visual Basic 編輯器」進入。Mac 版不支援 UserForm（自訂表單），其他基礎功能都能用。"},
    {"h":"第一個巨集", "type":"code","lang":"vba",
     "code":"Sub HelloWorld()\n    MsgBox \"哈囉，VBA！\"\nEnd Sub\n\nSub 寫入儲存格()\n    Range(\"A1\").Value = \"我是 VBA 寫的\"\n    Range(\"A2\").Value = 100\n    Range(\"A3\").Formula = \"=A2*1.05\"\nEnd Sub"},
    {"h":"變數與迴圈", "type":"code","lang":"vba",
     "code":"Sub 加總範例()\n    Dim i As Integer\n    Dim total As Double\n    total = 0\n    \n    For i = 2 To 100\n        If Cells(i, 2).Value > 0 Then\n            total = total + Cells(i, 2).Value\n        End If\n    Next i\n    \n    MsgBox \"加總：\" & total\nEnd Sub"},
    {"h":"執行巨集", "type":"p",
     "text":"按 F5 執行；或回 Excel → 工具 → 巨集 → 選一個來執行；或建立按鈕綁定巨集。檔案要存成 .xlsm（啟用巨集的活頁簿）。"}
  ],
  "tasks":["在 Mac 開啟 VB 編輯器","寫一個 MsgBox","寫一個 For 迴圈","存成 .xlsm 格式"],
},

"P5-02": {
  "tldr": "VBA 進階：陣列、字典、錯誤處理、事件",
  "intro": "進階 VBA 的精髓在「速度」——把 100 個操作壓成 1 個記憶體運算。",
  "blocks": [
    {"h":"陣列加速", "type":"code","lang":"vba",
     "code":"' 慢：直接讀寫儲存格（每次都跨應用程式）\nFor i = 2 To 10000\n    Cells(i, 2).Value = Cells(i, 1).Value * 2\nNext i\n\n' 快：先讀進陣列，運算完一次寫回（快 100 倍）\nDim arr As Variant\narr = Range(\"A2:A10000\").Value\nFor i = 1 To UBound(arr)\n    arr(i, 1) = arr(i, 1) * 2\nNext i\nRange(\"B2:B10000\").Value = arr"},
    {"h":"Dictionary（去重 / 計數神器）", "type":"code","lang":"vba",
     "code":"Dim dict As Object\nSet dict = CreateObject(\"Scripting.Dictionary\")\n\nFor i = 2 To 1000\n    Dim key As String\n    key = Cells(i, 1).Value\n    If dict.Exists(key) Then\n        dict(key) = dict(key) + 1\n    Else\n        dict.Add key, 1\n    End If\nNext i"},
    {"h":"錯誤處理", "type":"code","lang":"vba",
     "code":"Sub Safe()\n    On Error GoTo ErrHandler\n    ' 你的程式碼\n    Exit Sub\nErrHandler:\n    MsgBox \"錯誤：\" & Err.Description\nEnd Sub"},
    {"h":"事件驅動", "type":"p",
     "text":"工作表可以掛事件：Worksheet_Change（資料變更時觸發）、Worksheet_SelectionChange（選取變更時）。可以做「自動填入時間戳記」之類的功能。"}
  ],
  "tasks":["用陣列加速一段 VBA","用 Dictionary 去重","加上 On Error 錯誤處理","知道事件驅動是什麼"],
},

"P5-03": {
  "tldr": "VBA 實戰：自動產生月報、實務模板",
  "intro": "把前面學的全部組合起來，做一個真正能用的小工具。",
  "blocks": [
    {"h":"範例：自動清資料 + 產樞紐 + 寄報表", "type":"code","lang":"vba",
     "code":"Sub 月報自動化()\n    Application.ScreenUpdating = False  '關閉螢幕更新加速\n    \n    ' 1. 開啟原始資料\n    Workbooks.Open \"~/Desktop/原始資料.xlsx\"\n    \n    ' 2. 清空白列\n    Columns(\"A\").SpecialCells(xlCellTypeBlanks).EntireRow.Delete\n    \n    ' 3. 套用篩選\n    Range(\"A1\").AutoFilter Field:=2, Criteria1:=\">1000\"\n    \n    ' 4. 複製到月報\n    ' ... 你的處理邏輯\n    \n    ' 5. 儲存\n    ActiveWorkbook.SaveAs \"~/Desktop/月報_\" & Format(Date, \"yyyymm\") & \".xlsx\"\n    \n    Application.ScreenUpdating = True\n    MsgBox \"完成！\"\nEnd Sub"},
    {"h":"加速三件事", "type":"p",
     "text":"① Application.ScreenUpdating = False（關螢幕更新）② Application.Calculation = xlCalculationManual（暫停自動計算）③ 用陣列代替逐格操作。"},
    {"h":"按鈕綁巨集", "type":"p",
     "text":"開發人員 → 插入 → 按鈕 → 拖出來 → 指定巨集。使用者就能一鍵執行。"}
  ],
  "tasks":["寫一個自動化流程","加上 ScreenUpdating=False","建立一個按鈕綁巨集","存成 .xlsm"],
},

"P5-04": {
  "tldr": "綜合挑戰：跨章節整合題目練習",
  "intro": "把學過的全部用上！這些題目模擬真實工作場景，每題都會用到 3+ 個函數。",
  "blocks": [
    {"h":"挑戰 1：年資與獎金", "type":"p",
     "text":"員工表有「入職日期」和「年薪」兩欄。要算出：① 滿幾年（DATEDIF）② 年資 ≥ 3 年的人 ③ 每人應發獎金（IFS 分級）④ 全公司獎金加總（SUMIFS）。"},
    {"h":"挑戰 2：跨表查詢", "type":"p",
     "text":"訂單表只有「商品代號」，商品表有「代號→名稱→價格」。要：① 用 XLOOKUP 帶出商品名 ② 帶出單價 ③ 算小計 ④ 用樞紐做月報。"},
    {"h":"挑戰 3：動態儀表板", "type":"code","lang":"excel",
     "code":"# 給定銷售明細，做一個 Dashboard：\n=UNIQUE(銷售[業務員])           # 列出所有業務員\n=SUMIFS(銷售[金額], 銷售[業務員], A2)  # 個人業績\n=RANK(B2, B:B)                    # 排名\n=LARGE(銷售[金額], 1)             # 最高一筆\n# 加上條件式格式 + 走勢圖 = 完整儀表板"},
    {"h":"挑戰 4：自動化清單", "type":"p",
     "text":"每月收到 N 個分公司的 .csv → 用 Power Query 合併 → 用樞紐分析 → 用 VBA 一鍵更新 + 寄信。這是一個完整的「自動化報表系統」。"}
  ],
  "tasks":["完成挑戰 1（年資）","完成挑戰 2（跨表）","完成挑戰 3（儀表板）","挑戰 4 做一部分"],
},
}
