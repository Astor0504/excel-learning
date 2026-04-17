"""一次性腳本：為 lessons/*.html 補 meta description + og: + twitter:card。

從 index.html 的卡片抽出每課描述（已驗證映射），插在 <title> 後面。
冪等：若該頁已含 og:title 則跳過。
"""
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent.parent
LESSONS = ROOT / "lessons"
SITE = "Excel 學習地圖"

DESCRIPTIONS = {
    "P1-01": ("macOS 快捷鍵大全", "掌握 macOS 版 Excel 的核心快捷鍵，速度比滑鼠快 10 倍，含 WPS Office 對照"),
    "P1-02": ("基礎函式", "7 個最常用的統計函數（SUM/AVERAGE/COUNT/MAX/MIN/LARGE/SMALL），幾乎所有報表都從這裡開始"),
    "P1-03": ("條件判斷", "IF / IFS / SWITCH：Excel 的條件邏輯，學會這課才能寫複雜判斷"),
    "P2-01": ("條件統計", "SUMIF / COUNTIF / AVERAGEIF 系列：在條件下才加總/計算的「會挑食」函數"),
    "P2-02": ("查找比對", "VLOOKUP / XLOOKUP / INDEX+MATCH：跨表查資料的命脈函數，職場最常被要求的技能"),
    "P2-03": ("樞紐分析表", "樞紐分析表（Pivot Table）：5 秒做出 30 分鐘的報表，資料分析必備"),
    "P2-04": ("條件式格式化", "條件式格式化：讓重要的資料自動跳出來，色階、資料橫條、自訂規則"),
    "P3-01": ("資料驗證", "資料驗證（Data Validation）：讓使用者只能輸入合法的值，做出防呆表單"),
    "P3-02": ("文字與日期函式", "TEXT / LEFT / RIGHT / DATE / EOMONTH 等：清資料最常用的工具組"),
    "P3-03": ("圖表設計", "Excel 圖表設計：選對類型比畫得漂亮重要，柱狀/折線/圓餅/組合圖實戰"),
    "P3-04": ("命名範圍與表格", "命名範圍（Named Range）+ 表格（Table）：讓公式變得「會說人話」、自動擴張"),
    "P3-05": ("保護與安全", "保護工作表 / 活頁簿：避免別人亂改你的公式，搭配密碼與儲存格鎖定"),
    "P4-01": ("陣列與動態函式", "動態陣列函數（FILTER / SORT / UNIQUE / SEQUENCE）：M365 的革命，一個公式回傳整個範圍"),
    "P4-02": ("進階函式", "LAMBDA / LET / BYROW / REDUCE 等進階函數：解決傳統函數做不到的場景"),
    "P4-03": ("假設分析", "目標搜尋 / 規劃求解 / 分析藍本：給結果，反推輸入"),
    "P4-04": ("Power Query", "Power Query：Excel 內建的 ETL 神器，自動清資料、合併多檔、轉換格式"),
    "P4-05": ("Power Pivot 與資料模型", "Power Pivot + DAX：把 Excel 變成迷你 BI 工具，跨表關聯與計算量值"),
    "P5-01": ("VBA 基礎", "VBA 基礎：用程式碼自動化重複工作，巨集錄製、VBE、變數與迴圈入門"),
    "P5-02": ("VBA 進階", "VBA 進階：陣列、字典（Dictionary）、錯誤處理（On Error）、事件（Workbook/Worksheet Events）"),
    "P5-03": ("VBA 實戰專案", "VBA 實戰：自動產生月報、批次處理檔案、建立可複用的實務模板"),
    "P5-04": ("綜合挑戰", "綜合挑戰：跨章節整合題目練習，驗收 Excel 學習地圖 21 課的學習成果"),
}

META_TEMPLATE = (
    '<meta name="description" content="{desc}">\n'
    '<meta property="og:type" content="article">\n'
    '<meta property="og:title" content="{title}｜{site}">\n'
    '<meta property="og:description" content="{desc}">\n'
    '<meta property="og:locale" content="zh_TW">\n'
    '<meta name="twitter:card" content="summary">'
)

TITLE_RE = re.compile(r"<title>([^｜<]+)｜[^<]+</title>")


def process(path: Path) -> str:
    slug = path.stem  # e.g. "P1-01"
    if slug not in DESCRIPTIONS:
        return f"SKIP {slug} (no mapping)"

    text = path.read_text(encoding="utf-8")
    if 'property="og:title"' in text:
        return f"SKIP {slug} (already has og:)"

    expected_title, desc = DESCRIPTIONS[slug]
    m = TITLE_RE.search(text)
    if not m:
        return f"FAIL {slug} (title not found)"
    if m.group(1).strip() != expected_title:
        return f"FAIL {slug} (title mismatch: {m.group(1)!r} vs {expected_title!r})"

    meta_block = META_TEMPLATE.format(desc=desc, title=expected_title, site=SITE)
    new_text = text.replace(m.group(0), m.group(0) + "\n" + meta_block, 1)
    path.write_text(new_text, encoding="utf-8")
    return f"OK   {slug}"


if __name__ == "__main__":
    files = sorted(LESSONS.glob("P*.html"))
    for p in files:
        print(process(p))
    print(f"\n{len(files)} files processed.")
