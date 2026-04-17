
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

// Shared HTML escape — used by AI chat and search
function escHtml(s){ return (s||"").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c])); }

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
}
themeBtn?.addEventListener("click", () => {
  const cur = document.documentElement.dataset.theme === "dark" ? "" : "dark";
  document.documentElement.dataset.theme = cur;
  localStorage.setItem("theme", cur);
  syncTheme();
});
syncTheme();

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
    });
  });
  updateLessonProgress();
  markPageDone();
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
      <button type="button" class="btn lesson-kickoff-btn">開始任務</button>
    </div>
    <div class="lesson-kickoff-list">
      ${tasks.map((task, i) => `
        <div class="lesson-kickoff-item">
          <span class="lesson-kickoff-num">0${i + 1}</span>
          <span class="lesson-kickoff-text">${escHtml(task)}</span>
        </div>
      `).join("")}
    </div>
  `;

  box.querySelector(".lesson-kickoff-btn")?.addEventListener("click", () => {
    checklist.scrollIntoView({ behavior: "smooth", block: "start" });
  });

  const anchor = lesson.querySelector(".tldr") || lesson.querySelector(".progress-label");
  if (anchor) anchor.insertAdjacentElement("afterend", box);
})();

// Lesson compass
(function(){
  const lesson = document.querySelector(".lesson");
  if (!lesson || lesson.querySelector(".lesson-compass")) return;

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

  const anchor = lesson.querySelector(".lesson-kickoff") || lesson.querySelector(".tldr");
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
    if (weak.length) {
      pick = weak[0];
      prefix = "建議先複習";
    } else {
      const undone = idx.filter(e => !doneSet.has(e.u));
      if (undone.length) {
        pick = undone[0];
        prefix = "建議下一課";
      }
    }
    if (!pick) {
      helper.textContent = "目前進度很完整，去做一題快速測驗或複習筆記也不錯。";
      return;
    }
    helper.innerHTML = `${prefix}：<strong>${escHtml(pick.t)}</strong><span>${escHtml(pick.b || "")}</span>`;
  }
  updateHeroHelper();

  (function buildQuickNav(){
    if (!quickNav) return;
    const lastLesson = sessionStorage.getItem("last-lesson");
    const groups = [
      { key: "all", label: "全部", phases: [] },
      { key: "starter", label: "新手入門", phases: [1, 2] },
      { key: "work", label: "職場即用", phases: [2, 3] },
      { key: "auto", label: "自動化進階", phases: [4, 5] },
    ];
    const state = { active: "all" };

    const wrap = document.createElement("div");
    wrap.className = "hero-mode-switch";
    const title = document.createElement("div");
    title.className = "hero-mode-title";
    title.textContent = "學習捷徑";
    wrap.appendChild(title);

    const chips = document.createElement("div");
    chips.className = "hero-mode-chips";
    wrap.appendChild(chips);

    const continueWrap = document.createElement("div");
    continueWrap.className = "hero-continue";

    function applyMode(modeKey){
      state.active = modeKey;
      chips.querySelectorAll("button").forEach(function(btn){
        btn.classList.toggle("is-active", btn.dataset.mode === modeKey);
      });
      const selected = groups.find(g => g.key === modeKey) || groups[0];
      const allowed = new Set(selected.phases);
      document.querySelectorAll(".phase-section").forEach(function(section){
        const phaseMatch = section.className.match(/phase-(\d)/);
        const phaseNo = phaseMatch ? parseInt(phaseMatch[1], 10) : 0;
        const visible = !allowed.size || allowed.has(phaseNo);
        section.style.display = visible ? "" : "none";
        var next = section.nextElementSibling;
        while (next && !next.classList.contains("phase-section") && !next.classList.contains("phase-group")) {
          next.style.display = visible ? "" : "none";
          next = next.nextElementSibling;
        }
        const group = section.nextElementSibling;
        if (group && group.classList.contains("phase-group")) {
          group.style.display = visible ? "" : "none";
        }
      });
    }

    groups.forEach(function(group){
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "hero-mode-chip";
      btn.dataset.mode = group.key;
      btn.textContent = group.label;
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
  fab.id = 'ttsFab'; fab.title = '朗讀工具'; fab.textContent = '🔊';
  fab.onclick = () => panel.classList.toggle('open');
  document.body.appendChild(fab);

  window.addEventListener('beforeunload', () => synth.cancel());

  if (document.readyState === 'loading')
    document.addEventListener('DOMContentLoaded', injectButtons);
  else injectButtons();

  return { speak: speakChunks, splitText, stop: stopAll };
})();
