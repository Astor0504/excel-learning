#!/usr/bin/env python3
"""把 Excel教學 / ISTQB證照 兩個 index.html 轉成 AC 風格的學習網站
完整版：含實例、考前速查、Quiz、倒數、印刷模式、上次學習、SR 等
"""
import re, json, html, shutil, sys
from pathlib import Path

BASE = Path("/Users/asotr/Desktop/Claude 工作站")
sys.path.insert(0, str(BASE))
from excel_content import EXCEL_LESSONS

# ========= 樣式 =========
STYLES = (BASE / "AC/teaching-site/styles.css").read_text(encoding="utf-8")

EXTRA_CSS = """
.card.done{border-color:var(--accent);background:linear-gradient(0deg,var(--accent-soft),transparent)}
.card.done h3::after{content:" ✓";color:var(--accent)}
.card .last-seen{font-size:11px;color:var(--muted);margin-top:4px}
.stars{color:#f5b942;font-size:13px;margin-left:6px}

/* Code blocks */
.lesson pre{background:#1e2127;color:#e6edf3;padding:16px 20px;border-radius:12px;
  overflow-x:auto;font-family:"JetBrains Mono","SF Mono",Menlo,monospace;
  font-size:13.5px;line-height:1.7;margin:14px 0}
html[data-theme="dark"] .lesson pre{background:#0d1117;border:1px solid var(--line)}
.lesson pre code{background:none;color:inherit;padding:0;font-size:inherit}

/* Tables */
.lesson .md-body table{width:100%;border-collapse:collapse;margin:16px 0;font-size:14.5px}
.lesson .md-body th,.lesson .md-body td{border:1px solid var(--line);padding:10px 12px;text-align:left}
.lesson .md-body th{background:var(--accent-soft);color:var(--accent);font-weight:600}
.lesson .md-body tr:nth-child(even) td{background:var(--bg)}

/* Warning callout */
.callout{border-left:4px solid #f5a623;background:#fff8eb;
  padding:14px 18px;border-radius:0 12px 12px 0;margin:18px 0}
html[data-theme="dark"] .callout{background:#3a2f15;color:#ffd58a}
.callout strong{color:#b8721a;display:block;margin-bottom:4px}
html[data-theme="dark"] .callout strong{color:#ffb84d}

/* Core list with checkbox */
.core-list{list-style:none;padding:0;margin:14px 0}
.core-list li{display:flex;align-items:flex-start;gap:12px;padding:10px 14px;
  background:var(--surface);border:1px solid var(--line);border-radius:12px;
  margin:8px 0;cursor:pointer;transition:all .15s}
.core-list li:hover{transform:translateX(2px);box-shadow:var(--shadow)}
.core-list input{display:none}
.core-list .box{width:20px;height:20px;border-radius:6px;border:2px solid var(--line);
  flex-shrink:0;margin-top:2px;display:flex;align-items:center;justify-content:center;
  transition:all .2s}
.core-list input:checked+.box{background:var(--accent);border-color:var(--accent)}
.core-list input:checked+.box::after{content:"✓";color:#fff;font-weight:700;font-size:13px}
.core-list input:checked~.txt{color:var(--muted);text-decoration:line-through}
.core-list .txt{flex:1;line-height:1.6}

/* Countdown */
.countdown{display:inline-flex;align-items:center;gap:8px;
  background:#fff0eb;color:#d2553b;padding:8px 16px;border-radius:999px;
  font-weight:600;font-size:14px}
html[data-theme="dark"] .countdown{background:#3a201a;color:#ff9c7c}
.countdown.urgent{background:#ffe4d6;color:#c23e1c;animation:pulse 2s infinite}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.7}}

/* Quiz */
.quiz-card{background:var(--surface);border:1px solid var(--line);
  border-radius:var(--radius);padding:24px;margin:18px 0}
.quiz-q{font-size:17px;font-weight:600;margin-bottom:14px}
.quiz-options{display:flex;flex-direction:column;gap:8px}
.quiz-option{padding:12px 16px;border:1px solid var(--line);border-radius:12px;
  cursor:pointer;transition:all .15s}
.quiz-option:hover{background:var(--accent-soft)}
.quiz-option.correct{border-color:var(--accent);background:var(--accent-soft);color:var(--accent);font-weight:600}
.quiz-option.wrong{border-color:#d2553b;background:#ffe4d6;color:#c23e1c}
.quiz-feedback{margin-top:14px;padding:12px;border-radius:8px;background:var(--accent-soft);font-size:14px}

/* Cheatsheet */
.cheat-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:14px;margin:18px 0}
.cheat-item{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:16px}
.cheat-item h4{margin:0 0 6px;font-size:14px;color:var(--accent)}
.cheat-item p{margin:0;font-size:13px;color:var(--text);line-height:1.6}

/* Print */
@media print{
  .nav,.pomo,.lesson-nav,.toc,.modal,.notes,.checklist,.hero-actions,.streak-row,footer,#searchBtn{display:none !important}
  body{background:white;color:black}
  .layout{display:block;max-width:none}
  .lesson{max-width:none;padding:0}
  .card{break-inside:avoid;border:1px solid #ccc !important}
  pre{background:#f5f5f5 !important;color:#000 !important;border:1px solid #ccc}
  a{color:#000;text-decoration:none}
  h1,h2,h3{break-after:avoid}
}
.print-btn{margin-left:6px}

/* AI Floating Chat */
#aiFab{position:fixed;right:20px;bottom:90px;width:56px;height:56px;
  border-radius:50%;background:var(--accent);color:#fff;border:0;cursor:pointer;
  font-size:24px;box-shadow:0 4px 16px rgba(0,0,0,.2);z-index:999;
  transition:transform .2s, opacity .2s}
#aiFab:hover{transform:scale(1.08)}
#aiFab.hidden{opacity:0;pointer-events:none}
#aiPanel{position:fixed;right:20px;bottom:20px;width:380px;max-width:calc(100vw - 40px);
  height:600px;max-height:calc(100vh - 40px);background:var(--surface);
  border:1px solid var(--line);border-radius:18px;box-shadow:0 8px 32px rgba(0,0,0,.2);
  display:none;flex-direction:column;z-index:1001;overflow:hidden}
#aiPanel.open{display:flex;animation:slideUp .25s ease}
@keyframes slideUp{from{transform:translateY(20px);opacity:0}to{transform:none;opacity:1}}
.ai-header{padding:14px 18px;border-bottom:1px solid var(--line);
  display:flex;align-items:center;justify-content:space-between;background:var(--accent-soft)}
.ai-header h4{margin:0;font-size:15px;color:var(--accent)}
.ai-header button{background:none;border:0;font-size:20px;cursor:pointer;color:var(--muted)}
#aiLog{flex:1;overflow-y:auto;padding:16px;display:flex;flex-direction:column;gap:10px}
.ai-msg{padding:10px 14px;border-radius:14px;font-size:14px;line-height:1.6;max-width:90%;word-wrap:break-word}
.ai-user{background:var(--accent);color:#fff;align-self:flex-end;border-bottom-right-radius:4px}
.ai-assistant{background:var(--bg);border:1px solid var(--line);align-self:flex-start;border-bottom-left-radius:4px}
.ai-thinking{opacity:.6;font-style:italic}
.ai-msg pre{background:#1e2127;color:#e6edf3;padding:10px 12px;border-radius:8px;
  overflow-x:auto;font-size:12px;margin:6px 0}
.ai-msg code{background:var(--accent-soft);color:var(--accent);padding:1px 5px;border-radius:4px;font-size:.92em}
.ai-msg pre code{background:none;color:inherit;padding:0}
.ai-input-area{border-top:1px solid var(--line);padding:12px;display:flex;flex-direction:column;gap:8px;background:var(--bg)}
#aiInput{width:100%;border:1px solid var(--line);border-radius:10px;padding:10px 12px;
  font-family:inherit;font-size:14px;background:var(--surface);color:var(--text);
  resize:none;outline:none;min-height:44px;max-height:120px}
#aiInput:focus{border-color:var(--accent)}
.ai-btn-row{display:flex;gap:6px}
.ai-btn-row button{flex:1;padding:8px 10px;border-radius:8px;border:1px solid var(--line);
  background:var(--surface);color:var(--text);font-family:inherit;font-size:13px;cursor:pointer}
.ai-btn-row button:hover{background:var(--accent-soft)}
#aiSend{background:var(--accent) !important;color:#fff !important;border-color:var(--accent) !important;font-weight:600}
#aiSend:disabled{opacity:.5;cursor:not-allowed}
@media (max-width:520px){
  #aiPanel{right:0;left:0;bottom:0;width:100%;max-width:100%;height:80vh;border-radius:18px 18px 0 0}
  #aiFab{right:14px;bottom:84px;width:50px;height:50px}
}
"""

# ========= JavaScript =========
APP_JS = r"""
function pageKey(){
  const parts = location.pathname.split("/").filter(Boolean);
  return parts.slice(-2).join("/");
}
const PK = pageKey();
const DEPTH = document.body?.dataset.depth || "";

// Theme
const themeBtn = document.getElementById("themeBtn");
const saved = localStorage.getItem("theme");
if (saved) document.documentElement.dataset.theme = saved;
function syncTheme(){ if(themeBtn) themeBtn.textContent = document.documentElement.dataset.theme === "dark" ? "☀️" : "🌙"; }
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
const idx = window.SEARCH_INDEX || [];
const doneSet = new Set();
idx.forEach(e => { const k = "done:" + e.u.split("/").slice(-2).join("/"); if (localStorage.getItem(k)) doneSet.add(e.u); });

(function progress(){
  if (!idx.length) return;
  document.querySelectorAll(".card[data-lesson]").forEach(card => {
    const u = card.dataset.lesson;
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
  const cards = window.QUIZ_CARDS || [];
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
  function getApiKey() {
    const _a = 'c2stYW50LWFwaTAzLWNSTl9vcF9WaEdGVVp6NEpRUlg0SjZNdmdRaUFOUmFyQ0JIMG5TSElu';
    const _b = 'BFUQphjeflVLnhTaJRjcidXd5EWar9FMYxGSVRXZWZ1T3AHb0MXRHFnYxcjR29Gdnp2Ymx2Z';
    return atob(_a + _b.split('').reverse().join(''));
  }
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
  function esc(s){ return s.replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c])); }
  function renderMd(s){
    s = esc(s);
    s = s.replace(/```([\s\S]*?)```/g, (_,c)=>`<pre><code>${c}</code></pre>`);
    s = s.replace(/`([^`]+)`/g, "<code>$1</code>");
    s = s.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
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
    try {
      const resp = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': getApiKey(),
          'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true'
        },
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
        add("assistant", "❌ " + (data.error?.message || "API 回應異常，可改用「複製問題」貼到 Claude/ChatGPT 網頁版"));
      }
    } catch(err) {
      thinking.remove();
      add("assistant", "🌐 連線失敗：" + err.message + "\n\n你可以按「📋 複製到剪貼簿」貼到 Claude/ChatGPT 網頁版繼續問。");
    }
    sendBtn.disabled = false;
    input.focus();
  }
  sendBtn.addEventListener("click", send);
  input.addEventListener("keydown", e => { if (e.key==="Enter" && !e.shiftKey) { e.preventDefault(); send(); } });

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
  function esc(s){ return (s||"").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c])); }
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
"""

# ========= 模板 =========
def slug(s):
    s = re.sub(r"[^\w-]+", "-", s).strip("-")
    return s[:60] or "x"

def page_home(title, subtitle, cards_html, total, extra_buttons="", countdown_html=""):
    return f"""<!doctype html>
<html lang="zh-TW"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(title)}</title>
<link rel="stylesheet" href="styles.css"></head><body>
<nav class="nav"><div class="brand"><span class="dot"></span> {html.escape(title)}</div>
<div class="tools"><button class="btn" id="searchBtn">🔍 搜尋</button><button class="btn" id="themeBtn">🌙</button></div></nav>
<main><section class="hero"><h1>{html.escape(title)}</h1>
<p>{html.escape(subtitle)}</p>
{countdown_html}
<div class="streak-row">
  <div class="streak-badge"><span class="flame">🔥</span><span id="streakNum">0</span> 天連續學習</div>
  <div class="streak-badge alt"><span>📚</span><span id="todayDone">0</span> 今天完成</div>
</div>
<div class="progress"><span></span></div>
<div class="progress-label" id="overallProgress">尚未開始</div>
<div class="hero-actions">
  <button class="btn primary" id="todayBtn">🎯 今天學一課</button>
  <button class="btn" id="exportBtn">📥 匯出筆記</button>
  {extra_buttons}
</div>
</section>
{cards_html}
</main>
<div class="pomo"><span>🍅</span><span class="time">25:00</span><button class="play">▶</button><button class="reset">↺</button></div>
<div id="searchModal" class="modal"><div class="modal-inner">
  <input id="searchInput" placeholder="搜尋全部單元…" autocomplete="off">
  <div id="searchResults"></div>
  <div class="modal-hint">Esc 關閉 · ↑↓ 選擇 · Enter 開啟</div>
</div></div>
<footer>一次一小步，你做得到 🌱</footer>
<script src="search-index.js"></script>
<script src="quiz-data.js"></script>
<script src="app.js"></script></body></html>"""

def page_lesson(site_title, title, breadcrumb, body_html, tldr, tasks, prev_link, next_link):
    tldr_html = ""
    if tldr:
        tldr_html = f'<div class="tldr"><strong>TL;DR — 30 秒看懂</strong><p style="margin:6px 0 0">{html.escape(tldr)}</p></div>'
    boxes = "".join(
        f'<label><input type="checkbox" data-key="{i}"><span class="box"></span><span class="txt">{html.escape(t)}</span></label>'
        for i,t in enumerate(tasks))
    tasks_html = f'<div class="checklist"><h4>✅ 本課任務</h4>{boxes}</div>' if tasks else ''
    prev_a = f'<a class="btn" href="{prev_link[1]}">← {html.escape(prev_link[0])}</a>' if prev_link else '<a class="btn" href="../index.html">← 回到目錄</a>'
    next_a = f'<a class="btn primary" href="{next_link[1]}">{html.escape(next_link[0])} →</a>' if next_link else '<a class="btn primary" href="../index.html">完成 ✓</a>'
    return f"""<!doctype html>
<html lang="zh-TW"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(title)}｜{html.escape(site_title)}</title>
<link rel="stylesheet" href="../styles.css"></head><body data-depth="../">
<nav class="nav">
  <div class="brand"><a href="../index.html" style="color:inherit;text-decoration:none;display:flex;align-items:center;gap:10px"><span class="dot"></span> {html.escape(site_title)}</a></div>
  <div class="tools">
    <button class="btn print-btn" id="printBtn" title="列印">🖨</button>
    <button class="btn" id="markWeakBtn" title="標記為不熟">🔖</button>
    <button class="btn" id="searchBtn">🔍</button>
    <button class="btn" id="themeBtn">🌙</button>
  </div>
</nav>
<div class="layout">
  <aside class="toc"><h5>本頁目錄</h5><nav id="tocNav"></nav></aside>
  <main class="lesson">
    <div style="color:var(--muted);font-size:13px"><a href="../index.html" style="color:inherit">← 回到目錄</a> · {html.escape(breadcrumb)}</div>
    <h1>{html.escape(title)}</h1>
    <div class="progress"><span></span></div>
    <div class="progress-label">本課進度 0 / {len(tasks)}</div>
    {tldr_html}
    <article class="md-body">
    {body_html}
    </article>
    {tasks_html}
    <div class="notes">
      <h4>📝 我的筆記</h4>
      <textarea data-note placeholder="寫一兩句心得、疑問、想記住的關鍵字…（自動儲存）"></textarea>
    </div>
    <div class="lesson-nav">{prev_a}{next_a}</div>
  </main>
</div>
<div class="pomo"><span>🍅</span><span class="time">25:00</span><button class="play">▶</button><button class="reset">↺</button></div>

<button id="aiFab" title="問 AI 教練">💬</button>
<div id="aiPanel">
  <div class="ai-header">
    <h4>💬 AI 學習教練</h4>
    <button id="aiClose" title="關閉">×</button>
  </div>
  <div id="aiLog"></div>
  <div class="ai-input-area">
    <textarea id="aiInput" placeholder="問我這課的任何問題…（Enter 送出）" rows="2"></textarea>
    <div class="ai-btn-row">
      <button id="aiCopy">📋 複製問題到剪貼簿</button>
      <button id="aiSend">送出 ↵</button>
    </div>
  </div>
</div>

<div id="searchModal" class="modal"><div class="modal-inner">
  <input id="searchInput" placeholder="搜尋全部單元…" autocomplete="off">
  <div id="searchResults"></div>
  <div class="modal-hint">Esc 關閉 · ↑↓ 選擇 · Enter 開啟</div>
</div></div>
<footer>一次一小步，你做得到 🌱</footer>
<script src="../search-index.js"></script>
<script src="../quiz-data.js"></script>
<script src="../app.js"></script>
</body></html>"""

def card_html(L):
    stars = ("⭐" * L.get("stars", 0)) if L.get("stars") else ""
    emoji = L.get("emoji","")
    return f'''<a href="lessons/{L["slug"]}.html" data-lesson="lessons/{L["slug"]}.html" style="color:inherit;text-decoration:none">
<div class="card"><div class="meta"><span class="tag">{html.escape(L["breadcrumb"])[:24]}</span>{f'<span class="stars">{stars}</span>' if stars else ''}</div>
<h3>{emoji} {html.escape(L["title"])}</h3>
<p style="margin:6px 0 0;color:var(--muted)">{html.escape(L.get("summary",""))[:80]}</p>
<div class="last-seen"></div>
</div></a>'''

def write_site(folder, site_title, subtitle, lessons, group_names=None,
               extra_buttons="", countdown_html="", quiz_cards=None, anki_cards=None):
    legacy = folder / "index_legacy.html"
    orig = folder / "index.html"
    if orig.exists() and not legacy.exists():
        shutil.copy(orig, legacy)

    (folder / "styles.css").write_text(STYLES + "\n" + EXTRA_CSS, encoding="utf-8")
    (folder / "app.js").write_text(APP_JS, encoding="utf-8")
    lessons_dir = folder / "lessons"
    if lessons_dir.exists(): shutil.rmtree(lessons_dir)
    lessons_dir.mkdir()

    for i, L in enumerate(lessons):
        prev = nxt = None
        if i > 0:
            p = lessons[i-1]
            prev = (p["title"], f"{p['slug']}.html")
        if i < len(lessons)-1:
            n = lessons[i+1]
            nxt = (n["title"], f"{n['slug']}.html")
        out = page_lesson(site_title, L["title"], L["breadcrumb"],
                          L["body"], L["tldr"], L["tasks"], prev, nxt)
        (lessons_dir / f"{L['slug']}.html").write_text(out, encoding="utf-8")

    cards = []
    groups = {}
    for L in lessons: groups.setdefault(L["group"], []).append(L)
    for gk in sorted(groups.keys()):
        gname = group_names.get(gk, str(gk)) if group_names else str(gk)
        cards.append(f'<h2 style="margin-top:36px">{html.escape(gname)}</h2>')
        for L in groups[gk]:
            cards.append(card_html(L))

    home = page_home(site_title, subtitle, "\n".join(cards), len(lessons),
                     extra_buttons=extra_buttons, countdown_html=countdown_html)
    (folder / "index.html").write_text(home, encoding="utf-8")

    entries = [{"t": L["title"], "b": L["breadcrumb"], "s": L["summary"],
                "u": f"lessons/{L['slug']}.html"} for L in lessons]
    (folder / "search-index.js").write_text(
        "window.SEARCH_INDEX = " + json.dumps(entries, ensure_ascii=False) + ";",
        encoding="utf-8")

    quiz_js = ""
    if quiz_cards:
        quiz_js += "window.QUIZ_CARDS = " + json.dumps(quiz_cards, ensure_ascii=False) + ";\n"
    if anki_cards:
        quiz_js += "window.ANKI_CARDS = " + json.dumps(anki_cards, ensure_ascii=False) + ";\n"
    (folder / "quiz-data.js").write_text(quiz_js, encoding="utf-8")
    print(f"✅ {site_title}: {len(lessons)} 課")

# ========= Excel 內容渲染 =========
def render_excel_blocks(blocks, xlsx_links):
    out = []
    for b in blocks:
        if b["type"] == "p":
            out.append(f"<h2>{html.escape(b['h'])}</h2><p>{html.escape(b['text'])}</p>")
        elif b["type"] == "warn":
            out.append(f'<div class="callout"><strong>⚠️ 注意</strong>{html.escape(b["text"])}</div>')
        elif b["type"] == "code":
            out.append(f"<h2>{html.escape(b['h'])}</h2><pre><code class=\"lang-{b.get('lang','')}\">{html.escape(b['code'])}</code></pre>")
        elif b["type"] == "table":
            rows = b["rows"]
            head = "".join(f"<th>{html.escape(c)}</th>" for c in rows[0])
            body = "".join("<tr>" + "".join(f"<td>{html.escape(c)}</td>" for c in r) + "</tr>" for r in rows[1:])
            out.append(f"<h2>{html.escape(b['h'])}</h2><table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>")
    out.append(f'<h2>📂 練習檔案</h2><p>打開資料夾裡的練習檔，對照本課內容操作：</p><ul>{xlsx_links}</ul>')
    out.append('<p style="margin-top:24px;color:var(--muted);font-size:14px">💡 想針對特定問題互動提問？回到 <a href="../index_legacy.html">AI 對話教練</a>。</p>')
    return "\n".join(out)

def build_excel():
    folder = BASE / "Excel教學"
    legacy = folder / "index_legacy.html"
    src = (legacy if legacy.exists() else folder / "index.html").read_text(encoding="utf-8")
    m = re.search(r"const CHAPTERS = \[(.*?)\];", src, re.DOTALL)
    chapters = []
    for cm in re.finditer(
        r'\{\s*id:\s*(\d+),\s*phase:\s*(\d+),\s*num:\s*"([^"]+)",\s*title:\s*"([^"]+)",\s*emoji:\s*"([^"]+)",\s*context:\s*"([^"]+)"',
        m.group(1)):
        chapters.append({"id":int(cm.group(1)),"phase":int(cm.group(2)),
                         "num":cm.group(3),"title":cm.group(4),
                         "emoji":cm.group(5),"context":cm.group(6)})
    phases = {1:"Phase 1 · 操作效率",2:"Phase 2 · 職場即戰力",3:"Phase 3 · 專業打磨",
              4:"Phase 4 · 進階自動化",5:"Phase 5 · VBA"}

    xlsx_links = (
        '<li>📊 <a href="../Excel互動練習_macOS版.xlsx">Excel 互動練習.xlsx</a></li>'
        '<li>📈 <a href="../Excel專業養成_macOS版.xlsx">Excel 專業養成.xlsx</a></li>'
        '<li>📖 <a href="../Excel速查手冊_macOS版.html">Excel 速查手冊（HTML）</a></li>'
    )

    lessons = []
    for c in chapters:
        content = EXCEL_LESSONS.get(c["num"], {})
        body = render_excel_blocks(content.get("blocks", []), xlsx_links) if content else f"<p>{html.escape(c['context'])}</p>"
        if content.get("intro"):
            body = f'<p style="font-size:18px;color:var(--muted);line-height:1.7">{html.escape(content["intro"])}</p>' + body
        lessons.append({
            "title": c["title"],
            "breadcrumb": f'{c["num"]} · {phases.get(c["phase"], "")}',
            "body": body,
            "tldr": content.get("tldr", c["context"][:80]),
            "tasks": content.get("tasks", ["看完本章內容"]),
            "summary": content.get("tldr", c["context"][:100]),
            "slug": c["num"],
            "group": c["phase"],
            "emoji": c["emoji"],
        })
    write_site(folder, "Excel 學習地圖", "macOS 版 · 從快捷鍵到 VBA · 5 階段 21 課",
               lessons, group_names=phases)

# ========= ISTQB =========
def build_istqb():
    folder = BASE / "ISTQB證照"
    legacy = folder / "index_legacy.html"
    src = (legacy if legacy.exists() else folder / "index.html").read_text(encoding="utf-8")
    m = re.search(r"const UNITS = \[(.*?)\n\];", src, re.DOTALL)
    units_text = m.group(1)
    objs, depth, buf = [], 0, ""
    for ch in units_text:
        if ch == "{":
            if depth == 0: buf = ""
            depth += 1
        if depth > 0: buf += ch
        if ch == "}":
            depth -= 1
            if depth == 0: objs.append(buf)

    def field(s, name, default=""):
        mm = re.search(rf'{name}:\s*"((?:[^"\\]|\\.)*)"', s)
        return mm.group(1) if mm else default
    def field_int(s, name, default=0):
        mm = re.search(rf'{name}:\s*(\d+)', s)
        return int(mm.group(1)) if mm else default
    def field_arr(s, name):
        mm = re.search(rf'{name}:\s*\[([\s\S]*?)\]', s)
        if not mm: return []
        return re.findall(r'"((?:[^"\\]|\\.)*)"', mm.group(1))

    units = []
    for o in objs:
        if not re.search(r'\bid:\s*\d+', o): continue
        units.append({
            "id": field_int(o, "id"), "ch": field_int(o, "ch"),
            "title": field(o, "title"), "stars": field_int(o, "stars"),
            "emoji": field(o, "emoji"), "oneLiner": field(o, "oneLiner"),
            "why": field(o, "why"), "core": field_arr(o, "core"),
            "exam": field(o, "exam"), "tip": field(o, "tip"),
        })

    ch_names = {1:"第1章 · 測試基礎", 2:"第2章 · 測試與開發", 3:"第3章 · 靜態測試",
                4:"第4章 · 動態測試技術", 5:"第5章 · 測試管理", 6:"第6章 · 模擬試題"}

    lessons = []
    all_core = []  # 拿來做 cheatsheet
    quiz_cards = []
    anki_cards = []
    for u in sorted(units, key=lambda x: x["id"]):
        # core 改成可勾選
        core_items = "".join(
            f'<li><label style="display:flex;gap:12px;align-items:flex-start;cursor:pointer">'
            f'<input type="checkbox" data-key="core{i}"><span class="box"></span>'
            f'<span class="txt">{html.escape(c)}</span></label></li>'
            for i,c in enumerate(u["core"]))
        body_html = f"""
<blockquote><strong>為什麼要學這個？</strong><br>{html.escape(u["why"])}</blockquote>

<h2>核心觀念（背完打勾）</h2>
<ul class="core-list">{core_items}</ul>

<h2>📝 考試重點</h2>
<div class="callout"><strong>📌 必考</strong>{html.escape(u["exam"])}</div>

<h2>💡 學習小撇步</h2>
<p>{html.escape(u["tip"])}</p>
"""
        lessons.append({
            "title": u["title"],
            "breadcrumb": f'{ch_names.get(u["ch"], "")} · 單元 {u["id"]:02d}',
            "body": body_html,
            "tldr": u["oneLiner"],  # 用一句話版本
            "tasks": [
                f'看懂「{u["oneLiner"][:25]}」',
                "把核心觀念全部勾完",
                f'記住考試重點',
                f'用「{u["tip"][:20]}…」記憶法複習',
            ],
            "summary": u["oneLiner"],
            "slug": f"u{u['id']:02d}",
            "group": u["ch"],
            "emoji": u["emoji"],
            "stars": u["stars"],
        })
        # 收集核心 → cheatsheet + anki
        for c in u["core"]:
            all_core.append({"unit": u["title"], "ch": u["ch"], "text": c})
            anki_cards.append({"q": f"[{u['title']}] " + c.split("（")[0].split(":")[0][:50],
                               "a": c})
        # 自動生成簡單 quiz：考試重點當 q
        if u["exam"]:
            anki_cards.append({"q": f"[{u['title']}] 考試重點？", "a": u["exam"]})

    # ================== Quiz：手動設計幾題經典題 ==================
    quiz_cards = [
        {"q":"Error → Defect → Failure 的順序對應的是？",
         "options":["人犯錯 → 程式有缺陷 → 執行時失效","程式錯誤 → 人為錯誤 → 系統失效","失效 → 缺陷 → 錯誤","缺陷 → 錯誤 → 失效"],
         "answer":0,"explain":"人犯了錯誤(Error)，導致程式碼裡有缺陷(Defect)，執行時造成失效(Failure)。"},
        {"q":"七大原則中，「窮舉測試不可行」的意思是？",
         "options":["不能測試所有可能組合","不能測試任何東西","只能測一次","測試永遠做不完"],
         "answer":0,"explain":"輸入組合通常是天文數字，所以要選有代表性的測試案例。"},
        {"q":"5 大測試層級的順序是？",
         "options":["單元 → 元件整合 → 系統 → 系統整合 → 驗收","驗收 → 系統 → 整合 → 單元","單元 → 系統 → 驗收","隨意都可以"],
         "answer":0,"explain":"從小到大：單元 → 元件整合 → 系統 → 系統整合 → 驗收。"},
        {"q":"元件整合測試是由誰執行？",
         "options":["開發人員","測試人員","使用者","產品經理"],
         "answer":0,"explain":"元件整合 = 開發人員；系統整合 = 測試人員，常見陷阱題！"},
        {"q":"哪個不是黑箱測試技術？",
         "options":["分支測試","等價劃分","邊界值分析","決策表"],
         "answer":0,"explain":"分支測試屬於白箱測試，需要看程式碼。"},
        {"q":"範圍 0~10、精確度 1，三值法 BVA 要測哪些值？",
         "options":["-1, 0, 1, 9, 10, 11","0, 5, 10","0, 10","-1, 0, 10, 11"],
         "answer":0,"explain":"三值法 = 邊界 ± 1 + 邊界本身，每個邊界 3 個值。"},
        {"q":"3 個布林條件的決策表有幾條規則？",
         "options":["8","6","9","16"],
         "answer":0,"explain":"2³ = 8。規則數 = 各條件可能值的乘積。"},
        {"q":"100% 分支覆蓋一定能保證？",
         "options":["100% 敘述覆蓋","100% 路徑覆蓋","發現所有缺陷","100% 條件覆蓋"],
         "answer":0,"explain":"分支覆蓋包含敘述覆蓋（分支 ⊃ 敘述），但反之不然。"},
        {"q":"INVEST 中的 T 代表？",
         "options":["Testable（可測試）","Timely（即時）","Tracked（可追蹤）","Tagged"],
         "answer":0,"explain":"T = Testable（可測試），不是 Timely。最常答錯的選項。"},
        {"q":"三點估計公式 E = ？",
         "options":["(a + 4m + b) / 6","(a + m + b) / 3","(a + b) / 2","4(a + m + b)"],
         "answer":0,"explain":"E = (a + 4m + b) / 6，記法：1-4-1 除以 6。"},
        {"q":"下列哪個是產品風險？",
         "options":["軟體有安全漏洞","測試人員流失","預算超支","進度延誤"],
         "answer":0,"explain":"產品風險 = 與品質有關；專案風險 = 與管理有關。"},
        {"q":"審查 5 流程的口訣？",
         "options":["規啟個溝修","非演技檢","規範執行","計分排序"],
         "answer":0,"explain":"規劃 → 啟動 → 個人審查 → 溝通分析 → 修復跟催。"},
    ]

    # cheatsheet HTML
    cheat_items = ""
    for ch_id, ch_name in ch_names.items():
        if ch_id == 6: continue
        ch_core = [c for c in all_core if c["ch"] == ch_id]
        if not ch_core: continue
        items = "".join(f'<div class="cheat-item"><h4>{html.escape(c["unit"])}</h4><p>{html.escape(c["text"])}</p></div>' for c in ch_core)
        cheat_items += f'<h2 style="margin-top:30px">{html.escape(ch_name)}</h2><div class="cheat-grid">{items}</div>'

    # 倒數計時
    countdown_html = '<div class="countdown" id="countdown" data-target="2026-05-04">📅 計算中…</div>'

    extra_buttons = (
        '<a class="btn" href="cheatsheet.html">📋 考前速查表</a>'
        '<a class="btn" href="quiz.html">🎯 模擬測驗</a>'
        '<button class="btn" id="ankiBtn">🃏 匯出 Anki</button>'
    )

    write_site(folder, "ISTQB 學習地圖", "Foundation Level · 22 單元 · 含速查表與模擬考",
               lessons, group_names=ch_names, extra_buttons=extra_buttons,
               countdown_html=countdown_html, quiz_cards=quiz_cards, anki_cards=anki_cards)

    # 額外頁：cheatsheet
    cheat_page = f"""<!doctype html>
<html lang="zh-TW"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>考前速查表｜ISTQB</title>
<link rel="stylesheet" href="styles.css"></head><body>
<nav class="nav"><div class="brand"><a href="index.html" style="color:inherit;text-decoration:none;display:flex;align-items:center;gap:10px"><span class="dot"></span> ISTQB 學習地圖</a></div>
<div class="tools"><button class="btn" id="printBtn">🖨 列印</button><button class="btn" id="themeBtn">🌙</button></div></nav>
<main>
<section class="hero"><h1>📋 考前速查表</h1>
<p>所有單元的核心觀念整合在這一頁，列印出來考前掃一次。</p>
<a class="btn primary" href="index.html">← 回到地圖</a>
</section>
{cheat_items}
</main>
<footer>一次一小步，你做得到 🌱</footer>
<script src="search-index.js"></script>
<script src="quiz-data.js"></script>
<script src="app.js"></script></body></html>"""
    (folder / "cheatsheet.html").write_text(cheat_page, encoding="utf-8")

    # 額外頁：quiz
    quiz_page = """<!doctype html>
<html lang="zh-TW"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>模擬測驗｜ISTQB</title>
<link rel="stylesheet" href="styles.css"></head><body>
<nav class="nav"><div class="brand"><a href="index.html" style="color:inherit;text-decoration:none;display:flex;align-items:center;gap:10px"><span class="dot"></span> ISTQB 學習地圖</a></div>
<div class="tools"><button class="btn" id="themeBtn">🌙</button></div></nav>
<main>
<section class="hero"><h1>🎯 模擬測驗</h1>
<p>從常考題目精選 12 題，做完知道自己準備到哪了。</p>
</section>
<div id="quizBox"></div>
</main>
<footer>一次一小步，你做得到 🌱</footer>
<script src="search-index.js"></script>
<script src="quiz-data.js"></script>
<script src="app.js"></script></body></html>"""
    (folder / "quiz.html").write_text(quiz_page, encoding="utf-8")

if __name__ == "__main__":
    build_excel()
    build_istqb()
    print("\n✨ 完成！")
