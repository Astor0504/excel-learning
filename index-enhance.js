import { XLSX_CONTENT } from './xlsx-content.js';

/* index-enhance.js — 在首頁卡片上注入 XP/題數/難度，並算出總計顯示在 Hero。 */
(function(){
  if (!XLSX_CONTENT) return;
  var lessons = XLSX_CONTENT.lessons || {};

  // 1. 每張卡片加 chips
  var totalXp = 0, totalTasks = 0, totalMinutes = 0;
  document.querySelectorAll('a[href^="lessons/P"]').forEach(function(a){
    var m = a.getAttribute('href').match(/P\d-\d{2}/);
    if (!m) return;
    var slug = m[0];
    var data = lessons?.[slug];
    var meta = data?.meta;
    if (!meta) return;
    var card = a.querySelector('.card');
    if (!card) return;
    var doneKey = 'done:lessons/' + slug + '.html';
    var seenKey = 'seen:lessons/' + slug + '.html';
    var isDone = false;
    var seen = null;
    try {
      isDone = !!localStorage.getItem(doneKey);
      seen = localStorage.getItem(seenKey);
    } catch(e){}

    var chips = document.createElement('div');
    chips.className = 'idx-chips';
    function chip(emoji, txt, cls){
      if (!txt) return;
      var s = document.createElement('span');
      s.className = 'idx-chip' + (cls ? ' ' + cls : '');
      s.innerHTML = '<span class="idx-chip-emoji">'+emoji+'</span>'+txt;
      chips.appendChild(s);
    }
    if (isDone) {
      chip('✓', '已完成', 'is-done');
    } else if (seen) {
      var days = Math.floor((Date.now() - parseInt(seen, 10)) / 86400000);
      chip('↺', days <= 0 ? '今天看過' : days === 1 ? '昨天看過' : '最近看過', 'is-warm');
    }
    chip('🎯', meta.difficulty);
    chip('📝', meta.taskCount);
    chip('⏱', meta.time);
    chip('⚡', meta.xp);
    if (chips.children.length) card.appendChild(chips);

    // 累計
    var xpNum = parseInt((meta.xp||'').replace(/[^\d]/g,''))||0;
    var taskNum = parseInt((meta.taskCount||'').replace(/[^\d]/g,''))||0;
    var minNum = parseInt((meta.time||'').replace(/[^\d]/g,''))||0;
    totalXp += xpNum; totalTasks += taskNum; totalMinutes += minNum;
  });

  // 2. Hero 加總計條
  var heroActions = document.querySelector('.hero .hero-actions');
  if (heroActions && totalXp > 0){
    var summary = document.createElement('div');
    summary.className = 'idx-summary';
    summary.innerHTML =
      '<div class="idx-summary-item"><div class="idx-summary-num">'+totalXp.toLocaleString()+'</div><div class="idx-summary-lbl">總 XP</div></div>'+
      '<div class="idx-summary-item"><div class="idx-summary-num">'+totalTasks+'</div><div class="idx-summary-lbl">微任務</div></div>'+
      '<div class="idx-summary-item"><div class="idx-summary-num">'+Math.round(totalMinutes/60*10)/10+'</div><div class="idx-summary-lbl">小時</div></div>'+
      '<div class="idx-summary-item"><div class="idx-summary-num">21</div><div class="idx-summary-lbl">堂課</div></div>';
    heroActions.parentNode.insertBefore(summary, heroActions);
  }
})();
