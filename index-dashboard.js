/* index-dashboard.js — 首頁學習儀表板：徽章、學習策略、起點 */
(function(){
  if (!document.querySelector('.hero')) return;
  if (!window.XLSX_CONTENT) return;

  var D = window.XLSX_CONTENT;
  var badges = D.badges || [];
  var strategies = D.strategies || [];

  // 計算目前完成的階段數 / 累積 XP
  var stageXp = {};
  Object.keys(D.dashboard || {}).forEach(function(k){
    var item = D.dashboard[k];
    var xp = parseInt(String(item.xp||'').replace(/[^\d]/g,''),10) || 0;
    stageXp[k] = xp;
  });

  // 從 localStorage 讀 lesson 完成狀態（以 lesson-check-{slug} 是否有內容 = 已開始看作完成）
  var slugToXp = {
    'P1-01':150,'P1-02':100,'P1-03':100,
    'P2-01':150,'P2-02':200,'P2-03':200,'P2-04':150,
    'P3-01':150,'P3-02':150,'P3-03':200,'P3-04':150,'P3-05':100,
    'P4-01':200,'P4-02':250,'P4-03':200,'P4-04':300,
    'P5-01':300,'P5-02':300,'P5-03':400,'P5-04':500
  };
  var done = {};
  var totalXp = 0;
  Object.keys(slugToXp).forEach(function(slug){
    try {
      var v = localStorage.getItem('lesson-check-'+slug);
      if (v && JSON.parse(v||'[]').length){
        done[slug] = true;
        totalXp += slugToXp[slug];
      }
    } catch(e){}
  });

  // 徽章解鎖規則（以 Phase 完成度判斷）
  var phaseRules = [
    {name:'鍵盤戰士', slugs:['P1-01','P1-02','P1-03'], xp:350},
    {name:'即戰力新星', slugs:['P2-01','P2-02','P2-03','P2-04'], xp:1050},
    {name:'專業報表手', slugs:['P3-01','P3-02','P3-03','P3-04','P3-05'], xp:1800},
    {name:'資料工程師', slugs:['P4-01','P4-02','P4-03','P4-04'], xp:2750},
    {name:'自動化大師', slugs:['P5-01','P5-02'], xp:3350},
    {name:'全能導師', slugs:Object.keys(slugToXp), xp:3850},
  ];

  // 簡潔線條 SVG icons（避免 emoji/AI 感）
  var ICONS = {
    keyboard: '<svg viewBox="0 0 48 48" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="6" y="14" width="36" height="22" rx="3"/><line x1="12" y1="20" x2="14" y2="20"/><line x1="18" y1="20" x2="20" y2="20"/><line x1="24" y1="20" x2="26" y2="20"/><line x1="30" y1="20" x2="32" y2="20"/><line x1="36" y1="20" x2="38" y2="20"/><line x1="12" y1="26" x2="14" y2="26"/><line x1="18" y1="26" x2="20" y2="26"/><line x1="24" y1="26" x2="26" y2="26"/><line x1="30" y1="26" x2="32" y2="26"/><line x1="36" y1="26" x2="38" y2="26"/><line x1="16" y1="32" x2="32" y2="32"/></svg>',
    briefcase: '<svg viewBox="0 0 48 48" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="6" y="14" width="36" height="26" rx="3"/><path d="M18 14v-3a2 2 0 012-2h8a2 2 0 012 2v3"/><line x1="6" y1="24" x2="42" y2="24"/><circle cx="24" cy="24" r="2"/></svg>',
    chart: '<svg viewBox="0 0 48 48" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="8" y="6" width="32" height="36" rx="3"/><line x1="14" y1="34" x2="14" y2="26"/><line x1="20" y1="34" x2="20" y2="20"/><line x1="26" y1="34" x2="26" y2="14"/><line x1="32" y1="34" x2="32" y2="22"/></svg>',
    flow: '<svg viewBox="0 0 48 48" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="4"/><circle cx="36" cy="12" r="4"/><circle cx="24" cy="24" r="4"/><circle cx="12" cy="36" r="4"/><circle cx="36" cy="36" r="4"/><line x1="15" y1="14" x2="21" y2="21"/><line x1="33" y1="14" x2="27" y2="21"/><line x1="21" y1="27" x2="15" y2="33"/><line x1="27" y1="27" x2="33" y2="33"/></svg>',
    terminal: '<svg viewBox="0 0 48 48" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="6" y="10" width="36" height="28" rx="3"/><polyline points="14,20 19,24 14,28"/><line x1="23" y1="30" x2="33" y2="30"/></svg>',
    crown: '<svg viewBox="0 0 48 48" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M8 34l-2-16 8 6 10-14 10 14 8-6-2 16z"/><line x1="8" y1="40" x2="40" y2="40"/><circle cx="14" cy="18" r="1.5"/><circle cx="34" cy="18" r="1.5"/><circle cx="24" cy="10" r="1.5"/></svg>',
  };
  var iconList = [ICONS.keyboard, ICONS.briefcase, ICONS.chart, ICONS.flow, ICONS.terminal, ICONS.crown];

  // 找到 hero 區塊，接在後面
  var hero = document.querySelector('.hero');

  // ===== 成就徽章 =====
  var section = document.createElement('section');
  section.className = 'idx-dashboard';

  var bHead = document.createElement('div');
  bHead.className = 'idx-dash-head';
  bHead.innerHTML =
    '<h2 class="idx-dash-title">學習里程碑</h2>' +
    '<span class="idx-dash-sub">完成階段自動解鎖</span>';
  section.appendChild(bHead);

  var bGrid = document.createElement('div');
  bGrid.className = 'idx-badges';
  phaseRules.forEach(function(r, i){
    var doneCount = r.slugs.filter(function(s){return done[s];}).length;
    var total = r.slugs.length;
    var unlocked = doneCount === total;
    var card = document.createElement('div');
    card.className = 'idx-badge' + (unlocked ? ' is-unlocked' : '');
    card.innerHTML =
      '<div class="idx-badge-ico">' + iconList[i] + '</div>' +
      '<div class="idx-badge-name">' + r.name + '</div>' +
      '<div class="idx-badge-meta">' + doneCount + ' / ' + total + ' 階段 · ' +
      r.xp.toLocaleString() + ' XP</div>' +
      '<div class="idx-badge-bar"><span style="width:' + (doneCount/total*100) + '%"></span></div>';
    bGrid.appendChild(card);
  });
  section.appendChild(bGrid);

  // ===== 學習策略（可折疊）=====
  var sWrap = document.createElement('details');
  sWrap.className = 'idx-strategy';
  var sSum = document.createElement('summary');
  sSum.innerHTML = '<span class="idx-strategy-label">學習策略</span>' +
    '<span class="idx-strategy-hint">建議先讀過一次再開始</span>' +
    '<span class="idx-strategy-chev" aria-hidden="true">▾</span>';
  sWrap.appendChild(sSum);

  var sList = document.createElement('div');
  sList.className = 'idx-strategy-list';
  strategies.forEach(function(s){
    var title = String(s.title || '').replace(/^[^\u4e00-\u9fa5A-Za-z]+/, '').trim();
    var item = document.createElement('div');
    item.className = 'idx-strategy-item';
    item.innerHTML =
      '<div class="idx-strategy-title">' + escapeHtml(title) + '</div>' +
      '<div class="idx-strategy-desc">' + escapeHtml(s.desc) + '</div>';
    sList.appendChild(item);
  });
  sWrap.appendChild(sList);
  section.appendChild(sWrap);

  // 插到 hero 之後
  hero.parentNode.insertBefore(section, hero.nextSibling);

  function escapeHtml(s){
    return String(s||'').replace(/[&<>"']/g, function(c){
      return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];
    });
  }
})();
