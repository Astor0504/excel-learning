import { CHEATSHEET_DATA } from './cheatsheet-data.js';

/* 速查抽屜 (cheat drawer)
   在 lesson 頁右下角浮出「📘 速查」按鈕，點開從右側滑入。
   需要：cheatsheet-data.js 先載入、HTML body data-lesson-slug="P1-01" 用於預設篩選。
   若無 slug 則顯示全部，使用者可自行搜尋/篩選。
*/
(function(){
  if (!CHEATSHEET_DATA) {
    console.warn('drawer: CHEATSHEET_DATA 未載入');
    return;
  }
  var D = CHEATSHEET_DATA;
  var slug = document.body.dataset.lessonSlug || '';
  var lessonConf = D.lessonMap[slug] || {cats:[], title:'速查手冊'};

  var CAT_LABELS = {
    math:'數學/統計', condition:'條件', lookup:'查找', text:'文字',
    date:'日期', array:'陣列', advanced:'進階', shortcut:'快捷鍵',
    pivot:'樞紐', format:'格式', chart:'圖表', protect:'保護', pquery:'Power Query'
  };

  // ── 注入 UI 骨架 ──
  var fab = document.createElement('button');
  fab.id = 'cheatFab';
  fab.title = '打開速查';
  fab.setAttribute('aria-label', '打開速查手冊');
  fab.innerHTML = '📘<span class="cd-badge" style="display:none">0</span>';
  document.body.appendChild(fab);

  var LS_KEY = 'excelLearned';
  function updateBadge(){
    var learned = {};
    try { learned = JSON.parse(localStorage.getItem(LS_KEY) || '{}'); } catch(e){}
    var n = Object.keys(learned).filter(function(k){ return learned[k]; }).length;
    var badge = fab.querySelector('.cd-badge');
    if (!badge) return;
    if (n > 0) {
      badge.style.display = '';
      badge.textContent = n;
      fab.classList.add('has-progress');
    } else {
      badge.style.display = 'none';
      fab.classList.remove('has-progress');
    }
  }
  updateBadge();
  window.addEventListener('cheat-learned-changed', updateBadge);

  var panel = document.createElement('div');
  panel.id = 'cheatDrawer';
  panel.innerHTML = [
    '<div class="cd-head">',
    '  <strong id="cdTitle">'+lessonConf.title+'</strong>',
    '  <button id="cdClose" title="關閉">×</button>',
    '</div>',
    '<div class="cd-search">',
    '  <input id="cdSearch" placeholder="搜尋函式/快捷鍵…" autocomplete="off">',
    '  <div id="cdTabs" class="cd-tabs"></div>',
    '</div>',
    '<div id="cdBody" class="cd-body"></div>',
    '<a class="cd-full" href="../cheatsheet.html" target="_blank" rel="noopener">查看完整速查手冊 →</a>'
  ].join('');
  document.body.appendChild(panel);

  var backdrop = document.createElement('div');
  backdrop.id = 'cheatBackdrop';
  document.body.appendChild(backdrop);

  // ── 狀態 ──
  var state = {
    kind: (lessonConf.kinds && lessonConf.kinds[0]) || 'formula',
    cat: (lessonConf.cats && lessonConf.cats[0]) || 'all',
    q: ''
  };

  // ── Tabs（分類） ──
  var tabsEl = panel.querySelector('#cdTabs');
  function buildTabs(){
    var tabs = [];
    var cats = (lessonConf.cats && lessonConf.cats.length) ? lessonConf.cats : null;
    if (state.kind === 'formula') {
      tabs.push({v:'all', label:'全部'});
      var allCats = Array.from(new Set(D.formulas.map(function(f){return f.cat;})));
      // 先列 lesson 指定的分類
      if (cats) cats.forEach(function(c){ if (allCats.indexOf(c)>=0) tabs.push({v:c, label:CAT_LABELS[c]||c}); });
      // 再列其餘
      allCats.forEach(function(c){
        if (!cats || cats.indexOf(c)<0) tabs.push({v:c, label:CAT_LABELS[c]||c});
      });
    } else {
      tabs.push({v:'all', label:'全部'});
      var vbaCats = Array.from(new Set(D.vbaTemplates.map(function(v){return v.cat;})));
      vbaCats.forEach(function(c){ tabs.push({v:c, label:c}); });
    }
    tabsEl.innerHTML = tabs.map(function(t){
      return '<button class="cd-tab'+(t.v===state.cat?' active':'')+'" data-v="'+t.v+'">'+t.label+'</button>';
    }).join('');
    tabsEl.querySelectorAll('.cd-tab').forEach(function(b){
      b.addEventListener('click', function(){ state.cat = b.dataset.v; buildTabs(); render(); });
    });
  }

  // ── Kind 切換（公式 vs VBA） ──
  var hasVBA = (lessonConf.kinds && lessonConf.kinds.indexOf('vba')>=0);
  var kindBar = document.createElement('div');
  kindBar.className = 'cd-kind';
  kindBar.innerHTML =
    '<button data-k="formula"'+(state.kind==='formula'?' class="active"':'')+'>📘 函式/快捷</button>' +
    '<button data-k="vba"'+(state.kind==='vba'?' class="active"':'')+'>⚙️ VBA 範本</button>';
  panel.querySelector('.cd-search').insertBefore(kindBar, panel.querySelector('#cdTabs'));
  kindBar.querySelectorAll('button').forEach(function(b){
    b.addEventListener('click', function(){
      state.kind = b.dataset.k;
      state.cat = 'all';
      kindBar.querySelectorAll('button').forEach(function(x){x.classList.remove('active');});
      b.classList.add('active');
      buildTabs(); render();
    });
  });

  // ── 搜尋 ──
  panel.querySelector('#cdSearch').addEventListener('input', function(e){
    state.q = e.target.value.trim().toLowerCase();
    render();
  });

  // ── 渲染 ──
  var body = panel.querySelector('#cdBody');
  function render(){
    var items;
    if (state.kind === 'formula') {
      items = D.formulas.filter(function(f){
        if (state.cat !== 'all' && f.cat !== state.cat) return false;
        if (!state.q) return true;
        var hay = (f.name+' '+f.desc+' '+f.syntax+' '+(f.tags||[]).join(' ')).toLowerCase();
        return hay.indexOf(state.q) >= 0;
      });
      body.innerHTML = items.length ? items.map(formulaCard).join('') :
        '<p class="cd-empty">沒有符合的項目</p>';
    } else {
      items = D.vbaTemplates.filter(function(v){
        if (state.cat !== 'all' && v.cat !== state.cat) return false;
        if (!state.q) return true;
        return (v.title+' '+v.explain+' '+v.code).toLowerCase().indexOf(state.q) >= 0;
      });
      body.innerHTML = items.length ? items.map(vbaCard).join('') :
        '<p class="cd-empty">沒有符合的項目</p>';
    }
    body.querySelectorAll('.cd-card').forEach(function(c){
      c.addEventListener('click', function(e){
        if (e.target.tagName === 'BUTTON') return;
        c.classList.toggle('expanded');
      });
    });
    body.querySelectorAll('.cd-copy').forEach(function(b){
      b.addEventListener('click', function(e){
        e.stopPropagation();
        var code = b.dataset.code;
        navigator.clipboard.writeText(code).then(function(){
          b.textContent = '已複製';
          setTimeout(function(){ b.textContent = '複製'; }, 1500);
        });
      });
    });
  }

  function esc(s){ return String(s||'').replace(/[&<>]/g, function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c];}); }
  function formulaCard(f){
    return '<div class="cd-card lv'+f.lv+'">' +
      '<div class="cd-card-head"><span class="cd-name">'+esc(f.name)+'</span><span class="cd-lv">Lv.'+f.lv+'</span></div>' +
      '<div class="cd-desc">'+esc(f.desc)+'</div>' +
      '<code class="cd-syntax">'+esc(f.syntax)+'</code>' +
      '<div class="cd-detail"><div class="cd-ex"><b>範例：</b>'+esc(f.example)+'</div>' +
      '<div class="cd-tip">💡 '+esc(f.tip)+'</div>' +
      '<p>'+esc(f.detail).replace(/\\n/g,'<br>')+'</p></div>' +
    '</div>';
  }
  function vbaCard(v){
    return '<div class="cd-card cd-vba lv'+v.lv+'">' +
      '<div class="cd-card-head"><span class="cd-name">'+esc(v.title)+'</span><span class="cd-lv">Lv.'+v.lv+'</span></div>' +
      '<div class="cd-desc">'+esc(v.explain)+'</div>' +
      '<div class="cd-detail"><button class="cd-copy" data-code="'+esc(v.code).replace(/"/g,'&quot;')+'">複製</button>' +
      '<pre>'+esc(v.code)+'</pre></div>' +
    '</div>';
  }

  // ── 開關 ──
  function open(){
    panel.classList.add('open');
    backdrop.classList.add('show');
    setTimeout(function(){ panel.querySelector('#cdSearch').focus(); }, 200);
  }
  function close(){
    panel.classList.remove('open');
    backdrop.classList.remove('show');
  }
  fab.addEventListener('click', open);
  panel.querySelector('#cdClose').addEventListener('click', close);
  backdrop.addEventListener('click', close);
  document.addEventListener('keydown', function(e){
    if (e.key === 'Escape' && panel.classList.contains('open')) close();
  });

  // 若沒有 lesson-slug，hide kind bar 的 vba 按鈕除非有 VBA 資料
  if (!hasVBA && !/P5-/.test(slug)) {
    // 仍然允許手動切換到 VBA（以防使用者想看），只是不主動顯示
  }

  // 初始渲染
  buildTabs();
  render();

  // 課程頁把浮動工具收成同一條 rail，避免互相蓋住內容
  if (document.body.dataset.lessonSlug) {
    var rail = document.getElementById('lessonFabRail');
    if (!rail) {
      rail = document.createElement('div');
      rail.id = 'lessonFabRail';
      document.body.appendChild(rail);
    }
    ['ttsFab', 'aiFab', 'cheatFab'].forEach(function(id){
      var el = document.getElementById(id);
      if (el) rail.appendChild(el);
    });
  }
})();
