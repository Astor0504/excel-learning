/* practice.xlsx 內嵌瀏覽器（ADHD 友善版：自動載入 + 關鍵字高亮）
   用法：<div id="practiceViewer" data-src="../practice.xlsx"></div>
   高亮：從 body data-lesson-slug 讀對應 highlights。
*/
(function(){
  var SHEETJS_URL = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
  var mount = document.getElementById('practiceViewer');
  if (!mount) return;

  var src = mount.dataset.src || '../practice.xlsx';
  var slug = document.body.dataset.lessonSlug || '';
  var highlights = [];
  if (window.CHEATSHEET_DATA && window.CHEATSHEET_DATA.lessonMap[slug]) {
    highlights = window.CHEATSHEET_DATA.lessonMap[slug].highlights || [];
  }

  mount.innerHTML = [
    '<div class="pv-box">',
    '  <div class="pv-head">',
    '    <span class="pv-title">✏️ 練習簿即時預覽</span>',
    '    <div class="pv-actions">',
    '      <span class="pv-status">載入中…</span>',
    '      <a class="pv-btn" href="'+src+'" download>📥 下載原檔</a>',
    '    </div>',
    '  </div>',
    (highlights.length ? '  <div class="pv-hint">🎯 本課重點關鍵字：'+highlights.map(function(h){return '<span class="pv-chip">'+esc(h)+'</span>';}).join('')+'（練習簿中會自動高亮）</div>' : ''),
    '  <div class="pv-body"></div>',
    '</div>'
  ].join('');

  var body = mount.querySelector('.pv-body');
  var statusEl = mount.querySelector('.pv-status');

  // 自動載入（defer 到 idle 時間，避免阻塞首屏）
  var kick = function(){
    loadSheetJS(function(err){
      if (err) { setStatus('載入失敗', true); return; }
      fetchXlsx();
    });
  };
  if (window.requestIdleCallback) requestIdleCallback(kick, {timeout: 1500});
  else setTimeout(kick, 300);

  function setStatus(text, isErr){
    if (!statusEl) return;
    statusEl.textContent = text;
    statusEl.classList.toggle('pv-err-text', !!isErr);
  }

  function loadSheetJS(cb){
    if (window.XLSX) return cb(null);
    var s = document.createElement('script');
    s.src = SHEETJS_URL;
    s.onload = function(){ cb(null); };
    s.onerror = function(){ cb(new Error('script error')); };
    document.head.appendChild(s);
  }

  function fetchXlsx(){
    fetch(src).then(function(r){
      if (!r.ok) throw new Error('HTTP '+r.status);
      return r.arrayBuffer();
    }).then(function(buf){
      render(buf);
      setStatus('已載入');
      setTimeout(function(){ if (statusEl) statusEl.style.display = 'none'; }, 1200);
    }).catch(function(e){
      setStatus('讀取失敗：'+e.message, true);
    });
  }

  function render(buf){
    var wb = XLSX.read(buf, {type:'array'});
    var names = wb.SheetNames;

    // 預設切到和本課 highlights 最相關的 sheet（名稱含關鍵字的）
    var startIdx = 0;
    if (highlights.length) {
      for (var i = 0; i < names.length; i++) {
        var nm = names[i];
        if (highlights.some(function(h){ return nm.toUpperCase().indexOf(h.toUpperCase()) >= 0; })) {
          startIdx = i; break;
        }
      }
    }

    var tabsHtml = names.map(function(n, i){
      return '<button class="pv-tab'+(i===startIdx?' active':'')+'" data-i="'+i+'">'+esc(n)+'</button>';
    }).join('');
    body.innerHTML = '<div class="pv-tabs">'+tabsHtml+'</div><div class="pv-sheet"></div>';
    var sheetEl = body.querySelector('.pv-sheet');
    function show(i){
      var ws = wb.Sheets[names[i]];
      var html = XLSX.utils.sheet_to_html(ws, {editable:false});
      sheetEl.innerHTML = '<div class="pv-table">'+html+'</div>';
      applyHighlights(sheetEl);
    }
    body.querySelectorAll('.pv-tab').forEach(function(b){
      b.addEventListener('click', function(){
        body.querySelectorAll('.pv-tab').forEach(function(x){x.classList.remove('active');});
        b.classList.add('active');
        show(parseInt(b.dataset.i,10));
      });
    });
    show(startIdx);
  }

  function applyHighlights(container){
    if (!highlights.length) return;
    container.querySelectorAll('td').forEach(function(td){
      var txt = td.textContent || '';
      var matched = highlights.some(function(h){
        return txt.toUpperCase().indexOf(h.toUpperCase()) >= 0;
      });
      if (matched) td.classList.add('pv-hl');
    });
  }

  function esc(s){ return String(s||'').replace(/[&<>]/g, function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c];}); }
})();
