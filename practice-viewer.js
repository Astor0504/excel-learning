/* practice.xlsx 內嵌瀏覽器
   只顯示對應本課的工作表，其他 sheet 收進「其他工作表」disclosure。
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
    '    <div class="pv-head-l">',
    '      <span class="pv-title">本課練習簿</span>',
    '      <span class="pv-sheet-name"></span>',
    '    </div>',
    '    <div class="pv-actions">',
    '      <span class="pv-status">載入中…</span>',
    '      <a class="pv-btn" href="'+src+'" download>下載原檔</a>',
    '    </div>',
    '  </div>',
    (highlights.length ? '  <div class="pv-hint"><span class="pv-hint-label">本課重點</span>'+highlights.map(function(h){return '<span class="pv-chip">'+esc(h)+'</span>';}).join('')+'</div>' : ''),
    '  <div class="pv-body"></div>',
    '</div>'
  ].join('');

  var body = mount.querySelector('.pv-body');
  var statusEl = mount.querySelector('.pv-status');
  var sheetNameEl = mount.querySelector('.pv-sheet-name');

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

    var startIdx = 0;
    var slugMatch = -1;
    for (var j = 0; j < names.length; j++){
      if (slug && names[j].toUpperCase().indexOf(slug.toUpperCase()) >= 0){ slugMatch = j; break; }
    }
    if (slugMatch >= 0) startIdx = slugMatch;
    else if (highlights.length) {
      for (var i = 0; i < names.length; i++) {
        var nm = names[i];
        if (highlights.some(function(h){ return nm.toUpperCase().indexOf(h.toUpperCase()) >= 0; })) {
          startIdx = i; break;
        }
      }
    }

    body.innerHTML = '<div class="pv-sheet"></div>';
    var sheetEl = body.querySelector('.pv-sheet');

    function show(i){
      var ws = wb.Sheets[names[i]];
      var html = XLSX.utils.sheet_to_html(ws, {editable:false});
      sheetEl.innerHTML = '<div class="pv-table">'+html+'</div>';
      sheetNameEl.textContent = names[i];
      applyHighlights(sheetEl);
    }
    show(startIdx);

    if (names.length > 1) {
      var det = document.createElement('details');
      det.className = 'pv-other';
      var sum = document.createElement('summary');
      sum.innerHTML = '其他工作表 <span class="pv-other-count">' + (names.length-1) + '</span><span class="pv-other-chev" aria-hidden="true">▾</span>';
      det.appendChild(sum);
      var listEl = document.createElement('div');
      listEl.className = 'pv-other-list';
      names.forEach(function(n, i){
        if (i === startIdx) return;
        var btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'pv-other-link';
        btn.textContent = n;
        btn.addEventListener('click', function(){
          show(i);
          listEl.querySelectorAll('.pv-other-link').forEach(function(x){x.classList.remove('is-active');});
          btn.classList.add('is-active');
          var top = mount.getBoundingClientRect().top + window.pageYOffset - 80;
          window.scrollTo({top: top, behavior:'smooth'});
        });
        listEl.appendChild(btn);
      });
      det.appendChild(listEl);
      body.appendChild(det);
    }
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
