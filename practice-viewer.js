/* practice.xlsx 內嵌瀏覽器
   用法：在 lesson HTML 內放 <div id="practiceViewer" data-src="../practice.xlsx"></div>
   點開才懶載 SheetJS（避免首屏多下載 1MB）
*/
(function(){
  var SHEETJS_URL = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
  var mount = document.getElementById('practiceViewer');
  if (!mount) return;

  var src = mount.dataset.src || '../practice.xlsx';
  mount.innerHTML = [
    '<div class="pv-box">',
    '  <div class="pv-head">',
    '    <span class="pv-title">✏️ 練習簿（線上預覽）</span>',
    '    <div class="pv-actions">',
    '      <a class="pv-btn" href="'+src+'" download>📥 下載原檔</a>',
    '      <button class="pv-btn pv-load">載入預覽</button>',
    '    </div>',
    '  </div>',
    '  <div class="pv-hint">點「載入預覽」在本頁瀏覽所有工作表（不離開頁面）。首次載入約 1MB。</div>',
    '  <div class="pv-body"></div>',
    '</div>'
  ].join('');

  var body = mount.querySelector('.pv-body');
  var loadBtn = mount.querySelector('.pv-load');
  var loaded = false;

  loadBtn.addEventListener('click', function(){
    if (loaded) return;
    loadBtn.disabled = true;
    loadBtn.textContent = '載入中…';
    loadSheetJS(function(err){
      if (err) { body.innerHTML = '<p class="pv-err">SheetJS 載入失敗：'+err.message+'</p>'; loadBtn.disabled = false; loadBtn.textContent = '重試'; return; }
      fetchXlsx();
    });
  });

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
      loaded = true;
      loadBtn.style.display = 'none';
    }).catch(function(e){
      body.innerHTML = '<p class="pv-err">讀取失敗：'+e.message+'</p>';
      loadBtn.disabled = false; loadBtn.textContent = '重試';
    });
  }

  function render(buf){
    var wb = XLSX.read(buf, {type:'array'});
    var names = wb.SheetNames;
    var tabsHtml = names.map(function(n, i){
      return '<button class="pv-tab'+(i===0?' active':'')+'" data-i="'+i+'">'+esc(n)+'</button>';
    }).join('');
    body.innerHTML = '<div class="pv-tabs">'+tabsHtml+'</div><div class="pv-sheet"></div>';
    var sheetEl = body.querySelector('.pv-sheet');
    function show(i){
      var ws = wb.Sheets[names[i]];
      var html = XLSX.utils.sheet_to_html(ws, {editable:false});
      // 包一層方便套樣式 + 支援水平捲動
      sheetEl.innerHTML = '<div class="pv-table">'+html+'</div>';
    }
    body.querySelectorAll('.pv-tab').forEach(function(b){
      b.addEventListener('click', function(){
        body.querySelectorAll('.pv-tab').forEach(function(x){x.classList.remove('active');});
        b.classList.add('active');
        show(parseInt(b.dataset.i,10));
      });
    });
    show(0);
  }

  function esc(s){ return String(s||'').replace(/[&<>]/g, function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c];}); }
})();
