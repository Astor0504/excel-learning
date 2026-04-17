import { CHEATSHEET_DATA } from './cheatsheet-data.js';

/* 本課關鍵速查 — inline 大卡片
   需要：cheatsheet-data.js + <div id="cheatInline" data-slug="P1-01"></div>
   若 slug 無對應卡片則不渲染（例如未設 pick 的章節）
*/
(function(){
  if (!CHEATSHEET_DATA) return;
  var mount = document.getElementById('cheatInline');
  if (!mount) return;

  var D = CHEATSHEET_DATA;
  var slug = mount.dataset.slug || document.body.dataset.lessonSlug || '';
  var conf = D.lessonMap[slug];
  if (!conf) { mount.style.display = 'none'; return; }

  var picks = conf.pick || [];
  var pickVbas = conf.pickVba || [];
  var cards = [];
  picks.forEach(function(name){
    var f = D.formulas.find(function(x){ return x.name === name; });
    if (f) cards.push({kind:'formula', data:f});
  });
  pickVbas.forEach(function(name){
    var v = D.vbaTemplates.find(function(x){ return x.title === name; });
    if (v) cards.push({kind:'vba', data:v});
  });

  if (cards.length === 0) { mount.style.display = 'none'; return; }

  var LS_KEY = 'excelLearned';
  var learned = {};
  try { learned = JSON.parse(localStorage.getItem(LS_KEY) || '{}'); } catch(e){}

  function keyOf(card){ return card.kind === 'formula' ? card.data.name : 'VBA:' + card.data.title; }
  function isLearned(card){ return !!learned[keyOf(card)]; }

  function learnedCount(){
    return cards.filter(isLearned).length;
  }

  function esc(s){ return String(s||'').replace(/[&<>]/g, function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c];}); }

  function render(){
    var done = learnedCount();
    var pct = Math.round(done / cards.length * 100);

    var cardsHtml = cards.map(function(c, i){
      var d = c.data;
      var key = keyOf(c);
      var done = !!learned[key];
      if (c.kind === 'formula') {
        return '' +
          '<div class="lci-card lv'+d.lv+(done?' lci-done':'')+'" data-key="'+esc(key)+'">' +
            '<div class="lci-bar"></div>' +
            '<div class="lci-body">' +
              '<div class="lci-top">' +
                '<span class="lci-name">'+esc(d.name)+'</span>' +
                '<span class="lci-lv">Lv.'+d.lv+'</span>' +
              '</div>' +
              '<code class="lci-syntax">'+esc(d.syntax)+'</code>' +
              '<p class="lci-desc">'+esc(d.desc)+'</p>' +
              '<div class="lci-ex">'+esc(d.example)+'</div>' +
              '<div class="lci-tip">💡 '+esc(d.tip)+'</div>' +
              '<button class="lci-toggle" data-key="'+esc(key)+'">' + (done ? '✓ 已學會' : '標記學會') + '</button>' +
            '</div>' +
          '</div>';
      }
      // VBA
      return '' +
        '<div class="lci-card lv'+d.lv+' lci-vba'+(done?' lci-done':'')+'" data-key="'+esc(key)+'">' +
          '<div class="lci-bar"></div>' +
          '<div class="lci-body">' +
            '<div class="lci-top">' +
              '<span class="lci-name">⚙️ '+esc(d.title)+'</span>' +
              '<span class="lci-lv">Lv.'+d.lv+'</span>' +
            '</div>' +
            '<p class="lci-desc">'+esc(d.explain)+'</p>' +
            '<pre class="lci-code">'+esc(d.code)+'</pre>' +
            '<button class="lci-toggle" data-key="'+esc(key)+'">' + (done ? '✓ 已學會' : '標記學會') + '</button>' +
          '</div>' +
        '</div>';
    }).join('');

    mount.innerHTML = '' +
      '<section class="lci-wrap">' +
        '<div class="lci-header">' +
          '<div class="lci-headline">' +
            '<h3 class="lci-title">⚡ 本課關鍵速查</h3>' +
            '<span class="lci-subtitle">'+esc(conf.title)+' · 打勾標記已學會</span>' +
          '</div>' +
          '<div class="lci-progress">' +
            '<div class="lci-progress-bar"><span style="width:'+pct+'%"></span></div>' +
            '<div class="lci-progress-text">'+done+' / '+cards.length+' 已學會</div>' +
          '</div>' +
        '</div>' +
        '<div class="lci-grid">'+cardsHtml+'</div>' +
        '<p class="lci-hint">需要更多？點右下角 <b>📘 速查</b> 看全部 '+D.formulas.length+' 個函式</p>' +
      '</section>';

    // 綁勾選
    mount.querySelectorAll('.lci-toggle').forEach(function(b){
      b.addEventListener('click', function(e){
        e.stopPropagation();
        var k = b.dataset.key;
        learned[k] = !learned[k];
        try { localStorage.setItem(LS_KEY, JSON.stringify(learned)); } catch(e){}
        if (learned[k]) celebrate(b);
        render();
        // 通知 drawer 更新 FAB 徽章
        window.dispatchEvent(new CustomEvent('cheat-learned-changed'));
      });
    });
  }

  function celebrate(btn){
    // 小慶祝動畫：飛散的 ✓ 小點
    var rect = btn.getBoundingClientRect();
    for (var i = 0; i < 6; i++) {
      var dot = document.createElement('span');
      dot.className = 'lci-spark';
      dot.textContent = '✓';
      dot.style.left = (rect.left + rect.width/2) + 'px';
      dot.style.top = (rect.top + rect.height/2) + 'px';
      var angle = (Math.PI * 2 / 6) * i;
      dot.style.setProperty('--dx', Math.cos(angle) * 60 + 'px');
      dot.style.setProperty('--dy', Math.sin(angle) * 60 + 'px');
      document.body.appendChild(dot);
      setTimeout(function(el){ return function(){ el.remove(); }; }(dot), 700);
    }
  }

  render();
})();
