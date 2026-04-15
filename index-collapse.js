/* index-collapse.js — 首頁 Phase 折疊 / 進度聚合 */
(function(){
  if (!document.querySelector('.hero')) return;

  var main = document.querySelector('main');
  if (!main) return;

  // 找所有 Phase h2
  var h2s = Array.from(main.querySelectorAll('h2')).filter(function(h){
    return /Phase\s*\d/.test(h.textContent);
  });
  if (!h2s.length) return;

  // 把每個 h2 後面連續的 .card 包起來
  h2s.forEach(function(h2, idx){
    var group = document.createElement('div');
    group.className = 'phase-group';
    group.setAttribute('data-phase-idx', idx);
    var lessons = [];
    var node = h2.nextElementSibling;
    while (node && node.tagName !== 'H2'){
      var next = node.nextElementSibling;
      lessons.push(node);
      group.appendChild(node);
      node = next;
    }

    // 計算已完成數（看 .card 是否有 'done' class 或 localStorage）
    var total = lessons.length;
    var done = 0;
    lessons.forEach(function(a){
      var slug = (a.getAttribute('href')||'').match(/P\d-\d{2}/);
      if (!slug) return;
      try {
        var state = localStorage.getItem('lesson-check-'+slug[0]);
        if (state && JSON.parse(state||'[]').length) done++;
      } catch(e){}
    });

    // 改造 h2 為 accordion header
    var label = h2.textContent.trim();
    h2.className = 'phase-head';
    h2.innerHTML = '';
    h2.setAttribute('role','button');
    h2.setAttribute('tabindex','0');
    h2.setAttribute('aria-expanded','true');

    var title = document.createElement('span');
    title.className = 'phase-title';
    title.textContent = label;

    var stat = document.createElement('span');
    stat.className = 'phase-stat';
    stat.textContent = done + ' / ' + total;

    var chev = document.createElement('span');
    chev.className = 'phase-chev';
    chev.textContent = '▾';

    h2.appendChild(title);
    h2.appendChild(stat);
    h2.appendChild(chev);

    // 插入 group 到 h2 之後
    h2.insertAdjacentElement('afterend', group);

    // 預設展開 / 折疊
    var key = 'phase-open-'+idx;
    var saved = null;
    try { saved = localStorage.getItem(key); } catch(e){}
    // 規則：第一個 Phase 預設展開；已完成=100% 預設折疊
    var defaultOpen = (idx === 0) || (done > 0 && done < total);
    var open = saved === null ? defaultOpen : saved === '1';
    setOpen(open);

    function setOpen(o){
      open = o;
      group.classList.toggle('is-open', o);
      h2.setAttribute('aria-expanded', o ? 'true' : 'false');
      h2.classList.toggle('is-open', o);
      try { localStorage.setItem(key, o ? '1' : '0'); } catch(e){}
    }
    h2.addEventListener('click', function(){ setOpen(!open); });
    h2.addEventListener('keydown', function(e){
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setOpen(!open); }
    });
  });

  // 看看 referrer 來自哪個 lesson → 自動展開那個 Phase
  try {
    var last = sessionStorage.getItem('last-lesson');
    if (last) {
      var m = last.match(/P(\d)-/);
      if (m) {
        var phaseIdx = parseInt(m[1]) - 1;
        var grp = main.querySelector('.phase-group[data-phase-idx="'+phaseIdx+'"]');
        if (grp && !grp.classList.contains('is-open')){
          grp.classList.add('is-open');
          var hd = grp.previousElementSibling;
          if (hd && hd.classList.contains('phase-head')) {
            hd.classList.add('is-open');
            hd.setAttribute('aria-expanded','true');
          }
        }
      }
    }
  } catch(e){}
})();
