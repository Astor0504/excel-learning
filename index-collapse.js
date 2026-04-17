/* index-collapse.js — 首頁 Phase 折疊 / 進度聚合 */
(function(){
  if (!document.querySelector('.hero')) return;

  var main = document.querySelector('main');
  if (!main) return;

  var sections = Array.from(main.querySelectorAll('.phase-section'));
  if (!sections.length) return;

  sections.forEach(function(section, idx){
    var h2 = section.querySelector('h2');
    if (!h2) return;

    var group = document.createElement('div');
    group.className = 'phase-group';
    group.setAttribute('data-phase-idx', idx);
    var lessons = [];

    var node = section.nextElementSibling;
    while (node && !node.classList.contains('phase-section')){
      var next = node.nextElementSibling;
      if (node.matches('a[data-lesson]')) {
        lessons.push(node);
        group.appendChild(node);
      }
      node = next;
    }

    var total = lessons.length;
    var done = 0;
    lessons.forEach(function(a){
      var slug = (a.getAttribute('href')||'').match(/P\d-\d{2}/);
      if (!slug) return;
      try {
        if (localStorage.getItem('done:lessons/' + slug[0] + '.html')) done++;
      } catch(e){}
    });

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

    section.insertAdjacentElement('afterend', group);

    var key = 'phase-open-'+idx;
    var saved = null;
    try { saved = localStorage.getItem(key); } catch(e){}
    var defaultOpen = (idx === 0) || (done > 0 && done < total);
    var open = saved === null ? defaultOpen : saved === '1';
    setOpen(open);

    function setOpen(o){
      open = o;
      group.classList.toggle('is-open', o);
      h2.setAttribute('aria-expanded', o ? 'true' : 'false');
      h2.classList.toggle('is-open', o);
      section.classList.toggle('is-open', o);
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
          var sec = grp.previousElementSibling;
          var hd = sec && sec.querySelector('.phase-head');
          if (sec) sec.classList.add('is-open');
          if (hd) {
            hd.classList.add('is-open');
            hd.setAttribute('aria-expanded', 'true');
          }
        }
      }
    }
  } catch(e){}
})();
