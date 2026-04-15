/* lesson-tabs.js — 把 lesson 頁內容分成 4 個 tab（講解 / 練習 / 知識 / 參考）。
 * 執行順序：必須在 xlsx-integrator.js 之後才會看到注入內容，所以也用 defer
 * 並確保腳本引用順序正確。以 MutationObserver 備援處理延遲注入情況。
 */
(function(){
  function init(){
    var lesson = document.querySelector('main.lesson');
    if (!lesson) return;
    if (lesson.dataset.tabsInited) return;

    var mdBody = lesson.querySelector('.md-body');
    if (!mdBody) return;

    var slug = document.body.getAttribute('data-lesson-slug') || 'default';
    var KEY = 'lesson-tab-' + slug;
    try { sessionStorage.setItem('last-lesson', slug); } catch(e){}

    // 1. 蒐集 tab 內容
    var studyNodes = [];   // 原有講解（from .md-body 內容）
    var practiceNodes = [];// 互動 + 微任務
    var knowledgeNodes = [];// 深度知識 + VBA
    var referenceNodes = [];// 快捷鍵 + meta
    var metaNodes = [];    // 學習小檔案（永遠固定顯示）

    // 先把 .md-body 既有的子元素全部視為「講解」內容
    // 但排除後來 xlsx-integrator 插入的 .xc-section（這些會另外分派）
    Array.from(mdBody.children).forEach(function(n){
      if (n.classList && n.classList.contains('xc-section')) return;
      studyNodes.push(n);
    });

    // 分派 xlsx-integrator 的 sections（可能在 .lesson 下面，不一定在 .md-body 裡）
    lesson.querySelectorAll('.xc-section[data-xc-type]').forEach(function(sec){
      var t = sec.getAttribute('data-xc-type');
      if (t === 'meta') metaNodes.push(sec);
      else if (t === 'practice-inter' || t === 'practice-pro' || t === 'hands') practiceNodes.push(sec);
      else if (t === 'knowledge' || t === 'vba') knowledgeNodes.push(sec);
      else if (t === 'shortcuts') referenceNodes.push(sec);
    });

    // 清空 mdBody（準備重裝）
    while (mdBody.firstChild) mdBody.removeChild(mdBody.firstChild);

    // 把孤兒 .xc-integrated 容器（若存在）移除（內容已被搬走）
    lesson.querySelectorAll('.xc-integrated').forEach(function(c){
      if (!c.children.length) c.remove();
    });

    // 2. 組 tabs 清單（僅保留有內容的 tab）
    var tabs = [
      { id: 'study', label: '講解', icon: '📖', nodes: studyNodes },
      { id: 'practice', label: '練習', icon: '🎯', nodes: practiceNodes },
      { id: 'knowledge', label: '知識', icon: '📚', nodes: knowledgeNodes },
      { id: 'reference', label: '參考', icon: '⌨️', nodes: referenceNodes },
    ].filter(function(t){ return t.nodes.length > 0; });

    if (tabs.length <= 1) {
      // 單 tab 無必要切分，把全部內容塞回去即可
      studyNodes.concat(practiceNodes,knowledgeNodes,referenceNodes).forEach(function(n){
        mdBody.appendChild(n);
      });
      metaNodes.forEach(function(n){ mdBody.insertBefore(n, mdBody.firstChild); });
      return;
    }

    // 3. 先把 meta 永遠顯示區塊放進 mdBody 最上面
    metaNodes.forEach(function(n){ mdBody.appendChild(n); });

    // 4. 建 tab bar
    var bar = document.createElement('div');
    bar.className = 'lt-bar';
    var badges = {
      practice: countItems(practiceNodes, '.xc-task, .xc-hands-item'),
      knowledge: countItems(knowledgeNodes, '.xc-know, .xc-vba'),
      reference: countItems(referenceNodes, '.xc-sc-list li'),
    };

    tabs.forEach(function(t){
      var btn = document.createElement('button');
      btn.className = 'lt-tab';
      btn.type = 'button';
      btn.setAttribute('data-tab', t.id);
      var cnt = badges[t.id];
      btn.innerHTML = '<span class="lt-icon">' + t.icon + '</span>' +
        '<span class="lt-label">' + t.label + '</span>' +
        (cnt ? '<span class="lt-badge">' + cnt + '</span>' : '');
      btn.addEventListener('click', function(){ show(t.id, true); });
      bar.appendChild(btn);
    });
    mdBody.appendChild(bar);

    // 5. 建 panel 並塞入對應節點
    var panels = {};
    tabs.forEach(function(t){
      var panel = document.createElement('div');
      panel.className = 'lt-panel';
      panel.setAttribute('data-panel', t.id);
      t.nodes.forEach(function(n){ panel.appendChild(n); });
      mdBody.appendChild(panel);
      panels[t.id] = panel;
    });

    // 6. 顯示：從 localStorage 讀，或預設第一個（優先非 study，若進度為 0 則 study）
    var saved = null;
    try { saved = localStorage.getItem(KEY); } catch(e){}
    var initial = saved && panels[saved] ? saved : tabs[0].id;
    show(initial, false);

    function show(id, persist){
      Object.keys(panels).forEach(function(k){
        panels[k].classList.toggle('is-active', k === id);
      });
      bar.querySelectorAll('.lt-tab').forEach(function(b){
        b.classList.toggle('is-active', b.getAttribute('data-tab') === id);
      });
      if (persist) {
        try { localStorage.setItem(KEY, id); } catch(e){}
        // 切換時捲動到 tab bar 頂端
        var rect = bar.getBoundingClientRect();
        if (rect.top < 0 || rect.top > window.innerHeight * 0.6) {
          var y = window.pageYOffset + rect.top - 80;
          window.scrollTo({ top: y, behavior: 'smooth' });
        }
      }
    }

    function countItems(nodes, sel){
      var n = 0;
      nodes.forEach(function(x){ n += x.querySelectorAll(sel).length; });
      return n;
    }

    lesson.dataset.tabsInited = '1';
  }

  // 延後到 xlsx-integrator 完成後
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function(){
      setTimeout(init, 0);
    });
  } else {
    setTimeout(init, 0);
  }
})();
