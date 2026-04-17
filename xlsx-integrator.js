import { XLSX_CONTENT } from './xlsx-content.js';

/* xlsx-integrator.js
 * 把 Excel互動練習 / Excel專業養成 的內容深度整合進 lesson 頁。
 * 依 body[data-lesson-slug] 自動找對應內容並注入到主內容區末端。
 */
(function(){
  if (!XLSX_CONTENT) return;
  var slug = document.body.getAttribute('data-lesson-slug');
  if (!slug) return;
  var data = XLSX_CONTENT.lessons?.[slug];
  if (!data) return;

  // 工具
  function el(tag, props, children){
    var n = document.createElement(tag);
    if (props) for (var k in props){
      if (k === 'class') n.className = props[k];
      else if (k === 'html') n.innerHTML = props[k];
      else if (k === 'text') n.textContent = props[k];
      else if (k.indexOf('on') === 0) n.addEventListener(k.slice(2), props[k]);
      else n.setAttribute(k, props[k]);
    }
    if (children) (Array.isArray(children) ? children : [children]).forEach(function(c){
      if (c == null) return;
      n.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    });
    return n;
  }
  function escapeHtml(s){
    return String(s||'').replace(/[&<>"']/g, function(c){
      return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];
    });
  }
  // 公式正規化（比較用）
  function normFormula(s){
    return String(s||'').trim().toUpperCase()
      .replace(/\s+/g,'')
      .replace(/[，]/g,',')
      .replace(/[（]/g,'(').replace(/[）]/g,')')
      .replace(/[＝]/g,'=')
      .replace(/^=/,''); // 允許省略 =
  }

  // ---------- meta 條 ----------
  function buildMetaSection(){
    var m = data.meta;
    if (!m) return null;
    var chips = [];
    if (m.difficulty) chips.push({label:'難度', val:m.difficulty});
    if (m.taskCount)  chips.push({label:'題數', val:m.taskCount});
    if (m.time)       chips.push({label:'建議時間', val:m.time});
    if (m.xp)         chips.push({label:'經驗值', val:m.xp});
    if (m.topics)     chips.push({label:'涵蓋', val:m.topics});
    if (!chips.length) return null;
    var bar = el('div',{class:'xc-meta'},
      chips.map(function(c){
        return el('span',{class:'xc-chip'},[
          el('span',{class:'xc-chip-label',text:c.label+'：'}),
          document.createTextNode(c.val)
        ]);
      })
    );
    return bar;
  }

  // ---------- 練習資料表 ----------
  function buildDataTable(header, rows, caption){
    if (!rows || !rows.length) return null;
    var thead = el('thead',null,el('tr',null,
      header.filter(Boolean).map(function(h){return el('th',{text:h});})));
    var tbody = el('tbody',null,
      rows.map(function(r){
        return el('tr',null, r.map(function(c){
          return el('td',{text: c === '' ? '' : String(c)});
        }));
      }));
    var wrap = el('div',{class:'xc-data-wrap'});
    if (caption) wrap.appendChild(el('div',{class:'xc-data-cap',text:caption}));
    wrap.appendChild(el('table',{class:'xc-data'},[thead,tbody]));
    return wrap;
  }

  // ---------- 互動題卡（公式輸入 → 即時驗證）----------
  function buildInteractiveTask(t, progressUpdate){
    var card = el('div',{class:'xc-task'});
    var head = el('div',{class:'xc-task-head'},[
      el('span',{class:'xc-task-num',text:'#'+t.num}),
      el('span',{class:'xc-diff',text:t.difficulty||''}),
    ]);
    card.appendChild(head);
    card.appendChild(el('div',{class:'xc-task-desc',text:t.desc}));

    var input = el('input',{class:'xc-input', type:'text',
      placeholder:'輸入你的公式…例如 =SUM(A1:A10)', autocomplete:'off',
      spellcheck:'false'});
    var feedback = el('div',{class:'xc-feedback'});
    var hintBox = el('div',{class:'xc-reveal'},[
      el('span',{class:'xc-reveal-label',text:'💡 提示'}),
      document.createTextNode(' '+(t.hint||'(無提示)'))
    ]);
    var ansBox = el('div',{class:'xc-reveal'},[
      el('span',{class:'xc-reveal-label',text:'✅ 參考答案'}),
      el('code',{text:t.answer||''})
    ]);

    var checkBtn = el('button',{class:'xc-btn primary',type:'button',
      onclick:function(){ check(); }
    },'檢查');

    function check(){
      var v = input.value;
      if (!v.trim()) {
        feedback.textContent = '先輸入公式再檢查 🙂';
        feedback.className = 'xc-feedback';
        return;
      }
      // 數字答案：嘗試也比對直接的數值
      var ok = normFormula(v) === normFormula(t.answer);
      if (!ok) {
        // 若答案是純數字（互動練習偶有），也允許數字相等
        var num = parseFloat(v.replace(/[,，]/g,''));
        var ans = parseFloat(String(t.answer||'').replace(/[,，]/g,''));
        if (!isNaN(num) && !isNaN(ans) && Math.abs(num-ans) < 0.001) ok = true;
      }
      if (ok){
        input.className = 'xc-input is-correct';
        card.classList.add('is-correct');
        card.classList.remove('is-wrong');
        feedback.textContent = '✅ 正確！再試下一題';
        feedback.className = 'xc-feedback is-correct';
        ansBox.classList.add('show');
        progressUpdate(t.num, true);
      } else {
        input.className = 'xc-input is-wrong';
        card.classList.add('is-wrong');
        card.classList.remove('is-correct');
        feedback.textContent = '❌ 還不對，看一下提示？';
        feedback.className = 'xc-feedback is-wrong';
      }
    }
    input.addEventListener('keydown', function(e){
      if (e.key === 'Enter') check();
    });
    input.addEventListener('input', function(){
      if (input.classList.contains('is-wrong')) {
        input.classList.remove('is-wrong');
        feedback.textContent = '';
      }
    });

    card.appendChild(el('div',{class:'xc-input-row'},[input, checkBtn]));
    card.appendChild(feedback);
    card.appendChild(el('div',{class:'xc-extra'},[
      el('button',{class:'xc-btn ghost',type:'button',
        onclick:function(){ hintBox.classList.toggle('show'); }
      },'💡 看提示'),
      el('button',{class:'xc-btn ghost',type:'button',
        onclick:function(){ ansBox.classList.toggle('show'); }
      },'👁 顯示答案'),
    ]));
    card.appendChild(hintBox);
    card.appendChild(ansBox);
    return card;
  }

  // ---------- 專業養成題卡（含原理解說）----------
  function buildProTask(t){
    var card = el('div',{class:'xc-task'});
    card.appendChild(el('div',{class:'xc-task-head'},[
      el('span',{class:'xc-task-num',text:t.numLabel||''}),
      el('span',{class:'xc-task-time',text:t.time||''})
    ]));
    card.appendChild(el('div',{class:'xc-task-desc',text:t.desc}));

    var hintBox = el('div',{class:'xc-reveal'},[
      el('span',{class:'xc-reveal-label',text:'💡 提示'}),
      el('code',{text:t.hint||''})
    ]);
    var ansBox = el('div',{class:'xc-reveal'},[
      el('span',{class:'xc-reveal-label',text:'✅ 解答'}),
      el('code',{text:t.answer||''}),
      t.explain ? el('div',{class:'xc-know-extra',text:'🧠 '+t.explain}) : null
    ].filter(Boolean));

    card.appendChild(el('div',{class:'xc-extra'},[
      el('button',{class:'xc-btn ghost',type:'button',
        onclick:function(){ hintBox.classList.toggle('show'); }
      },'💡 提示'),
      el('button',{class:'xc-btn ghost',type:'button',
        onclick:function(){ ansBox.classList.toggle('show'); }
      },'👁 答案 + 解析'),
    ]));
    card.appendChild(hintBox);
    card.appendChild(ansBox);
    return card;
  }

  // ---------- 動手練習（checklist 型：只打勾、不驗證答案）----------
  function buildHandsOn(tasks){
    if (!tasks || !tasks.length) return null;
    var KEY = 'xc-hands-'+slug;
    var saved = {};
    try { saved = JSON.parse(localStorage.getItem(KEY) || '{}'); } catch(e){}
    var wrap = el('div',{class:'xc-hands'});
    var bar = el('div',{class:'xc-hands-bar'},[
      el('span',{class:'xc-hands-count'}),
      el('div',{class:'xc-hands-track'}, el('span'))
    ]);
    wrap.appendChild(bar);
    function updateBar(){
      var done = 0;
      tasks.forEach(function(t){ if (saved[t.num]) done++; });
      wrap.querySelector('.xc-hands-count').textContent = done + ' / ' + tasks.length + ' 完成';
      wrap.querySelector('.xc-hands-track > span').style.width = (done/tasks.length*100)+'%';
    }
    tasks.forEach(function(t){
      var item = el('label',{class:'xc-hands-item'});
      var cb = el('input',{type:'checkbox'});
      if (saved[t.num]) { cb.checked = true; item.classList.add('is-done'); }
      cb.addEventListener('change', function(){
        saved[t.num] = cb.checked;
        item.classList.toggle('is-done', cb.checked);
        try { localStorage.setItem(KEY, JSON.stringify(saved)); } catch(e){}
        updateBar();
      });
      // 難度標籤：把 emoji 換成色點
      var diff = String(t.difficulty||'').trim();
      var dotClass = 'is-easy';
      if (/🔵|標準|基礎/.test(diff)) dotClass = 'is-normal';
      else if (/🟡|變化|進階/.test(diff)) dotClass = 'is-harder';
      else if (/🟠|高階/.test(diff)) dotClass = 'is-hard';
      else if (/🔴|挑戰|專家/.test(diff)) dotClass = 'is-challenge';
      var diffText = diff.replace(/^[🟢🔵🟡🟠🔴]\s*/u,'') || '暖身';
      item.appendChild(cb);
      item.appendChild(el('span',{class:'xc-hands-box'}));
      item.appendChild(el('span',{class:'xc-hands-num',text:String(t.num).padStart(2,'0')}));
      item.appendChild(el('span',{class:'xc-diff-pill '+dotClass},[
        el('span',{class:'xc-diff-dot'}),
        document.createTextNode(diffText)
      ]));
      item.appendChild(el('span',{class:'xc-hands-desc',text:t.desc}));
      wrap.appendChild(item);
    });
    updateBar();
    return wrap;
  }

  // ---------- 知識型 sections ----------
  function buildKnowledge(k){
    var nodes = [];
    if (k.subtitle) nodes.push(el('div',{class:'xc-sub',text:k.subtitle}));
    (k.sections || []).forEach(function(sec){
      var card = el('div',{class:'xc-know'});
      card.appendChild(el('h3',null,sec.title));
      var ul = el('ul',{class:'xc-know-list'});
      (sec.items || []).forEach(function(item){
        if (!item || !item.length) return;
        var key = item[0];
        var rest = item.slice(1).filter(function(x){return x!=null&&x!=='';});
        var bodyHtml = rest.length ? rest.map(escapeHtml).join('<span class="xc-know-extra">') + (rest.length>1 ? '</span>'.repeat(rest.length-1) : '') : '';
        var li = el('li',null,[
          el('span',{class:'xc-know-key',text:key}),
          el('div',{class:'xc-know-body',html:bodyHtml || ''})
        ]);
        ul.appendChild(li);
      });
      card.appendChild(ul);
      nodes.push(card);
    });
    return nodes;
  }

  // ---------- 快捷鍵 groups ----------
  function buildShortcuts(s){
    var nodes = [];
    if (s.intro) nodes.push(el('div',{class:'xc-sub',text:s.intro}));
    (s.groups || []).forEach(function(g){
      if (!g.items || !g.items.length) {
        if (g.title) nodes.push(el('div',{class:'xc-sub',text:g.title}));
        return;
      }
      var card = el('div',{class:'xc-sc-group'});
      card.appendChild(el('h3',null,g.title));
      var ul = el('ul',{class:'xc-sc-list'});
      g.items.forEach(function(it){
        var li = el('li',null,[
          el('span',{class:'xc-kbd',text:it.key}),
          el('div',{class:'xc-sc-desc'},[
            document.createTextNode(it.desc||''),
            it.note ? el('span',{class:'xc-sc-note',text:it.note}) : null
          ].filter(Boolean))
        ]);
        ul.appendChild(li);
      });
      card.appendChild(ul);
      nodes.push(card);
    });
    return nodes;
  }

  // ---------- VBA microTasks + practiceTasks ----------
  function buildVBA(v){
    var nodes = [];
    if (v.warning) nodes.push(
      el('div',{class:'callout',html:'<strong>⚠️ 注意</strong>'+escapeHtml(v.warning)})
    );
    if (v.quickOps && v.quickOps.length){
      var ops = el('div',{class:'xc-know'});
      ops.appendChild(el('h3',null,'⌨️ 快速操作指南'));
      var ul = el('ul',{class:'xc-know-list'});
      v.quickOps.forEach(function(op){
        ul.appendChild(el('li',null,[
          el('span',{class:'xc-know-key',text:op.op}),
          el('div',{class:'xc-know-body',text:op.how})
        ]));
      });
      ops.appendChild(ul);
      nodes.push(ops);
    }
    (v.microTasks || []).forEach(function(t){
      var box = el('div',{class:'xc-vba'});
      var copyBtn = el('button',{class:'xc-copy',type:'button'}, '📋 複製');
      var codeLines = Array.isArray(t.code) ? t.code : [];
      var pre = el('pre',null,[el('code',{text: codeLines.join('\n')})]);
      copyBtn.addEventListener('click', function(){
        navigator.clipboard.writeText(codeLines.join('\n')).then(function(){
          copyBtn.textContent = '✓ 已複製';
          copyBtn.classList.add('copied');
          setTimeout(function(){
            copyBtn.textContent = '📋 複製';
            copyBtn.classList.remove('copied');
          }, 1600);
        });
      });
      box.appendChild(el('div',{class:'xc-vba-head'},[
        el('div',{class:'xc-vba-title',text:t.title}),
        el('span',{class:'xc-vba-time',text:t.time||''}),
        copyBtn
      ]));
      box.appendChild(pre);
      if (t.tip) box.appendChild(el('div',{class:'xc-vba-tip',text:t.tip}));
      nodes.push(box);
    });
    if (v.practiceTasks && v.practiceTasks.length){
      var p = el('div',{class:'xc-know'});
      p.appendChild(el('h3',null,'🎯 動手練習'));
      var ul2 = el('ul',{class:'xc-know-list'});
      v.practiceTasks.forEach(function(pt){
        ul2.appendChild(el('li',null,[
          el('span',{class:'xc-know-key',text:'#'+pt.num+' '+(pt.difficulty||'')}),
          el('div',{class:'xc-know-body',text:pt.desc})
        ]));
      });
      p.appendChild(ul2);
      nodes.push(p);
    }
    return nodes;
  }

  // ---------- 組裝 sections ----------
  var article = document.querySelector('.lesson .md-body') || document.querySelector('.lesson');
  if (!article) return;

  var container = el('div', {class:'xc-integrated'});

  var meta = buildMetaSection();
  if (meta) {
    var mwrap = el('div',{class:'xc-section','data-xc-type':'meta'});
    mwrap.appendChild(el('h2',null,[el('span',{class:'xc-emoji',text:'📊'}), document.createTextNode(' 學習小檔案')]));
    mwrap.appendChild(meta);
    container.appendChild(mwrap);
  }

  // 1. 快捷鍵（P1-01 專屬）
  if (data.shortcuts){
    var sec = el('div',{class:'xc-section','data-xc-type':'shortcuts'});
    sec.appendChild(el('h2',null,[el('span',{class:'xc-emoji',text:'⌨️'}),document.createTextNode(' macOS 完整快捷鍵清單（120+ 條）')]));
    buildShortcuts(data.shortcuts).forEach(function(n){ sec.appendChild(n); });
    container.appendChild(sec);
  }

  // 2. 知識型 sections
  if (data.knowledge){
    var sec = el('div',{class:'xc-section','data-xc-type':'knowledge'});
    sec.appendChild(el('h2',null,[el('span',{class:'xc-emoji',text:'📚'}),document.createTextNode(' 深度知識 — 步驟、技巧、常見錯誤')]));
    buildKnowledge(data.knowledge).forEach(function(n){ sec.appendChild(n); });
    container.appendChild(sec);

    // 2b. 動手練習 checklist（若該 sheet 有）
    if (data.knowledge?.handsTasks?.length){
      var hsec = el('div',{class:'xc-section','data-xc-type':'hands'});
      hsec.appendChild(el('h2',null,[el('span',{class:'xc-emoji',text:'🎯'}),document.createTextNode(' 動手練習 — 跟著做、做完打勾')]));
      hsec.appendChild(el('div',{class:'xc-sub',text:'以下是循序漸進的實作任務。打開 Excel 跟著做，做完就打勾。進度會自動記住。'}));
      var h = buildHandsOn(data.knowledge?.handsTasks);
      if (h) hsec.appendChild(h);
      container.appendChild(hsec);
    }
  }

  // 3. VBA
  if (data.vba){
    var sec = el('div',{class:'xc-section','data-xc-type':'vba'});
    sec.appendChild(el('h2',null,[el('span',{class:'xc-emoji',text:'⚙️'}),document.createTextNode(' VBA 程式碼範例（複製即用）')]));
    buildVBA(data.vba).forEach(function(n){ sec.appendChild(n); });
    container.appendChild(sec);
  }

  // 4. 互動練習（含資料表 + 即時驗證）
  if (data.inter){
    var interTasks = data.inter?.tasks || [];
    var interTaskCount = data.inter?.tasks?.length ?? 0;
    var sec = el('div',{class:'xc-section','data-xc-type':'practice-inter'});
    sec.appendChild(el('h2',null,[el('span',{class:'xc-emoji',text:'🎯'}),document.createTextNode(' 互動練習 — 輸入公式即時驗證')]));
    sec.appendChild(el('div',{class:'xc-sub',text:'規則：在輸入框寫公式（可省略開頭 =）→ 按 Enter 或「檢查」→ 答對自動變綠 ✅，答錯變紅 ❌'}));
    var table = buildDataTable(data.inter?.dataHeader || [], data.inter?.dataRows || [], '【練習資料】');
    if (table) sec.appendChild(table);
    var progress = el('div',{class:'xc-progress'},[
      el('span',{text:'答對 '}),
      el('span',{class:'xc-progress-count',text:'0'}),
      el('span',{text:' / '+interTaskCount}),
      el('div',{class:'xc-progress-bar'},el('span'))
    ]);
    sec.appendChild(progress);
    var done = {};
    function update(num, ok){
      done[num] = ok;
      var n = Object.keys(done).filter(function(k){return done[k];}).length;
      sec.querySelector('.xc-progress-count').textContent = n;
      sec.querySelector('.xc-progress-bar > span').style.width = interTaskCount ? (n/interTaskCount*100)+'%' : '0%';
      // 持久化
      try {
        localStorage.setItem('xc-inter-'+slug, JSON.stringify(done));
      } catch(e){}
    }
    // 載入既有進度
    try {
      var saved = JSON.parse(localStorage.getItem('xc-inter-'+slug)||'{}');
      done = saved;
      var n = Object.keys(done).filter(function(k){return done[k];}).length;
      sec.querySelector('.xc-progress-count').textContent = n;
      sec.querySelector('.xc-progress-bar > span').style.width = interTaskCount ? (n/interTaskCount*100)+'%' : '0%';
    } catch(e){}
    interTasks.forEach(function(t){
      sec.appendChild(buildInteractiveTask(t, update));
    });
    container.appendChild(sec);
  }

  // 5. 專業養成題目（含解析）
  if (data.pro){
    var sec = el('div',{class:'xc-section','data-xc-type':'practice-pro'});
    sec.appendChild(el('h2',null,[el('span',{class:'xc-emoji',text:'🧪'}),document.createTextNode(' 微任務練習 — 含原理解說')]));
    sec.appendChild(el('div',{class:'xc-sub',text:'每題提供「💡 提示」與「✅ 答案 + 🧠 解析」按鈕。建議先自己想、再翻答案'}));
    var table = buildDataTable(data.pro?.dataHeader || [], data.pro?.dataRows || [], '【練習資料】');
    if (table) sec.appendChild(table);
    (data.pro?.tasks || []).forEach(function(t){
      sec.appendChild(buildProTask(t));
    });
    container.appendChild(sec);
  }

  // 插入到 lesson-nav 之前
  var nav = document.querySelector('.lesson-nav');
  if (nav) {
    nav.parentNode.insertBefore(container, nav);
  } else {
    article.appendChild(container);
  }
  document.dispatchEvent(new CustomEvent('xlsx-integrator:ready'));
})();
