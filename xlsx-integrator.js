import { XLSX_CONTENT } from './xlsx-content.js';
import { LESSON_DEMOS } from './lesson-demos-data.js';

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
  function clearNode(node){
    while (node.firstChild) node.removeChild(node.firstChild);
  }
  function forceSwap(node){
    node.classList.remove('is-swapping');
    void node.offsetWidth;
    node.classList.add('is-swapping');
  }
  function applyToneClasses(node, tone){
    String(tone || '').split(/\s+/).forEach(function(token){
      if (token) node.classList.add('is-' + token);
    });
  }
  function columnLabelToIndex(label){
    var s = String(label || '').toUpperCase().replace(/[^A-Z]/g,'');
    var n = 0;
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n ? n - 1 : -1;
  }
  function parseCellRef(ref){
    var m = String(ref || '').toUpperCase().match(/^([A-Z]+)(\d+)$/);
    if (!m) return null;
    return { col: columnLabelToIndex(m[1]), row: Math.max(0, parseInt(m[2], 10) - 1) };
  }
  function syncSheetScene(card){
    var stage = card.querySelector('.xc-demo-sheet-stage');
    var win = card.querySelector('.xc-demo-sheet-window');
    if (!stage || !win) return;
    var active = win.querySelector('.xc-demo-sheet-cell.is-active');
    var rowHeads = win.querySelectorAll('.xc-demo-sheet-rowhead');
    var colHeads = win.querySelectorAll('.xc-demo-sheet-colhead');
    rowHeads.forEach(function(node){ node.classList.remove('is-active'); });
    colHeads.forEach(function(node){ node.classList.remove('is-active'); });

    if (active) {
      var rowIndex = active.parentNode ? (active.parentNode.rowIndex - 1) : -1;
      var colIndex = typeof active.cellIndex === 'number' ? active.cellIndex - 1 : -1;
      if (rowHeads[rowIndex]) rowHeads[rowIndex].classList.add('is-active');
      if (colHeads[colIndex]) colHeads[colIndex].classList.add('is-active');

      var cursor = win.querySelector('.xc-demo-sheet-cursor');
      if (cursor) {
        var winRect = win.getBoundingClientRect();
        var cellRect = active.getBoundingClientRect();
        cursor.style.setProperty('--cursor-x', Math.round(cellRect.left - winRect.left + cellRect.width - 6) + 'px');
        cursor.style.setProperty('--cursor-y', Math.round(cellRect.top - winRect.top + cellRect.height - 6) + 'px');
      }
    }

    stage.classList.remove('is-ready');
    win.classList.remove('is-ready');
    requestAnimationFrame(function(){
      stage.classList.add('is-ready');
      win.classList.add('is-ready');
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

  // ---------- 動畫 / 操作 demo ----------
  function buildSheetPanel(panel){
    var card = el('div',{class:'xc-demo-panel xc-demo-sheet-panel'});
    if (panel.title) card.appendChild(el('div',{class:'xc-demo-panel-title',text:panel.title}));

    var shell = el('div',{class:'xc-demo-sheet-shell'});
    var window = el('div',{class:'xc-demo-sheet-window'});
    var toolbar = el('div',{class:'xc-demo-sheet-toolbar'},[
      el('div',{class:'xc-demo-sheet-dots'},[
        el('span'), el('span'), el('span')
      ]),
      el('div',{class:'xc-demo-sheet-ribbon'},[
        el('span',{class:'xc-demo-sheet-ribbon-tab is-active',text:panel.sheetName || 'Sheet1'}),
        panel.context ? el('span',{class:'xc-demo-sheet-context',text:panel.context}) : null
      ].filter(Boolean))
    ]);
    var formulaCopy = el('span',{class:'xc-demo-sheet-formula-copy',text:panel.formula || '準備操作這張工作表'});
    var formula = el('div',{class:'xc-demo-sheet-formula'},[
      el('div',{class:'xc-demo-sheet-namebox',text:panel.nameBox || panel.activeCell || 'A1'}),
      el('div',{class:'xc-demo-sheet-fx',text:'fx'}),
      el('div',{class:'xc-demo-sheet-formula-text'}, formulaCopy)
    ]);
    formula.querySelector('.xc-demo-sheet-formula-text').style.setProperty('--formula-chars', Math.max(12, String(panel.formula || '準備操作這張工作表').length));
    var table = el('table',{class:'xc-demo-sheet-grid'});
    var cols = panel.columns || [];

    table.appendChild(el('thead',null,el('tr',null,[
      el('th',{class:'xc-demo-sheet-corner',text:''})
    ].concat(cols.map(function(col){
      return el('th',{class:'xc-demo-sheet-colhead',text:col});
    })))));

    table.appendChild(el('tbody',null,(panel.rows || []).map(function(row, rowIdx){
      return el('tr',null,[
        el('th',{class:'xc-demo-sheet-rowhead',text:String(rowIdx + 1)})
      ].concat((row || []).map(function(cell, colIdx){
        var meta = (cell && typeof cell === 'object' && !Array.isArray(cell)) ? cell : { text: cell };
        var colRef = cols[colIdx] || String.fromCharCode(65 + colIdx);
        var ref = colRef + (rowIdx + 1);
        var td = el('td',{class:'xc-demo-sheet-cell',text:meta.text == null ? '' : String(meta.text)});
        if (meta.tone) applyToneClasses(td, meta.tone);
        if (meta.align === 'right' || (/^-?[\d,.%]+$/.test(String(meta.text || '')) && String(meta.text || '').length)) {
          td.classList.add('is-right');
        }
        if (meta.strong) td.classList.add('is-strong');
        if ((panel.activeCell && String(panel.activeCell).toUpperCase() === ref.toUpperCase()) || meta.active) {
          td.classList.add('is-active');
        }
        return td;
      })));
    })));

    window.appendChild(toolbar);
    window.appendChild(formula);
    window.appendChild(el('div',{class:'xc-demo-sheet-scroll'},table));
    window.appendChild(el('div',{class:'xc-demo-sheet-tabs'},[
      el('span',{class:'xc-demo-sheet-tab is-active',text:panel.sheetName || 'Sheet1'})
    ]));
    window.appendChild(el('div',{class:'xc-demo-sheet-cursor','aria-hidden':'true'}));

    var stageChildren = [window];
    if (panel.sidebar && (panel.sidebar.title || (panel.sidebar.items || []).length)){
      stageChildren.push(
        el('aside',{class:'xc-demo-sheet-side'},[
          panel.sidebar.title ? el('div',{class:'xc-demo-sheet-side-title',text:panel.sidebar.title}) : null,
          panel.sidebar.items && panel.sidebar.items.length
            ? el('ul',{class:'xc-demo-sheet-side-list'},
                panel.sidebar.items.map(function(item){ return el('li',{text:item}); }))
            : null
        ].filter(Boolean))
      );
    }

    shell.appendChild(el('div',{class:'xc-demo-sheet-stage' + (stageChildren.length > 1 ? ' has-side' : '')},stageChildren));
    if (panel.footer) shell.appendChild(el('div',{class:'xc-demo-sheet-footer',text:panel.footer}));
    card.appendChild(shell);
    requestAnimationFrame(function(){ syncSheetScene(card); });
    setTimeout(function(){ syncSheetScene(card); }, 160);
    return card;
  }

  function buildDemoMedia(media){
    if (!media || !media.src) return null;
    var fig = el('figure',{class:'xc-demo-media'});
    var headBits = [];
    if (media.kicker) headBits.push(el('span',{class:'xc-demo-media-badge',text:media.kicker}));
    if (media.title) headBits.push(el('div',{class:'xc-demo-media-title',text:media.title}));
    if (headBits.length) fig.appendChild(el('figcaption',{class:'xc-demo-media-head'},headBits));

    var picture = el('picture',{class:'xc-demo-media-frame'});
    if (media.webp) picture.appendChild(el('source',{type:'image/webp',srcset:media.webp}));
    picture.appendChild(el('img',{
      src: media.src,
      alt: media.alt || media.title || '操作示範短片',
      loading: 'lazy',
      decoding: 'async'
    }));
    fig.appendChild(picture);
    if (media.note) fig.appendChild(el('div',{class:'xc-demo-media-note',text:media.note}));
    return fig;
  }

  function buildDemoPanel(panel){
    if (panel.type === 'sheet') return buildSheetPanel(panel);
    var card = el('div',{class:'xc-demo-panel'});
    if (panel.title) card.appendChild(el('div',{class:'xc-demo-panel-title',text:panel.title}));
    if (panel.columns && panel.rows){
      var tableWrap = el('div',{class:'xc-demo-mini'});
      var table = el('table',{class:'xc-demo-table'});
      var thead = el('thead',null,el('tr',null,(panel.columns || []).map(function(col){
        return el('th',{text:col});
      })));
      var tbody = el('tbody',null,(panel.rows || []).map(function(row){
        return el('tr',null,(row || []).map(function(cell){
          return el('td',{text:cell == null ? '' : String(cell)});
        }));
      }));
      table.appendChild(thead);
      table.appendChild(tbody);
      tableWrap.appendChild(table);
      card.appendChild(tableWrap);
    } else if (panel.lines && panel.lines.length){
      card.appendChild(el('ul',{class:'xc-demo-lines'},
        panel.lines.map(function(line){ return el('li',{text:line}); })
      ));
    }
    return card;
  }

  function createBoardView(board){
    var shell = el('div',{class:'xc-demo-board'});
    if (board.title) shell.appendChild(el('div',{class:'xc-demo-board-title',text:board.title}));
    var table = el('table',{class:'xc-demo-table xc-demo-table-board'});
    var refs = [];

    if (board.columns && board.columns.length){
      table.appendChild(el('thead',null,el('tr',null,
        board.columns.map(function(col){ return el('th',{text:col}); })
      )));
    }

    table.appendChild(el('tbody',null,(board.rows || []).map(function(row){
      return el('tr',null,(row || []).map(function(cell){
        var meta = (cell && typeof cell === 'object' && !Array.isArray(cell))
          ? cell
          : { text: cell };
        var td = el('td',{text:meta.text == null ? '' : String(meta.text)});
        refs.push({ node: td, marks: meta.marks || {} });
        return td;
      }));
    })));

    shell.appendChild(table);

    function setStep(stepNum){
      refs.forEach(function(ref){
        ref.node.classList.remove('is-focus','is-match','is-result');
        var mark = ref.marks[String(stepNum)] || ref.marks[stepNum];
        if (mark) ref.node.classList.add('is-' + mark);
      });
    }

    return { node: shell, setStep: setStep };
  }

  function createDemoControls(totalSteps, onChange){
    var index = 0;
    var timer = null;
    var prevBtn = el('button',{class:'xc-btn ghost',type:'button'},'← 上一步');
    var playBtn = el('button',{class:'xc-btn primary',type:'button','aria-pressed':'false'},'▶ 播放');
    var nextBtn = el('button',{class:'xc-btn ghost',type:'button'},'下一步 →');
    var meterFill = el('span');
    var meter = el('div',{class:'xc-demo-meter'},meterFill);
    var status = el('span',{class:'xc-demo-status'});
    var shell = el('div',{class:'xc-demo-controls'},[
      el('div',{class:'xc-demo-transport'},[prevBtn, playBtn, nextBtn]),
      el('div',{class:'xc-demo-controls-side'},[meter, status])
    ]);

    function updateMeter(){
      meterFill.style.width = totalSteps ? (((index + 1) / totalSteps) * 100) + '%' : '0%';
      meter.classList.toggle('is-playing', !!timer);
    }
    function stop(){
      if (timer) {
        clearInterval(timer);
        timer = null;
      }
      playBtn.textContent = '▶ 播放';
      playBtn.setAttribute('aria-pressed','false');
      shell.classList.remove('is-playing');
      updateMeter();
    }
    function setIndex(nextIndex){
      index = Math.max(0, Math.min(totalSteps - 1, nextIndex));
      onChange(index, { stop: stop });
      status.textContent = '步驟 ' + (index + 1) + ' / ' + totalSteps;
      prevBtn.disabled = index === 0;
      nextBtn.disabled = index === totalSteps - 1;
      updateMeter();
    }
    function step(delta){
      stop();
      setIndex(index + delta);
    }
    prevBtn.addEventListener('click', function(){ step(-1); });
    nextBtn.addEventListener('click', function(){ step(1); });
    playBtn.addEventListener('click', function(){
      if (timer) {
        stop();
        return;
      }
      if (index === totalSteps - 1) setIndex(0);
      playBtn.textContent = '⏸ 暫停';
      playBtn.setAttribute('aria-pressed','true');
      shell.classList.add('is-playing');
      updateMeter();
      timer = setInterval(function(){
        if (index >= totalSteps - 1) {
          stop();
          return;
        }
        setIndex(index + 1);
      }, 2400);
    });
    setIndex(0);
    return {
      node: shell,
      setIndex: function(nextIndex){
        stop();
        setIndex(nextIndex);
      },
      stop: stop
    };
  }

  function buildFormulaDemo(demo){
    var card = el('section',{class:'xc-demo xc-demo-formula-card'});
    var summary = el('div',{class:'xc-demo-summary','aria-live':'polite'});
    var summaryStep = el('div',{class:'xc-demo-step-label'});
    var summaryTitle = el('div',{class:'xc-demo-summary-title'});
    var summaryText = el('div',{class:'xc-demo-summary-text'});
    var summaryCaption = el('div',{class:'xc-demo-summary-caption'});
    summary.appendChild(summaryStep);
    summary.appendChild(summaryTitle);
    summary.appendChild(summaryText);
    summary.appendChild(summaryCaption);

    var partButtons = [];
    var formulaBar = el('div',{class:'xc-demo-formula'});
    (demo.formulaParts || []).forEach(function(part){
      var cls = 'xc-demo-part';
      if (part.tone) cls += ' ' + part.tone;
      if (!part.step) cls += ' is-passive';
      var btn = el('button',{
        class:cls,
        type:'button',
        text:part.text || '',
        'data-step': part.step || ''
      });
      if (!part.step) btn.disabled = true;
      formulaBar.appendChild(btn);
      partButtons.push(btn);
    });

    var stepsRail = el('div',{class:'xc-demo-steps'});
    var stepButtons = [];
    (demo.steps || []).forEach(function(step, idx){
      var btn = el('button',{class:'xc-demo-stepbtn ' + (step.tone || ''),type:'button'},[
        el('span',{class:'xc-demo-stepnum',text:String(idx + 1).padStart(2,'0')}),
        el('span',{class:'xc-demo-stepcopy',text:step.label || ('步驟 ' + (idx + 1))})
      ]);
      stepsRail.appendChild(btn);
      stepButtons.push(btn);
    });

    var boardView = createBoardView(demo.board || {});
    var rightRail = el('div',{class:'xc-demo-side'},[
      demo.outcome ? el('div',{class:'xc-demo-outcome',text:demo.outcome}) : null,
      boardView.node
    ].filter(Boolean));

    var body = el('div',{class:'xc-demo-body'},[
      el('div',{class:'xc-demo-main'},[
        formulaBar,
        stepsRail,
        summary
      ]),
      rightRail
    ]);

    card.appendChild(el('div',{class:'xc-demo-head'},[
      el('div',{class:'xc-demo-kicker'},[
        demo.badge ? el('span',{class:'xc-demo-badge',text:demo.badge}) : null,
        demo.title ? el('h3',{class:'xc-demo-title',text:demo.title}) : null,
        demo.subtitle ? el('p',{class:'xc-demo-subtitle',text:demo.subtitle}) : null
      ].filter(Boolean))
    ]));
    if (demo.media){
      var mediaNode = buildDemoMedia(demo.media);
      if (mediaNode) card.appendChild(mediaNode);
    }
    card.appendChild(body);

    var controls = createDemoControls((demo.steps || []).length, function(idx, api){
      var step = (demo.steps || [])[idx];
      if (!step) return;
      partButtons.forEach(function(btn){
        var btnStep = Number(btn.getAttribute('data-step') || 0);
        btn.classList.toggle('is-active', btnStep === idx + 1);
      });
      stepButtons.forEach(function(btn, sidx){
        btn.classList.toggle('is-active', sidx === idx);
        btn.classList.toggle('is-done', sidx < idx);
      });
      summary.className = 'xc-demo-summary ' + (step.tone || '');
      summaryStep.textContent = '步驟 ' + (idx + 1);
      summaryTitle.textContent = step.label || '';
      summaryText.textContent = step.text || '';
      summaryCaption.textContent = step.caption || '';
      boardView.setStep(idx + 1);
      forceSwap(summary);
      if (idx === (demo.steps || []).length - 1) api.stop();
    });
    card.appendChild(controls.node);

    partButtons.forEach(function(btn){
      var btnStep = Number(btn.getAttribute('data-step') || 0);
      if (!btnStep) return;
      btn.addEventListener('click', function(){ controls.setIndex(btnStep - 1); });
    });
    stepButtons.forEach(function(btn, idx){
      btn.addEventListener('click', function(){ controls.setIndex(idx); });
    });

    return card;
  }

  function buildWorkflowDemo(demo){
    var card = el('section',{class:'xc-demo xc-demo-workflow-card'});
    var stageRail = el('div',{class:'xc-demo-stage-rail'});
    var stageButtons = [];
    (demo.stages || []).forEach(function(stage, idx){
      var btn = el('button',{class:'xc-demo-stage ' + (stage.tone || ''),type:'button'},[
        el('span',{class:'xc-demo-stage-num',text:String(idx + 1).padStart(2,'0')}),
        el('span',{class:'xc-demo-stage-text',text:stage.text || ''})
      ]);
      stageRail.appendChild(btn);
      stageButtons.push(btn);
    });

    var bodyTitle = el('div',{class:'xc-demo-summary-title'});
    var bodyText = el('div',{class:'xc-demo-summary-text'});
    var bodyCaption = el('div',{class:'xc-demo-summary-caption'});
    var bodyBox = el('div',{class:'xc-demo-summary','aria-live':'polite'},[
      el('div',{class:'xc-demo-step-label'}),
      bodyTitle,
      bodyText,
      bodyCaption
    ]);
    var bodyStep = bodyBox.querySelector('.xc-demo-step-label');

    var panelHost = el('div',{class:'xc-demo-panels'});
    var body = el('div',{class:'xc-demo-body xc-demo-body-single'},[
      el('div',{class:'xc-demo-main'},[
        stageRail,
        bodyBox,
        panelHost
      ]),
      demo.outcome ? el('aside',{class:'xc-demo-side'},[
        el('div',{class:'xc-demo-outcome',text:demo.outcome})
      ]) : null
    ].filter(Boolean));

    card.appendChild(el('div',{class:'xc-demo-head'},[
      el('div',{class:'xc-demo-kicker'},[
        demo.badge ? el('span',{class:'xc-demo-badge',text:demo.badge}) : null,
        demo.title ? el('h3',{class:'xc-demo-title',text:demo.title}) : null,
        demo.subtitle ? el('p',{class:'xc-demo-subtitle',text:demo.subtitle}) : null
      ].filter(Boolean))
    ]));
    if (demo.media){
      var mediaNode = buildDemoMedia(demo.media);
      if (mediaNode) card.appendChild(mediaNode);
    }
    card.appendChild(body);

    var controls = createDemoControls((demo.steps || []).length, function(idx, api){
      var step = (demo.steps || [])[idx];
      if (!step) return;
      var activeStage = Math.max(0, (step.stage || idx + 1) - 1);
      var activeStageMeta = (demo.stages || [])[activeStage] || {};
      stageButtons.forEach(function(btn, sidx){
        btn.classList.toggle('is-active', sidx === activeStage);
        btn.classList.toggle('is-done', sidx < activeStage);
      });
      bodyBox.className = 'xc-demo-summary ' + (step.tone || '');
      bodyStep.textContent = '階段 ' + (activeStage + 1) + ' · ' + (activeStageMeta.text || ('步驟 ' + (idx + 1)));
      bodyTitle.textContent = step.label || ('步驟 ' + (idx + 1));
      bodyText.textContent = step.text || '';
      bodyCaption.textContent = step.caption || '';
      clearNode(panelHost);
      (step.panels || []).forEach(function(panel){
        panelHost.appendChild(buildDemoPanel(panel));
      });
      forceSwap(bodyBox);
      forceSwap(panelHost);
      if (idx === (demo.steps || []).length - 1) api.stop();
    });
    card.appendChild(controls.node);

    stageButtons.forEach(function(btn, idx){
      btn.addEventListener('click', function(){ controls.setIndex(idx); });
    });

    return card;
  }

  function buildMediaDemo(demo){
    var card = el('section',{class:'xc-demo xc-demo-media-card'});
    card.appendChild(el('div',{class:'xc-demo-head'},[
      el('div',{class:'xc-demo-kicker'},[
        demo.badge ? el('span',{class:'xc-demo-badge',text:demo.badge}) : null,
        demo.title ? el('h3',{class:'xc-demo-title',text:demo.title}) : null,
        demo.subtitle ? el('p',{class:'xc-demo-subtitle',text:demo.subtitle}) : null
      ].filter(Boolean))
    ]));
    if (demo.media){
      var mediaNode = buildDemoMedia(demo.media);
      if (mediaNode) card.appendChild(mediaNode);
    }
    if (demo.outcome || demo.followup){
      card.appendChild(el('div',{class:'xc-demo-outcome xc-demo-outcome-compact'},[
        demo.outcome ? el('div',{class:'xc-demo-outcome-primary',text:demo.outcome}) : null,
        demo.followup ? el('div',{class:'xc-demo-outcome-next',text:demo.followup}) : null
      ].filter(Boolean)));
    }
    return card;
  }

  function buildDemoSection(demos){
    if (!demos || !demos.length) return null;
    var sec = el('div',{class:'xc-section','data-xc-type':'demo'});
    sec.appendChild(el('h2',null,[el('span',{class:'xc-emoji',text:'🎬'}),document.createTextNode(' 實作動畫範例')]));
    sec.appendChild(el('div',{class:'xc-sub',text:'操作課先看慢速短片抓節奏，公式課看拆解動畫抓邏輯。先看一次，再回 Excel 自己做，會比一開始就硬讀規則更容易進狀況。'}));
    demos.forEach(function(demo){
      if (demo.kind === 'formula') sec.appendChild(buildFormulaDemo(demo));
      else if (demo.kind === 'media') sec.appendChild(buildMediaDemo(demo));
      else sec.appendChild(buildWorkflowDemo(demo));
    });
    return sec;
  }

  // ---------- 組裝 sections ----------
  var article = document.querySelector('.lesson .md-body') || document.querySelector('.lesson');
  if (!article) return;

  var container = el('div', {class:'xc-integrated'});
  var demos = LESSON_DEMOS[slug] || [];

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

  // 2c. 操作動畫 / 公式拆解
  if (demos.length){
    var dsec = buildDemoSection(demos);
    if (dsec) container.appendChild(dsec);
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
