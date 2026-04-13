/* Mac 鍵盤可視化共用元件
   用法（在 lesson HTML 內）：
   <div class="mac-kb" data-shortcuts='[
     {"keys":"cmd,s","label":"⌘ S · 儲存"},
     {"keys":"cmd,shift,z","label":"⌘ ⇧ Z · 重做"}
   ]'></div>
   <script src="../keyboard.js"></script>
*/
(function(){
  var ROWS = [
    [['esc'],['f1','F1'],['f2','F2'],['f3','F3'],['f4','F4'],['f5','F5'],['f6','F6'],['f7','F7'],['f8','F8'],['f9','F9'],['f10','F10'],['f11','F11'],['f12','F12']],
    [['`'],['1'],['2'],['3'],['4'],['5'],['6'],['7'],['8'],['9'],['0'],['-'],['='],['del','⌫','wide']],
    [['tab','tab','wide'],['q','Q'],['w','W'],['e','E'],['r','R'],['t','T'],['y','Y'],['u','U'],['i','I'],['o','O'],['p','P'],['['],[']'],['\\','\\']],
    [['caps','caps','wide'],['a','A'],['s','S'],['d','D'],['f','F'],['g','G'],['h','H'],['j','J'],['k','K'],['l','L'],[';'],["'"],['enter','return ↩','xwide']],
    [['shift','⇧ shift','xwide'],['z','Z'],['x','X'],['c','C'],['v','V'],['b','B'],['n','N'],['m','M'],[','],['.'],['/'],['shift','⇧ shift','xwide']],
    [['fn','fn'],['ctrl','⌃'],['opt','⌥'],['cmd','⌘','wide'],['space','space','space'],['cmd','⌘','wide'],['opt','⌥'],['left','←'],['up','↑'],['down','↓'],['right','→']]
  ];

  function buildKeyboard(){
    var html = '<div class="kb">';
    ROWS.forEach(function(row){
      html += '<div class="kbRow">';
      row.forEach(function(k){
        var id = k[0], label = k[1] || k[0], cls = k[2] || '';
        html += '<div class="k '+cls+'" data-k="'+id+'">'+label+'</div>';
      });
      html += '</div>';
    });
    html += '</div>';
    return html;
  }

  function init(el){
    if (el.dataset.kbInit) return;
    el.dataset.kbInit = '1';
    var shortcuts = [];
    try { shortcuts = JSON.parse(el.dataset.shortcuts || '[]'); }
    catch(e){ console.warn('mac-kb: invalid JSON', e); }

    el.classList.add('kbWrap');
    var chipsHtml = '<div class="kbChips">';
    shortcuts.forEach(function(s){
      chipsHtml += '<div class="chip" data-keys="'+s.keys+'">'+s.label+'</div>';
    });
    chipsHtml += '</div>';
    el.innerHTML = buildKeyboard() + chipsHtml;

    var kb = el.querySelector('.kb');
    var chips = el.querySelector('.kbChips');
    function clear(){ kb.querySelectorAll('.k.lit').forEach(function(e){e.classList.remove('lit');}); }
    function light(keys){
      clear();
      keys.split(',').forEach(function(k){
        kb.querySelectorAll('.k[data-k="'+k+'"]').forEach(function(e){e.classList.add('lit');});
      });
    }
    chips.querySelectorAll('.chip').forEach(function(c){
      c.addEventListener('mouseenter', function(){ light(c.dataset.keys); });
      c.addEventListener('click', function(){
        chips.querySelectorAll('.chip.active').forEach(function(e){e.classList.remove('active');});
        c.classList.add('active'); light(c.dataset.keys);
      });
    });
    el.addEventListener('mouseleave', function(){
      var a = chips.querySelector('.chip.active');
      clear(); if (a) light(a.dataset.keys);
    });
  }

  function scan(){ document.querySelectorAll('.mac-kb').forEach(init); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', scan);
  else scan();
  window.MacKb = { init: init, scan: scan };
})();
