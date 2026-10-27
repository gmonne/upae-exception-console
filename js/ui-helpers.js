// Shared interface helpers: paging for long lists, and Show / Hide for whole panels. Loaded first so every screen can use them.
function pagerHtml(ns, total, page, size){
  var pages = Math.max(1, Math.ceil(total / size)), from = total ? (page - 1) * size + 1 : 0, to = Math.min(total, page * size);
  var b = function(label, to2, dis){ return '<button class="btn ghost" type="button" data-pg="' + ns + '" data-to="' + to2 + '"' + (dis ? ' disabled' : '') + '>' + label + '</button>'; };
  return '<div class="hrs-bar pager" style="flex-wrap:wrap;">' + b('« First', 1, page <= 1) + b('‹ Prev', page - 1, page <= 1) +
    '<span class="panel-sub">Page ' + page + ' of ' + pages + ' &middot; ' + from + '–' + to + ' of ' + total + '</span>' + b('Next ›', page + 1, page >= pages) + b('Last »', pages, page >= pages) +
    '<span style="margin-left:auto;" class="panel-sub">Rows per page <select data-pgsize="' + ns + '" style="padding:4px 6px; border:1px solid var(--line); border-radius:6px; background:var(--bg); color:var(--ink); font:inherit; font-size:12.5px;">' +
    [10, 25, 50, 100, 250].map(function(n){ return '<option value="' + n + '"' + (n === size ? ' selected' : '') + '>' + n + '</option>'; }).join('') + '</select></span></div>';
}
function pagerWire(root, ns, onChange){
  Array.prototype.forEach.call(root.querySelectorAll('[data-pg="' + ns + '"]'), function(b){ b.addEventListener('click', function(){ onChange(parseInt(b.getAttribute('data-to'), 10), null); }); });
  Array.prototype.forEach.call(root.querySelectorAll('[data-pgsize="' + ns + '"]'), function(s){ s.addEventListener('change', function(){ onChange(1, parseInt(s.value, 10)); }); });
}
// A panel marked data-collapsible gets a Hide / Show button in its header; everything under the header folds away.
function upaeInitPanelToggles(){
  Array.prototype.forEach.call(document.querySelectorAll('.panel[data-collapsible]'), function(panel){
    var head = panel.querySelector('.panel-head'); if(!head || head.querySelector('.panel-toggle')) return;
    var key = 'upae.panel.' + panel.getAttribute('data-collapsible');
    var btn = document.createElement('button'); btn.type = 'button'; btn.className = 'btn ghost panel-toggle';
    btn.style.cssText = 'padding:4px 10px; font-size:12px;' + (head.querySelector('[style*="margin-left:auto"]') ? '' : ' margin-left:auto;');
    function apply(hide){ panel.classList.toggle('collapsed', hide); btn.textContent = hide ? 'Show' : 'Hide'; btn.setAttribute('aria-expanded', hide ? 'false' : 'true'); }
    var start = false; try{ start = sessionStorage.getItem(key) === '1'; }catch(e){}
    apply(start);
    btn.addEventListener('click', function(){ var h = !panel.classList.contains('collapsed'); apply(h); try{ sessionStorage.setItem(key, h ? '1' : '0'); }catch(e){} });
    head.appendChild(btn);
  });
}
