/* ══════════════════════════════════════════════════════════════
   DASHBOARD PRODUÇÃO — app.js
   Seções:
     1. CONFIG & ESTADO GLOBAL
     2. DEFINIÇÃO DE COLUNAS
     3. INICIALIZAÇÃO
     4. PROCESSAMENTO DE DADOS
     5. MULTISELECT ENGINE
     6. FILTROS & ORDENAÇÃO
     7. RENDER DA TABELA & PAGINAÇÃO
     8. KPIs
     9. GRÁFICOS (CHARTS)
    10. MODAIS
    11. UTILITÁRIOS
══════════════════════════════════════════════════════════════ */

/* ══════════════════════════════════════════
   1. CONFIG & ESTADO GLOBAL
══════════════════════════════════════════ */
var allData = [], filteredData = [];
var sortCol = '', sortDir = 1;
var curPage = 1;
var grpFiltro = 'todos';
var PAGE = 50;
var modalSortCol = '', modalSortDir= 1;
var modalRows = [];

/* ══════════════════════════════════════════
   2. DEFINIÇÃO DE COLUNAS
══════════════════════════════════════════ */
var COLS = [
  { key: 'OP_KEY',             label: 'OP PAI',        cls: 'mono op-c' },
  { key: 'COD_PRODUTO_PAI',    label: 'Cód. Pai',      cls: 'mono',  style: 'color:var(--text2);font-size:12px' },
  { key: 'DESC_PRODUTO_PAI',   label: 'Produto Pai',   cls: '',      maxw: 160 },
  { key: 'PESO_PAI',           label: 'Peso Pai',      cls: 'mono',  style: 'color:var(--cyan)',   fmt: 'kg' },
  { key: 'SITUACAO_PAI',       label: 'Sit. Pai',      cls: '',      chip: true },
  { key: 'USO_PAI',            label: 'Uso',           cls: '',      uso: true },
  { key: 'PREPARACAO',         label: 'Preparação',    cls: '',      prep: true },
  { key: 'SEMANA',             label: 'Sem.',          cls: 'mono',  style: 'font-size:12px;color:var(--text2)', prefix: 'Sem ' },
  { key: 'RECURSO',            label: 'Recurso',       cls: '',      style: 'font-size:12px;color:var(--text2)', maxw: 120 },
  { key: 'OP_FILHO',           label: 'OP FILHO',      cls: 'mono op-o' },
  { key: 'COD_PRODUTO_FILHO',  label: 'Cód. Filho',    cls: 'mono',  style: 'color:var(--text2);font-size:12px' },
  { key: 'DESC_PRODUTO_FILHO', label: 'Produto Filho', cls: '',      maxw: 160 },
  { key: 'PESO_FILHO',         label: 'Peso Filho',    cls: 'mono',  style: 'color:var(--yellow)', fmt: 'kg' },
  { key: 'SITUACAO_FILHO',     label: 'Sit. Filho',    cls: '',      chip: true },
  { key: 'ANALISE',            label: 'Análise',       cls: '',      anl: true },
  { key: 'PERC_ENCERRADO',     label: '% Enc.',        cls: 'mono',  percClr: true },
  { key: 'RECURSO_PENDENTE',   label: 'Rec Pendente',  cls: '',      pend: true },
];

/* Cria o cabeçalho da tabela principal no load */
(function buildTableHead() {
  var tr = document.getElementById('tblHead');
  COLS.forEach(function (c, i) {
    var th = document.createElement('th');
    th.textContent = c.label;
    th.onclick = function () { sortBy(c.key, i); };
    tr.appendChild(th);
  });
})();

/* ══════════════════════════════════════════
   3. INICIALIZAÇÃO
══════════════════════════════════════════ */
window.onload = function () {
  /* Upload de arquivo */
  document.getElementById('fileInput').onchange = function (e) {
    var file = e.target.files[0];
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function (ev) {
      try {
        var wb = XLSX.read(ev.target.result, { type: 'binary' });
        var ws = wb.Sheets[wb.SheetNames[0]];
        var json = XLSX.utils.sheet_to_json(ws, { defval: '' });
        if (!json.length) { showErr('Planilha vazia!'); return; }
        processData(json, file.name);
      } catch (err) { showErr('Erro ao ler arquivo: ' + err.message); }
    };
    reader.onerror = function () { showErr('Erro ao abrir o arquivo.'); };
    reader.readAsBinaryString(file);
  };

  /* Trocar arquivo */
  document.getElementById('reloadBtn').onclick = function () {
    allData = []; filteredData = [];
    msState = {};
    sortCol = ''; sortDir = 1;  
    curPage = 1;  
    grpFiltro = 'todos'; 
    document.getElementById('searchInput').value = '';   
    document.getElementById('fileInput').value = '';
    document.getElementById('dashboard').style.display = 'none';
    document.getElementById('uploadScreen').style.display = 'flex';
  };

  /* Fecha dropdowns ao clicar fora */
  document.addEventListener('click', function (e) {
    document.querySelectorAll('.ms-dropdown.open').forEach(function (d) {
      if (!d.closest('.ms-wrap').contains(e.target)) closeAllMs();
    });
  });

  document.getElementById('searchInput').oninput = applyFilters;
  document.getElementById('modalClose').onclick = closeModal;
  document.getElementById('modalExpand').onclick = toggleExpand;
  document.getElementById('modalOverlay').onclick = function (e) { if (e.target === this) closeModal(); };

  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') {
      if (document.getElementById('modalOverlay').classList.contains('open')) closeModal();
    }
    if (e.key === 'F11' && document.getElementById('modalOverlay').classList.contains('open')) {
      e.preventDefault(); toggleExpand();
    }
  });
};

/* ══════════════════════════════════════════
   4. PROCESSAMENTO DE DADOS
══════════════════════════════════════════ */
function processData(json, fname) {
  msState = {};
  sortCol = ''; sortDir = 1;
  curPage = 1;
  grpFiltro = 'todos';
  document.getElementById('searchInput').value = '';

  ['fSemana','fSitPai','fSitFilho','fAnalise','fRecursoPend','fUso','fPendente','fPreparacao'].forEach(function(id) {
    var btn = document.getElementById('btn-' + id);
    if (!btn) return;
    var count = btn.querySelector('.ms-count');
    if (count) count.remove();
    btn.style.borderColor = '';
    btn.style.color = '';
    var lbl = document.getElementById('lbl-' + id);
    if (lbl) lbl.textContent = getMsDefault(id);
  });

  allData = json.map(function (row) {
    var norm = {};
    Object.keys(row).forEach(function (k) { norm[k.trim().toUpperCase()] = row[k]; });
    return {
      OP_KEY:              s(norm['OP_PAI']            || norm['OP_KEY']),
      COD_PRODUTO_PAI:     s(norm['COD_PRODUTO_PAI']),
      DESC_PRODUTO_PAI:    s(norm['DESC_PRODUTO_PAI']),
      PESO_PAI:            n(norm['PESO_PAI']),
      SITUACAO_PAI:        s(norm['SITUACAO_PAI']      || norm['SITUAÇÃO_PAI']),
      USO_PAI:             s(norm['USO_PAI']),
      GRUPO_PAI:           s(norm['GRUPO_PAI']),
      DATA_PRF:            s(norm['DATA_PRF']),
      PREPARACAO:          s(norm['PREPARACAO']        || norm['PREPARAÇÃO']),
      SEMANA:              s(norm['SEMANA']),
      RECURSO:             s(norm['RECURSO']           || norm['MONTAGEM']),
      OP_FILHO:            s(norm['OP_FILHO']),
      COD_PRODUTO_FILHO:   s(norm['COD_PRODUTO_FILHO']),
      DESC_PRODUTO_FILHO:  s(norm['DESC_PRODUTO_FILHO']),
      PESO_FILHO:          n(norm['PESO_FILHO']),
      SITUACAO_FILHO:      s(norm['SITUACAO_FILHO']    || norm['SITUAÇÃO_FILHO']),
      ANALISE:             s(norm['ANALISE']           || norm['ANÁLISE']),
      PERC_ENCERRADO:      '',
      RECURSO_PENDENTE:    s(norm['RECURSO_PENDENTE']),
    };
  });

  // --- CÁLCULO DO PERC_ENCERRADO ---

  // Bloco 1 — mapear filhos únicos por OP Pai
  var percMap = {};
  allData.forEach(function(r) {
    if (!r.OP_KEY || !r.OP_FILHO) return;
    if (!percMap[r.OP_KEY]) percMap[r.OP_KEY] = { filhos: {} };
    percMap[r.OP_KEY].filhos[r.OP_FILHO] = r.SITUACAO_FILHO;
  });

  // Bloco 2 — calcular percentual por OP Pai
  var percResult = {};
  Object.keys(percMap).forEach(function(opKey) {
    var filhos = Object.values(percMap[opKey].filhos);
    var total  = filhos.length;
    var enc    = filhos.filter(function(s) { return s === 'Enc'; }).length;
    percResult[opKey] = total === 0 ? 'Sem Filhos' : Math.round(enc / total * 100) + '%';
  });

  // Bloco 3 — atribuir a cada linha
  allData.forEach(function(r) {
    r.PERC_ENCERRADO = percResult[r.OP_KEY] || 'Sem Filhos';
  });

  var nomeDash = fname.replace(/\.[^.]+$/, '').replace(/^dash_/i, '').replace(/_/g, ' ').toUpperCase();
  document.getElementById('dashTitle').innerHTML =
  '<span style="color:var(--cyan)">F1</span>' +
  '<span style="color:var(--cyan);margin:0 8px">◈</span>' +
  '<span style="color:var(--cyan)">DASHBOARD ' + nomeDash + '</span>';
  document.querySelector('.file-info').textContent = fname;
  document.getElementById('uploadScreen').style.display = 'none';
  document.getElementById('dashboard').style.display = 'block';
  populateFilters();
  applyFilters();
  openChangelog();
}
/* ══════════════════════════════════════════
   5. MULTISELECT ENGINE
══════════════════════════════════════════ */
var msState = {};

function initMs(id, values, labelDefault, isRadio) {
  msState[id] = new Set();
  var container = document.getElementById('items-' + id);
  container.innerHTML = '';
  values.forEach(function (v) {
    var item = document.createElement('div');
    item.className = 'ms-item' + (isRadio ? ' radio' : '');
    item.dataset.val = v;
    var uid = id + '_' + v.replace(/\W/g, '_');
    item.innerHTML = '<input type="checkbox" id="' + uid + '" value="' + v + '"><label for="' + uid + '">' + v + '</label>';
    item.querySelector('input').onchange = function () {
      if (isRadio) {
        container.querySelectorAll('input').forEach(function (cb) { if (cb !== this) cb.checked = false; }.bind(this));
        msState[id] = this.checked ? new Set([v]) : new Set();
      } else {
        if (this.checked) msState[id].add(v); else msState[id].delete(v);
      }
      updateMsLabel(id, labelDefault);
      applyFilters();
    };
    container.appendChild(item);
  });
}

function toggleMs(id) {
  var drop = document.getElementById('drop-' + id);
  var btn  = document.getElementById('btn-' + id);
  var isOpen = drop.classList.contains('open');
  closeAllMs();
  if (!isOpen) { drop.classList.add('open'); btn.classList.add('open'); }
}

function closeAllMs() {
  document.querySelectorAll('.ms-dropdown').forEach(function (d) { d.classList.remove('open'); });
  document.querySelectorAll('.ms-btn').forEach(function (b) { b.classList.remove('open'); });
}

function filterMsItems(id, q) {
  document.querySelectorAll('#items-' + id + ' .ms-item').forEach(function (item) {
    item.style.display = item.dataset.val.toLowerCase().includes(q.toLowerCase()) ? '' : 'none';
  });
}

function selectAllMs(id) {
  document.querySelectorAll('#items-' + id + ' .ms-item').forEach(function (item) {
    if (item.style.display !== 'none') {
      item.querySelector('input').checked = true;
      msState[id].add(item.dataset.val);
    }
  });
  updateMsLabel(id, getMsDefault(id));
  applyFilters();
}

function clearMs(id) {
  msState[id] = new Set();
  document.querySelectorAll('#items-' + id + ' input').forEach(function (cb) { cb.checked = false; });
  updateMsLabel(id, getMsDefault(id));
  applyFilters();
}

function getMsDefault(id) {
  var map = {
    fSemana:      'Semana: todas',
    fSitPai:      'Sit. Pai: todas',
    fSitFilho:    'Sit. Filho: todas',
    fAnalise:     'Análise: todas',
    fRecursoPend: 'Recurso Pend.: todos',
    fUso:         'Uso: todos',
    fPreparacao:  'Preparação: todas',
  };
  return map[id] || 'todas';
}

function updateMsLabel(id, labelDefault) {
  var sel = msState[id];
  var btn = document.getElementById('btn-' + id);
  var lbl = document.getElementById('lbl-' + id);
  var old = btn.querySelector('.ms-count');
  if (old) old.remove();
  if (!sel || sel.size === 0) {
    lbl.textContent = labelDefault;
    btn.style.borderColor = ''; btn.style.color = '';
  } else if (sel.size === 1) {
    lbl.textContent = [...sel][0];
    btn.style.borderColor = 'var(--cyan)'; btn.style.color = 'var(--cyan)';
  } else {
    lbl.textContent = labelDefault.split(':')[0] + ':';
    var badge = document.createElement('span');
    badge.className = 'ms-count'; badge.textContent = sel.size;
    btn.insertBefore(badge, btn.querySelector('.ms-arrow'));
    btn.style.borderColor = 'var(--cyan)'; btn.style.color = 'var(--cyan)';
  }
}

function populateFilters() {
  var uniq = function (key) {
    return [...new Set(allData.map(function (r) { return r[key]; }))].filter(Boolean).sort();
  };

  /* ── Semanas: ordenação numérica (1, 2, 3 … 10, 11 …) ── */
  var semanas = [...new Set(allData.map(function (r) { return r['SEMANA']; }))]
    .filter(Boolean)
    .sort(function (a, b) { return parseFloat(a) - parseFloat(b); });

  initMs('fSemana',      semanas,                  'Semana: todas');
  initMs('fSitPai',      uniq('SITUACAO_PAI'),     'Sit. Pai: todas');
  initMs('fSitFilho',    uniq('SITUACAO_FILHO'),   'Sit. Filho: todas');
  initMs('fAnalise',     uniq('ANALISE'),          'Análise: todas');
  initMs('fRecursoPend', uniq('RECURSO_PENDENTE'), 'Recurso Pend.: todos');
  initMs('fUso',         uniq('USO_PAI'),          'Uso: todos');
  initMs('fPendente',    ['Com pendência', 'Sem pendência'], 'Pendente: todos', true);
  initMs('fPreparacao',  ['Liberada', 'Em Andamento', 'Não Iniciada'], 'Preparação: todas');
}

function clearAllFilters() {
  ['fSemana', 'fSitPai', 'fSitFilho', 'fAnalise', 'fRecursoPend', 'fUso', 'fPendente', 'fPreparacao'].forEach(function (id) {
    clearMs(id);
  });
  document.getElementById('searchInput').value = '';
  applyFilters();
}

/* ══════════════════════════════════════════
   6. FILTROS & ORDENAÇÃO
══════════════════════════════════════════ */
function applyFilters() {
  var sem      = msState['fSemana']      || new Set();
  var sitP     = msState['fSitPai']      || new Set();
  var sitF     = msState['fSitFilho']    || new Set();
  var anl      = msState['fAnalise']     || new Set();
  var pend     = msState['fPendente']    || new Set();
  var recPend  = msState['fRecursoPend'] || new Set();
  var uso      = msState['fUso']         || new Set();
  var prep     = msState['fPreparacao']  || new Set();
  var srch     = document.getElementById('searchInput').value.toLowerCase().trim();

  filteredData = allData.filter(function (r) {
    if (sem.size     && !sem.has(r.SEMANA))            return false;
    if (sitP.size    && !sitP.has(r.SITUACAO_PAI))     return false;
    if (sitF.size    && !sitF.has(r.SITUACAO_FILHO))   return false;
    if (anl.size     && !anl.has(r.ANALISE))           return false;

    var semPendencia = r.PERC_ENCERRADO === '100%'
                    || String(r.PERC_ENCERRADO).trim() === 'Sem Filhos';

    if (pend.has('Sem pendência') && !semPendencia) return false;
    if (pend.has('Com pendência') &&  semPendencia) return false;

    if (recPend.size && !recPend.has(r.RECURSO_PENDENTE)) return false;
    if (uso.size     && !uso.has(r.USO_PAI))           return false;

    /* Filtro Preparação — deduplica por OP_KEY (status vem do pai) */
    if (prep.size) {
      var prepVal = r.PREPARACAO || 'Não Iniciada';
      if (!prep.has(prepVal)) return false;
    }

    if (srch) {
      var hay = (r.OP_KEY + r.DESC_PRODUTO_PAI + r.OP_FILHO + r.DESC_PRODUTO_FILHO + r.RECURSO_PENDENTE).toLowerCase();
      var terms = srch.split(/[\s,\n]+/).filter(Boolean);
      var rowTerms = [r.OP_KEY, r.OP_FILHO].filter(Boolean).map(function (t) { return t.toLowerCase(); });
      var match = terms.some(function (t) { return rowTerms.some(function (rt) { return rt.includes(t); }); });
      if (!match && !hay.includes(srch)) return false;
    }
    return true;
  });

  if (sortCol) doSort();
  curPage = 1;
  renderTable();
  updateKPIs();
  updateCharts();
}

function sortBy(col, idx) {
  if (sortCol === col) sortDir *= -1; else { sortCol = col; sortDir = 1; }
  document.querySelectorAll('.tbl thead th').forEach(function (th) { th.classList.remove('asc', 'desc'); });
  var ths = document.querySelectorAll('.tbl thead th');
  if (ths[idx]) ths[idx].classList.add(sortDir === 1 ? 'asc' : 'desc');
  doSort(); curPage = 1; renderTable();
}

function doSort() {
  filteredData.sort(function (a, b) {
    var va = a[sortCol], vb = b[sortCol];
    if (typeof va === 'number') return (va - vb) * sortDir;
    return String(va).localeCompare(String(vb)) * sortDir;
  });
}

/* ══════════════════════════════════════════
   7. RENDER DA TABELA & PAGINAÇÃO
══════════════════════════════════════════ */
function renderTable() {
  var tbody  = document.getElementById('mainBody');
  var total  = filteredData.length;
  var pages  = Math.max(1, Math.ceil(total / PAGE));
  if (curPage > pages) curPage = pages;
  var start  = (curPage - 1) * PAGE;
  var slice  = filteredData.slice(start, start + PAGE);

  document.getElementById('rowCount').textContent = total + ' linha' + (total !== 1 ? 's' : '');

  if (!slice.length) {
    tbody.innerHTML = '<tr><td colspan="' + COLS.length + '" class="no-data">Nenhum resultado encontrado</td></tr>';
    renderPag(pages, total, start);
    return;
  }

  var lastPai = null;
  tbody.innerHTML = '';

  slice.forEach(function (r) {
    var tr = document.createElement('tr');
    if (r.OP_KEY !== lastPai) { tr.classList.add('sep'); lastPai = r.OP_KEY; }

    var html = '';
    COLS.forEach(function (c) {
      var v     = r[c.key];
      var style = c.style || '';
      var mw    = c.maxw ? 'max-width:' + c.maxw + 'px;overflow:hidden;text-overflow:ellipsis;' : '';
      var title = c.maxw ? 'title="' + v + '"' : '';

      if (c.chip)
        html += '<td><span class="chip ' + chipCls(v) + '">' + (v || '—') + '</span></td>';
      else if (c.uso) {
        var uCls = v === 'Expedido' ? 'color:var(--cyan)' : 'color:var(--orange)';
        html += '<td><span class="chip" style="' + uCls + ';background:rgba(0,0,0,.2);border:1px solid currentColor;font-size:10px;display:inline-flex;justify-content:center;width:72px">' + (v || '—') + '</span></td>';
      }
      else if (c.prep)
        html += '<td><span class="chip ' + prepCls(v) + '">' + (v || 'Não Iniciada') + '</span></td>';
      else if (c.anl)
        html += '<td>' + (v ? '<span class="anl ' + anlCls(v) + '">' + v + '</span>' : '—') + '</td>';
      else if (c.fmt === 'kg')
        html += '<td class="' + c.cls + '" style="' + style + '">' + fmtKg(v) + '</td>';
      else if (c.percClr)
        html += '<td class="mono" style="color:' + percClr(v) + ';font-size:11px">' + (v || '—') + '</td>';
      else if (c.pend)
        html += '<td style="font-size:12px;color:' + (v ? 'var(--red)' : 'var(--green)') + ';">' + (v || '✔') + '</td>';
      else {
        var disp = c.prefix && v ? c.prefix + v : (v || '—');
        html += '<td class="' + c.cls + '" style="' + style + mw + '" ' + title + '>' + disp + '</td>';
      }
    });

    tr.innerHTML = html;
    tr.ondblclick = function () {
      var opRows = filteredData.filter(function (x) { return x.OP_KEY === r.OP_KEY; });
      openModalOP(r.OP_KEY, opRows, r);
    };
    tbody.appendChild(tr);
  });

  renderPag(pages, total, start);
}

function renderPag(pages, total, start) {
  document.getElementById('pagInfo').textContent =
    (start + 1) + '–' + Math.min(start + PAGE, total) + ' de ' + total;

  var btns = document.getElementById('pagBtns');
  btns.innerHTML = '';

  function btn(lbl, pg, disabled, active) {
    var b = document.createElement('button');
    b.className = 'pag-btn' + (active ? ' active' : '');
    b.textContent = lbl;
    b.disabled = disabled;
    b.onclick = function () { curPage = pg; renderTable(); document.querySelector('.tbl-wrap').scrollTop = 0; };
    btns.appendChild(b);
  }

  btn('‹', curPage - 1, curPage === 1, false);
  var s = Math.max(1, curPage - 2), e = Math.min(pages, s + 4);
  for (var p = s; p <= e; p++) btn(p, p, false, p === curPage);
  btn('›', curPage + 1, curPage === pages, false);
}

/* ══════════════════════════════════════════
   8. KPIs
══════════════════════════════════════════ */
function updateKPIs() {
  var data = filteredData;

  var paiMap = {};
  data.forEach(function (r) { if (!paiMap[r.OP_KEY]) paiMap[r.OP_KEY] = r.PESO_PAI; });
  var pesoPai = Object.values(paiMap).reduce(function (a, b) { return a + b; }, 0);

  var filhoMap = {};
  data.forEach(function (r) {
    if (r.OP_FILHO && r.SITUACAO_FILHO !== 'Enc' && !filhoMap[r.OP_FILHO])
      filhoMap[r.OP_FILHO] = r.PESO_FILHO || 0;
  });
  var pesoFilho = Object.values(filhoMap).reduce(function (a, b) { return a + b; }, 0);

  var filhosUnicos = {};
  data.forEach(function (r) { if (r.OP_FILHO) filhosUnicos[r.OP_FILHO] = true; });

  document.getElementById('kOPsPai').textContent    = Object.keys(paiMap).length;
  document.getElementById('kSub').textContent       = data.length + ' LINHAS TOTAIS';
  document.getElementById('kFilhos').textContent    = Object.keys(filhosUnicos).length;
  document.getElementById('kPesoPai').textContent   = fmtTon(pesoPai);
  document.getElementById('kPesoFilho').textContent = fmtTon(pesoFilho);

  var filhosPendentes = {};
  data.forEach(function (r) {
    if (r.OP_FILHO && r.SITUACAO_FILHO !== 'Enc' && r.RECURSO_PENDENTE && r.RECURSO_PENDENTE.trim() !== '')
      filhosPendentes[r.OP_FILHO] = true;
  });
  document.getElementById('kPendentes').textContent = Object.keys(filhosPendentes).length;
}

function openModalOPsPai() {
  var box = document.getElementById('modalBox');
  box.className = 'modal modal-op';

  var paiMap = {};
  filteredData.forEach(function (r) {
    if (!paiMap[r.OP_KEY]) paiMap[r.OP_KEY] = r;
  });
  modalRows = Object.values(paiMap);
  modalSortCol = ''; modalSortDir = 1;

  document.getElementById('modalTitle').textContent = '◈  OPs Pai Filtradas';
  document.getElementById('modalSub').textContent   = modalRows.length + ' OP(s) pai únicas';

  setModalHead(['OP PAI', 'Produto Pai', 'Peso Pai', 'Situação', 'Uso', 'Preparação', 'Semana', '% Enc.']);

  var ths = document.querySelectorAll('#modalThead th');
  var colKeys = ['OP_KEY', 'DESC_PRODUTO_PAI', 'PESO_PAI', 'SITUACAO_PAI', 'USO_PAI', 'PREPARACAO', 'SEMANA', 'PERC_ENCERRADO'];
  ths.forEach(function (th, i) {
    th.style.cursor = 'pointer';
    th.title = 'Clique para ordenar';
    th.onclick = function () {
      if (modalSortCol === colKeys[i]) modalSortDir *= -1;
      else { modalSortCol = colKeys[i]; modalSortDir = 1; }
      ths.forEach(function (t) { t.classList.remove('asc', 'desc'); });
      th.classList.add(modalSortDir === 1 ? 'asc' : 'desc');
      renderModalOPsPai();
    };
  });

  renderModalOPsPai();
  openModal();
}

function renderModalOPsPai() {
  var sorted = modalRows.slice().sort(function (a, b) {
    if (!modalSortCol) return String(a.OP_KEY).localeCompare(String(b.OP_KEY));
    var va = a[modalSortCol], vb = b[modalSortCol];
    if (modalSortCol === 'PESO_PAI') return (va - vb) * modalSortDir;
    if (modalSortCol === 'SEMANA')   return (parseFloat(va) - parseFloat(vb)) * modalSortDir;
    if (modalSortCol === 'PERC_ENCERRADO') {
      var na = parseFloat(String(va).replace('%', '').replace(',', '.')) || 0;
      var nb = parseFloat(String(vb).replace('%', '').replace(',', '.')) || 0;
      return (na - nb) * modalSortDir;
    }
    return String(va).localeCompare(String(vb)) * modalSortDir;
  });

  document.getElementById('modalBody').innerHTML = sorted.map(function (r) {
    return '<tr>' +
      '<td class="mono op-c">' + r.OP_KEY + '</td>' +
      '<td style="max-width:180px;overflow:hidden;text-overflow:ellipsis" title="' + r.DESC_PRODUTO_PAI + '">' + r.DESC_PRODUTO_PAI + '</td>' +
      '<td class="mono" style="color:var(--cyan)">' + fmtKg(r.PESO_PAI) + '</td>' +
      '<td><span class="chip ' + chipCls(r.SITUACAO_PAI) + '">' + r.SITUACAO_PAI + '</span></td>' +
      '<td style="font-size:11px">' + (r.USO_PAI || '—') + '</td>' +
      '<td><span class="chip ' + prepCls(r.PREPARACAO) + '">' + (r.PREPARACAO || 'Não Iniciada') + '</span></td>' +
      '<td class="mono" style="color:var(--text2);font-size:11px">' + (r.SEMANA ? 'Sem ' + r.SEMANA : '—') + '</td>' +
      '<td class="mono" style="color:' + percClr(r.PERC_ENCERRADO) + ';font-size:11px">' + (r.PERC_ENCERRADO || '—') + '</td>' +
      '</tr>';
  }).join('');

  document.getElementById('modalFooter').textContent = modalRows.length + ' OPs pai  ·  clique nas colunas para ordenar';
}

/* ══════════════════════════════════════════
   9. GRÁFICOS
══════════════════════════════════════════ */
function updateCharts() {
  updateChartRecurso();
  updateChartGrupo();
}

function updateChartRecurso() {
  var recMap = {};
  filteredData.forEach(function (r) {
    if (!r.RECURSO_PENDENTE) return;
    if (!recMap[r.RECURSO_PENDENTE]) recMap[r.RECURSO_PENDENTE] = { peso: 0, ops: [], opFilhosVisto: {} };
    if (!recMap[r.RECURSO_PENDENTE].opFilhosVisto[r.OP_FILHO]) {
      recMap[r.RECURSO_PENDENTE].opFilhosVisto[r.OP_FILHO] = true;
      recMap[r.RECURSO_PENDENTE].peso += (r.PESO_FILHO || 0);
    }
    recMap[r.RECURSO_PENDENTE].ops.push(r);
  });

  var entries = Object.entries(recMap).sort(function (a, b) { return b[1].peso - a[1].peso; });
  var maxR    = entries.length ? Math.max.apply(null, entries.map(function (e) { return e[1].peso; })) : 1;
  var el      = document.getElementById('recursosBars');

  if (!entries.length) { el.innerHTML = '<div class="no-data">Sem pendências nos dados filtrados</div>'; return; }

  el.innerHTML = entries.map(function (e) {
    var recurso = e[0], peso = e[1].peso, qtd = Object.keys(e[1].opFilhosVisto).length;
    return '<div class="hbar-row clickable" data-recurso="' + encodeURIComponent(recurso) + '">' +
      '<div class="hbar-lbl" title="' + recurso + '">' + recurso + '</div>' +
      '<div class="hbar-out"><div class="hbar-fill" style="width:' + (peso / maxR * 100) + '%;background:var(--yellow)"></div></div>' +
      '<div class="hbar-val" style="color:var(--yellow)">' + fmtTon(peso) + '</div>' +
      '<div class="hbar-val" style="color:var(--red);font-size:12px">' + qtd + ' OP' + (qtd !== 1 ? 's' : '') + '</div>' +
      '<div class="hbar-icon" title="clique para detalhes">⚑</div>' +
      '</div>';
  }).join('');

  el.querySelectorAll('.hbar-row.clickable').forEach(function (row) {
    row.onclick = function () {
      var recurso = decodeURIComponent(row.getAttribute('data-recurso'));
      openModalRecurso(recurso, recMap[recurso]);
    };
  });
}

function updateChartGrupo() {
  var grpMap = {};
  filteredData.forEach(function (r) {
    if (!r.GRUPO_PAI || !r.PESO_FILHO || !r.OP_FILHO) return;
    if (!grpMap[r.GRUPO_PAI]) grpMap[r.GRUPO_PAI] = { exp: 0, cons: 0, filhosVistoExp: {}, filhosVistoCons: {}, ops: { all: [], exp: [], cons: [] } };

    var g = grpMap[r.GRUPO_PAI];
    if (r.USO_PAI === 'Expedido') {
      if (!g.filhosVistoExp[r.OP_FILHO]) { g.filhosVistoExp[r.OP_FILHO] = true; g.exp += r.PESO_FILHO; }
      g.ops.exp.push(r);
    } else {
      if (!g.filhosVistoCons[r.OP_FILHO]) { g.filhosVistoCons[r.OP_FILHO] = true; g.cons += r.PESO_FILHO; }
      g.ops.cons.push(r);
    }
    g.ops.all.push(r);
  });

  function grpVal(d) {
    return grpFiltro === 'Expedido' ? d.exp : grpFiltro === 'Consumido' ? d.cons : d.exp + d.cons;
  }

  var entries  = Object.entries(grpMap).filter(function (e) { return grpVal(e[1]) > 0; })
                       .sort(function (a, b) { return grpVal(b[1]) - grpVal(a[1]); });
  var maxG     = entries.length ? Math.max.apply(null, entries.map(function (e) { return grpVal(e[1]); })) : 1;
  var showExp  = grpFiltro === 'todos' || grpFiltro === 'Expedido';
  var showCons = grpFiltro === 'todos' || grpFiltro === 'Consumido';
  var el       = document.getElementById('sitBars');

  if (!entries.length) { el.innerHTML = '<div class="no-data">sem dados</div>'; return; }

  var legend = '<div class="grp-legend">';
  if (showExp)  legend += '<div class="grp-dot" style="background:var(--cyan)"></div><span class="grp-leg-lbl">Expedido</span>';
  if (showCons) legend += '<div class="grp-dot" style="background:var(--orange)"></div><span class="grp-leg-lbl">Consumido</span>';
  legend += '</div>';

  el.innerHTML = legend + entries.map(function (e) {
    var grp = e[0], d = e[1], val = grpVal(d);
    var bars = '';
    if (showExp)  bars += '<div class="hbar-out-sm"><div class="hbar-fill-sm" style="width:' + (maxG ? d.exp / maxG * 100 : 0) + '%;background:var(--cyan)"></div></div>';
    if (showCons) bars += '<div class="hbar-out-sm"><div class="hbar-fill-sm" style="width:' + (maxG ? d.cons / maxG * 100 : 0) + '%;background:var(--orange)"></div></div>';
    var valTxt = grpFiltro === 'todos'
      ? '<span style="color:var(--cyan);font-size:12px">' + fmtTon(d.exp) + '</span> <span style="color:var(--muted);font-size:12px">/</span> <span style="color:var(--orange);font-size:12px">' + fmtTon(d.cons) + '</span>'
      : '<span style="color:' + (grpFiltro === 'Expedido' ? 'var(--cyan)' : 'var(--orange)') + '">' + fmtTon(val) + '</span>';

    return '<div class="grp-row" data-grp="' + encodeURIComponent(grp) + '">' +
      '<div class="hbar-lbl" title="' + grp + '">' + grp + '</div>' +
      '<div class="grp-vals">' + bars + '</div>' +
      '<div class="grp-total" style="font-size:12px;line-height:1.4">' + valTxt + '</div>' +
      '</div>';
  }).join('');

  el.querySelectorAll('.grp-row').forEach(function (row) {
    row.onclick = function () {
      var grp = decodeURIComponent(row.getAttribute('data-grp'));
      var opsModal = grpFiltro === 'Expedido' ? grpMap[grp].ops.exp : grpFiltro === 'Consumido' ? grpMap[grp].ops.cons : grpMap[grp].ops.all;
      openModalGrp(grp, { ops: opsModal, exp: grpMap[grp].exp, cons: grpMap[grp].cons });
    };
  });
}

function setGrpFiltro(val) {
  grpFiltro = val;
  document.getElementById('grpAll').classList.toggle('active',  val === 'todos');
  document.getElementById('grpExp').classList.toggle('active',  val === 'Expedido');
  document.getElementById('grpCons').classList.toggle('active', val === 'Consumido');
  updateCharts();
}

function openModalPesoFilho() {
  var box = document.getElementById('modalBox');
  box.className = 'modal';
  var filhosVisto = {};
  var rows = [];
  filteredData.forEach(function (r) {
    if (r.OP_FILHO && r.SITUACAO_FILHO !== 'Enc' && !filhosVisto[r.OP_FILHO]) {
      filhosVisto[r.OP_FILHO] = true;
      rows.push(r);
    }
  });
  var pesoTotal = rows.reduce(function (a, r) { return a + (r.PESO_FILHO || 0); }, 0);
  document.getElementById('modalTitle').textContent = '⚖  Filhos em Aberto';
  document.getElementById('modalSub').textContent   =
    rows.length + ' OP(s) filho únicas  ·  Peso total: ' + fmtTon(pesoTotal);
  setModalHead(['OP PAI', 'Produto Pai', 'OP FILHO', 'Produto Filho', 'Peso Filho', 'Sit. Filho', 'Análise', 'Uso', 'Recurso Pend.', 'Semana']);
  document.getElementById('modalBody').innerHTML = rows
    .sort(function (a, b) { return String(a.OP_KEY).localeCompare(String(b.OP_KEY)); })
    .map(function (r) {
      var pendCor = r.RECURSO_PENDENTE ? 'color:var(--red)' : 'color:var(--green)';
      return '<tr>' +
        '<td class="mono op-c">' + r.OP_KEY + '</td>' +
        '<td style="max-width:160px;overflow:hidden;text-overflow:ellipsis" title="' + r.DESC_PRODUTO_PAI + '">' + r.DESC_PRODUTO_PAI + '</td>' +
        '<td class="mono op-o">' + r.OP_FILHO + '</td>' +
        '<td style="max-width:160px;overflow:hidden;text-overflow:ellipsis" title="' + r.DESC_PRODUTO_FILHO + '">' + r.DESC_PRODUTO_FILHO + '</td>' +
        '<td class="mono" style="color:var(--yellow)">' + fmtKg(r.PESO_FILHO) + '</td>' +
        '<td><span class="chip ' + chipCls(r.SITUACAO_FILHO) + '">' + r.SITUACAO_FILHO + '</span></td>' +
        '<td>' + (r.ANALISE ? '<span class="anl ' + anlCls(r.ANALISE) + '">' + r.ANALISE + '</span>' : '—') + '</td>' +
        '<td style="font-size:11px">' + (r.USO_PAI || '—') + '</td>' +
        '<td style="font-size:12px;' + pendCor + '">' + (r.RECURSO_PENDENTE || '✔') + '</td>' +
        '<td class="mono" style="color:var(--text2);font-size:11px">' + (r.SEMANA ? 'Sem ' + r.SEMANA : '—') + '</td>' +
        '</tr>';
    }).join('');
  document.getElementById('modalFooter').textContent =
    rows.length + ' filhos em aberto  ·  Peso total: ' + fmtTon(pesoTotal);
  openModal();
}

/* ══════════════════════════════════════════
   10. MODAIS
══════════════════════════════════════════ */
function setModalHead(cols) {
  document.getElementById('modalThead').innerHTML =
    '<tr>' + cols.map(function (c) { return '<th>' + c + '</th>'; }).join('') + '</tr>';
}

function openModalOP(opKey, rows, refRow) {
  var box = document.getElementById('modalBox');
  box.className = 'modal modal-op';

  var pesoPai        = refRow.PESO_PAI;
  var pesoFilhoTotal = rows.reduce(function (a, r) { return a + (r.PESO_FILHO || 0); }, 0);
  var enc            = rows.filter(function (r) { return r.SITUACAO_FILHO === 'Enc'; }).length;

  document.getElementById('modalTitle').textContent = '◈  OP PAI: ' + opKey;
  document.getElementById('modalSub').textContent   =
    (refRow.DESC_PRODUTO_PAI || '') +
    '  ·  ' + rows.length + ' filho(s)' +
    '  ·  ' + enc + ' encerrado(s)' +
    '  ·  Preparação: ' + (refRow.PREPARACAO || 'Não Iniciada');

  setModalHead(['OP FILHO', 'Produto Filho', 'Peso Filho', 'Sit. Filho', 'Recurso Pend.', 'Análise', 'Uso', '% Enc.', 'Semana', 'Peso Pai']);

  document.getElementById('modalBody').innerHTML = rows
    .sort(function (a, b) { return String(a.OP_FILHO).localeCompare(String(b.OP_FILHO)); })
    .map(function (r) {
      var pendCor = r.RECURSO_PENDENTE ? 'color:var(--red)' : 'color:var(--green)';
      return '<tr>' +
        '<td class="mono op-o">' + r.OP_FILHO + '</td>' +
        '<td style="max-width:180px;overflow:hidden;text-overflow:ellipsis" title="' + r.DESC_PRODUTO_FILHO + '">' + r.DESC_PRODUTO_FILHO + '</td>' +
        '<td class="mono" style="color:var(--yellow)">' + fmtKg(r.PESO_FILHO) + '</td>' +
        '<td><span class="chip ' + chipCls(r.SITUACAO_FILHO) + '">' + r.SITUACAO_FILHO + '</span></td>' +
        '<td style="font-size:12px;' + pendCor + '">' + (r.RECURSO_PENDENTE || '✔') + '</td>' +
        '<td>' + (r.ANALISE ? '<span class="anl ' + anlCls(r.ANALISE) + '">' + r.ANALISE + '</span>' : '—') + '</td>' +
        '<td style="font-size:11px">' + (r.USO_PAI || '—') + '</td>' +
        '<td class="mono" style="color:' + percClr(r.PERC_ENCERRADO) + ';font-size:11px">' + (r.PERC_ENCERRADO || '—') + '</td>' +
        '<td class="mono" style="color:var(--text2);font-size:11px">' + (r.SEMANA ? 'Sem ' + r.SEMANA : '—') + '</td>' +
        '<td><span class="peso-pai-badge">' + fmtKg(pesoPai) + '</span></td>' +
        '</tr>';
    }).join('');

  document.getElementById('modalFooter').textContent =
    rows.length + ' filho(s)  ·  Peso Filho Total: ' + fmtTon(pesoFilhoTotal) + '  ·  Peso Pai: ' + fmtKg(pesoPai);

  openModal();
}

function openModalRecurso(recurso, data) {
  var box = document.getElementById('modalBox');
  box.className = 'modal';

  document.getElementById('modalTitle').textContent = '⚑  ' + recurso;
  document.getElementById('modalSub').textContent   =
    Object.keys(data.opFilhosVisto).length + ' OP(s) filhas pendentes  ·  Peso total (Filho): ' + fmtTon(data.peso);

  setModalHead(['OP PAI', 'Produto Pai', 'Peso Pai', 'OP FILHO', 'Produto Filho', 'Peso Filho', 'Sit. Filho', 'Análise', 'Uso', 'Semana']);

  document.getElementById('modalBody').innerHTML = data.ops
    .sort(function (a, b) { return String(a.OP_KEY).localeCompare(String(b.OP_KEY)); })
    .map(function (r) {
      return '<tr>' +
        '<td class="mono op-c">' + r.OP_KEY + '</td>' +
        '<td style="max-width:160px;overflow:hidden;text-overflow:ellipsis" title="' + r.DESC_PRODUTO_PAI + '">' + r.DESC_PRODUTO_PAI + '</td>' +
        '<td class="peso-pai-cell">' + fmtKg(r.PESO_PAI) + '</td>' +
        '<td class="mono op-o">' + r.OP_FILHO + '</td>' +
        '<td style="max-width:160px;overflow:hidden;text-overflow:ellipsis" title="' + r.DESC_PRODUTO_FILHO + '">' + r.DESC_PRODUTO_FILHO + '</td>' +
        '<td class="mono" style="color:var(--yellow)">' + fmtKg(r.PESO_FILHO) + '</td>' +
        '<td><span class="chip ' + chipCls(r.SITUACAO_FILHO) + '">' + r.SITUACAO_FILHO + '</span></td>' +
        '<td>' + (r.ANALISE ? '<span class="anl ' + anlCls(r.ANALISE) + '">' + r.ANALISE + '</span>' : '—') + '</td>' +
        '<td style="font-size:11px">' + (r.USO_PAI || '—') + '</td>' +
        '<td class="mono" style="color:var(--text2);font-size:11px">' + (r.SEMANA ? 'Sem ' + r.SEMANA : '—') + '</td>' +
        '</tr>';
    }).join('');

  document.getElementById('modalFooter').textContent =
    data.ops.length + ' registros  ·  ' + fmtTon(data.peso) + ' (Peso Filho) pendentes em ' + recurso;

  openModal();
}

function openModalGrp(grp, data) {
  var box = document.getElementById('modalBox');
  box.className = 'modal';

  var ops = data.ops, exp = data.exp, cons = data.cons, total = exp + cons;
  document.getElementById('modalTitle').textContent = '▪  ' + grp;
  document.getElementById('modalSub').textContent   =
    ops.length + ' OP' + (ops.length !== 1 ? 's' : '') +
    '  ·  Exp: ' + fmtTon(exp) + '  ·  Cons: ' + fmtTon(cons) + '  ·  Total: ' + fmtTon(total);

  setModalHead(['OP PAI', 'Produto Pai', 'Peso Pai', 'OP FILHO', 'Produto Filho', 'Peso Filho', 'Sit. Filho', 'Uso', 'Análise', 'Semana']);

  document.getElementById('modalBody').innerHTML = ops
    .sort(function (a, b) { return String(a.OP_KEY).localeCompare(String(b.OP_KEY)); })
    .map(function (r) {
      var usoCor = r.USO_PAI === 'Expedido' ? 'color:var(--cyan)' : 'color:var(--orange)';
      return '<tr>' +
        '<td class="mono op-c">' + r.OP_KEY + '</td>' +
        '<td style="max-width:160px;overflow:hidden;text-overflow:ellipsis" title="' + r.DESC_PRODUTO_PAI + '">' + r.DESC_PRODUTO_PAI + '</td>' +
        '<td class="peso-pai-cell">' + fmtKg(r.PESO_PAI) + '</td>' +
        '<td class="mono op-o">' + r.OP_FILHO + '</td>' +
        '<td style="max-width:160px;overflow:hidden;text-overflow:ellipsis" title="' + r.DESC_PRODUTO_FILHO + '">' + r.DESC_PRODUTO_FILHO + '</td>' +
        '<td class="mono" style="color:var(--yellow)">' + fmtKg(r.PESO_FILHO) + '</td>' +
        '<td><span class="chip ' + chipCls(r.SITUACAO_FILHO) + '">' + r.SITUACAO_FILHO + '</span></td>' +
        '<td style="font-size:11px;' + usoCor + '">' + r.USO_PAI + '</td>' +
        '<td>' + (r.ANALISE ? '<span class="anl ' + anlCls(r.ANALISE) + '">' + r.ANALISE + '</span>' : '—') + '</td>' +
        '<td class="mono" style="color:var(--text2);font-size:11px">' + (r.SEMANA ? 'Sem ' + r.SEMANA : '—') + '</td>' +
        '</tr>';
    }).join('');

  document.getElementById('modalFooter').textContent =
    ops.length + ' registros  ·  Expedido: ' + fmtTon(exp) + '  ·  Consumido: ' + fmtTon(cons);

  openModal();
}

function openModal() {
  document.getElementById('modalOverlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeModal() {
  document.getElementById('modalOverlay').classList.remove('open');
  document.getElementById('modalBox').classList.remove('expanded', 'modal-op');
  document.getElementById('modalExpand').textContent = '⛶ EXPANDIR';
  document.body.style.overflow = '';
}

function toggleExpand() {
  var box      = document.getElementById('modalBox');
  var btn      = document.getElementById('modalExpand');
  var expanded = box.classList.toggle('expanded');
  btn.textContent = expanded ? '⛶ RECOLHER' : '⛶ EXPANDIR';
}

/* ══════════════════════════════════════════
   11. UTILITÁRIOS
══════════════════════════════════════════ */
function showErr(msg) {
  var el = document.getElementById('errMsg');
  el.textContent = '⚠ ' + msg;
  el.style.display = 'block';
}

function n(v) { return parseFloat(String(v || 0).replace(',', '.')) || 0; }
function s(v) { return String(v || '').trim(); }

function chipCls(s) {
  var m = {
    'Enc':      'c-enc',
    'F.M.P':    'c-fmp',
    'F.MAQ':    'c-fmp',
    'Prog':     'c-prog',
    'Imp':      'c-imp',
    'Espera':   'c-esp',
    'Não.Imp':  'c-imp',
    'Env.Prod': 'c-prog',
    'Rec.Prod': 'c-prog',
    'Ret.PCP':  'c-esp',
  };
  return m[s] || 'c-risk';
}

/** CSS class para badge de preparação (STATUS_ENTREPOSTO) */
function prepCls(v) {
  if (v === 'Liberada')     return 'c-enc';       /* verde */
  if (v === 'Em Andamento') return 'c-prog';      /* amarelo */
  return 'c-imp';                                  /* cinza — Não Iniciada */
}

function anlCls(a) {
  if (a === 'F1')    return 'anl-f1';
  if (a === 'F2')    return 'anl-f2';
  if (a === 'G')     return 'anl-g';
  return 'anl-f2g';
}

function fmtKg(v) {
  return v.toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 1 }) + ' kg';
}

function fmtTon(v) {
  if (v >= 1000) return (v / 1000).toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' t';
  return fmtKg(v);
}

function percClr(p) {
  var v = parseFloat(p);
  if (isNaN(v))  return 'var(--muted)';
  if (v >= 75)   return 'var(--green)';
  if (v >= 40)   return 'var(--yellow)';
  return 'var(--red)';
}

/* ══════════════════════════════════════════
   CHANGELOG
══════════════════════════════════════════ */
function openChangelog() {
  document.getElementById('changelogOverlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeChangelog() {
  document.getElementById('changelogOverlay').classList.remove('open');
  document.body.style.overflow = '';
}
