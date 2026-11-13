/* Interface. Todo o processamento ocorre localmente; nenhum dado é enviado. */
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const CDN_URLS = [
    'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js',
    'https://unpkg.com/xlsx-js-style@1.2.0/dist/xlsx.bundle.js'
  ]; // alternativas caso uma CDN esteja indisponível
  let file = null, url = null;

  function loadXlsx() {
    if (window.XLSX) return Promise.resolve();
    return new Promise(async (resolve, reject) => {
      for (const src of CDN_URLS) {
        try {
          await new Promise((ok, fail) => {
            const script = document.createElement('script');
            script.src = src;
            script.onload = () => window.XLSX ? ok() : fail(new Error('Biblioteca não inicializou'));
            script.onerror = () => fail(new Error('Falha ao acessar ' + new URL(src).host));
            document.head.appendChild(script);
          });
          resolve();
          return;
        } catch (_) { /* tenta a próxima CDN */ }
      }
      reject(new Error('Não foi possível carregar a biblioteca Excel. Verifique a conexão ou peça à equipe de TI a liberação de cdn.jsdelivr.net e unpkg.com.'));
    });
  }
  const showError = m => { $('error').textContent = m; $('error').hidden = !m; };
  function setFile(f) {
    showError(''); $('result').hidden = true;
    if (!f) return;
    if (!/\.(xls|xlsx|xlsm)$/i.test(f.name)) return showError('Selecione um arquivo .xls ou .xlsx.');
    file = f; $('fname').textContent = f.name; $('gerar').disabled = false;
  }
  $('file').addEventListener('change', e => setFile(e.target.files[0]));
  const dz = $('drop');
  ['dragenter', 'dragover'].forEach(t => dz.addEventListener(t, e => { e.preventDefault(); dz.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(t => dz.addEventListener(t, e => { e.preventDefault(); dz.classList.remove('over'); }));
  dz.addEventListener('drop', e => setFile(e.dataTransfer.files[0]));

  function li(parent, text) { const x = document.createElement('li'); x.textContent = text; parent.appendChild(x); return x; }

  function render(r, name) {
    const st = r.stats;
    $('done').textContent = 'Programação gerada com sucesso.';
    $('t-all').textContent = st.analyzed; $('t-rem').textContent = st.removed; $('t-kept').textContent = st.kept;
    const locs = $('locs'); locs.textContent = '';
    const max = Math.max(1, ...Object.values(st.perSheet));
    Processor.SHEETS.forEach(n => {
      const x = document.createElement('li'), a = document.createElement('span'), bar = document.createElement('span'), b = document.createElement('b'), c = document.createElement('span');
      a.textContent = n; bar.className = 'bar'; b.style.width = (st.perSheet[n] / max * 100) + '%'; bar.appendChild(b);
      c.className = 'n'; c.textContent = st.perSheet[n];
      x.append(a, bar, c); locs.appendChild(x);
    });
    const notes = $('notes'); notes.textContent = '';
    li(notes, r.integrity.ok ? 'Conferência: a soma das abas é igual às OS aproveitadas e nenhuma OS aparece em duas abas.' : 'ATENÇÃO: a conferência de totais falhou. Não use este arquivo.');
    const rs = Object.entries(r.removedSubjects);
    if (rs.length) li(notes, 'Assuntos excluídos: ' + rs.map(([k, v]) => k + ' (' + v + ')').join(', ') + '.');
    r.warnings.forEach(w => li(notes, w));
    if (r.unused.length) li(notes, 'Colunas do relatório não utilizadas: ' + r.unused.join(', ') + '.');
    $('result').hidden = false;
  }

  $('gerar').addEventListener('click', async () => {
    showError(''); $('result').hidden = true; $('gerar').disabled = true; $('busy').hidden = false;
    try {
      await loadXlsx();
      await new Promise(r => setTimeout(r, 30));
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      let result = null, err = null;
      for (const n of wb.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '', blankrows: false });
        try { result = Processor.processRows(rows); break; } catch (e) { err = e; }
      }
      if (!result) throw err || new Error('Planilha vazia.');
      const buf = XLSX.write(Processor.buildWorkbook(result), { bookType: 'xlsx', type: 'array' });
      if (url) URL.revokeObjectURL(url);
      url = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const d = new Date(), p = n => String(n).padStart(2, '0');
      const a = $('download'); a.href = url; a.download = 'PROGRAMACAO_' + p(d.getDate()) + '-' + p(d.getMonth() + 1) + '-' + d.getFullYear() + '.xlsx';
      render(result);
    } catch (e) { showError('Erro: ' + e.message); }
    finally { $('busy').hidden = true; $('gerar').disabled = false; }
  });
})();
