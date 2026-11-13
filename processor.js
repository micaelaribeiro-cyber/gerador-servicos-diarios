/* Lógica de processamento (sem dependência de biblioteca). Roda no navegador e no Node (testes). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Processor = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const SHEETS = ['Cidades Externas', 'Itabatã', 'Mucuri', 'Nova Viçosa', 'Posto da Mata', 'Nanuque', 'Serra dos Aimorés'];
  const EXTERNAS = SHEETS[0];
  const LOCAL = { itabata: 'Itabatã', mucuri: 'Mucuri', 'nova vicosa': 'Nova Viçosa', 'posto da mata': 'Posto da Mata', nanuque: 'Nanuque', 'serra dos aimores': 'Serra dos Aimorés' };
  const FMT_DT = 'dd/mm/yyyy hh:mm', FMT_D = 'dd/mm/yyyy';

  // Ordem final das colunas (A..N) e nomes aceitos no relatório original (já normalizados).
  const COLS = [
    { key: 'filial', label: 'Filial', aliases: ['filial'] },
    { key: 'abertura', label: 'Abertura', type: 'date', aliases: ['abertura', 'data abertura', 'data de abertura'] },
    { key: 'status', label: 'Status', aliases: ['status', 'situacao'] },
    { key: 'agendamento', label: 'Agendamento', type: 'date', aliases: ['agendamento', 'data agendamento', 'data de agendamento', 'data agenda'] },
    { key: 'cliente', label: 'Cliente', aliases: ['cliente', 'nome cliente', 'nome do cliente', 'razao social'] },
    { key: 'assunto', label: 'Assunto', aliases: ['assunto'] },
    { key: 'colaborador', label: 'Colaborador', aliases: ['colaborador', 'tecnico'] },
    { key: 'endereco', label: 'Endereço', aliases: ['endereco', 'logradouro'] },
    { key: 'complemento', label: 'Complemento', aliases: ['complemento'] },
    { key: 'bairro', label: 'Bairro', aliases: ['bairro'] },
    { key: 'celular', label: 'Telefone celular', type: 'phone', aliases: ['telefone celular', 'celular', 'tel celular'] },
    { key: 'whatsapp', label: 'Whatsapp', type: 'phone', aliases: ['whatsapp', 'whats app'] },
    { key: 'residencial', label: 'Telefone residencial', type: 'phone', aliases: ['telefone residencial', 'tel residencial', 'telefone fixo'] },
    { key: 'comercial', label: 'Telefone comercial', type: 'phone', aliases: ['telefone comercial', 'tel comercial'] }
  ];

  // Minúsculas, sem acentos, sem pontuação, espaços únicos.
  const norm = v => String(v == null ? '' : v).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();

  function mapColumns(cells) {
    const normed = cells.map(norm), used = new Set(), map = {};
    COLS.forEach(c => {
      let idx = -1;
      for (const a of c.aliases) {
        idx = normed.findIndex((x, j) => x === a && !used.has(j));
        if (idx >= 0) break;
      }
      if (idx >= 0) used.add(idx);
      map[c.key] = idx;
    });
    return map;
  }

  function detectHeader(rows) {
    let best = null;
    for (let r = 0; r < Math.min(rows.length, 40); r++) {
      const map = mapColumns(rows[r] || []);
      const n = Object.values(map).filter(i => i >= 0).length;
      if (n >= 3 && (!best || n > best.n)) best = { row: r, map, n };
    }
    return best;
  }

  // Converte número serial do Excel, Date ou texto (dd/mm/aaaa [hh:mm[:ss]] ou aaaa-mm-dd) em serial.
  function toSerial(v) {
    if (typeof v === 'number') return v >= 1 && v < 2958466 ? v : null;
    if (v instanceof Date && !isNaN(v)) return Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), v.getHours(), v.getMinutes(), v.getSeconds()) / 864e5 + 25569;
    if (typeof v !== 'string') return null;
    const s = v.trim();
    let m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})(?:[ T,]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/), d, mo, y;
    if (m) { d = +m[1]; mo = +m[2]; y = +m[3]; }
    else if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/))) { y = +m[1]; mo = +m[2]; d = +m[3]; }
    else return null;
    if (y < 100) y += 2000;
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    return Date.UTC(y, mo - 1, d, +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)) / 864e5 + 25569;
  }

  function makeCell(v, col) {
    if (v == null || v === '') return null;
    if (col.type === 'date') {
      const s = toSerial(v);
      if (s == null) return { t: 's', v: String(v) }; // não inventa: mantém o texto original
      return { t: 'n', v: s, z: s - Math.floor(s) > 1e-6 ? FMT_DT : FMT_D };
    }
    if (typeof v === 'number') return col.type === 'phone' ? { t: 's', v: String(v) } : { t: 'n', v };
    if (typeof v === 'boolean') return { t: 'b', v };
    return { t: 's', v: String(v) };
  }

  const isRetirada = assunto => /\bretirada\b/.test(norm(assunto));

  function processRows(rows) {
    const hdr = detectHeader(rows);
    if (!hdr) throw new Error('Não encontrei a linha de cabeçalhos (Assunto, Complemento, Cliente...). Confira se o arquivo é o relatório de OS.');
    const map = hdr.map;
    if (map.assunto < 0) throw new Error('A coluna "Assunto" não foi encontrada no relatório.');
    if (map.complemento < 0) throw new Error('A coluna "Complemento" não foi encontrada no relatório.');

    const sheets = {}, srcIdx = {};
    SHEETS.forEach(n => { sheets[n] = []; srcIdx[n] = []; });
    const removedSubjects = {};
    let analyzed = 0, removed = 0;

    for (let i = hdr.row + 1; i < rows.length; i++) {
      const row = rows[i] || [];
      if (row.every(c => c === '' || c == null)) continue;
      analyzed++;
      const assunto = row[map.assunto];
      if (isRetirada(assunto)) {
        removed++;
        const k = String(assunto).trim();
        removedSubjects[k] = (removedSubjects[k] || 0) + 1;
        continue;
      }
      const rec = COLS.map(c => (map[c.key] >= 0 ? makeCell(row[map[c.key]], c) : null));
      const dest = LOCAL[norm(row[map.complemento])] || EXTERNAS;
      if (dest !== EXTERNAS) rec[8] = { t: 's', v: dest }; // complemento padronizado
      sheets[dest].push(rec);
      srcIdx[dest].push(i);
    }

    const val = (rec, i) => (rec[i] ? rec[i].v : '');
    SHEETS.forEach(n => sheets[n].sort((a, b) => {
      const sa = a[3] && a[3].t === 'n' ? a[3].v : Infinity, sb = b[3] && b[3].t === 'n' ? b[3].v : Infinity;
      return (sa === sb ? 0 : sa < sb ? -1 : 1) || String(val(a, 6)).localeCompare(String(val(b, 6)), 'pt-BR') || String(val(a, 4)).localeCompare(String(val(b, 4)), 'pt-BR');
    }));

    const perSheet = {};
    let kept = 0;
    SHEETS.forEach(n => { perSheet[n] = sheets[n].length; kept += sheets[n].length; });
    const all = [].concat(...SHEETS.map(n => srcIdx[n]));
    const integrity = { ok: kept === analyzed - removed && new Set(all).size === all.length };

    const warnings = [];
    const missing = COLS.filter(c => map[c.key] < 0).map(c => c.label);
    if (missing.length) warnings.push('Colunas não encontradas no relatório (ficaram em branco): ' + missing.join(', ') + '.');
    const seen = new Map();
    SHEETS.forEach(n => sheets[n].forEach(r => { const k = JSON.stringify(r); seen.set(k, (seen.get(k) || 0) + 1); }));
    const dup = [...seen.values()].filter(c => c > 1).reduce((a, c) => a + c, 0);
    if (dup) warnings.push(dup + ' linhas são idênticas a outra linha do relatório (todas as colunas). Foram mantidas.');

    const usedCols = new Set(Object.values(map));
    const unused = (rows[hdr.row] || []).map((h, j) => (!usedCols.has(j) && String(h).trim() ? String(h).trim() : null)).filter(Boolean);

    return { map, headerRow: hdr.row, sheets, stats: { analyzed, removed, kept, perSheet }, removedSubjects, warnings, unused, integrity };
  }

  const HEADER_STYLE = {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb: '0A5560' } },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: { bottom: { style: 'medium', color: { rgb: 'E0531F' } } }
  };
  const letter = i => String.fromCharCode(65 + i);

  function buildSheet(records) {
    const ws = {}, widths = COLS.map(c => c.label.length + 4);
    COLS.forEach((c, i) => { ws[letter(i) + '1'] = { t: 's', v: c.label, s: HEADER_STYLE }; });
    records.forEach((rec, r) => rec.forEach((cell, i) => {
      if (!cell) return;
      ws[letter(i) + (r + 2)] = cell;
      widths[i] = Math.max(widths[i], (cell.z ? 17 : String(cell.v).length) + 2);
    }));
    const ref = 'A1:N' + (records.length + 1);
    ws['!ref'] = ref;
    ws['!cols'] = widths.map(w => ({ wch: Math.min(w, 50) }));
    ws['!autofilter'] = { ref };
    ws['!freeze'] = { xSplit: '0', ySplit: '1', topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' };
    return ws;
  }

  function buildWorkbook(result) {
    const wb = { SheetNames: SHEETS.slice(), Sheets: {} };
    SHEETS.forEach(n => { wb.Sheets[n] = buildSheet(result.sheets[n]); });
    return wb;
  }

  return { SHEETS, COLS, norm, toSerial, processRows, buildWorkbook };
});
