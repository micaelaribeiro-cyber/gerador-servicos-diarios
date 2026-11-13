// Executar: node tests/test-processor.js  — usa DADOS SINTÉTICOS (não é o relatório real).
const assert = require('assert');
const P = require('../js/processor.js');

const H = ['Filial','Abertura','Status','Agendamento','Cliente','Assunto','Colaborador','Endereço','Complemento','Bairro','Telefone celular','Whatsapp','Telefone residencial','Telefone comercial','Coluna Extra'];
const d = (day, h) => 46000 + day + h / 24; // serial do Excel
const rows = [
  ['Relatório de OS', '', ''], [], H,
  ['F1', d(0,8), 'Agendada', d(2,9),  'Ana',   'Instalação', 'Téc1', 'Rua A', 'NOVA VIÇOSA ', 'Centro', 73988887777, '', '', '', 'x'],
  ['F1', d(0,8), 'Agendada', d(2,8),  'Bia',   'RETIRADA',   'Téc1', 'Rua B', 'Mucuri',       'Centro', '', '', '', '', ''],
  ['F1', d(0,8), 'Agendada', '14/11/2026 10:30', 'Caio', 'Suporte', 'Téc2', 'Rua C', 'nova vicosa', 'Centro', '', '', '', '', ''],
  ['F1', d(0,8), 'Agendada', d(3,9),  'Duda',  'Retirada',   'Téc2', 'Rua D', 'Itabatã',      'Centro', '', '', '', '', ''],
  ['F1', d(0,8), 'Agendada', d(3,10), 'Edu',   'Mudança',    'Téc2', 'Rua E', 'itabata',      'Centro', '', '', '', '', ''],
  ['F1', d(0,8), 'Agendada', d(3,11), 'Fabi',  'Instalação', 'Téc3', 'Rua F', 'Posto  da Mata','Centro', '', '', '', '', ''],
  ['F1', d(0,8), 'Agendada', d(4,9),  'Gil',   'Instalação', 'Téc3', 'Rua G', 'NANUQUE',       'Centro', '', '', '', '', ''],
  ['F1', d(0,8), 'Agendada', d(4,10), 'Hugo',  'Instalação', 'Téc3', 'Rua H', 'Serra dos Aimores','Centro','', '', '', '', ''],
  ['F1', d(0,8), 'Agendada', d(4,11), 'Iara',  'Instalação', 'Téc3', 'Rua I', 'Teixeira de Freitas','Centro','', '', '', '', ''],
  ['F1', d(0,8), 'Agendada', d(4,12), 'Joca',  'Instalação', 'Téc3', 'Rua J', '',              'Centro', '', '', '', '', ''],
  [], ['', '', '']
];
const r = P.processRows(rows);
const s = r.stats;

assert.deepStrictEqual(P.SHEETS.length, 7);
assert.strictEqual(s.analyzed, 10);  assert.strictEqual(s.removed, 2);  assert.strictEqual(s.kept, 8);
assert.deepStrictEqual(s.perSheet, { 'Cidades Externas': 2, 'Itabatã': 1, 'Mucuri': 0, 'Nova Viçosa': 2, 'Posto da Mata': 1, 'Nanuque': 1, 'Serra dos Aimorés': 1 });
assert.ok(r.integrity.ok);
assert.ok(!Object.values(r.sheets).flat().some(rec => /retirada/i.test(rec[5].v)));
const nv = r.sheets['Nova Viçosa'];
assert.strictEqual(nv[0][4].v, 'Ana'); assert.strictEqual(nv[0][3].v, d(2,9));          // ordenado por agendamento
assert.strictEqual(nv[0][10].v, '73988887777'); assert.strictEqual(nv[0][8].v, 'Nova Viçosa');
assert.strictEqual(nv[1][3].t, 'n');
assert.strictEqual(P.toSerial('14/11/2026 10:30'), Date.UTC(2026,10,14,10,30)/864e5 + 25569);
assert.strictEqual(r.sheets['Cidades Externas'].map(x => x[4].v).sort().join(), 'Iara,Joca');
const wb = P.buildWorkbook(r);
assert.deepStrictEqual(wb.SheetNames, P.SHEETS);
assert.strictEqual(wb.Sheets['Mucuri']['!ref'], 'A1:N1');                               // aba vazia existe
assert.strictEqual(wb.Sheets['Nova Viçosa']['!ref'], 'A1:N3');
assert.strictEqual(wb.Sheets['Nova Viçosa']['D2'].z, 'dd/mm/yyyy hh:mm');
assert.strictEqual(wb.Sheets['Nova Viçosa']['N2'], undefined);                           // vazio não vira texto inventado
assert.deepStrictEqual(r.unused, ['Coluna Extra']);
console.log('OK — todos os testes sintéticos passaram.', JSON.stringify(s));
