// Contagem de estoque - backend (Google Apps Script)
const SS_ID = '1fpaZKJNYKyOUaVOU7MadLyLFkQrOsqYoLLeQ3NkBh90';
const ABA_CADASTRO = 'Cadastro';   // cabeçalho na linha 4, dados a partir da linha 5
const ABA_REGISTROS = 'Registros'; // cada bipe vira uma linha aqui
const CHAVE = 'troque-esta-senha'; // senha usada pela página hospedada (celular). TROQUE!

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Contagem de Estoque')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// API para a página hospedada fora do Google (necessária para usar a câmera)
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.chave !== CHAVE) throw new Error('Senha incorreta');
    const fns = {getCadastro, registrar, importarCadastro};
    if (!fns[req.fn]) throw new Error('Função inválida');
    out = {dados: fns[req.fn](req.arg)};
  } catch (err) {
    out = {erro: String(err.message || err)};
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function getCadastro() {
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(ABA_CADASTRO);
  const n = sh.getLastRow() - 4;
  const map = {};
  if (n < 1) return map;
  sh.getRange(5, 1, n, 3).getDisplayValues().forEach(r => {
    const cod = String(r[0]).trim();
    if (cod) map[cod] = [r[1], r[2]];
  });
  return map;
}

// linhas = [[codigo, sku, descricao], ...]. Código novo -> adiciona; código existente -> atualiza SKU/descrição
function importarCadastro(linhas) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(ABA_CADASTRO);
    const col = sh.getRange(5, 1, Math.max(sh.getMaxRows() - 4, 1), 1).getDisplayValues();
    const pos = {};
    let fim = 4;
    col.forEach((r, i) => { const c = String(r[0]).trim(); if (c) { pos[c] = 5 + i; fim = 5 + i; } });
    const novos = [];
    let upd = 0, dup = 0;
    linhas.forEach(l => {
      const c = String(l[0]).trim();
      if (!c) return;
      const sku = String(l[1] || '').trim(), desc = String(l[2] || '').trim();
      if (pos[c] > 0) { sh.getRange(pos[c], 2, 1, 2).setValues([[sku, desc]]); upd++; }
      else if (!pos[c]) { pos[c] = -1; novos.push([c, sku, desc]); }
      else dup++;
    });
    if (novos.length) {
      const falta = fim + novos.length - sh.getMaxRows();
      if (falta > 0) sh.insertRowsAfter(sh.getMaxRows(), falta);
      sh.getRange(fim + 1, 1, novos.length, 2).setNumberFormat('@'); // preserva zeros à esquerda
      sh.getRange(fim + 1, 1, novos.length, 3).setValues(novos);
    }
    return {add: novos.length, upd: upd, dup: dup};
  } finally {
    lock.releaseLock();
  }
}

function registrar(lote) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const cad = getCadastro();
    const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(ABA_REGISTROS);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, 6).setValues([['Data/Hora', 'Código de barras', 'SKU', 'Descrição', 'Qtd', 'Status']]);
      sh.setFrozenRows(1);
    }
    const linhas = lote.map(b => {
      const c = cad[b.code];
      return [new Date(b.ts), b.code, c ? c[0] : '', c ? c[1] : '', b.qty, c ? 'OK' : 'NÃO CADASTRADO'];
    });
    const ini = sh.getLastRow() + 1;
    sh.getRange(ini, 2, linhas.length, 1).setNumberFormat('@');
    sh.getRange(ini, 1, linhas.length, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sh.getRange(ini, 1, linhas.length, 6).setValues(linhas);
    return linhas.length;
  } finally {
    lock.releaseLock();
  }
}
