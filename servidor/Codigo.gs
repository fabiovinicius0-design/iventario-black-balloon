/**
 * Iventário Estoque Black Balloon — servidor da planilha (Google Apps Script)
 * Criado por Fabio Ferraz
 *
 * COMO INSTALAR
 * 1. Troque o PIN abaixo por um código só seu (vai dentro do link da equipe).
 * 2. No menu acima, escolha a função "configurar" e clique em Executar. Autorize.
 *    Isso cria as abas ESTOQUE e HISTORICO com todos os produtos.
 * 3. Implantar > Nova implantação > tipo "App da Web"
 *      Executar como: Eu
 *      Quem pode acessar: Qualquer pessoa
 *    Copie o endereço que termina em /exec e cole no site.
 *
 * Se mudar este código depois: Implantar > Gerenciar implantações > editar > Nova versão.
 */
const PIN = '__PIN__';

// Planilha "Estoque Black Balloon". Deixe vazio se o script foi criado pelo menu Extensões da própria planilha.
const PLANILHA_ID = '__PLANILHA_ID__';

function planilha_() {
  return PLANILHA_ID ? SpreadsheetApp.openById(PLANILHA_ID) : SpreadsheetApp.getActive();
}

const ABA = 'ESTOQUE';
const ABA_HIST = 'HISTORICO';
const LOCAIS_PADRAO = ['BAR PISTA', 'VIP01', 'ESTOQUE', 'VIP02'];
const COL_ID = 1;
const COL_CAT = 2;
const COL_NOME = 3;
const COL_LOCAL1 = 4;

// Produtos iniciais (aba 1709 da planilha original): [categoria, nome, custo, [BAR PISTA, VIP01, ESTOQUE, VIP02]]
const PRODUTOS_INICIAIS = [
  ["Cerveja", "Budweiser (343ml)", 4.99, [0, 0, 0, 0]],
  ["Cerveja", "Budweiser zero (343ml)", 4.99, [8, 0, 0, 0]],
  ["Cerveja", "Stella Artois (275ml)", 5.85, [14, 0, 0, 0]],
  ["Cerveja", "Geezer - Premium Pilsener (355ml)", 0, [0, 0, 0, 0]],
  ["Cerveja", "Heineken (330ml)", 6.29, [18, 0, 0, 0]],
  ["Cerveja", "Corona (355ml)", 7.99, [8, 0, 0, 0]],
  ["Cerveja", "Geezer - American IPA (355ml)", 0, [0, 0, 0, 0]],
  ["DRINK PRONTO", "Smirnoff Ice (Long Neck)", 6.65, [0, 0, 0, 0]],
  ["DRINK PRONTO", "Smirnoff Ice (LATA)", 0, [0, 0, 0, 0]],
  ["DRINK PRONTO", "Smirnoff Ice Rasperry (Long Neck)", 0, [0, 0, 0, 0]],
  ["DRINK PRONTO", "Skol Beats Red mix (long neck)", 6.4, [0, 0, 0, 0]],
  ["DRINK PRONTO", "Skol Beats sensse (long Neck 269ml)", 6.4, [0, 0, 0, 0]],
  ["DRINK PRONTO", "Skol Beats G&T (Long Neck 269ml)", 6.4, [0, 0, 0, 0]],
  ["DRINK PRONTO", "Smirnoff Ice Tropical (Long Neck)", 0, [0, 0, 0, 0]],
  ["DRINK PRONTO", "Smirnoff Ice Green (Long Neck)", 0, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Água sem gás", 1.19, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Água com Gás", 2.99, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "SPRITE", 3.82, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "COCA COLA", 3.99, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Coca Cola zero", 3.99, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Fanta", 3.82, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Guarana antartica", 3.82, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Scheweppes tônica", 3.82, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Schweppers cítrus", 3.82, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Del Valle uva", 2.99, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Del vale pêssego", 2.99, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Del vale maracujá", 2.99, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Red Bull tropical", 7.25, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Red Bull melancia", 7.25, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Red bull normal", 7.25, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Red bull Spring (rosa)", 7.25, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Red bull Pomelo (roxo)", 7.25, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Red bull melão e maracujá (verde)", 7.25, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Red bull Maça", 7.25, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Red bull sugar free", 7.25, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "RED BUL Summer necatarine", 7.25, [0, 0, 0, 0]],
  ["SEM ÁLCOOL", "Monster", 0, [0, 0, 0, 0]],
  ["TEQUILAS", "José Cuervo Prata", 120, [1.2, 0, 0, 0]],
  ["TEQUILAS", "José cuervo ouro", 120, [1.5, 0, 0, 0]],
  ["TEQUILAS", "Tequiloka", 9.99, [0, 0, 0, 0]],
  ["CACHACAS", "Ypioca Reserva", 49.99, [4, 0, 0, 0]],
  ["CACHACAS", "Seleta", 45.99, [0.8, 0, 0, 0]],
  ["CACHACAS", "velho barreiro", 15.99, [4, 0, 0, 0]],
  ["CACHACAS", "Cachaça com Mel (consultar)", 29.9, [0, 0, 0, 0]],
  ["VODKA", "Smirnoff", 27.9, [4.8, 0, 0, 0]],
  ["VODKA", "Absolut", 56.27, [3.7, 0, 0, 0]],
  ["VODKA", "Ketel One", 99, [0.8, 0, 0, 0]],
  ["VODKA", "Ciroc Regular", 159, [4.2, 0, 0, 0]],
  ["VODKA", "Ciroc Flavors", 159, [0, 0, 0, 0]],
  ["WHISKY", "Black & White", 89, [0.9, 0, 0, 0]],
  ["WHISKY", "Johnnie Walker Red Label", 89, [0.15, 0, 0, 0]],
  ["WHISKY", "Johnnie Walker Black Label", 159, [1.8, 0, 0, 0]],
  ["WHISKY", "Johnnie Walker Reserve", 0, [0, 0, 0, 0]],
  ["WHISKY", "Jim Beam", 0, [0, 0, 0, 0]],
  ["WHISKY", "Jim Beam Honey", 0, [0, 0, 0, 0]],
  ["WHISKY", "Jack Daniels", 109.9, [0.85, 0, 0, 0]],
  ["WHISKY", "Jack Daniels Honey", 109.9, [4, 0, 0, 0]],
  ["WHISKY", "Jack BlackBerry", 139.9, [1, 0, 0, 0]],
  ["APERITIVO", "Domeg", 34.99, [0, 0, 0, 0]],
  ["APERITIVO", "Bacardi Big Apple", 29.99, [0.9, 0, 0, 0]],
  ["APERITIVO", "Bacardi Ouro/Prata", 43.99, [3.3, 0, 0, 0]],
  ["APERITIVO", "Campari", 59.9, [1.8, 0, 0, 0]],
  ["APERITIVO", "APEROL", 0, [2.5, 0, 0, 0]],
  ["APERITIVO", "Gin Tanqueray", 88.9, [1.7, 0, 0, 0]],
  ["APERITIVO", "Gin Gordons", 54.9, [0.75, 0, 0, 0]],
  ["APERITIVO", "Jaggermeister", 139.99, [0.7, 0, 0, 0]],
  ["APERITIVO", "Chandon", 99.99, [1, 0, 0, 0]],
  ["APERITIVO", "Licor 43", 139.99, [0.4, 0, 0, 0]],
  ["INSUMOS BAR", "Copos de 330ml", 0, [0, 0, 0, 0]],
  ["INSUMOS BAR", "SAL", 0, [0, 0, 0, 0]],
  ["INSUMOS BAR", "GROSELHA", 0, [0, 0, 0, 0]],
  ["INSUMOS BAR", "PIMENTA", 0, [0, 0, 0, 0]],
];

function configurar() {
  const aba = obterAba_();
  obterHist_();
  planilha_().setActiveSheet(aba);
  Logger.log('Pronto: ' + (aba.getLastRow() - 1) + ' produtos na aba ' + ABA + '.');
}

function doGet(e) {
  return responder_(() => executar_((e && e.parameter) || {}));
}

function doPost(e) {
  return responder_(() => executar_(JSON.parse((e && e.postData && e.postData.contents) || '{}')));
}

function responder_(fn) {
  let saida;
  try {
    saida = fn();
  } catch (err) {
    saida = { ok: false, erro: 'servidor', mensagem: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(saida)).setMimeType(ContentService.MimeType.JSON);
}

function executar_(p) {
  if (PIN && String(p.pin || '') !== PIN) {
    return { ok: false, erro: 'pin', mensagem: 'Código de acesso (PIN) incorreto.' };
  }
  switch (p.acao) {
    case 'listar': return listar_(p);
    case 'mover': return comTrava_(() => umaVez_(p, () => alterarQtd_(p, 'mover')));
    case 'definir': return comTrava_(() => umaVez_(p, () => alterarQtd_(p, 'definir')));
    case 'salvarProduto': return comTrava_(() => umaVez_(p, () => salvarProduto_(p)));
    case 'removerProduto': return comTrava_(() => umaVez_(p, () => removerProduto_(p)));
    case 'zerar': return comTrava_(() => umaVez_(p, () => zerar_(p)));
    default: return { ok: false, erro: 'acao', mensagem: 'Ação desconhecida: ' + p.acao };
  }
}

// ---------- infraestrutura ----------
function comTrava_(fn) {
  const trava = LockService.getScriptLock();
  trava.waitLock(20000);
  try {
    const r = fn();
    SpreadsheetApp.flush();
    return r;
  } finally {
    trava.releaseLock();
  }
}

// Evita aplicar duas vezes a mesma alteração quando o celular reenvia depois de perder a conexão.
function umaVez_(p, fn) {
  const cache = CacheService.getScriptCache();
  const chave = p.uid ? 'op_' + String(p.uid).slice(0, 100) : null;
  if (chave) {
    const antes = cache.get(chave);
    if (antes) return JSON.parse(antes);
  }
  const r = fn();
  if (chave && r.ok) cache.put(chave, JSON.stringify(r), 21600);
  return r;
}

function versao_() {
  return PropertiesService.getScriptProperties().getProperty('versao') || '0';
}

function tocar_() {
  PropertiesService.getScriptProperties().setProperty('versao', String(Date.now()));
}

const arred_ = (n) => Math.round(Number(n) * 1000) / 1000;

function slug_(s) {
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase()
    .replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 60) || 'produto';
}

// ---------- abas ----------
function obterAba_() {
  const ss = planilha_();
  let aba = ss.getSheetByName(ABA);
  if (aba) return aba;

  aba = ss.insertSheet(ABA, 0);
  const cab = ['ID', 'CATEGORIA', 'PRODUTOS'].concat(LOCAIS_PADRAO, ['TOTAL', 'CUSTO', 'VALOR TOTAL', 'ALTERADO EM']);
  aba.getRange(1, 1, 1, cab.length).setValues([cab]);

  const usados = {};
  const linhas = PRODUTOS_INICIAIS.map((x) => {
    let id = slug_(x[1]);
    let n = 2;
    while (usados[id]) id = slug_(x[1]) + '-' + n++;
    usados[id] = true;
    return [id, x[0], x[1]].concat(x[3], ['', x[2] || '', '', '']);
  });
  if (linhas.length) {
    aba.getRange(2, 1, linhas.length, cab.length).setValues(linhas);
    aplicarFormulas_(aba, 2, linhas.length, LOCAIS_PADRAO.length);
  }
  formatar_(aba);
  aba.hideColumns(COL_ID);
  aba.setFrozenRows(1);
  aba.setFrozenColumns(COL_NOME);

  const padrao = ss.getSheets().find((s) => s.getName() !== ABA && s.getLastRow() === 0 && s.getLastColumn() === 0);
  if (padrao && ss.getSheets().length > 1) ss.deleteSheet(padrao);
  tocar_();
  return aba;
}

function obterHist_() {
  const ss = planilha_();
  let h = ss.getSheetByName(ABA_HIST);
  if (h) return h;
  h = ss.insertSheet(ABA_HIST);
  const cab = ['DATA/HORA', 'AÇÃO', 'PRODUTO', 'LOCAL', 'ALTERAÇÃO', 'ANTES', 'DEPOIS', 'QUEM'];
  h.getRange(1, 1, 1, cab.length).setValues([cab])
    .setBackground('#1A1A1A').setFontColor('#C9971E').setFontWeight('bold').setFontFamily('Arial');
  h.setFrozenRows(1);
  h.setColumnWidth(1, 150);
  h.setColumnWidth(3, 260);
  h.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  return h;
}

function historico_(acao, produto, local, alteracao, antes, depois, quem) {
  obterHist_().appendRow([new Date(), acao, produto, local, alteracao, antes, depois, String(quem || '').slice(0, 60)]);
}

function aplicarFormulas_(aba, linha, qtdLinhas, nLocais) {
  const colTotal = COL_LOCAL1 + nLocais;
  aba.getRange(linha, colTotal, qtdLinhas, 1).setFormulaR1C1('=SUM(RC[-' + nLocais + ']:RC[-1])');
  aba.getRange(linha, colTotal + 2, qtdLinhas, 1).setFormulaR1C1('=RC[-2]*RC[-1]');
}

function formatar_(aba) {
  const nLinhas = Math.max(aba.getLastRow(), 2);
  const nCols = aba.getLastColumn();
  const cab = aba.getRange(1, 1, 1, nCols).getValues()[0].map(String);
  const iTotal = cab.indexOf('TOTAL') + 1;
  const iCusto = cab.indexOf('CUSTO') + 1;
  const iAlt = cab.indexOf('ALTERADO EM') + 1;

  aba.getBandings().forEach((b) => b.remove());
  aba.getRange(1, 1, nLinhas, nCols).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor('#1A1A1A').setFirstRowColor('#FFFFFF').setSecondRowColor('#F7F5EF');
  aba.getRange(1, 1, 1, nCols).setFontColor('#C9971E').setFontWeight('bold').setFontFamily('Arial')
    .setFontSize(11).setVerticalAlignment('middle');
  aba.setRowHeight(1, 28);
  aba.getRange(2, 1, nLinhas - 1, nCols).setFontFamily('Arial').setFontSize(10);
  aba.getRange(2, COL_CAT, nLinhas - 1, 1).setFontColor('#8F6A10').setFontWeight('bold');
  aba.getRange(1, COL_LOCAL1, nLinhas, iTotal - COL_LOCAL1 + 1).setHorizontalAlignment('center');
  aba.getRange(2, iTotal, nLinhas - 1, 1).setFontWeight('bold');
  aba.getRange(2, iCusto, nLinhas - 1, 2).setNumberFormat('"R$" #,##0.00');
  if (iAlt) aba.getRange(2, iAlt, nLinhas - 1, 1).setNumberFormat('dd/MM HH:mm').setFontColor('#8A8A8A');
  aba.setColumnWidth(COL_CAT, 120);
  aba.setColumnWidth(COL_NOME, 260);
  for (let c = COL_LOCAL1; c <= iTotal; c++) aba.setColumnWidth(c, 95);
  aba.setColumnWidth(iCusto, 95);
  aba.setColumnWidth(iCusto + 1, 115);
}

// ---------- leitura ----------
function ler_() {
  const aba = obterAba_();
  const dados = aba.getDataRange().getValues();
  const cab = dados[0].map(String);
  const iTotal = cab.indexOf('TOTAL');
  const locais = cab.slice(COL_LOCAL1 - 1, iTotal);
  const iCusto = cab.indexOf('CUSTO');
  const produtos = [];
  for (let r = 1; r < dados.length; r++) {
    const l = dados[r];
    if (!l[COL_ID - 1]) continue;
    const qtd = {};
    locais.forEach((loc, k) => { qtd[loc] = Number(l[COL_LOCAL1 - 1 + k]) || 0; });
    produtos.push({
      id: String(l[COL_ID - 1]), categoria: String(l[COL_CAT - 1] || 'OUTROS'), nome: String(l[COL_NOME - 1]),
      qtd: qtd, custo: Number(l[iCusto]) || 0, linha: r + 1,
    });
  }
  return { aba: aba, cab: cab, locais: locais, produtos: produtos };
}

function listar_(p) {
  const v = versao_();
  if (p.v && String(p.v) === v) return { ok: true, igual: true, versao: v };
  const d = ler_();
  return {
    ok: true, versao: v, locais: d.locais, planilha: planilha_().getUrl(),
    produtos: d.produtos.map((x) => ({ id: x.id, categoria: x.categoria, nome: x.nome, qtd: x.qtd, custo: x.custo })),
  };
}

// ---------- escrita ----------
function alterarQtd_(p, tipo) {
  const d = ler_();
  const prod = d.produtos.find((x) => x.id === String(p.id));
  if (!prod) return { ok: false, erro: 'produto', mensagem: 'Esse produto não existe mais na planilha.' };
  const k = d.locais.indexOf(String(p.local));
  if (k < 0) return { ok: false, erro: 'local', mensagem: 'Local desconhecido: ' + p.local };
  const antes = prod.qtd[p.local];
  const n = Number(tipo === 'mover' ? p.delta : p.valor);
  if (!isFinite(n)) return { ok: false, erro: 'valor', mensagem: 'Quantidade inválida.' };
  const depois = Math.max(0, arred_(tipo === 'mover' ? antes + n : n));
  d.aba.getRange(prod.linha, COL_LOCAL1 + k).setValue(depois);
  const iAlt = d.cab.indexOf('ALTERADO EM');
  if (iAlt >= 0) d.aba.getRange(prod.linha, iAlt + 1).setValue(new Date());
  historico_(tipo === 'mover' ? (n > 0 ? '+' : '−') : 'definir', prod.nome, p.local,
    arred_(depois - antes), antes, depois, p.quem);
  tocar_();
  return { ok: true, id: prod.id, local: p.local, valor: depois };
}

function salvarProduto_(p) {
  const nome = String(p.nome || '').trim();
  const categoria = String(p.categoria || '').trim();
  const custo = Math.max(0, Math.round((Number(p.custo) || 0) * 100) / 100);
  if (!nome || !categoria) return { ok: false, erro: 'dados', mensagem: 'Informe o nome e a categoria.' };
  const d = ler_();
  const iCusto = d.cab.indexOf('CUSTO') + 1;
  const existente = p.id ? d.produtos.find((x) => x.id === String(p.id)) : null;

  if (existente && existente.categoria === categoria) {
    d.aba.getRange(existente.linha, COL_NOME).setValue(nome);
    d.aba.getRange(existente.linha, iCusto).setValue(custo || '');
    historico_('editar produto', nome, '', '', '', '', p.quem);
    tocar_();
    return { ok: true, id: existente.id };
  }

  let id = existente ? existente.id : slug_(nome);
  if (!existente) {
    let n = 2;
    while (d.produtos.some((x) => x.id === id)) id = slug_(nome) + '-' + n++;
  }
  const qtd = d.locais.map((l) => (existente ? existente.qtd[l] : 0));
  if (existente) d.aba.deleteRow(existente.linha);

  // insere no fim da categoria, para a planilha continuar agrupada
  const atual = ler_();
  const daCategoria = atual.produtos.filter((x) => x.categoria === categoria);
  let linha;
  if (daCategoria.length) {
    const ultima = daCategoria[daCategoria.length - 1].linha;
    atual.aba.insertRowAfter(ultima);
    linha = ultima + 1;
  } else {
    linha = Math.max(atual.aba.getLastRow(), 1) + 1;
  }
  const valores = [id, categoria, nome].concat(qtd, ['', custo || '', '', new Date()]);
  atual.aba.getRange(linha, 1, 1, valores.length).setValues([valores]);
  aplicarFormulas_(atual.aba, linha, 1, atual.locais.length);
  formatar_(atual.aba);
  historico_(existente ? 'editar produto' : 'novo produto', nome, '', '', '', '', p.quem);
  tocar_();
  return { ok: true, id: id };
}

function removerProduto_(p) {
  const d = ler_();
  const prod = d.produtos.find((x) => x.id === String(p.id));
  if (!prod) return { ok: true, id: p.id };
  d.aba.deleteRow(prod.linha);
  historico_('remover produto', prod.nome, '', '', '', '', p.quem);
  tocar_();
  return { ok: true, id: prod.id };
}

// Guarda a contagem atual numa aba com a data (ex.: 3009), igual à planilha original, e zera tudo.
function zerar_(p) {
  const ss = planilha_();
  const d = ler_();
  const base = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'ddMM');
  let nome = base;
  let n = 2;
  while (ss.getSheetByName(nome)) nome = base + ' (' + n++ + ')';
  const copia = d.aba.copyTo(ss).setName(nome);
  copia.getDataRange().setValues(copia.getDataRange().getValues()); // congela os números
  if (d.produtos.length) d.aba.getRange(2, COL_LOCAL1, d.aba.getLastRow() - 1, d.locais.length).setValue(0);
  historico_('zerar contagem', 'TODOS', '', '', '', '', p.quem);
  tocar_();
  return { ok: true, aba: nome };
}
