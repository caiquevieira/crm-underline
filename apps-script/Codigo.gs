/**
 * CRM — API da Planilha Google para o index.html publicado no GitHub Pages.
 * Faz o mesmo que o server.py faz com o crm.xlsx: lê e grava as abas desta planilha.
 *
 * Instalação (uma vez): veja o README.md do repositório.
 *  1. Extensões > Apps Script, cole este arquivo e salve.
 *  2. Rode a função configurarChave() e copie a chave que aparece no registro de execução.
 *  3. Implantar > Nova implantação > App da Web: executar como "Eu", acesso "Qualquer pessoa".
 */
const ABA_CONFIG = '_Config';
const CAB_CONFIG = ['Lista', 'Item', 'Nome', 'Tipo', 'Opções', 'Ordem', 'Visível', 'Cor', 'Oferta', 'Produção', 'Limite dias', 'Padrão'];
const ABA_HIST = '_Historico'; // mudanças de etapa: uma linha por mudança
const CAB_HIST = ['Lista', '_id', 'Data', 'De', 'Para'];
const ABA_ITENS = '_Servicos'; // serviços vendidos: uma linha por item de cada registro
const CAB_ITENS = ['Lista', '_id', 'Serviço', 'Cobrança', 'Valor'];
const ABA_CONTATOS = '_Contatos'; // contatos de parceiros e clientes: uma linha por contato
const CAB_CONTATOS = ['Tipo', '_id do registro', 'Nome', 'Cargo', 'Telefone', 'E-mail', 'Observação', '_id'];
const CAMPOS_CONTATO = ['nome', 'cargo', 'telefone', 'email', 'obs'];
const COBRANCAS = { unico: 'Único', mensal: 'Mensal' };
// "vinculo": guarda o _id de um registro de outra lista (ex.: o cliente aponta para o parceiro)
const TIPOS_COLUNA = ['texto', 'numero', 'data', 'selecao', 'url', 'telefone', 'caixa', 'servicos', 'vinculo'];
// Colunas que toda lista tem. O tipo identifica a coluna, então o nome pode ser trocado.
const COLUNAS_SISTEMA = [['Etapa', 'etapa'], ['Próxima ação', 'proxima_acao'], ['Data da próxima ação', 'data_proxima_acao'],
  ['Etapa desde', 'etapa_desde'], ['Dias na etapa', 'dias_na_etapa']];
const TIPOS_SISTEMA = COLUNAS_SISTEMA.map(c => c[1]);
const TIPOS_DATA = ['data', 'etapa_desde', 'data_proxima_acao'];
const TIPOS_ETAPA = ['aberta', 'ganha', 'perdida'];
const PROPS = PropertiesService.getScriptProperties();
const VERSAO_API = 2; // o index.html só cria Parcerias/Clientes/Contatos com esta versão ou mais nova

// ---------- instalação ----------

function configurarChave() {
  const chave = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  PROPS.setProperty('CHAVE', chave);
  Logger.log('Chave de acesso (guarde e digite no CRM ao conectar): ' + chave);
}

// Edição feita à mão na planilha muda a versão: o CRM pergunta antes de gravar por cima.
function onEdit() {
  novaVersao();
}

// ---------- HTTP ----------

function doGet() {
  return ContentService.createTextOutput('CRM: API ativa. Abra o CRM pelo endereço do GitHub Pages.');
}

// Corpo em texto JSON (Content-Type text/plain evita a verificação CORS que o Apps Script não atende):
// { chave, acao: 'ler' } ou { chave, acao: 'salvar', mtime, forcar, listas, contatos }
// O Apps Script sempre responde 200; erros vêm em { codigo, erro } com os mesmos códigos do server.py.
function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return resposta({ codigo: 400, erro: 'Dados inválidos.' });
  }
  const chave = PROPS.getProperty('CHAVE');
  if (!chave || req.chave !== chave) return resposta({ codigo: 403, erro: 'Chave de acesso inválida.' });
  const trava = LockService.getScriptLock();
  if (!trava.tryLock(25000)) return resposta({ codigo: 423, erro: 'A planilha está ocupada com outra gravação.' });
  try {
    if (req.acao === 'ler') return resposta(ler());
    if (req.acao !== 'salvar') return resposta({ codigo: 404, erro: 'Ação desconhecida.' });
    if (!req.forcar && String(req.mtime) !== versao())
      return resposta({ codigo: 409, erro: 'A planilha foi alterada em outro lugar (outro aparelho ou no Google Planilhas) depois da última leitura.' });
    try {
      gravar(req);
    } catch (err) {
      return resposta({ codigo: err.validacao ? 400 : 500, erro: err.validacao ? err.message : 'Erro ao gravar: ' + err.message });
    }
    return resposta({ mtime: novaVersao() });
  } catch (err) {
    return resposta({ codigo: 500, erro: 'Não consegui ler a planilha: ' + err.message });
  } finally {
    trava.releaseLock();
  }
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function versao() {
  return PROPS.getProperty('VERSAO') || '0';
}

function novaVersao() {
  const v = String(Date.now());
  PROPS.setProperty('VERSAO', v);
  return v;
}

// ---------- utilidades ----------

function col(nome, tipo, opcoes, visivel) {
  return { nome: nome, tipo: tipo, opcoes: opcoes || [], visivel: visivel !== false };
}

function novaLista(nome) {
  return {
    nome: nome, colunas: [col('Nome', 'texto')].concat(COLUNAS_SISTEMA.map(c => col(c[0], c[1]))),
    etapas: [{ nome: 'Novo', tipo: 'aberta', cor: '#60a5fa' }, { nome: 'Ganho', tipo: 'ganha', cor: '#22c55e' },
      { nome: 'Perdido', tipo: 'perdida', cor: '#ef4444' }],
    registros: []
  };
}

function novoId() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 10);
}

function cobranca(v) {
  return String(v == null ? '' : v).trim().toLowerCase().startsWith('mens') ? 'mensal' : 'unico';
}

function brl(v) {
  const s = v.toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return 'R$ ' + (s.endsWith(',00') ? s.slice(0, -3) : s);
}

// Texto da célula: "Landing Page R$ 500; Tráfego Pago R$ 800/mês". Os itens ficam na aba _Servicos.
function resumoServicos(itens) {
  return itens.map(it => it.servico + ' ' + (typeof it.valor === 'number'
    ? brl(it.valor) + (it.cobranca === 'mensal' ? '/mês' : '') : '(a definir)')).join('; ');
}

const ehData = v => Object.prototype.toString.call(v) === '[object Date]';
const vazio = v => v === '' || v == null || (typeof v === 'string' && !v.trim());

// ---------- conversão de valores planilha <-> JSON ----------

function lerValor(v, tipo, tz) {
  if (vazio(v)) return null;
  if (tipo === 'dias_na_etapa' || tipo === 'servicos') return null; // calculado / vem da aba _Servicos
  if (TIPOS_DATA.includes(tipo)) {
    if (ehData(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    const m = /^\s*(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})\s*$/.exec(String(v));
    if (m) {
      let a = +m[3];
      if (a < 100) a += 2000;
      const d = new Date(a, m[2] - 1, m[1]);
      if (d.getMonth() === m[2] - 1 && d.getDate() === +m[1])
        return a + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[1]).padStart(2, '0');
    }
    return String(v).trim();
  }
  if (tipo === 'caixa') {
    if (typeof v === 'boolean') return v;
    return ['sim', 's', 'x', 'true', 'verdadeiro', '1', 'ok'].includes(String(v).trim().toLowerCase());
  }
  if (tipo === 'numero') {
    if (typeof v === 'number') return v;
    let t = String(v).trim();
    if (t.includes(',')) t = t.replace(/\./g, '').replace(',', '.');
    const n = Number(t);
    return isFinite(n) ? n : String(v).trim();
  }
  if (ehData(v)) return Utilities.formatDate(v, tz, 'dd/MM/yyyy');
  return String(v).trim();
}

// ---------- leitura ----------

function linhas(ss, aba, largura) {
  const ws = ss.getSheetByName(aba);
  if (!ws || ws.getLastRow() < 2) return [];
  return ws.getRange(2, 1, ws.getLastRow() - 1, Math.max(largura, ws.getLastColumn())).getValues().map(l => l.slice(0, largura));
}

function lerConfig(ss) {
  const cfg = {};
  const ordem = l => { const n = parseFloat(l[5]); return isNaN(n) ? 9999 : n; };
  const sim = v => ['true', 'verdadeiro', 'sim', '1'].includes(String(v).trim().toLowerCase());
  const numero = v => { if (vazio(v) || typeof v === 'boolean') return null; const n = Number(String(v).replace(',', '.')); return isFinite(n) ? n : null; };
  for (const l of linhas(ss, ABA_CONFIG, 12).sort((a, b) => ordem(a) - ordem(b))) {
    let [lista, item, nome, tipo, opcoes, , visivel, cor, oferta, producao, limite, padrao] = l;
    if (vazio(lista) || vazio(nome)) continue;
    lista = String(lista).trim();
    const c = cfg[lista] || (cfg[lista] = { colunas: [], etapas: [], servicos: [] });
    nome = String(nome).trim(); tipo = String(tipo).trim().toLowerCase(); item = String(item).trim().toLowerCase();
    if (item === 'papel') { // papel fixo da lista (parcerias, clientes): vale mesmo se a lista for renomeada
      c.papel = nome;
    } else if (item === 'servico' || item === 'serviço') { // catálogo da coluna de serviços; Padrão = preço sugerido
      if (!c.servicos.some(s => s.nome === nome)) c.servicos.push({ nome: nome, cobranca: cobranca(tipo), preco: numero(padrao) });
    } else if (item === 'etapa') {
      if (!c.etapas.some(e => e.nome === nome))
        c.etapas.push({ nome: nome, tipo: TIPOS_ETAPA.includes(tipo) ? tipo : 'aberta', cor: vazio(cor) ? '#94a3b8' : String(cor),
          oferta: sim(oferta), producao: sim(producao), limite: numero(limite) });
    } else if (!c.colunas.some(x => x.nome === nome)) {
      const ops = String(opcoes).split(';').map(o => o.trim()).filter(Boolean);
      const vis = !['false', 'falso', 'não', 'nao', '0'].includes(String(visivel).trim().toLowerCase());
      const x = col(nome, TIPOS_COLUNA.concat(TIPOS_SISTEMA).includes(tipo) ? tipo : 'texto', ops, vis);
      if (!vazio(padrao)) x.padrao = padrao;
      c.colunas.push(x);
    }
  }
  return cfg;
}

// Garante colunas de sistema, etapas e opções para valores digitados à mão na planilha.
function completar(lista) {
  let mudou = false;
  const tipos = lista.colunas.map(c => c.tipo);
  for (const [n, tipo] of COLUNAS_SISTEMA) {
    if (tipos.includes(tipo)) continue;
    let nome = n;
    while (lista.colunas.some(c => c.nome === nome)) nome += ' (2)';
    lista.colunas.push(col(nome, tipo));
    mudou = true;
  }
  if (!lista.etapas.length) { lista.etapas = novaLista('').etapas; mudou = true; }
  const cs = lista.colunas.find(c => c.tipo === 'servicos');
  if (cs) { // serviço digitado à mão na aba _Servicos entra no catálogo
    cs.servicos = cs.servicos || [];
    for (const r of lista.registros)
      for (const it of r[cs.nome] || [])
        if (!cs.servicos.some(s => s.nome === it.servico)) { cs.servicos.push({ nome: it.servico, cobranca: it.cobranca, preco: null }); mudou = true; }
  }
  const nomeEtapa = lista.colunas.find(c => c.tipo === 'etapa').nome;
  for (const r of lista.registros) {
    const v = r[nomeEtapa];
    if (v && !lista.etapas.some(e => e.nome === v)) {
      const abertas = lista.etapas.filter(e => e.tipo === 'aberta').length;
      lista.etapas.splice(abertas, 0, { nome: v, tipo: 'aberta', cor: '#94a3b8' });
      mudou = true;
    }
    for (const c of lista.colunas)
      if (c.tipo === 'selecao' && r[c.nome] && !c.opcoes.includes(r[c.nome])) { c.opcoes.push(r[c.nome]); mudou = true; }
  }
  return mudou;
}

function ler() {
  const ss = SpreadsheetApp.getActive(), tz = ss.getSpreadsheetTimeZone();
  const cfg = lerConfig(ss);
  const hist = {}, itens = {};
  for (const [lista, rid, data, de, para] of linhas(ss, ABA_HIST, 5))
    if (!vazio(lista) && !vazio(rid))
      (hist[lista] || (hist[lista] = [])).push({ id: String(rid), data: lerValor(data, 'data', tz), de: String(de), para: String(para) });
  for (const [lista, rid, servico, cob, valor] of linhas(ss, ABA_ITENS, 5)) {
    if (vazio(lista) || vazio(rid) || vazio(servico)) continue;
    const it = { servico: String(servico).trim(), cobranca: cobranca(cob) };
    const v = lerValor(valor, 'numero', tz);
    if (typeof v === 'number') it.valor = v;
    const k = lista + '\u0000' + String(rid).trim();
    (itens[k] || (itens[k] = [])).push(it);
  }
  const contatos = [], idsContato = new Set();
  let mudou = false;
  for (const l of linhas(ss, ABA_CONTATOS, CAB_CONTATOS.length)) {
    const [tipo, ref] = l, k = { tipo: String(tipo).trim(), ref: String(ref).trim() };
    if (vazio(tipo) || vazio(ref)) continue;
    CAMPOS_CONTATO.forEach((f, i) => { if (!vazio(l[i + 2])) k[f] = lerValor(l[i + 2], 'texto', tz); });
    let id = vazio(l[7]) ? '' : String(l[7]).trim();
    if (!id || idsContato.has(id)) { id = novoId(); mudou = true; }
    idsContato.add(id);
    k._id = id;
    contatos.push(k);
  }
  const listas = [];
  const porNome = {};
  for (const [n, t] of COLUNAS_SISTEMA) porNome[n.toLowerCase()] = t;
  for (const ws of ss.getSheets()) {
    const titulo = ws.getName();
    if (titulo.startsWith('_')) continue;
    let c = cfg[titulo];
    if (!c) { c = { colunas: [], etapas: [], servicos: [] }; mudou = true; }
    const valores = ws.getLastRow() ? ws.getRange(1, 1, ws.getLastRow(), Math.max(1, ws.getLastColumn())).getValues() : [[]];
    const cab = valores[0].map(x => String(x).trim());
    for (const h of cab) { // coluna criada à mão na planilha
      if (!h || h === '_id' || c.colunas.some(x => x.nome === h)) continue;
      let tipo = porNome[h.toLowerCase()] || 'texto';
      if (c.colunas.some(x => x.tipo === tipo)) tipo = 'texto';
      c.colunas.push(col(h, tipo));
      mudou = true;
    }
    const idx = {};
    cab.forEach((h, i) => { if (h && !(h in idx)) idx[h] = i; });
    const cs = c.colunas.find(x => x.tipo === 'servicos');
    if (cs) cs.servicos = c.servicos || [];
    const registros = [], ids = new Set();
    for (const linha of valores.slice(1)) {
      const r = {};
      for (const x of c.colunas) {
        const i = idx[x.nome];
        const v = i == null ? null : lerValor(linha[i], x.tipo, tz);
        if (v != null) r[x.nome] = v;
      }
      if (!Object.keys(r).length) continue;
      let rid = idx._id == null || vazio(linha[idx._id]) ? '' : String(linha[idx._id]).trim();
      if (!rid || ids.has(rid)) { rid = novoId(); mudou = true; } // vazio ou linha copiada
      ids.add(rid);
      r._id = rid;
      if (cs && itens[titulo + '\u0000' + rid]) r[cs.nome] = itens[titulo + '\u0000' + rid];
      registros.push(r);
    }
    const lista = { nome: titulo, colunas: c.colunas, etapas: c.etapas, registros: registros, historico: hist[titulo] || [] };
    if (c.papel) lista.papel = c.papel;
    mudou = completar(lista) || mudou;
    listas.push(lista);
  }
  if (!listas.length) { listas.push(novaLista('Leads')); mudou = true; }
  return { listas: listas, contatos: contatos, mtime: versao(), precisaSalvar: mudou, versaoApi: VERSAO_API };
}

// ---------- gravação ----------

function erroValidacao(msg) {
  const e = new Error(msg);
  e.validacao = true;
  return e;
}

function validar(estado) {
  const listas = estado.listas || [];
  if (!listas.length) throw erroValidacao('É preciso ter pelo menos uma lista.');
  const vistos = new Set(), papeis = new Set();
  for (const l of listas) {
    const n = l.nome || '';
    if (!n || n.length > 31 || /[\[\]:*?\/\\]/.test(n) || n.startsWith('_') || n.startsWith("'") || n.endsWith("'"))
      throw erroValidacao(`Nome de lista inválido: "${n}" (até 31 caracteres, sem [ ] : * ? / \\ e sem começar com _).`);
    if (vistos.has(n.toLowerCase())) throw erroValidacao(`Lista repetida: "${n}".`);
    vistos.add(n.toLowerCase());
    if (l.papel && papeis.has(l.papel)) throw erroValidacao(`Duas listas com o papel "${l.papel}".`);
    if (l.papel) papeis.add(l.papel);
    const nomes = l.colunas.map(c => c.nome.toLowerCase());
    if (l.colunas.some(c => !c.nome || c.nome.startsWith('_')) || new Set(nomes).size !== nomes.length)
      throw erroValidacao(`Lista "${n}": nomes de coluna vazios, repetidos ou começando com _.`);
    const tipos = l.colunas.map(c => c.tipo), conta = t => tipos.filter(x => x === t).length;
    if (tipos.some(t => !TIPOS_COLUNA.includes(t) && !TIPOS_SISTEMA.includes(t)) || TIPOS_SISTEMA.some(t => conta(t) !== 1) || conta('servicos') > 1 || conta('vinculo') > 1)
      throw erroValidacao(`Lista "${n}": tipos de coluna inválidos.`);
    const etapas = l.etapas.map(e => e.nome.toLowerCase());
    if (!etapas.length || new Set(etapas).size !== etapas.length || l.etapas.some(e => !TIPOS_ETAPA.includes(e.tipo)))
      throw erroValidacao(`Lista "${n}": etapas vazias, repetidas ou com tipo inválido.`);
  }
}

const letra = n => { let s = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; };

// Reescreve a aba inteira. formatos[i]: formato da coluna i ('@' texto puro, 'dd/mm/yyyy', '0', ou null = automático).
// Colunas de texto ficam em "texto puro" para a planilha não transformar "001" ou "12e5" em número.
function escreverAba(ss, nome, cabecalho, dados, formatos) {
  const ws = ss.getSheetByName(nome) || ss.insertSheet(nome);
  ws.clear();
  const largura = cabecalho.length;
  ws.getRange(1, 1, 1, largura).setValues([cabecalho]).setFontWeight('bold').setBackground('#e5e7eb');
  if (dados.length) {
    formatos.forEach((f, i) => { if (f) ws.getRange(2, i + 1, dados.length, 1).setNumberFormat(f); });
    ws.getRange(2, 1, dados.length, largura).setValues(dados);
  }
  ws.setFrozenRows(1);
  return ws;
}

function gravar(estado) {
  validar(estado);
  const ss = SpreadsheetApp.getActive();
  const config = [], itens = [], hist = [];
  const formatoDe = t => TIPOS_DATA.includes(t) ? 'dd/mm/yyyy' : t === 'dias_na_etapa' ? '0' : ['numero', 'caixa'].includes(t) ? null : '@';
  estado.listas.forEach(l => {
    const cols = l.colunas;
    const desde = letra(cols.findIndex(c => c.tipo === 'etapa_desde') + 1);
    const dados = l.registros.map((r, k) => cols.map(c => {
      if (c.tipo === 'dias_na_etapa') return `=IF(${desde}${k + 2}="","",TODAY()-${desde}${k + 2})`;
      if (c.tipo === 'servicos') {
        const lista = (r[c.nome] || []).filter(it => it && it.servico);
        for (const it of lista)
          itens.push([l.nome, String(r._id), it.servico, COBRANCAS[cobranca(it.cobranca)], typeof it.valor === 'number' ? it.valor : '']);
        return resumoServicos(lista);
      }
      return r[c.nome] == null ? '' : r[c.nome]; // datas vão como "aaaa-mm-dd" e a planilha converte em data
    }).concat([String(r._id)]));
    if (l.papel) config.push([l.nome, 'papel', l.papel, '', '', 0, '', '', '', '', '', '']);
    const ws = escreverAba(ss, l.nome, cols.map(c => c.nome).concat(['_id']), dados, cols.map(c => formatoDe(c.tipo)).concat(['@']));
    ws.getRange(1, cols.length + 1, Math.max(1, dados.length + 1), 1).setFontColor('#9ca3af');
    cols.forEach((c, i) => {
      config.push([l.nome, 'coluna', c.nome, c.tipo, (c.opcoes || []).join('; '), i + 1, !!c.visivel, '', '', '', '', c.padrao == null ? '' : c.padrao]);
      if (c.tipo === 'servicos')
        (c.servicos || []).forEach((s, j) => config.push([l.nome, 'servico', s.nome, cobranca(s.cobranca), '', j + 1, '', '', '', '', '',
          typeof s.preco === 'number' ? s.preco : '']));
    });
    l.etapas.forEach((e, i) => config.push([l.nome, 'etapa', e.nome, e.tipo, '', i + 1, '', e.cor, !!e.oferta, !!e.producao,
      e.limite == null ? '' : e.limite, '']));
    for (const h of l.historico || []) hist.push([l.nome, String(h.id || ''), h.data || '', h.de || '', h.para || '']);
  });
  escreverAba(ss, ABA_HIST, CAB_HIST, hist, ['@', '@', 'dd/mm/yyyy', '@', '@']);
  escreverAba(ss, ABA_ITENS, CAB_ITENS, itens, ['@', '@', '@', '@', null]);
  // contatos: só regrava quando vierem (um CRM de versão antiga não manda e não pode apagá-los)
  if (Array.isArray(estado.contatos))
    escreverAba(ss, ABA_CONTATOS, CAB_CONTATOS, estado.contatos.filter(k => k && !vazio(k.tipo) && !vazio(k.ref))
      .map(k => [String(k.tipo), String(k.ref)].concat(CAMPOS_CONTATO.map(f => k[f] == null ? '' : String(k[f])), [String(k._id || novoId())])),
      CAB_CONTATOS.map(() => '@'));
  escreverAba(ss, ABA_CONFIG, CAB_CONFIG, config, ['@', '@', '@', '@', '@', null, null, '@', null, null, null, null]);
  // abas de listas apagadas ou renomeadas saem; a ordem das abas segue a ordem das listas
  const nomes = estado.listas.map(l => l.nome);
  for (const ws of ss.getSheets()) if (!ws.getName().startsWith('_') && !nomes.includes(ws.getName())) ss.deleteSheet(ws);
  nomes.forEach((n, i) => { ss.setActiveSheet(ss.getSheetByName(n)); ss.moveActiveSheet(i + 1); });
  SpreadsheetApp.flush();
}
