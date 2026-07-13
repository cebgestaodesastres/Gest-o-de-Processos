/**
 * Controle de Processos — Web App (Google Apps Script)
 *
 * Os dados ficam gravados numa planilha Google. No primeiro acesso o script
 * cria automaticamente a planilha "Controle de Processos" e importa os dados
 * iniciais (DADOS_INICIAIS). Para usar uma planilha existente, defina a
 * propriedade de script SPREADSHEET_ID em: Configurações do projeto >
 * Propriedades do script.
 */

var NOME_ABA = 'Processos';

// Cabeçalhos na ordem usada ao semear uma planilha nova. Planilhas antigas são
// migradas por nome de coluna (ver obterMapaColunas_), ganhando as colunas que
// faltarem — nenhum dado existente é perdido.
var CABECALHOS = [
  'Processo', 'Link', 'Descrição', 'Status',
  'Data de mudança de Status', 'Prazo', 'Links relacionados', 'Obs', 'Arquivado'
];

// Mapa campo (cliente) -> nome do cabeçalho (planilha).
var COL = {
  processo: 'Processo',
  link: 'Link',
  descricao: 'Descrição',
  status: 'Status',
  dataStatus: 'Data de mudança de Status',
  prazo: 'Prazo',
  linksRelacionados: 'Links relacionados',
  obs: 'Obs',
  arquivado: 'Arquivado'
};

// Colunas que guardam datas (formato dd/mm/yyyy).
var COLS_DATA = ['Data de mudança de Status', 'Prazo'];

// Dados importados da planilha Controle_de_Processos.xlsx (aba "Tabela").
// Estrutura: [processo, descrição, status, dataStatus, prazo, obs].
// Datas no formato yyyy-MM-dd; "-" indica campo sem valor.
var DADOS_INICIAIS = [
  ['1400.01.0031537/2026-36', 'Of 344/206-LIGABOM Solic.para participação de militares integrantes da Câm.Técnica do Projeto RESPAD', 'Em Andamento', '2026-07-08', '2026-07-09', 'Email enviado para a LIGABOM, Ofício assinado pela CG. Enviar email resposta Maj Lucas Pacheco'],
  ['GD NAC', 'Proposta de treinamento de GD para NACs', 'Em Andamento', '2026-06-26', '-', 'Feito e repassado para o Cap Gomes em 26/06'],
  ['1400.01.0019793/2026-31', 'Relatórios de Período Chuvoso dos COBs e o Relatório Consolidado do CEB 2025/2026', 'Em Andamento', '2026-07-06', '2026-07-10', 'Comissão para revisão da Resolução do NAC (TC Patrick e Cap Tiago Costa)'],
  ['1400.01.0036499/2026-19', 'Proposta de alteração da Resolução dos NAC', 'Em Andamento', '2026-06-25', '-', 'Minuta em confecção. LER  e produzir relatório para uso pessoal'],
  ['1400.01.0037182/2026-08', 'Proposta de Alteração da Resolução do CEB', 'Em Andamento', '2026-06-25', '2026-07-08', 'Em confecção. LER  e produzir relatório para uso pessoal. PRAZO 30/06'],
  ['1400.01.0017270/2026-58', 'abc@itamaraty.gov.br Cooperação humanitária. Brasil-ONU. Desastres. OCHA. INSARAG. IEC. QUITO', 'Consulta', '2026-06-26', '-', 'Manter sob controle'],
  ['1400.01.0053751/2024-15', 'Indicadores para atividade especializada', 'Consulta', '2026-06-26', '-', 'Manter sob controle'],
  ['1400.01.0028716/2026-58', 'Semana da Prevenção 2026', 'Em Andamento', '2026-06-26', '-', 'Aguardar ordem para finalizar'],
  ['1400.01.0038621/2026-52', 'Empenho em Missão Internacional de Busca e Salvamento VENEZUELA (Caixa SEI INSARAG)', 'Em Andamento', '2026-06-26', '-', 'Acompanhar e juntar informações em DRIVE específico'],
  ['Venezuela 06/26', 'Drive da Missão Venezuela (1400.01.0038621/2026-52)', 'Em Andamento', '2026-06-26', '-', 'Juntar informações e fotos'],
  ['1400.01.0040469/2026-14', 'Minuta de Memorando sobre Banco de Gestão de Capacidades Operacionais (nome sob júdice)', 'Em Andamento', '2026-07-03', '-', 'Despachado com TC Patrick'],
  ['1400.01.0019619/2026-73', 'Solicitação de Informações Institucionais – Projeto RESPAD (Formalização de ACTs estaduais)', 'Não iniciado', '2026-07-06', '-', 'Conhecer'],
  ['1400.01.0041371/2026-07', 'Pedido de recompensa - Semana da Prevenção', 'Em Andamento', '2026-07-08', '-', 'Enviado para assinatura do TC Dias'],
  ['1400.01.0041106/2026-81', 'Seminário em Brumadinho', 'Em Andamento', '2026-07-08', '2026-07-10', 'OS Feita, pendente efetivo do CEB e BEMAD e viatura'],
  ['-', 'Caminhão Roll On Roll Off', 'Não iniciado', '2026-07-07', '-', 'Verificar especificação']
];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Controle de Processos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Inclui o conteúdo de outro arquivo HTML dentro do template principal.
 * Usado em Index.html via <?!= include('Gestao'); ?> para manter cada aba
 * (Gestao, Planner, Skills) em seu próprio arquivo.
 */
function include(nome) {
  return HtmlService.createHtmlOutputFromFile(nome).getContent();
}

/**
 * Retorna a planilha de dados, criando e semeando na primeira execução.
 * Ordem de resolução:
 *   1) SPREADSHEET_ID salvo nas propriedades do script (planilha externa);
 *   2) planilha à qual este script está vinculado (script "bound" criado via
 *      Extensões > Apps Script dentro da própria planilha) — evita criar uma
 *      planilha nova quando o código já mora junto dos dados;
 *   3) cria uma planilha nova (uso como script standalone, primeiro acesso).
 */
function obterPlanilha_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {
      // ID inválido ou planilha excluída: tenta os fallbacks abaixo.
    }
  }
  var vinculada = SpreadsheetApp.getActiveSpreadsheet();
  if (vinculada) {
    props.setProperty('SPREADSHEET_ID', vinculada.getId());
    return vinculada;
  }
  var ss = SpreadsheetApp.create('Controle de Processos');
  var aba = ss.getSheets()[0].setName(NOME_ABA);
  semearDados_(aba);
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function obterAba_() {
  var ss = obterPlanilha_();
  var aba = ss.getSheetByName(NOME_ABA);
  if (!aba) {
    aba = ss.insertSheet(NOME_ABA);
    semearDados_(aba);
  }
  return aba;
}

/**
 * Mapa nome-do-cabeçalho -> índice da coluna (1-based). Garante que toda coluna
 * de CABECALHOS exista, criando as que faltarem ao final (migração segura de
 * planilhas antigas, sem perder dados já gravados).
 */
function obterMapaColunas_(aba) {
  var ultCol = Math.max(aba.getLastColumn(), 1);
  var cabec = aba.getRange(1, 1, 1, ultCol).getValues()[0];
  var mapa = {};
  cabec.forEach(function (nome, i) {
    var n = String(nome).trim();
    if (n) mapa[n] = i + 1;
  });
  CABECALHOS.forEach(function (nome) {
    if (!mapa[nome]) {
      ultCol++;
      aba.getRange(1, ultCol).setValue(nome).setFontWeight('bold');
      mapa[nome] = ultCol;
    }
  });
  return mapa;
}

function semearDados_(aba) {
  aba.getRange(1, 1, 1, CABECALHOS.length).setValues([CABECALHOS]).setFontWeight('bold');
  aba.setFrozenRows(1);
  if (DADOS_INICIAIS.length) {
    var linhas = DADOS_INICIAIS.map(function (r) {
      // [processo, link, descrição, status, dataStatus, prazo, linksRel, obs, arquivado]
      return [r[0], '', r[1], r[2], paraCelula_(r[3]), paraCelula_(r[4]), '', r[5], false];
    });
    aba.getRange(2, 1, linhas.length, CABECALHOS.length).setValues(linhas);
  }
  var mapa = {};
  CABECALHOS.forEach(function (n, i) { mapa[n] = i + 1; });
  COLS_DATA.forEach(function (n) {
    aba.getRange(2, mapa[n], Math.max(DADOS_INICIAIS.length, 1000), 1).setNumberFormat('dd/mm/yyyy');
  });
  aba.setColumnWidth(mapa['Processo'], 190);
  aba.setColumnWidth(mapa['Link'], 120);
  aba.setColumnWidth(mapa['Descrição'], 420);
  aba.setColumnWidth(mapa['Status'], 120);
  aba.setColumnWidth(mapa['Data de mudança de Status'], 120);
  aba.setColumnWidth(mapa['Prazo'], 100);
  aba.setColumnWidth(mapa['Links relacionados'], 240);
  aba.setColumnWidth(mapa['Obs'], 420);
  aba.setColumnWidth(mapa['Arquivado'], 90);
}

/** 'yyyy-MM-dd' -> Date (para gravar na célula); qualquer outro valor vira texto. */
function paraCelula_(v) {
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)) {
    var p = v.split('-');
    return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  }
  return v === '' || v === null || v === undefined ? '-' : v;
}

/** Valor da célula -> string para o cliente (datas viram 'yyyy-MM-dd'). */
function deCelula_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v === null || v === undefined ? '' : String(v);
}

/** Interpreta o valor da coluna "Arquivado" como booleano. */
function ehVerdadeiro_(v) {
  if (v === true) return true;
  var s = String(v).trim().toLowerCase();
  return s === 'true' || s === 'sim' || s === 'verdadeiro' || s === '1';
}

/** Lê todos os processos. Chamado pelo cliente. */
function obterProcessos() {
  var aba = obterAba_();
  var mapa = obterMapaColunas_(aba);
  var ultCol = aba.getLastColumn();
  var ultima = aba.getLastRow();
  var processos = [];
  if (ultima >= 2) {
    var valores = aba.getRange(2, 1, ultima - 1, ultCol).getValues();
    valores.forEach(function (r, i) {
      var vazio = r.every(function (c) { return c === '' || c === null; });
      if (vazio) return;
      processos.push({
        linha: i + 2,
        processo: deCelula_(r[mapa[COL.processo] - 1]),
        link: deCelula_(r[mapa[COL.link] - 1]),
        descricao: deCelula_(r[mapa[COL.descricao] - 1]),
        status: deCelula_(r[mapa[COL.status] - 1]),
        dataStatus: deCelula_(r[mapa[COL.dataStatus] - 1]),
        prazo: deCelula_(r[mapa[COL.prazo] - 1]),
        linksRelacionados: deCelula_(r[mapa[COL.linksRelacionados] - 1]),
        obs: deCelula_(r[mapa[COL.obs] - 1]),
        arquivado: ehVerdadeiro_(r[mapa[COL.arquivado] - 1])
      });
    });
  }
  return {
    processos: processos,
    urlPlanilha: obterPlanilha_().getUrl()
  };
}

/**
 * Cria (sem p.linha) ou atualiza (com p.linha) um processo.
 * Campos: processo, link, descricao, status, dataStatus, prazo,
 * linksRelacionados, obs. O estado "arquivado" é preservado na atualização.
 */
function salvarProcesso(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var aba = obterAba_();
    var mapa = obterMapaColunas_(aba);
    var ultCol = aba.getLastColumn();
    var linha = Number(p.linha);
    var novo = !(linha >= 2);
    if (novo) linha = aba.getLastRow() + 1;

    var faixa = aba.getRange(linha, 1, 1, ultCol);
    var row = novo ? novaLinhaVazia_(ultCol) : faixa.getValues()[0];
    function set(campo, valor) { row[mapa[COL[campo]] - 1] = valor; }

    set('processo', p.processo || '-');
    set('link', p.link || '');
    set('descricao', p.descricao || '');
    set('status', p.status || 'Não iniciado');
    set('dataStatus', paraCelula_(p.dataStatus || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')));
    set('prazo', paraCelula_(p.prazo));
    set('linksRelacionados', p.linksRelacionados || '');
    set('obs', p.obs || '');
    if (novo) set('arquivado', false);

    faixa.setValues([row]);
    if (novo) {
      COLS_DATA.forEach(function (n) {
        aba.getRange(linha, mapa[n], 1, 1).setNumberFormat('dd/mm/yyyy');
      });
    }
    return obterProcessos();
  } finally {
    lock.releaseLock();
  }
}

function novaLinhaVazia_(n) {
  var a = [];
  for (var i = 0; i < n; i++) a.push('');
  return a;
}

/** Troca rápida de status: grava o status e a data de mudança (hoje). */
function alterarStatus(linha, status) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var aba = obterAba_();
    var mapa = obterMapaColunas_(aba);
    linha = Number(linha);
    if (!(linha >= 2)) throw new Error('Linha inválida.');
    aba.getRange(linha, mapa[COL.status]).setValue(status);
    aba.getRange(linha, mapa[COL.dataStatus]).setValue(new Date()).setNumberFormat('dd/mm/yyyy');
    return obterProcessos();
  } finally {
    lock.releaseLock();
  }
}

/** Arquiva (arquivar=true) ou desarquiva (arquivar=false) um processo. */
function arquivarProcesso(linha, arquivar) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var aba = obterAba_();
    var mapa = obterMapaColunas_(aba);
    linha = Number(linha);
    if (!(linha >= 2)) throw new Error('Linha inválida.');
    aba.getRange(linha, mapa[COL.arquivado]).setValue(arquivar ? true : false);
    return obterProcessos();
  } finally {
    lock.releaseLock();
  }
}

/** Exclui a linha do processo. */
function excluirProcesso(linha) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var aba = obterAba_();
    linha = Number(linha);
    if (!(linha >= 2)) throw new Error('Linha inválida.');
    aba.deleteRow(linha);
    return obterProcessos();
  } finally {
    lock.releaseLock();
  }
}

/* ==========================================================================
 * Infraestrutura genérica de abas (usada por Planner e Banco de Skills).
 * Cada módulo tem sua própria aba na mesma planilha, criada sob demanda no
 * primeiro acesso. As colunas são mapeadas por nome (obterMapaColunasPor_),
 * o que permite migrar abas antigas sem perder dados já gravados.
 * ======================================================================== */

/** Retorna a aba `nome`, criando-a com `cabecalhos` (em negrito) se não existir. */
function obterAbaPor_(nome, cabecalhos) {
  var ss = obterPlanilha_();
  var aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
    aba.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]).setFontWeight('bold');
    aba.setFrozenRows(1);
  }
  return aba;
}

/** Como obterMapaColunas_, mas para uma lista de cabeçalhos arbitrária. */
function obterMapaColunasPor_(aba, cabecalhos) {
  var ultCol = Math.max(aba.getLastColumn(), 1);
  var cabec = aba.getRange(1, 1, 1, ultCol).getValues()[0];
  var mapa = {};
  cabec.forEach(function (nome, i) {
    var n = String(nome).trim();
    if (n) mapa[n] = i + 1;
  });
  cabecalhos.forEach(function (nome) {
    if (!mapa[nome]) {
      ultCol++;
      aba.getRange(1, ultCol).setValue(nome).setFontWeight('bold');
      mapa[nome] = ultCol;
    }
  });
  return mapa;
}

/* ==========================================================================
 * Planner — agenda de atividades por data (aba "Planner").
 * ======================================================================== */

var NOME_ABA_PLANNER = 'Planner';
var CABECALHOS_PLANNER = ['Data', 'Título', 'Descrição', 'Início', 'Fim', 'Categoria', 'Concluída'];
var COL_PLANNER = {
  data: 'Data', titulo: 'Título', descricao: 'Descrição',
  inicio: 'Início', fim: 'Fim', categoria: 'Categoria', concluida: 'Concluída'
};

/** Lê todas as atividades do Planner. Chamado pelo cliente. */
function obterAtividades() {
  var aba = obterAbaPor_(NOME_ABA_PLANNER, CABECALHOS_PLANNER);
  var mapa = obterMapaColunasPor_(aba, CABECALHOS_PLANNER);
  var ultCol = aba.getLastColumn();
  var ultima = aba.getLastRow();
  var itens = [];
  if (ultima >= 2) {
    var valores = aba.getRange(2, 1, ultima - 1, ultCol).getValues();
    valores.forEach(function (r, i) {
      var vazio = r.every(function (c) { return c === '' || c === null; });
      if (vazio) return;
      itens.push({
        linha: i + 2,
        data: deCelula_(r[mapa[COL_PLANNER.data] - 1]),
        titulo: deCelula_(r[mapa[COL_PLANNER.titulo] - 1]),
        descricao: deCelula_(r[mapa[COL_PLANNER.descricao] - 1]),
        inicio: deCelula_(r[mapa[COL_PLANNER.inicio] - 1]),
        fim: deCelula_(r[mapa[COL_PLANNER.fim] - 1]),
        categoria: deCelula_(r[mapa[COL_PLANNER.categoria] - 1]),
        concluida: ehVerdadeiro_(r[mapa[COL_PLANNER.concluida] - 1])
      });
    });
  }
  return { atividades: itens };
}

/**
 * Cria (sem a.linha) ou atualiza (com a.linha) uma atividade.
 * Campos: data (yyyy-MM-dd), titulo, descricao, inicio ('HH:mm'), fim ('HH:mm'),
 * categoria, concluida.
 */
function salvarAtividade(a) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var aba = obterAbaPor_(NOME_ABA_PLANNER, CABECALHOS_PLANNER);
    var mapa = obterMapaColunasPor_(aba, CABECALHOS_PLANNER);
    var ultCol = aba.getLastColumn();
    var linha = Number(a.linha);
    var novo = !(linha >= 2);
    if (novo) linha = aba.getLastRow() + 1;

    // Início/Fim guardados como texto para o Sheets não converter "08:00" em hora.
    aba.getRange(linha, mapa[COL_PLANNER.inicio]).setNumberFormat('@');
    aba.getRange(linha, mapa[COL_PLANNER.fim]).setNumberFormat('@');
    aba.getRange(linha, mapa[COL_PLANNER.data]).setNumberFormat('dd/mm/yyyy');

    var faixa = aba.getRange(linha, 1, 1, ultCol);
    var row = novo ? novaLinhaVazia_(ultCol) : faixa.getValues()[0];
    function set(campo, valor) { row[mapa[COL_PLANNER[campo]] - 1] = valor; }

    set('data', paraCelula_(a.data));
    set('titulo', a.titulo || '');
    set('descricao', a.descricao || '');
    set('inicio', a.inicio || '');
    set('fim', a.fim || '');
    set('categoria', a.categoria || 'Geral');
    set('concluida', a.concluida ? true : false);

    faixa.setValues([row]);
    return obterAtividades();
  } finally {
    lock.releaseLock();
  }
}

/** Marca/desmarca a atividade como concluída. */
function alternarConcluidaAtividade(linha, concluida) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var aba = obterAbaPor_(NOME_ABA_PLANNER, CABECALHOS_PLANNER);
    var mapa = obterMapaColunasPor_(aba, CABECALHOS_PLANNER);
    linha = Number(linha);
    if (!(linha >= 2)) throw new Error('Linha inválida.');
    aba.getRange(linha, mapa[COL_PLANNER.concluida]).setValue(concluida ? true : false);
    return obterAtividades();
  } finally {
    lock.releaseLock();
  }
}

/** Exclui a atividade. */
function excluirAtividade(linha) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var aba = obterAbaPor_(NOME_ABA_PLANNER, CABECALHOS_PLANNER);
    linha = Number(linha);
    if (!(linha >= 2)) throw new Error('Linha inválida.');
    aba.deleteRow(linha);
    return obterAtividades();
  } finally {
    lock.releaseLock();
  }
}

/* ==========================================================================
 * Banco de Skills — cartões de skills do Claude (aba "Skills").
 * O upload de um .md ou .zip é analisado no servidor: extraímos "name" e
 * "description" do front-matter YAML (padrão SKILL.md do Claude). Sem
 * front-matter, usamos o nome do arquivo e o primeiro parágrafo como fallback.
 * ======================================================================== */

var NOME_ABA_SKILLS = 'Skills';
var CABECALHOS_SKILLS = ['Nome', 'Descrição', 'Conteúdo', 'Arquivo', 'Data de upload'];
var COL_SKILLS = {
  nome: 'Nome', descricao: 'Descrição', conteudo: 'Conteúdo',
  arquivo: 'Arquivo', dataUpload: 'Data de upload'
};
// Célula da planilha aceita ~50 mil caracteres; deixamos margem.
var LIMITE_CONTEUDO_SKILL = 45000;

/** Lê todas as skills. Chamado pelo cliente. */
function obterSkills() {
  var aba = obterAbaPor_(NOME_ABA_SKILLS, CABECALHOS_SKILLS);
  var mapa = obterMapaColunasPor_(aba, CABECALHOS_SKILLS);
  var ultCol = aba.getLastColumn();
  var ultima = aba.getLastRow();
  var itens = [];
  if (ultima >= 2) {
    var valores = aba.getRange(2, 1, ultima - 1, ultCol).getValues();
    valores.forEach(function (r, i) {
      var vazio = r.every(function (c) { return c === '' || c === null; });
      if (vazio) return;
      itens.push({
        linha: i + 2,
        nome: deCelula_(r[mapa[COL_SKILLS.nome] - 1]),
        descricao: deCelula_(r[mapa[COL_SKILLS.descricao] - 1]),
        conteudo: deCelula_(r[mapa[COL_SKILLS.conteudo] - 1]),
        arquivo: deCelula_(r[mapa[COL_SKILLS.arquivo] - 1]),
        dataUpload: deCelula_(r[mapa[COL_SKILLS.dataUpload] - 1])
      });
    });
  }
  return { skills: itens };
}

/**
 * Recebe um arquivo enviado pelo cliente, extrai nome/descrição e grava a skill.
 * payload: { arquivo: nome do arquivo, tipo: 'md'|'zip', base64: conteúdo }.
 */
function salvarSkillArquivo(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var bytes = Utilities.base64Decode(payload.base64 || '');
    var conteudo = '';
    var tipo = (payload.tipo || '').toLowerCase();
    var nomeArquivo = payload.arquivo || 'skill';

    if (tipo === 'zip' || /\.zip$/i.test(nomeArquivo)) {
      var blob = Utilities.newBlob(bytes, 'application/zip', nomeArquivo);
      var arquivos = Utilities.unzip(blob);
      var alvo = null;
      // Preferimos SKILL.md (padrão do Claude), depois qualquer .md.
      arquivos.forEach(function (f) {
        if (!alvo && /(^|\/)SKILL\.md$/i.test(f.getName())) alvo = f;
      });
      if (!alvo) {
        arquivos.forEach(function (f) {
          if (!alvo && /\.(md|markdown|txt)$/i.test(f.getName())) alvo = f;
        });
      }
      conteudo = alvo ? alvo.getDataAsString('UTF-8') : '';
    } else {
      conteudo = Utilities.newBlob(bytes).getDataAsString('UTF-8');
    }

    var meta = extrairMetaSkill_(conteudo);
    var nome = meta.nome || nomeBaseArquivo_(nomeArquivo);
    var descricao = meta.descricao || primeiroParagrafo_(conteudo) || 'Sem descrição.';
    if (conteudo.length > LIMITE_CONTEUDO_SKILL) {
      conteudo = conteudo.slice(0, LIMITE_CONTEUDO_SKILL) + '\n\n… (conteúdo truncado)';
    }

    var aba = obterAbaPor_(NOME_ABA_SKILLS, CABECALHOS_SKILLS);
    var mapa = obterMapaColunasPor_(aba, CABECALHOS_SKILLS);
    var ultCol = aba.getLastColumn();
    var linha = aba.getLastRow() + 1;
    var row = novaLinhaVazia_(ultCol);
    row[mapa[COL_SKILLS.nome] - 1] = nome;
    row[mapa[COL_SKILLS.descricao] - 1] = descricao;
    row[mapa[COL_SKILLS.conteudo] - 1] = conteudo;
    row[mapa[COL_SKILLS.arquivo] - 1] = nomeArquivo;
    row[mapa[COL_SKILLS.dataUpload] - 1] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    aba.getRange(linha, 1, 1, ultCol).setValues([row]);
    return obterSkills();
  } finally {
    lock.releaseLock();
  }
}

/** Edita manualmente o nome/descrição de uma skill já cadastrada. */
function atualizarSkill(linha, nome, descricao) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var aba = obterAbaPor_(NOME_ABA_SKILLS, CABECALHOS_SKILLS);
    var mapa = obterMapaColunasPor_(aba, CABECALHOS_SKILLS);
    linha = Number(linha);
    if (!(linha >= 2)) throw new Error('Linha inválida.');
    aba.getRange(linha, mapa[COL_SKILLS.nome]).setValue(nome || '');
    aba.getRange(linha, mapa[COL_SKILLS.descricao]).setValue(descricao || '');
    return obterSkills();
  } finally {
    lock.releaseLock();
  }
}

/** Exclui a skill. */
function excluirSkill(linha) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var aba = obterAbaPor_(NOME_ABA_SKILLS, CABECALHOS_SKILLS);
    linha = Number(linha);
    if (!(linha >= 2)) throw new Error('Linha inválida.');
    aba.deleteRow(linha);
    return obterSkills();
  } finally {
    lock.releaseLock();
  }
}

/** Extrai name/description do front-matter YAML (entre '---' no topo). */
function extrairMetaSkill_(texto) {
  var res = { nome: '', descricao: '' };
  if (!texto) return res;
  var t = texto.replace(/^﻿/, '');
  var m = t.match(/^---\s*\r?\n([\s\S]*?)\r?\n---/);
  if (!m) return res;
  var bloco = m[1];
  var nm = bloco.match(/^\s*name:\s*(.+?)\s*$/mi);
  var dm = bloco.match(/^\s*description:\s*(.+?)\s*$/mi);
  if (nm) res.nome = limparYaml_(nm[1]);
  if (dm) res.descricao = limparYaml_(dm[1]);
  return res;
}

/** Remove aspas envolventes e espaços de um valor YAML simples. */
function limparYaml_(v) {
  v = String(v == null ? '' : v).trim();
  if ((v.charAt(0) === '"' && v.slice(-1) === '"') ||
      (v.charAt(0) === "'" && v.slice(-1) === "'")) {
    v = v.slice(1, -1);
  }
  return v.trim();
}

/** Nome do arquivo sem extensão nem caminho (fallback para o nome da skill). */
function nomeBaseArquivo_(nome) {
  var base = String(nome || '').split('/').pop().split('\\').pop();
  return base.replace(/\.(md|markdown|txt|zip)$/i, '') || 'Skill';
}

/** Primeiro parágrafo de texto útil de um markdown (fallback de descrição). */
function primeiroParagrafo_(texto) {
  if (!texto) return '';
  var t = texto.replace(/^﻿/, '').replace(/^---\s*\r?\n[\s\S]*?\r?\n---\s*/, '');
  var linhas = t.split(/\r?\n/);
  for (var i = 0; i < linhas.length; i++) {
    var l = linhas[i].trim();
    if (!l) continue;
    if (l.charAt(0) === '#') { l = l.replace(/^#+\s*/, '').trim(); if (!l) continue; }
    if (l.length > 300) l = l.slice(0, 297) + '…';
    return l;
  }
  return '';
}