// =============================================================================
// CONFIGURAÇÕES
// =============================================================================

const VR_SHEET_ID   = '1spDbC6FRrImECVzNG6lEXsuJFalQP1tIgPG9jR8F00Q';
const FUNC_SHEET_ID = '1BDiPjv0FqRJp5EwcvLdYXVvEAWesvwdEgbhYdnTlqPY';
const HUB_SS_ID     = '1eZPbzhzjhjHoPwMhAW5YvOZgYiAvlTYc07dRan6Lyoc';
const MEU_ACESSO    = 'webvr';

// Índices das colunas na planilha de funcionários (base 0)
const COL = {
  NOME:         2,   // C
  FUNCAO:       5,   // F
  ATIVO:        10,  // K
  MATRICULA:    27,  // AB
  UNIDADE:      21,  // V
  UNIDADE_SEC:  30   // AE
};

// Normaliza texto para comparação: minúsculo, sem acento, sem espaços nas bordas
function norm_(s) {
  return String(s || '').trim().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '');
}

// Extrai o número do mês mesmo quando a célula guarda texto como "06 Junho" (em vez de 6)
function parseMes_(v) {
  return parseInt(String(v).trim(), 10) || 0;
}

// =============================================================================
// ENTRY POINT
// =============================================================================

function doGet(e) {
  const token = (e && e.parameter && e.parameter.s) ? e.parameter.s : '';
  const tmpl = HtmlService.createTemplateFromFile('Index');
  tmpl.token = token;
  return tmpl.evaluate()
    .setTitle('Vale Refeição — BRASAS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =============================================================================
// AUTENTICAÇÃO
// =============================================================================

// Colunas SESSOES: TOKEN(0)|EMAIL(1)|NOME(2)|ROLE(3)|UNIDADE(4)|CRIADO_EM(5)|EXPIRA_EM(6)|ACESSOS(7)
// UNIDADE pode ser pipe-separado (ex: "BG|FG"). Vazio = acesso a todas.

function getUserFromHub(token) {
  if (!token) throw new Error('Token não fornecido.');

  const user = getSessionUser_(token);
  if (!user) throw new Error('Sessão inválida ou expirada. Acesse novamente pelo Hub.');

  return user;
}

function getSessionUser_(token) {
  if (!token) return null;
  try {
    const ss       = SpreadsheetApp.openById(HUB_SS_ID);
    const sesSheet = ss.getSheetByName('SESSOES');
    if (!sesSheet) return null;

    const tok   = String(token).trim();
    const found = sesSheet.getRange(1, 1, sesSheet.getLastRow(), 1)
      .createTextFinder(tok).matchEntireCell(true).findNext();
    if (!found) return null;

    // [TOKEN, EMAIL, NOME, ROLE, UNIDADE, CRIADO_EM, EXPIRA_EM, ACESSOS]
    const row = sesSheet.getRange(found.getRow(), 1, 1, 8).getValues()[0];

    if (row[6] && new Date(row[6]) < new Date()) return null; // expirado

    const email = String(row[1] || '').trim().toLowerCase();
    if (!email) return null;

    // Verifica acesso a este dashboard na coluna ACESSOS
    const acessos = String(row[7] || '').toLowerCase()
      .split(',').map(function(a) { return a.trim(); });
    if (!acessos.includes(MEU_ACESSO)) {
      throw new Error('Você não tem permissão para acessar o Vale Refeição. Contacte o administrador.');
    }

    // UNIDADE: vazio = todas; pipe-separado = restringe a essas
    const unidadeRaw = String(row[4] || '').trim();
    const units = unidadeRaw
      ? unidadeRaw.split('|').map(function(u) { return u.trim(); }).filter(Boolean)
      : [];

    return {
      email:    email,
      nome:     String(row[2] || '').trim(),
      role:     String(row[3] || '').trim().toLowerCase(),
      unidade:  units[0] || '',
      units:    units  // [] = acesso total; preenchido = só essas unidades
    };
  } catch (e) {
    if (e.message && e.message.includes('permissão')) throw e;
    Logger.log('getSessionUser_: ' + e);
    return null;
  }
}

function isUserAllowedUnit_(user, unit) {
  if (!user.units || !user.units.length) return true; // acesso total
  return user.units.some(function(u) {
    return u.toLowerCase().trim() === unit.toLowerCase().trim();
  });
}

// Todas as unidades que o usuário pode ver: as dele (se restrito) ou todas que existem
// (com funcionário ativo cadastrado OU já com lançamento na planilha de VR)
function getAllowedUnidades_(user) {
  if (user.units && user.units.length > 0) return user.units;

  const set = {};

  const funcSheet = SpreadsheetApp.openById(FUNC_SHEET_ID).getSheetByName('RJ - UNIDADES');
  if (!funcSheet) throw new Error('Aba "RJ - UNIDADES" não encontrada.');
  const funcRows = funcSheet.getDataRange().getValues();
  for (let i = 1; i < funcRows.length; i++) {
    const nome = String(funcRows[i][COL.NOME] || '').trim();
    if (!nome) continue;
    const ativoRaw = norm_(funcRows[i][COL.ATIVO]);
    if (ativoRaw === 'false' || ativoRaw === 'nao' || ativoRaw === 'no' || ativoRaw === '0') continue;
    const u = String(funcRows[i][COL.UNIDADE] || '').trim();
    if (u) set[u] = true;
  }

  const vrSs = SpreadsheetApp.openById(VR_SHEET_ID);
  ['ADMINISTRATIVO', 'DOCENTE'].forEach(function(sheetName) {
    const sheet = vrSs.getSheetByName(sheetName);
    if (!sheet) return;
    const vrRows = sheet.getDataRange().getValues();
    for (let i = 1; i < vrRows.length; i++) {
      const u = String(vrRows[i][0] || '').trim();
      if (u) set[u] = true;
    }
  });

  return Object.keys(set).sort();
}

// Retorna lista de unidades disponíveis para o usuário
function getUnidades(token) {
  const user = getSessionUser_(token);
  if (!user) throw new Error('Sessão inválida.');
  return getAllowedUnidades_(user);
}

// =============================================================================
// PERÍODO VIGENTE
// =============================================================================

function getCurrentPeriod(token) {
  const now  = new Date();
  const mes  = now.getMonth() + 1;
  const ano  = now.getFullYear();

  // Previsto = próximo mês; Efetivo = mês atual
  let previsoMes = mes + 1;
  let previsoAno = ano;
  if (previsoMes > 12) { previsoMes = 1; previsoAno++; }

  // DEV: bloqueio desativado — restaurar para: let locked = now.getDate() > 11;
  let locked = false;

  // Liberação temporária (24h) concedida por um admin ignora o bloqueio para esse usuário
  if (locked) {
    const user = getSessionUser_(token);
    if (user && hasActiveLiberacao_(user.email)) locked = false;
  }

  return {
    efetivo:  { mes: mes,       ano: ano       },  // preenchido no próprio mês
    previsto: { mes: previsoMes, ano: previsoAno }, // preenchido um mês antes
    locked:   locked
  };
}

// =============================================================================
// LIBERAÇÕES TEMPORÁRIAS DE EDIÇÃO (24h) — restrito a admins
// =============================================================================

// Colunas: Email | Liberado Por | Criado Em | Expira Em
function getLiberacoesSheet_() {
  const ss = SpreadsheetApp.openById(VR_SHEET_ID);
  let sheet = ss.getSheetByName('LIBERACOES');
  if (!sheet) {
    sheet = ss.insertSheet('LIBERACOES');
    sheet.appendRow(['Email', 'Liberado Por', 'Criado Em', 'Expira Em']);
  }
  return sheet;
}

function requireAdmin_(token) {
  const user = getSessionUser_(token);
  if (!user) throw new Error('Sessão inválida ou expirada. Acesse novamente pelo Hub.');
  if (user.role !== 'admin') throw new Error('Acesso restrito a administradores.');
  return user;
}

function hasActiveLiberacao_(email) {
  if (!email) return false;
  const emailNorm = norm_(email);
  const now  = new Date();
  const rows = getLiberacoesSheet_().getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (norm_(r[0]) === emailNorm && r[3] && new Date(r[3]) > now) return true;
  }
  return false;
}

// Lista todas as liberações já concedidas (mais recentes primeiro) — só para admins
function getLiberacoes(token) {
  requireAdmin_(token);
  const rows = getLiberacoesSheet_().getDataRange().getValues();

  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    list.push({ email: String(r[0]).trim(), liberadoPor: String(r[1]).trim(), criadoEm: r[2], expiraEm: r[3] });
  }
  list.sort(function(a, b) { return new Date(b.criadoEm) - new Date(a.criadoEm); });
  return list;
}

// Concede 24h de edição liberada para um e-mail — só admins podem chamar
function criarLiberacao(token, email) {
  const admin = requireAdmin_(token);

  email = String(email || '').trim().toLowerCase();
  if (!email || email.indexOf('@') === -1) throw new Error('Informe um e-mail válido.');

  const now    = new Date();
  const expira = new Date(now.getTime() + 24 * 60 * 60 * 1000);

  getLiberacoesSheet_().appendRow([email, admin.email, now, expira]);
  return getLiberacoes(token);
}

// =============================================================================
// FUNCIONÁRIOS
// =============================================================================

// Retorna os funcionários de TODAS as unidades que o usuário pode ver.
// Um funcionário com unidade principal + secundária aparece uma vez para cada uma
// (desde que esteja entre as unidades permitidas), pois cada uma é um contexto de lançamento distinto.
function getFuncionarios(token) {
  const user = getSessionUser_(token);
  if (!user) throw new Error('Sessão inválida.');
  const allowedNorm = getAllowedUnidades_(user).map(norm_);

  const ss    = SpreadsheetApp.openById(FUNC_SHEET_ID);
  const sheet = ss.getSheetByName('RJ - UNIDADES');
  if (!sheet) throw new Error('Aba "RJ - UNIDADES" não encontrada na planilha de funcionários.');
  const rows  = sheet.getDataRange().getValues();

  const administrativo = [];
  const docente        = [];

  for (let i = 1; i < rows.length; i++) {
    const row  = rows[i];
    const nome = String(row[COL.NOME]).trim();
    if (!nome) continue;

    const ativoRaw = norm_(row[COL.ATIVO]);
    if (ativoRaw === 'false' || ativoRaw === 'nao' || ativoRaw === 'no' || ativoRaw === '0') continue;

    const matricula = String(row[COL.MATRICULA]).trim();
    if (!matricula) continue;

    const funcao = String(row[COL.FUNCAO]).trim().toUpperCase();
    const list   = funcao === 'PROFESSOR' ? docente : administrativo;

    // Unidade principal + secundária, sem repetir se forem iguais
    const unidades = [String(row[COL.UNIDADE]).trim(), String(row[COL.UNIDADE_SEC]).trim()]
      .filter(function(u, idx, arr) { return u && arr.indexOf(u) === idx; });

    unidades.forEach(function(u) {
      if (allowedNorm.indexOf(norm_(u)) === -1) return;
      list.push({ nome: nome, matricula: matricula, unidade: u });
    });
  }

  administrativo.sort(function(a, b) { return a.nome.localeCompare(b.nome, 'pt-BR') || a.unidade.localeCompare(b.unidade, 'pt-BR'); });
  docente.sort(function(a, b)        { return a.nome.localeCompare(b.nome, 'pt-BR') || a.unidade.localeCompare(b.unidade, 'pt-BR'); });

  return { administrativo: administrativo, docente: docente };
}

// =============================================================================
// LEITURA DO VR — todas as linhas de todas as unidades permitidas, estilo planilha
// =============================================================================

// Colunas ADMINISTRATIVO: Unidade|Mes|Ano|Mat|Nome|Prev8|Prev6|PrevSab|Efet8|Efet6|EfetSab
// Colunas DOCENTE:        Unidade|Mes|Ano|Mat|Nome|CHPrev|CHEfet|DiasPrev|DiasEfet
// Retorna TODAS as linhas já lançadas nas unidades que o usuário pode ver, na ordem cronológica
function getVRData(token) {
  const user = getSessionUser_(token);
  if (!user) throw new Error('Sessão inválida.');
  const allowedNorm = getAllowedUnidades_(user).map(norm_);

  const ss           = SpreadsheetApp.openById(VR_SHEET_ID);
  const adminSheet   = ss.getSheetByName('ADMINISTRATIVO');
  const docenteSheet = ss.getSheetByName('DOCENTE');

  const administrativo = [];
  const adminRows = adminSheet.getDataRange().getValues();
  for (let i = 1; i < adminRows.length; i++) {
    const r = adminRows[i];
    if (allowedNorm.indexOf(norm_(r[0])) === -1) continue;
    const mes = parseMes_(r[1]), ano = Number(r[2]);
    if (!mes || !ano) continue;
    administrativo.push({
      unidade: String(r[0]).trim(), mes: mes, ano: ano,
      matricula: String(r[3]).trim(), nome: String(r[4]).trim(),
      previsto8hs: r[5] || 0, previsto6hs: r[6] || 0, previstoSab: r[7] || 0,
      efetivo8hs:  r[8] || 0, efetivo6hs:  r[9] || 0, efetivoSab:  r[10] || 0
    });
  }
  administrativo.sort(function(a, b) {
    return (a.ano - b.ano) || (a.mes - b.mes) || a.unidade.localeCompare(b.unidade, 'pt-BR') || a.nome.localeCompare(b.nome, 'pt-BR');
  });

  const docente = [];
  const docenteRows = docenteSheet.getDataRange().getValues();
  for (let i = 1; i < docenteRows.length; i++) {
    const r = docenteRows[i];
    if (allowedNorm.indexOf(norm_(r[0])) === -1) continue;
    const mes = parseMes_(r[1]), ano = Number(r[2]);
    if (!mes || !ano) continue;
    docente.push({
      unidade: String(r[0]).trim(), mes: mes, ano: ano,
      matricula: String(r[3]).trim(), nome: String(r[4]).trim(),
      chPrevisto: r[5] || 0, chEfetivo: r[6] || 0,
      diasPrevisto: r[7] || 0, diasEfetivo: r[8] || 0
    });
  }
  docente.sort(function(a, b) {
    return (a.ano - b.ano) || (a.mes - b.mes) || a.unidade.localeCompare(b.unidade, 'pt-BR') || a.nome.localeCompare(b.nome, 'pt-BR');
  });

  return { administrativo: administrativo, docente: docente };
}

// =============================================================================
// SALVAMENTO DO VR — cada item do payload já traz sua própria unidade/mês/ano
// =============================================================================

function saveVRData(payload) {
  // DEV: validação de bloqueio desativada — restaurar após testes:
  // const period = getCurrentPeriod();
  // if (period.locked) throw new Error('O período está bloqueado. Prazo encerrado no dia 11.');

  const user = getSessionUser_(payload.token);
  if (!user) throw new Error('Sessão inválida ou expirada. Acesse novamente pelo Hub.');

  const ss           = SpreadsheetApp.openById(VR_SHEET_ID);
  const adminSheet   = ss.getSheetByName('ADMINISTRATIVO');
  const docenteSheet = ss.getSheetByName('DOCENTE');

  // Nunca confia na unidade vinda do cliente sem checar permissão
  const adminEntries   = (payload.administrativo || []).filter(function(e) { return isUserAllowedUnit_(user, e.unidade); });
  const docenteEntries = (payload.docente        || []).filter(function(e) { return isUserAllowedUnit_(user, e.unidade); });

  _upsertRows_(adminSheet, adminEntries, function(e) {
    return [Number(e.previsto8hs)||0, Number(e.previsto6hs)||0, Number(e.previstoSab)||0,
            Number(e.efetivo8hs)||0,  Number(e.efetivo6hs)||0,  Number(e.efetivoSab)||0];
  });

  _upsertRows_(docenteSheet, docenteEntries, function(e) {
    return [Number(e.chPrevisto)||0, Number(e.chEfetivo)||0,
            Number(e.diasPrevisto)||0, Number(e.diasEfetivo)||0];
  });

  return { success: true };
}

// Atualiza a linha existente (unidade+mes+ano+matricula) ou cria uma nova, para cada item
function _upsertRows_(sheet, entries, valuesFn) {
  if (!entries || !entries.length) return;

  const allRows = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < allRows.length; i++) {
    const r = allRows[i];
    map[norm_(r[0]) + '|' + parseMes_(r[1]) + '|' + Number(r[2]) + '|' + String(r[3]).trim()] = i + 1;
  }

  entries.forEach(function(e) {
    const mat    = String(e.matricula).trim();
    const key    = norm_(e.unidade) + '|' + Number(e.mes) + '|' + Number(e.ano) + '|' + mat;
    const values = valuesFn(e);
    if (map[key]) {
      sheet.getRange(map[key], 6, 1, values.length).setValues([values]);
    } else {
      sheet.appendRow([e.unidade, e.mes, e.ano, mat, e.nome].concat(values));
      map[key] = sheet.getLastRow();
    }
  });
}

// =============================================================================
// DIAGNÓSTICO — rode no editor do Apps Script e veja os logs (Ctrl+Enter)
// =============================================================================

function diagnosticoVR() {
  const period = getCurrentPeriod();
  const ss     = SpreadsheetApp.openById(VR_SHEET_ID);
  const admin  = ss.getSheetByName('ADMINISTRATIVO');
  const doc    = ss.getSheetByName('DOCENTE');

  Logger.log('=== PERÍODO ATUAL ===');
  Logger.log('Efetivo : mês %s / ano %s', period.efetivo.mes,  period.efetivo.ano);
  Logger.log('Previsto: mês %s / ano %s', period.previsto.mes, period.previsto.ano);

  Logger.log('\n=== LINHAS EM ADMINISTRATIVO ===');
  admin.getDataRange().getValues().slice(1).forEach(function(r, i) {
    Logger.log('Linha %s → unidade="%s" mes=%s ano=%s mat="%s" vals=%s',
      i + 2, r[0], r[1], r[2], r[3], JSON.stringify(r.slice(5)));
  });

  Logger.log('\n=== LINHAS EM DOCENTE ===');
  doc.getDataRange().getValues().slice(1).forEach(function(r, i) {
    Logger.log('Linha %s → unidade="%s" mes=%s ano=%s mat="%s" vals=%s',
      i + 2, r[0], r[1], r[2], r[3], JSON.stringify(r.slice(5)));
  });
}
