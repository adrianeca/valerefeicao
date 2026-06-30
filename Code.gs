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

// Retorna lista de unidades disponíveis para o usuário
function getUnidades(token) {
  const user = getSessionUser_(token);
  if (!user) throw new Error('Sessão inválida.');

  // Usuário com unidades específicas no Hub: retorna só as dele
  if (user.units && user.units.length > 0) return user.units;

  // Acesso total: descobre unidades ativas na planilha de funcionários
  const ss    = SpreadsheetApp.openById(FUNC_SHEET_ID);
  const sheet = ss.getSheetByName('RJ - UNIDADES');
  if (!sheet) throw new Error('Aba "RJ - UNIDADES" não encontrada.');
  const rows  = sheet.getDataRange().getValues();

  const norm = function(s) {
    return String(s || '').trim().toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '');
  };

  const set = {};
  for (let i = 1; i < rows.length; i++) {
    const nome = String(rows[i][COL.NOME] || '').trim();
    if (!nome) continue;
    const ativoRaw = norm(rows[i][COL.ATIVO]);
    if (ativoRaw === 'false' || ativoRaw === 'nao' || ativoRaw === 'no' || ativoRaw === '0') continue;
    const u = String(rows[i][COL.UNIDADE] || '').trim();
    if (u) set[u] = true;
  }
  return Object.keys(set).sort();
}

// =============================================================================
// PERÍODO VIGENTE
// =============================================================================

function getCurrentPeriod() {
  const now  = new Date();
  const mes  = now.getMonth() + 1;
  const ano  = now.getFullYear();

  // Previsto = próximo mês; Efetivo = mês atual
  let previsoMes = mes + 1;
  let previsoAno = ano;
  if (previsoMes > 12) { previsoMes = 1; previsoAno++; }

  // DEV: bloqueio desativado — restaurar para: const locked = now.getDate() > 11;
  const locked = false;

  return {
    efetivo:  { mes: mes,       ano: ano       },  // preenchido no próprio mês
    previsto: { mes: previsoMes, ano: previsoAno }, // preenchido um mês antes
    locked:   locked
  };
}

// =============================================================================
// FUNCIONÁRIOS
// =============================================================================

function getFuncionarios(unidade) {
  if (!unidade) throw new Error('Selecione uma unidade antes de carregar os funcionários.');

  const ss    = SpreadsheetApp.openById(FUNC_SHEET_ID);
  const sheet = ss.getSheetByName('RJ - UNIDADES');
  if (!sheet) throw new Error('Aba "RJ - UNIDADES" não encontrada na planilha de funcionários.');
  const rows  = sheet.getDataRange().getValues();

  const norm = function(s) {
    return String(s || '').trim().toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '');
  };

  const unidadeNorm    = norm(unidade);
  const administrativo = [];
  const docente        = [];

  for (let i = 1; i < rows.length; i++) {
    const row  = rows[i];
    const nome = String(row[COL.NOME]).trim();
    if (!nome) continue;

    const ativoRaw = norm(row[COL.ATIVO]);
    if (ativoRaw === 'false' || ativoRaw === 'nao' || ativoRaw === 'no' || ativoRaw === '0') continue;

    // Pertence à unidade (principal ou secundária) — comparação sem acento e sem case
    const unidadePrinc = norm(row[COL.UNIDADE]);
    const unidadeSec   = norm(row[COL.UNIDADE_SEC]);
    if (unidadePrinc !== unidadeNorm && unidadeSec !== unidadeNorm) continue;

    const funcao    = String(row[COL.FUNCAO]).trim().toUpperCase();
    const matricula = String(row[COL.MATRICULA]).trim();
    if (!matricula) continue;

    // Inclui a unidade principal do funcionário para desambiguar matrículas duplicadas entre unidades
    const emp = {
      nome:     nome,
      matricula: matricula,
      unidade:  String(row[COL.UNIDADE]).trim()
    };

    if (funcao === 'PROFESSOR') {
      docente.push(emp);
    } else {
      administrativo.push(emp);
    }
  }

  administrativo.sort(function(a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });
  docente.sort(function(a, b)        { return a.nome.localeCompare(b.nome, 'pt-BR'); });

  return { administrativo: administrativo, docente: docente };
}

// =============================================================================
// LEITURA DO VR
// =============================================================================

// Lê dados de DOIS períodos: previsto (próximo mês) e efetivo (mês atual)
// Colunas ADMINISTRATIVO: Unidade|Mes|Ano|Mat|Nome|Prev8|Prev6|PrevSab|Efet8|Efet6|EfetSab
// Colunas DOCENTE:        Unidade|Mes|Ano|Mat|Nome|CHPrev|CHEfet|DiasPrev|DiasEfet
function getVRData(unidade, efetivoMes, efetivoAno, previsoMes, previsoAno) {
  const ss           = SpreadsheetApp.openById(VR_SHEET_ID);
  const adminSheet   = ss.getSheetByName('ADMINISTRATIVO');
  const docenteSheet = ss.getSheetByName('DOCENTE');

  const adminMap   = {};
  const docenteMap = {};

  const adminRows = adminSheet.getDataRange().getValues();
  for (let i = 1; i < adminRows.length; i++) {
    const r = adminRows[i];
    if (String(r[0]).trim() !== unidade) continue;
    const rMes = Number(r[1]);
    const rAno = Number(r[2]);
    const mat  = String(r[3]).trim();
    if (!adminMap[mat]) adminMap[mat] = { previsto8hs:0, previsto6hs:0, previstoSab:0, efetivo8hs:0, efetivo6hs:0, efetivoSab:0 };
    if (rMes === previsoMes && rAno === previsoAno) {
      adminMap[mat].previsto8hs = r[5]  || 0;
      adminMap[mat].previsto6hs = r[6]  || 0;
      adminMap[mat].previstoSab = r[7]  || 0;
    }
    if (rMes === efetivoMes && rAno === efetivoAno) {
      adminMap[mat].efetivo8hs  = r[8]  || 0;
      adminMap[mat].efetivo6hs  = r[9]  || 0;
      adminMap[mat].efetivoSab  = r[10] || 0;
    }
  }

  const docenteRows = docenteSheet.getDataRange().getValues();
  for (let i = 1; i < docenteRows.length; i++) {
    const r = docenteRows[i];
    if (String(r[0]).trim() !== unidade) continue;
    const rMes = Number(r[1]);
    const rAno = Number(r[2]);
    const mat  = String(r[3]).trim();
    if (!docenteMap[mat]) docenteMap[mat] = { chPrevisto:0, chEfetivo:0, diasPrevisto:0, diasEfetivo:0 };
    if (rMes === previsoMes && rAno === previsoAno) {
      docenteMap[mat].chPrevisto   = r[5] || 0;
      docenteMap[mat].diasPrevisto = r[7] || 0;
    }
    if (rMes === efetivoMes && rAno === efetivoAno) {
      docenteMap[mat].chEfetivo   = r[6] || 0;
      docenteMap[mat].diasEfetivo = r[8] || 0;
    }
  }

  return { adminMap: adminMap, docenteMap: docenteMap };
}

// =============================================================================
// SALVAMENTO DO VR
// =============================================================================

function saveVRData(payload) {
  // DEV: validação de bloqueio desativada — restaurar após testes:
  // const period = getCurrentPeriod();
  // if (period.locked) throw new Error('O período está bloqueado. Prazo encerrado no dia 11.');

  const ss           = SpreadsheetApp.openById(VR_SHEET_ID);
  const adminSheet   = ss.getSheetByName('ADMINISTRATIVO');
  const docenteSheet = ss.getSheetByName('DOCENTE');

  const { unidade, efetivoMes, efetivoAno, previsoMes, previsoAno } = payload;

  // Previsto → linha do próximo mês (só atualiza colunas de previsto, preserva efetivo)
  _upsertPrevisto(adminSheet,   unidade, previsoMes, previsoAno, payload.administrativo, 'admin');
  _upsertPrevisto(docenteSheet, unidade, previsoMes, previsoAno, payload.docente,        'docente');

  // Efetivo → linha do mês atual (só atualiza colunas de efetivo, preserva previsto)
  _upsertEfetivo(adminSheet,   unidade, efetivoMes, efetivoAno, payload.administrativo, 'admin');
  _upsertEfetivo(docenteSheet, unidade, efetivoMes, efetivoAno, payload.docente,        'docente');

  return { success: true };
}

function _buildExistingMap_(sheet, unidade, mes, ano) {
  const allRows = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < allRows.length; i++) {
    const r = allRows[i];
    if (String(r[0]).trim() === unidade && Number(r[1]) === mes && Number(r[2]) === ano) {
      map[String(r[3]).trim()] = { rowNum: i + 1 };
    }
  }
  return map;
}

// Salva apenas os campos de previsto — preserva efetivo se a linha já existir
function _upsertPrevisto(sheet, unidade, mes, ano, employees, type) {
  if (!employees || !employees.length) return;
  const existing = _buildExistingMap_(sheet, unidade, mes, ano);

  for (const emp of employees) {
    const mat = String(emp.matricula).trim();
    if (type === 'admin') {
      if (existing[mat]) {
        // Atualiza só colunas 6-8 (Prev8, Prev6, PrevSab)
        sheet.getRange(existing[mat].rowNum, 6, 1, 3).setValues([[
          Number(emp.previsto8hs)||0, Number(emp.previsto6hs)||0, Number(emp.previstoSab)||0
        ]]);
      } else {
        sheet.appendRow([unidade, mes, ano, mat, emp.nome,
          Number(emp.previsto8hs)||0, Number(emp.previsto6hs)||0, Number(emp.previstoSab)||0,
          0, 0, 0]);
      }
    } else {
      if (existing[mat]) {
        sheet.getRange(existing[mat].rowNum, 6, 1, 1).setValues([[Number(emp.chPrevisto)||0]]);
        sheet.getRange(existing[mat].rowNum, 8, 1, 1).setValues([[Number(emp.diasPrevisto)||0]]);
      } else {
        sheet.appendRow([unidade, mes, ano, mat, emp.nome,
          Number(emp.chPrevisto)||0, 0, Number(emp.diasPrevisto)||0, 0]);
      }
    }
  }
}

// Salva apenas os campos de efetivo — preserva previsto se a linha já existir
function _upsertEfetivo(sheet, unidade, mes, ano, employees, type) {
  if (!employees || !employees.length) return;
  const existing = _buildExistingMap_(sheet, unidade, mes, ano);

  for (const emp of employees) {
    const mat = String(emp.matricula).trim();
    if (type === 'admin') {
      if (existing[mat]) {
        // Atualiza só colunas 9-11 (Efet8, Efet6, EfetSab)
        sheet.getRange(existing[mat].rowNum, 9, 1, 3).setValues([[
          Number(emp.efetivo8hs)||0, Number(emp.efetivo6hs)||0, Number(emp.efetivoSab)||0
        ]]);
      } else {
        sheet.appendRow([unidade, mes, ano, mat, emp.nome,
          0, 0, 0,
          Number(emp.efetivo8hs)||0, Number(emp.efetivo6hs)||0, Number(emp.efetivoSab)||0]);
      }
    } else {
      if (existing[mat]) {
        sheet.getRange(existing[mat].rowNum, 7, 1, 1).setValues([[Number(emp.chEfetivo)||0]]);
        sheet.getRange(existing[mat].rowNum, 9, 1, 1).setValues([[Number(emp.diasEfetivo)||0]]);
      } else {
        sheet.appendRow([unidade, mes, ano, mat, emp.nome,
          0, Number(emp.chEfetivo)||0, 0, Number(emp.diasEfetivo)||0]);
      }
    }
  }
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

  // Testa a leitura real para a primeira unidade encontrada
  const frows = SpreadsheetApp.openById(FUNC_SHEET_ID)
    .getSheetByName('RJ - UNIDADES').getDataRange().getValues();
  var unidade = '';
  for (var i = 1; i < frows.length; i++) {
    var u = String(frows[i][COL.UNIDADE] || '').trim();
    if (u) { unidade = u; break; }
  }
  Logger.log('\n=== getVRData para unidade "%s" ===', unidade);
  const result = getVRData(unidade,
    period.efetivo.mes,  period.efetivo.ano,
    period.previsto.mes, period.previsto.ano);
  Logger.log('adminMap: %s',   JSON.stringify(result.adminMap));
  Logger.log('docenteMap: %s', JSON.stringify(result.docenteMap));
}
