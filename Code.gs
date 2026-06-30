// =============================================================================
// CONFIGURAÇÕES
// =============================================================================

const VR_SHEET_ID    = '1spDbC6FRrImECVzNG6lEXsuJFalQP1tIgPG9jR8F00Q';
const FUNC_SHEET_ID  = '1BDiPjv0FqRJp5EwcvLdYXVvEAWesvwdEgbhYdnTlqPY';
const USERS_SHEET_ID = '1eZPbzhzjhjHoPwMhAW5YvOZgYiAvlTYc07dRan6Lyoc';

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

function getUserFromHub(token) {
  if (!token) throw new Error('Token não fornecido.');

  const ss  = SpreadsheetApp.openById(USERS_SHEET_ID);

  // ── 1. Valida sessão na aba SESSOES ──────────────────────────────────────
  // Colunas: TOKEN(0) | EMAIL(1) | NOME(2) | ROLE(3) | UNIDADE(4) | CRIADO_EM(5) | EXPIRA_EM(6) | ACESSOS(7)
  const sesSheet = ss.getSheetByName('SESSOES');
  if (!sesSheet) throw new Error('Configuração inválida. Contacte o administrador.');

  const sesData = sesSheet.getDataRange().getValues();
  const now     = new Date();
  let sesRow    = null;

  for (let i = 1; i < sesData.length; i++) {
    if (String(sesData[i][0]) !== String(token)) continue;
    const expira = sesData[i][6] ? new Date(sesData[i][6]) : null;
    if (!expira || expira < now) throw new Error('Sessão expirada. Acesse novamente pelo Hub.');
    sesRow = sesData[i];
    break;
  }

  if (!sesRow) throw new Error('Sessão não encontrada. Acesse novamente pelo Hub.');

  const email   = String(sesRow[1] || '').trim().toLowerCase();
  const nome    = String(sesRow[2] || '').trim();
  const role    = String(sesRow[3] || '').trim();
  const unidade = String(sesRow[4] || '').trim();

  // ── 2. Verifica acesso ao VR na aba USUARIOS ─────────────────────────────
  // Colunas: E-MAIL(0) | NOME(1) | ROLE(2) | UNIDADE(3) | ativo(4) | extra_dashboards(5) | acessos_dashboards(6)
  const usuSheet = ss.getSheetByName('USUARIOS');
  if (!usuSheet) throw new Error('Configuração inválida. Contacte o administrador.');

  const usuData   = usuSheet.getDataRange().getValues();
  let hasAccess   = false;

  for (let i = 1; i < usuData.length; i++) {
    const rowEmail = String(usuData[i][0] || '').trim().toLowerCase();
    if (rowEmail !== email) continue;

    const ativo = String(usuData[i][4] || '').trim().toUpperCase();
    if (ativo === 'FALSE') throw new Error('Usuário inativo no sistema.');

    // Verifica "webvr" nas colunas F (extra_dashboards) e G (acessos_dashboards)
    const colF = String(usuData[i][5] || '').toLowerCase();
    const colG = String(usuData[i][6] || '').toLowerCase();
    hasAccess  = (colF + ',' + colG)
      .split(',')
      .map(function(s) { return s.trim(); })
      .some(function(a) { return a === 'webvr' || a === 'vr'; });
    break;
  }

  if (!hasAccess) {
    throw new Error('Você não tem permissão para acessar o Vale Refeição. Contacte o administrador.');
  }

  return { email: email, nome: nome, role: role, unidade: unidade };
}

// =============================================================================
// PERÍODO VIGENTE
// =============================================================================

function getCurrentPeriod() {
  const now   = new Date();
  const month = now.getMonth() + 1; // 1–12
  const year  = now.getFullYear();

  let targetMonth = month + 1;
  let targetYear  = year;
  if (targetMonth > 12) { targetMonth = 1; targetYear++; }

  // DEV: bloqueio desativado — restaurar para: const locked = now.getDate() > 11;
  const locked = false;

  return { mes: targetMonth, ano: targetYear, locked: locked };
}

// =============================================================================
// FUNCIONÁRIOS
// =============================================================================

function getFuncionarios(unidade) {
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

    // Pula se explicitamente inativo — aceita qualquer valor que não seja falso
    const ativoRaw = norm(row[COL.ATIVO]);
    if (ativoRaw === 'false' || ativoRaw === 'nao' || ativoRaw === 'no' || ativoRaw === '0') continue;

    // Pertence à unidade (principal ou secundária) — comparação sem acento e sem case
    const unidadePrinc = norm(row[COL.UNIDADE]);
    const unidadeSec   = norm(row[COL.UNIDADE_SEC]);
    if (unidadePrinc !== unidadeNorm && unidadeSec !== unidadeNorm) continue;

    const funcao    = String(row[COL.FUNCAO]).trim().toUpperCase();
    const matricula = String(row[COL.MATRICULA]).trim();
    if (!matricula) continue;

    const emp = { nome: nome, matricula: matricula };

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

function getVRData(unidade, mes, ano) {
  const ss          = SpreadsheetApp.openById(VR_SHEET_ID);
  const adminSheet  = ss.getSheetByName('ADMINISTRATIVO');
  const docenteSheet = ss.getSheetByName('DOCENTE');

  const adminMap   = {};
  const docenteMap = {};

  // ADMINISTRATIVO: Unidade|Mês|Ano|Matrícula|Nome|Prev8|Prev6|PrevSab|Efet8|Efet6|EfetSab
  const adminRows = adminSheet.getDataRange().getValues();
  for (let i = 1; i < adminRows.length; i++) {
    const r = adminRows[i];
    if (String(r[0]).trim() !== unidade) continue;
    if (Number(r[1]) !== mes || Number(r[2]) !== ano) continue;
    const mat = String(r[3]).trim();
    adminMap[mat] = {
      previsto8hs:  r[5]  || 0,
      previsto6hs:  r[6]  || 0,
      previstoSab:  r[7]  || 0,
      efetivo8hs:   r[8]  || 0,
      efetivo6hs:   r[9]  || 0,
      efetivoSab:   r[10] || 0
    };
  }

  // DOCENTE: Unidade|Mês|Ano|Matrícula|Nome|CHPrev|CHEfet|DiasPrev|DiasEfet
  const docenteRows = docenteSheet.getDataRange().getValues();
  for (let i = 1; i < docenteRows.length; i++) {
    const r = docenteRows[i];
    if (String(r[0]).trim() !== unidade) continue;
    if (Number(r[1]) !== mes || Number(r[2]) !== ano) continue;
    const mat = String(r[3]).trim();
    docenteMap[mat] = {
      chPrevisto:   r[5] || 0,
      chEfetivo:    r[6] || 0,
      diasPrevisto: r[7] || 0,
      diasEfetivo:  r[8] || 0
    };
  }

  return { adminMap: adminMap, docenteMap: docenteMap };
}

// =============================================================================
// SALVAMENTO DO VR
// =============================================================================

function saveVRData(payload) {
  const period = getCurrentPeriod();

  // DEV: validação de bloqueio desativada — restaurar após testes
  // if (period.locked) throw new Error('O período está bloqueado. Prazo encerrado no dia 11.');
  // if (Number(payload.mes) !== period.mes || Number(payload.ano) !== period.ano)
  //   throw new Error('Período inválido para edição.');

  const ss           = SpreadsheetApp.openById(VR_SHEET_ID);
  const adminSheet   = ss.getSheetByName('ADMINISTRATIVO');
  const docenteSheet = ss.getSheetByName('DOCENTE');

  _upsertRows(adminSheet,   payload.unidade, payload.mes, payload.ano, payload.administrativo, 'admin');
  _upsertRows(docenteSheet, payload.unidade, payload.mes, payload.ano, payload.docente,        'docente');

  return { success: true };
}

function _upsertRows(sheet, unidade, mes, ano, employees, type) {
  if (!employees || employees.length === 0) return;

  const allRows = sheet.getDataRange().getValues();

  // Mapeia matrícula → número da linha na planilha (base 1)
  const existing = {};
  for (let i = 1; i < allRows.length; i++) {
    const r = allRows[i];
    if (String(r[0]).trim() === unidade && Number(r[1]) === mes && Number(r[2]) === ano) {
      existing[String(r[3]).trim()] = i + 1;
    }
  }

  for (let j = 0; j < employees.length; j++) {
    const emp = employees[j];
    const mat = String(emp.matricula).trim();
    let newRow;

    if (type === 'admin') {
      newRow = [
        unidade, mes, ano, mat, emp.nome,
        Number(emp.previsto8hs) || 0,
        Number(emp.previsto6hs) || 0,
        Number(emp.previstoSab) || 0,
        Number(emp.efetivo8hs)  || 0,
        Number(emp.efetivo6hs)  || 0,
        Number(emp.efetivoSab)  || 0
      ];
    } else {
      newRow = [
        unidade, mes, ano, mat, emp.nome,
        Number(emp.chPrevisto)   || 0,
        Number(emp.chEfetivo)    || 0,
        Number(emp.diasPrevisto) || 0,
        Number(emp.diasEfetivo)  || 0
      ];
    }

    if (existing[mat]) {
      sheet.getRange(existing[mat], 1, 1, newRow.length).setValues([newRow]);
    } else {
      sheet.appendRow(newRow);
    }
  }
}
