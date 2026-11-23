/**
 * Operação Curupira-Mirim II 2026 — Amazonas
 * Combate à violência contra crianças e adolescentes
 * Google Apps Script — Backend Web App v1
 *
 * ═══════════════════════════════════════════════════════════════
 * SETUP (primeira vez):
 *  1. Acesse script.google.com → Novo projeto (sem vínculo)
 *  2. Cole este código inteiro
 *  3. Execute setupSheet() uma vez → cria os cabeçalhos na planilha
 *  4. Implantar → Nova implantação → App da Web
 *       Executar como: Eu
 *       Quem tem acesso: Qualquer pessoa
 *  5. Copie a URL da implantação → cole no HTML (variável GAS_URL)
 * ═══════════════════════════════════════════════════════════════
 */

const SPREADSHEET_ID = '1_X26wbOHR-YFYpiAm1b9jPW0Dv8keUt49MfO-NEIAk0';
const SHEET_NAME     = 'Respostas';
const ADMIN_PWD      = 'admin1311';

// ── COLUMN ORDER (28 colunas A → AB) ────────────────────────────
const COLUMN_ORDER = [
  // Identificação A-E
  'timestamp',              // A
  'data',                   // B
  'orgao',                  // C
  'responsavel',            // D
  'whatsapp',               // E
  // Abrangência F-K
  'efetivo_policial',       // F  PC/PM
  'viaturas_empregadas',    // G  PC/PM
  'efetivo_geral',          // H  outros órgãos
  'carros_empregados',      // I  outros órgãos
  'municipios',             // J  nº
  'municipios_lista',       // K  CSV
  // Ações Preventivas e Educativas L-O
  'ac_ed_realizadas',       // L
  'comunidades_visitadas',  // M
  'locais_fiscalizados',    // N
  'pessoas_alcancadas',     // O
  // Diligências P-S (PC)
  'denuncias_recebidas',    // P
  'denuncias_apuradas',     // Q
  'vitimas_atendidas',      // R
  'locais_mapeados',        // S
  // Ações Repressivas T-V (PC)
  'presos_flagrante',       // T
  'presos_mandado',         // U
  'mba_cumpridos',          // V
  // Ações Cartorárias W-Y (PC)
  'bos_registrados',        // W
  'ips_instaurados',        // X
  'ips_concluidos',         // Y
  // Extra Z-AB
  'orgao_depto',            // Z  Departamento/Subcategoria
  'orgao_outro',            // AA
  'resumo_acao',            // AB
  'id',                     // AC
];

const HEADERS = {
  timestamp:              'Timestamp',
  data:                   'Data da Ação',
  orgao:                  'Órgão',
  responsavel:            'Responsável',
  whatsapp:               'WhatsApp',
  efetivo_policial:       'Efetivo Policial (PC/PM)',
  viaturas_empregadas:    'Viaturas Empregadas (PC/PM)',
  efetivo_geral:          'Efetivo Empregado (outros órgãos)',
  carros_empregados:      'Carros Empregados (outros órgãos)',
  municipios:             'Nº Municípios',
  municipios_lista:       'Municípios (lista)',
  ac_ed_realizadas:       'Ações Educativas Realizadas',
  comunidades_visitadas:  'Comunidades Indígenas e Ribeirinhas Visitadas',
  locais_fiscalizados:    'Locais Fiscalizados',
  pessoas_alcancadas:     'Pessoas Alcançadas com as Ações Educativas',
  denuncias_recebidas:    'Denúncias Recebidas',
  denuncias_apuradas:     'Denúncias Apuradas',
  vitimas_atendidas:      'Vítimas Atendidas',
  locais_mapeados:        'Locais Vulneráveis Mapeados',
  presos_flagrante:       'Presos em Flagrante Delito',
  presos_mandado:         'Presos por Mandado de Prisão',
  mba_cumpridos:          'MBA Cumpridos',
  bos_registrados:        'BOs Registrados',
  ips_instaurados:        'IPs Instaurados',
  ips_concluidos:         'IPs Concluídos com Autoria Identificada',
  orgao_depto:            'Departamento / Subcategoria',
  orgao_outro:            'Órgão (especificado)',
  resumo_acao:            'Resumo da Ação',
  id:                     'ID',
};

// Grupos de cabeçalho (linha 1, cores)
const HEADER_GROUPS = [
  {label:'IDENTIFICAÇÃO',               from:1,  to:5,  bg:'#3D1A6B'},
  {label:'ABRANGÊNCIA',                 from:6,  to:11, bg:'#1E40AF'},
  {label:'AÇÕES PREVENTIVAS E EDUCATIVAS', from:12, to:15, bg:'#065F46'},
  {label:'DILIGÊNCIAS (PC)',            from:16, to:19, bg:'#92400E'},
  {label:'AÇÕES REPRESSIVAS (PC)',      from:20, to:22, bg:'#7F1D1D'},
  {label:'AÇÕES CARTORÁRIAS (PC)',      from:23, to:25, bg:'#4C1D95'},
  {label:'EXTRA',                       from:26, to:29, bg:'#374151'},
];

const ID_COL_INDEX = COLUMN_ORDER.indexOf('id');

// ── SETUP — execute uma vez para criar os cabeçalhos ────────────

function setupSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  } else {
    sheet.getRange(1, 1, 2, Math.max(sheet.getLastColumn(), COLUMN_ORDER.length)).clearContent().clearFormat();
  }

  const numCols = COLUMN_ORDER.length;

  // Linha 1: grupos
  const groupRow = Array(numCols).fill('');
  HEADER_GROUPS.forEach(g => { groupRow[g.from - 1] = g.label; });
  sheet.getRange(1, 1, 1, numCols).setValues([groupRow]);
  HEADER_GROUPS.forEach(g => {
    const r = sheet.getRange(1, g.from, 1, g.to - g.from + 1);
    r.setBackground(g.bg); r.setFontColor('#FFFFFF');
    r.setFontWeight('bold'); r.setFontSize(8);
    if (g.from < g.to) r.merge();
  });

  // Linha 2: nomes dos campos
  const headers = COLUMN_ORDER.map(k => HEADERS[k] || k);
  const hRange = sheet.getRange(2, 1, 1, numCols);
  hRange.setValues([headers]);
  hRange.setBackground('#5B2D8E'); hRange.setFontColor('#FFFFFF');
  hRange.setFontWeight('bold'); hRange.setFontSize(9);

  sheet.setFrozenRows(2);
  sheet.setColumnWidth(1, 140);   // timestamp
  sheet.setColumnWidth(3, 160);   // orgao
  sheet.setColumnWidth(4, 160);   // responsavel
  sheet.setColumnWidth(11, 300);  // municipios_lista
  sheet.setColumnWidth(COLUMN_ORDER.indexOf('resumo_acao') + 1, 350);
  sheet.getRange(2, 1, 1, numCols).createFilter();

  Logger.log('✅ setupSheet concluído. ' + ss.getName() + ' | aba ' + SHEET_NAME);
}

// ── HELPERS ─────────────────────────────────────────────────────

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Aba "' + SHEET_NAME + '" não encontrada. Execute setupSheet() primeiro.');
  return sheet;
}

// ── POST ────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data._action === 'clearAll') {
      if (data._pwd !== ADMIN_PWD) return jsonOut({success:false, error:'Senha incorreta'});
      const sheet = getSheet();
      const last = sheet.getLastRow();
      if (last > 2) sheet.deleteRows(3, last - 2);
      return jsonOut({success:true, cleared:true});
    }

    if (data._action === 'deleteOne') {
      if (data._pwd !== ADMIN_PWD) return jsonOut({success:false, error:'Senha incorreta'});
      const sheet = getSheet();
      const last = sheet.getLastRow();
      if (last <= 2) return jsonOut({success:true, deleted:false});
      const col = ID_COL_INDEX + 1;
      const values = sheet.getRange(3, col, last - 2, 1).getValues();
      for (let i = 0; i < values.length; i++) {
        if (String(values[i][0]) === String(data.id)) {
          sheet.deleteRow(i + 3);
          return jsonOut({success:true, deleted:true});
        }
      }
      return jsonOut({success:true, deleted:false, reason:'id not found'});
    }

    // Salvar registro
    const sheet = getSheet();
    const row = COLUMN_ORDER.map(k =>
      (data[k] !== undefined && data[k] !== null && data[k] !== '') ? data[k] : ''
    );
    sheet.appendRow(row);
    const lastRow = sheet.getLastRow();
    if ((lastRow - 2) % 2 === 0) {
      sheet.getRange(lastRow, 1, 1, COLUMN_ORDER.length).setBackground('#F3EEF9');
    }
    return jsonOut({success:true, row:lastRow});

  } catch (err) {
    return jsonOut({success:false, error:err.message});
  }
}

// ── GET ─────────────────────────────────────────────────────────

function doGet(e) {
  if (e.parameter.test === 'true') {
    return jsonOut({status:'online', sheet:SHEET_NAME, spreadsheet:SPREADSHEET_ID, version:'v1', cols:COLUMN_ORDER.length});
  }

  if (e.parameter.action === 'getData') {
    try {
      const sheet = getSheet();
      const lastRow = sheet.getLastRow();
      if (lastRow <= 2) return jsonOut({success:true, data:[]});
      const lastCol = sheet.getLastColumn();
      const numCols = Math.min(lastCol, COLUMN_ORDER.length);
      const values = sheet.getRange(3, 1, lastRow - 2, numCols).getValues();
      const data = values.map(row => {
        const obj = {};
        for (let i = 0; i < numCols; i++) { obj[COLUMN_ORDER[i]] = row[i]; }
        return obj;
      }).filter(obj => obj.orgao);
      return jsonOut({success:true, data});
    } catch (err) {
      return jsonOut({success:false, error:err.message});
    }
  }

  return ContentService.createTextOutput('Curupira-Mirim II 2026 — AMAZONAS · online v1.').setMimeType(ContentService.MimeType.TEXT);
}
