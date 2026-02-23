/**
 * Web App de escalas - Google Apps Script
 */
const CONFIG = {
  SPREADSHEET_ID: '1mc3nNSeW6GI2rXudQ30c2bzIlDtccheEdsTG85n_Y4g',
  TIMEZONE: 'America/Sao_Paulo',
  PDF_FOLDER_ID: '1UzyIn1fsiVIatfgQQK-GyGIeJI4Z-AFs',
  SPECIAL_STATUSES: ['Folga', 'Abatimento', 'Dom. mês', 'Licença', 'Férias'],
  MAX_ROWS: 100,
  DEFAULT_ROWS: 2,
  WEEKDAYS_PT: ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'],
};

const ESCAL_HEADERS = [
  'Protocolo', 'Empresa', 'Mês', 'Ano', 'EmailAnalise', 'ValorTaxi', 'ValorRefeicao',
  'ValorDobra', 'ValorHoraFreelancer', 'DuracaoIntervalo', 'MaxSemIntervalo',
  'PayloadJSON', 'LinkPdfAnalise', 'LinkPdfColaboradores', 'LinkPdfCopia', 'CriadoEm', 'AtualizadoEm',
];

function getBestLock_() {
  try { const l = LockService.getScriptLock(); if (l) return l; } catch (e) {}
  try { const l = LockService.getDocumentLock(); if (l) return l; } catch (e) {}
  return null;
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('WebApp de Escalas')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getBootstrapData() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const now = new Date();
  const ano = Number(getNamedSingleValue_(ss, 'ANO_AQUI')) || now.getFullYear();
  const mes = now.getMonth() + 1;

  const empRange = ss.getRangeByName('EMP_COMP');
  if (!empRange) throw new Error('Intervalo nomeado não encontrado: EMP_COMP');
  const empresas = uniqueKeepOrder_(empRange.getValues().map(r => String(r[0] || '').trim()).filter(Boolean));
  const colabData = getNamedRangeValues_(ss, 'COLAB');
  const funcData = getNamedRangeValues_(ss, 'FUNC');

  const colaboradores = colabData.filter(r => String(r[0] || '').trim()).map(r => ({
    nome: String(r[0] || '').trim(),
    funcao: String(r[1] || '').trim(),
    empresa: String(r[2] || '').trim(),
    escala: String(r[6] || '').trim(),
    jornada: toHHMM_(r[7]),
    historicoDobras: String(r[8] || '').trim(),
  }));

  const funcoes = funcData.filter(r => String(r[0] || '').trim()).map(r => ({
    nome: String(r[0] || '').trim(),
    empresasTokens: String(r[1] || '').split(',').map(s => s.trim()).filter(Boolean),
  }));

  return {
    empresas,
    colaboradores,
    funcoes,
    feriados: readDateDescMapByNamedRangeColumns_(ss, 'FERIAD', 24, 25),
    eventos: readDateDescMapByNamedRangeColumns_(ss, 'EVENT', 0, 1),
    prevFat: readDateValueMapByNamedRangeColumns_(ss, 'PREV_FAT', 0, 2),
    weather: getWeatherForTiradentes_(ano, mes),
    config: {
      specialStatuses: CONFIG.SPECIAL_STATUSES,
      maxRows: CONFIG.MAX_ROWS,
      defaultRows: CONFIG.DEFAULT_ROWS,
      timezone: CONFIG.TIMEZONE,
      yearDefault: ano,
      monthDefault: mes,
      weekdayNames: CONFIG.WEEKDAYS_PT,
    },
  };
}

function loadProtocols(protocolList) {
  const parsed = String(protocolList || '').split(',').map(s => s.trim()).filter(Boolean);
  if (!parsed.length) return { merged: null, warnings: ['Nenhum protocolo informado'] };

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  ensureEscalHeader_(ss);
  const escal = getNamedRangeValues_(ss, 'ESCAL');

  const byProtocol = new Map();
  escal.forEach(row => {
    const p = String(row[0] || '').trim();
    if (p && p !== ESCAL_HEADERS[0]) byProtocol.set(p, row);
  });

  const warnings = [];
  const rows = parsed.map(p => byProtocol.get(p) || null).filter((x, idx) => {
    if (!x) warnings.push(`Protocolo não encontrado: ${parsed[idx]}`);
    return !!x;
  });
  if (!rows.length) return { merged: null, warnings };

  let mergedPayload = null;
  rows.forEach(r => {
    const payload = safeJsonParse_(r[11]);
    if (!payload) return;
    mergedPayload = mergedPayload ? mergePayloads_(mergedPayload, payload) : payload;
  });
  return { merged: mergedPayload, warnings };
}

function saveAndGeneratePdfAnalise(payload) {
  validatePayload_(payload, { strictForSubmit: true });
  const lock = getBestLock_();
  if (lock) lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    ensureEscalHeader_(ss);
    updateAnoAqui_(ss, payload.ano);

    const protocol = generateProtocol_();
    payload.protocol = protocol;

    const analysis = computeAnalysis_(payload);
    const html = buildAnaliseHtml_(payload, analysis, { showPrevFat: true, showDobraText: true });
    const pdfLink = saveHtmlPdf_(html, `analise_${protocol}.pdf`);
    const emailErr = sendEmailWithAttachment_(payload.emailAnalise, `Análise de Escala ${protocol}`,
      `Segue anexo PDF de análise.\n\nProtocolo: ${protocol}\nEmpresa: ${payload.empresa}`,
      html, `analise_${protocol}.pdf`);

    appendEscal_(ss, { protocol, payload, linkAnalise: pdfLink });
    appendDobrasHistory_(ss, payload, analysis);

    return { ok: true, protocol, pdfLink, emailError: emailErr || null };
  } finally {
    if (lock) lock.releaseLock();
  }
}

function generatePdfCopia(payload) {
  validatePayload_(payload, { strictForSubmit: false });
  const html = buildWeeklyHtml_(payload, { title: 'Salvar cópia', showPrevFat: false, hideEscalaJornadaAfterName: false, hideDobraText: false, transformForColab: false });
  return { ok: true, pdfLink: saveHtmlPdf_(html, `copia_${payload.protocol || generateProtocol_()}.pdf`) };
}

function generatePdfColaboradores(payload) {
  validatePayload_(payload, { strictForSubmit: false });
  const html = buildWeeklyHtml_(payload, { title: 'Envio colaboradores', showPrevFat: false, hideEscalaJornadaAfterName: true, hideDobraText: true, transformForColab: true });
  return { ok: true, pdfLink: saveHtmlPdf_(html, `colaboradores_${payload.protocol || generateProtocol_()}.pdf`) };
}

function getNamedRangeValues_(ss, rangeName) {
  const r = ss.getRangeByName(rangeName);
  if (!r) throw new Error(`Intervalo nomeado não encontrado: ${rangeName}`);
  return r.getValues();
}
function getNamedSingleValue_(ss, rangeName) { const v = getNamedRangeValues_(ss, rangeName); return (v[0] || [])[0] || ''; }
function updateAnoAqui_(ss, ano) { const r = ss.getRangeByName('ANO_AQUI'); if (!r) throw new Error('ANO_AQUI não encontrado'); r.setValue(Number(ano)); }

function setAnoAqui(ano) {
  const lock = getBestLock_();
  if (lock) lock.waitLock(5000);
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    updateAnoAqui_(ss, Number(ano));
    return { ok: true };
  } finally {
    if (lock) lock.releaseLock();
  }
}

function readDateDescMapByNamedRangeColumns_(ss, namedRange, dateOffset, descOffset) {
  const range = ss.getRangeByName(namedRange);
  if (!range) return {};

  let values = range.getValues();
  const maxOffset = Math.max(dateOffset, descOffset);
  if (range.getNumColumns() <= maxOffset) {
    const sheet = range.getSheet();
    values = sheet.getRange(range.getRow(), 1, range.getNumRows(), maxOffset + 1).getValues();
  }

  const out = {};
  values.forEach(r => {
    const date = r[dateOffset];
    const desc = String(r[descOffset] || '').trim();
    if (!date || !desc) return;
    const key = dateKey_(date);
    if (!out[key]) out[key] = [];
    out[key].push(desc);
  });
  return out;
}

function readDateValueMapByNamedRangeColumns_(ss, namedRange, dateOffset, valueOffset) {
  const range = ss.getRangeByName(namedRange);
  if (!range) return {};
  const values = range.getValues();
  const out = {};
  values.forEach(r => {
    const date = r[dateOffset];
    if (!date) return;
    out[dateKey_(date)] = r[valueOffset];
  });
  return out;
}

function ensureEscalHeader_(ss) {
  const range = ss.getRangeByName('ESCAL');
  if (!range) throw new Error('Intervalo nomeado ESCAL não encontrado');
  const headRange = range.offset(0, 0, 1, ESCAL_HEADERS.length);
  const first = headRange.getValues()[0];
  if (!first.some(Boolean)) headRange.setValues([ESCAL_HEADERS]);
}

function appendEscal_(ss, params) {
  const range = ss.getRangeByName('ESCAL');
  const values = range.getValues();
  let target = -1;
  for (let i = 1; i < values.length; i++) {
    if (!String(values[i][0] || '').trim()) { target = i; break; }
  }
  if (target < 0) throw new Error('ESCAL sem linha vazia disponível no intervalo nomeado.');

  const now = new Date();
  const row = [
    params.protocol, params.payload.empresa, Number(params.payload.mes), Number(params.payload.ano), params.payload.emailAnalise,
    Number(params.payload.valorTaxi), Number(params.payload.valorRefeicao), Number(params.payload.valorDobra), Number(params.payload.valorHoraFreelancer),
    params.payload.duracaoIntervalo, params.payload.maxSemIntervalo, JSON.stringify(params.payload), params.linkAnalise || '', '', '', now, now,
  ];
  range.offset(target, 0, 1, ESCAL_HEADERS.length).setValues([row]);
}

function appendDobrasHistory_(ss, payload, analysis) {
  const lock = getBestLock_();
  if (lock) lock.waitLock(5000);
  try {
  const colabRange = ss.getRangeByName('COLAB');
  const sheet = colabRange.getSheet();
  const values = colabRange.getValues();
  const realNames = new Set((payload.rows || []).filter(r => !r.isFreelancer).map(r => r.nome));
  const mapDobras = analysis.totalDobrasByPerson || {};

  const updates = [];
  values.forEach((r, idx) => {
    const nome = String(r[0] || '').trim();
    if (!nome || !realNames.has(nome) || !(nome in mapDobras)) return;
    const oldVal = String(r[8] || '').trim();
    const block = `{${payload.protocol || ''}, ${payload.mes}, ${payload.ano}, ${toMoney_(mapDobras[nome])}}`;
    const newVal = oldVal ? `${oldVal}, ${block}` : block;
    if (newVal !== oldVal) updates.push({ idx, newVal });
  });

  updates.forEach(u => sheet.getRange(colabRange.getRow() + u.idx, colabRange.getColumn() + 8).setValue(u.newVal));
  } finally {
    if (lock) lock.releaseLock();
  }
}

function mergePayloads_(a, b) {
  const out = JSON.parse(JSON.stringify(a));
  const byName = new Map();
  (out.rows || []).forEach(r => byName.set(r.nome, r));
  (b.rows || []).forEach(r => byName.set(r.nome, r));
  out.rows = Array.from(byName.values());
  return out;
}

function validatePayload_(payload, opts) {
  if (!payload) throw new Error('Payload vazio');
  const required = ['empresa', 'valorTaxi', 'valorRefeicao', 'valorDobra', 'emailAnalise', 'duracaoIntervalo', 'maxSemIntervalo', 'mes', 'ano', 'valorHoraFreelancer'];
  required.forEach(k => { if (payload[k] === null || payload[k] === undefined || payload[k] === '') throw new Error(`Campo obrigatório: ${k}`); });
  if (!/^\S+@\S+\.\S+$/.test(payload.emailAnalise)) throw new Error('Email inválido');
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const empRange = ss.getRangeByName('EMP_COMP');
  const empresasValidas = empRange ? uniqueKeepOrder_(empRange.getValues().map(r => String(r[0] || '').trim()).filter(Boolean)) : [];
  if (!empresasValidas.includes(String(payload.empresa || '').trim())) throw new Error('Empresa inválida para EMP_COMP');

  const dur = hhmmToMin_(payload.duracaoIntervalo);
  const maxSem = hhmmToMin_(payload.maxSemIntervalo);
  if (dur < 0 || maxSem < 0) throw new Error('Campos de duração devem estar em HH:MM');

  const names = new Set();
  (payload.rows || []).forEach(row => {
    if (!row.nome) throw new Error('Linha sem nome');
    if (names.has(row.nome)) throw new Error(`Colaborador repetido no protocolo: ${row.nome}`);
    names.add(row.nome);

    (row.days || []).forEach((day, idx) => {
      const n = normalizeDayCell_(day, row);
      day.entrada = n.entrada; day.intervalo = n.intervalo; day.saida = n.saida;

      const isSpecialDay = n.isSpecialDay;
      if (!isSpecialDay) {
        if ([n.entrada, n.intervalo, n.saida].some(v => v && !isHHMMAllowOver24_(v))) throw new Error(`Horário inválido em ${row.nome} dia ${idx + 1}`);
        if (n.entrada && n.intervalo && n.saida) validateTimeOrder_(n.entrada, n.intervalo, n.saida, payload.duracaoIntervalo);
      }

      if (row.isFreelancer) return;
      if (!n.entrada || !n.intervalo || !n.saida) return;
      if (isSpecialDay) return;

      const en = hhmmToMin_(n.entrada);
      const it = hhmmToMin_(n.intervalo);
      const ex = hhmmToMin_(n.saida);
      const jornada = hhmmToMin_(row.jornada || '00:00');

      if (!day.exceptionMaxSemInt && !day.isDobra) {
        if ((it - en) > maxSem || (ex - (it + dur)) > maxSem) throw new Error('Excedeu máximo sem intervalo (campo 6). Use exceção “M” se necessário.');
      }
      if (!day.exceptionExtra && !day.isDobra) {
        const total = ex - en - dur;
        if (total > jornada + 120) throw new Error('Excedeu Jornada + 02:00. Use exceção “+2” ou “Dobra”.');
      }
    });

    if (opts && opts.strictForSubmit && !row.isFreelancer) {
      (row.days || []).forEach((d, idx) => {
        if (!d.entrada || !d.intervalo || !d.saida) throw new Error(`Preencha todos os dias para não freelancers: ${row.nome} dia ${idx + 1}`);
      });
    }
  });
}

function normalizeDayCell_(day, row) {
  const entrada = normalizeStatusOrTime_(day.entrada || '');
  let intervalo = normalizeStatusOrTime_(day.intervalo || '');
  let saida = normalizeStatusOrTime_(day.saida || '');

  if (isHHMMAllowOver24_(entrada) && isHHMMAllowOver24_(intervalo) && hhmmToMin_(intervalo) < hhmmToMin_(entrada)) intervalo = minToHHMM_(hhmmToMin_(intervalo) + 1440);
  if (isHHMMAllowOver24_(entrada) && isHHMMAllowOver24_(saida) && hhmmToMin_(saida) < hhmmToMin_(entrada)) saida = minToHHMM_(hhmmToMin_(saida) + 1440);

  const isSpecialDay = CONFIG.SPECIAL_STATUSES.includes(entrada) && entrada === intervalo && entrada === saida;
  if (CONFIG.SPECIAL_STATUSES.includes(entrada) && (!intervalo || !saida)) {
    intervalo = entrada;
    saida = entrada;
  }
  return { entrada, intervalo, saida, isSpecialDay };
}

function normalizeStatusOrTime_(raw) {
  const s = String(raw || '').trim();
  if (!s) return '';
  if (isHHMMAllowOver24_(s)) return s;
  const k = s.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[()]/g, '').replace(/\./g, '').replace(/\s+/g, ' ').trim();
  const map = {
    'folga': 'Folga',
    'abatimento': 'Abatimento',
    'dom mes': 'Dom. mês',
    'domingo do mes': 'Dom. mês',
    'domingo mes': 'Dom. mês',
    'dom meso': 'Dom. mês',
    'licenca': 'Licença',
    'ferias': 'Férias',
  };
  return map[k] || s;
}

function validateTimeOrder_(entrada, intervalo, saida, duracaoIntervalo) {
  const en = hhmmToMin_(entrada);
  const it = hhmmToMin_(intervalo);
  const ex = hhmmToMin_(saida);
  const dur = hhmmToMin_(duracaoIntervalo);
  if (!(en < it && it < ex)) throw new Error('Entrada precisa ser menor que Intervalo e Saída');
  if (it + dur > ex) throw new Error('Intervalo + duração excede a Saída');
}

function computeAnalysis_(payload) {
  const dur = hhmmToMin_(payload.duracaoIntervalo);
  const daysInMonth = new Date(payload.ano, payload.mes, 0).getDate();
  const out = {
    totalTaxi: 0,
    totalRefeicao: 0,
    totalDobrasMes: 0,
    totalDobrasByPerson: {},
    freelancersMinutes: 0,
    freelancersValue: 0,
    freelancersMinutesByPerson: {},
    freelancersValueByPerson: {},
    bancoHorasByPerson: {},
    bancoHorasTotal: 0,
    totalHorasFreelancers: '00:00',
  };

  for (let d = 1; d <= daysInMonth; d++) {
    let taxiCount = 0;
    (payload.rows || []).forEach(row => {
      const raw = row.days[d - 1] || {};
      const c = normalizeDayCell_(raw, row);
      const e = c.entrada;
      const sOut = c.saida;

      if (isHHMMAllowOver24_(e) && isHHMMAllowOver24_(sOut)) {
        const en = hhmmToMin_(e);
        const ex = hhmmToMin_(sOut);
        if (ex > 1380) taxiCount += 1;
        if (en < 900) out.totalRefeicao += 1;
        if (ex >= 1080) out.totalRefeicao += 1;

        if (row.isFreelancer) {
          const mins = Math.max(0, ex - en - dur);
          out.freelancersMinutes += mins;
          out.freelancersMinutesByPerson[row.nome] = (out.freelancersMinutesByPerson[row.nome] || 0) + mins;
        } else {
          const jornada = hhmmToMin_(row.jornada || '00:00');
          let extra = ex - en - dur - jornada;
          if (raw.isDobra || (CONFIG.SPECIAL_STATUSES.includes(e) && e !== 'Abatimento')) extra = 0;
          if (e === 'Abatimento') extra = -jornada;
          out.bancoHorasByPerson[row.nome] = (out.bancoHorasByPerson[row.nome] || 0) + extra;
        }
      }

      if (raw.isDobra) {
        out.totalDobrasMes += Number(payload.valorDobra);
        if (!row.isFreelancer) out.totalDobrasByPerson[row.nome] = (out.totalDobrasByPerson[row.nome] || 0) + Number(payload.valorDobra);
      }
    });

    out.totalTaxi += Math.ceil(taxiCount / 4) * Number(payload.valorTaxi);
  }

  out.freelancersValue = (out.freelancersMinutes / 60) * Number(payload.valorHoraFreelancer);
  Object.keys(out.freelancersMinutesByPerson).forEach(n => {
    out.freelancersValueByPerson[n] = (out.freelancersMinutesByPerson[n] / 60) * Number(payload.valorHoraFreelancer);
  });
  out.bancoHorasTotal = Object.values(out.bancoHorasByPerson).reduce((a, b) => a + b, 0);
  out.totalHorasFreelancers = minToHHMM_(out.freelancersMinutes);
  return out;
}

function saveHtmlPdf_(html, filename) {
  const blob = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(filename);
  return DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID).createFile(blob).getUrl();
}
function sendEmailWithAttachment_(to, subject, body, html, filename) {
  try {
    const attachment = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(filename);
    GmailApp.sendEmail(to, subject, body, { attachments: [attachment] });
    return null;
  } catch (e) { return String(e && e.message || e); }
}

function buildAnaliseHtml_(payload, analysis, opts) {
  const daysInMonth = new Date(payload.ano, payload.mes, 0).getDate();
  let blocks = '';
  for (let d = 1; d <= daysInMonth; d++) {
    blocks += `<section class="day-block"><h3>${formatDateFullPt_(payload.ano, payload.mes, d)} ${buildHeaderObs_(payload, d, opts.showPrevFat)}</h3>${buildDayGridHtml_(payload, d, opts.showDobraText)}</section>`;
  }

  const dobraRows = Object.entries(analysis.totalDobrasByPerson || {}).filter(([,v]) => v > 0);
  const bancoRows = Object.entries(analysis.bancoHorasByPerson || {});
  const freeRows = Object.entries(analysis.freelancersMinutesByPerson || {});

  const dobraTable = dobraRows.length
    ? dobraRows.map(([n,v]) => `<tr><td>${n}</td><td>${toMoney_(v)}</td></tr>`).join('')
    : '<tr><td colspan="2">Sem dados</td></tr>';

  const bancoTable = bancoRows.length
    ? bancoRows.map(([n,v]) => `<tr><td>${n}</td><td>${minToHHMM_(v)}</td></tr>`).join('')
    : '<tr><td colspan="2">Sem dados</td></tr>';

  const freeTable = freeRows.length
    ? freeRows.map(([n,m]) => `<tr><td>${n}</td><td>${minToHHMM_(m)}</td><td>${toMoney_(analysis.freelancersValueByPerson[n] || 0)}</td></tr>`).join('')
    : '<tr><td colspan="3">Sem dados</td></tr>';

  return `<html><head><meta charset="utf-8"/><style>${pdfCss_()}</style></head><body>
  <section class="first">
    <h1>Análise de Escala - ${payload.empresa}</h1>
    ${buildHeaderFieldsHtml_(payload, analysis)}
    <h3>Dobras por colaborador</h3>
    <table><tr><th>Nome</th><th>Valor</th></tr>${dobraTable}</table>
    <h3>Banco de horas (não freelancers)</h3>
    <table><tr><th>Nome</th><th>HH:MM</th></tr>${bancoTable}</table>
    <h3>Horas por freelancer</h3>
    <table><tr><th>Nome freelancer</th><th>HH:MM</th><th>R$</th></tr>${freeTable}</table>
  </section>${blocks}</body></html>`;
}


function buildWeeklyHtml_(payload, options) {
  const daysInMonth = new Date(payload.ano, payload.mes, 0).getDate();
  let out = `<html><head><meta charset="utf-8"/><style>${pdfCss_()}</style></head><body><h1>${options.title} - ${payload.empresa}</h1>`;
  for (let start = 1; start <= daysInMonth; start += 7) {
    const end = Math.min(daysInMonth, start + 6);
    out += `<section class="week-block"><h2>Dias ${start} a ${end}</h2><table><tr><th>Nome</th>`;
    for (let d = start; d <= end; d++) out += `<th>${formatDateFullPt_(payload.ano, payload.mes, d)} ${buildHeaderObs_(payload, d, options.showPrevFat)}</th>`;
    out += '</tr>';
    (payload.rows || []).forEach(row => {
      const displayName = options.hideEscalaJornadaAfterName ? row.nome : `${row.nome} (${row.escala || '-'} / ${row.jornada || '-'})`;
      out += `<tr><td>${displayName}</td>`;
      for (let d = start; d <= end; d++) {
        let c = row.days[d - 1] || {};
        if (options.transformForColab) c = transformCellForColabPdf_(row, c, payload);
        let text = [c.entrada || '-', c.intervalo || '-', c.saida || '-'].join(' / ');
        if (c.isDobra && !options.hideDobraText) text += ' (dobra)';
        out += `<td>${text}</td>`;
      }
      out += '</tr>';
    });
    out += '</table></section>';
  }
  return out + '</body></html>';
}

function transformCellForColabPdf_(row, cell, payload) {
  const c = JSON.parse(JSON.stringify(cell || {}));
  const norm = normalizeDayCell_(c, row);
  c.entrada = norm.entrada;
  c.intervalo = norm.intervalo;
  c.saida = norm.saida;

  if (row.isFreelancer) {
    c.isDobra = false;
    return c;
  }

  if (c.exceptionEscala) c.entrada = c.intervalo = c.saida = 'Folga';

  if (c.exceptionExtra && isHHMMAllowOver24_(c.entrada) && isHHMM_(row.jornada || '')) {
    const calc = hhmmToMin_(c.entrada) + hhmmToMin_(row.jornada) + hhmmToMin_(payload.duracaoIntervalo);
    c.saida = minToHHMM_(calc);
  }

  if (c.exceptionMaxSemInt && isHHMMAllowOver24_(c.entrada) && isHHMM_(row.jornada || '')) {
    const mid = hhmmToMin_(c.entrada) + Math.round((hhmmToMin_(row.jornada) / 2) / 30) * 30;
    c.intervalo = minToHHMM_(mid);
  }

  if (isHHMMAllowOver24_(c.entrada) && isHHMMAllowOver24_(c.saida) && hhmmToMin_(c.saida) < hhmmToMin_(c.entrada)) {
    c.saida = minToHHMM_(hhmmToMin_(c.saida) + 1440);
  }

  c.isDobra = false;
  return c;
}

function buildHeaderFieldsHtml_(payload, analysis) {
  const rows = [
    ['Empresa', payload.empresa, 'Taxi total', toMoney_(analysis.totalTaxi)],
    ['Mês/Ano', `${payload.mes}/${payload.ano}`, 'Refeição total', `${toMoney_(analysis.totalRefeicao * Number(payload.valorRefeicao))} (${analysis.totalRefeicao} refeições)`],
    ['Email análise', payload.emailAnalise, 'Dobras total', toMoney_(analysis.totalDobrasMes)],
    ['Intervalo', payload.duracaoIntervalo, 'Freelancers total', `${analysis.totalHorasFreelancers} (${toMoney_(analysis.freelancersValue)})`],
    ['Max sem intervalo', payload.maxSemIntervalo, 'Banco de horas total', minToHHMM_(analysis.bancoHorasTotal)],
    ['Protocolo', payload.protocol || '-', '', ''],
  ];
  return `<table>${rows.map(r => `<tr><th>${r[0]}</th><td>${r[1]}</td><th>${r[2]}</th><td>${r[3]}</td></tr>`).join('')}</table>`;
}

function buildHeaderObs_(payload, day, showPrevFat) {
  const k = dateKey_(new Date(payload.ano, payload.mes - 1, day));
  const obs = [ ...((payload.feriados && payload.feriados[k]) || []), ...((payload.eventos && payload.eventos[k]) || []) ];
  const p = showPrevFat && payload.prevFat ? payload.prevFat[k] : '';
  return `${obs.length ? `(${obs.join(', ')})` : ''} ${p ? `[${p}]` : ''}`.trim();
}
function buildDayGridHtml_(payload, day, showDobraText) {
  const dur = hhmmToMin_(payload.duracaoIntervalo);
  const rows = payload.rows || [];

  let minStart = Infinity;
  let maxEnd = -Infinity;
  rows.forEach(r => {
    const c = normalizeDayCell_(r.days[day - 1] || {}, r);
    if (isHHMMAllowOver24_(c.entrada) && isHHMMAllowOver24_(c.saida)) {
      minStart = Math.min(minStart, hhmmToMin_(c.entrada));
      maxEnd = Math.max(maxEnd, hhmmToMin_(c.saida));
    }
  });

  if (!isFinite(minStart) || !isFinite(maxEnd) || minStart >= maxEnd) {
    minStart = 8 * 60;
    maxEnd = 18 * 60;
  }

  minStart = Math.floor(minStart / 30) * 30;
  maxEnd = Math.ceil(maxEnd / 30) * 30;

  const slots = [];
  for (let t = minStart; t < maxEnd; t += 30) slots.push(t);

  const head = slots.map(t => `<th class="slot-head">${minToHHMM_(t)}</th>`).join('');
  let html = `<table class="grid30"><tr><th>Nome</th>${head}<th>Extra/Total</th></tr>`;

  rows.forEach(r => {
    const raw = r.days[day - 1] || {};
    const c = normalizeDayCell_(raw, r);
    const isSpecial = c.isSpecialDay;
    const isFreela = !!r.isFreelancer;
    const en = isHHMMAllowOver24_(c.entrada) ? hhmmToMin_(c.entrada) : NaN;
    const it = isHHMMAllowOver24_(c.intervalo) ? hhmmToMin_(c.intervalo) : NaN;
    const ex = isHHMMAllowOver24_(c.saida) ? hhmmToMin_(c.saida) : NaN;
    const jornada = hhmmToMin_(r.jornada || '00:00');

    let extraText = '-';
    if (!isNaN(en) && !isNaN(ex)) {
      if (isFreela) {
        extraText = minToHHMM_(Math.max(0, ex - en - dur));
      } else {
        let extra = ex - en - dur - jornada;
        if (raw.isDobra) extra = 0;
        if (['Folga', 'Férias', 'Licença', 'Dom. mês'].includes(c.entrada)) extra = 0;
        if (c.entrada === 'Abatimento') extra = -jornada;
        extraText = minToHHMM_(extra);
      }
    }

    html += `<tr><td>${r.nome}</td>`;
    slots.forEach(t => {
      let txt = "";
      let bg = "#ffffff";
      if (isSpecial) {
        bg = "#ffe0e0";
        txt = c.entrada;
      } else if (!isNaN(en) && !isNaN(ex) && t >= en && t < ex) {
        bg = isFreela ? "#fff3bf" : "#dff2d8";
        if (!isNaN(it) && t >= it && t < (it + dur)) bg = "#dbeafe";
        if (raw.isDobra && showDobraText) txt = "dobra";
      }
      html += `<td class="slot" style="background-color:${bg}">${txt}</td>`;
    });
    html += `<td>${extraText}</td></tr>`;
  });

  html += '</table>';
  return html;
}
function pdfCss_() {
  return `*{-webkit-print-color-adjust:exact;print-color-adjust:exact}body{font-family:Arial,sans-serif;font-size:10px}table{width:100%;border-collapse:collapse}th,td{border:1px solid #bbb;padding:4px}.first{page-break-after:always}.day-block{page-break-before:always;break-inside:avoid;page-break-inside:avoid}.day-block:first-of-type{page-break-before:auto}.week-block{break-inside:avoid;page-break-inside:avoid}.week-block + .week-block{page-break-before:always}.grid30 th,.grid30 td{font-size:8px;padding:2px}.grid30 .slot-head{background:#f6e9d6}.grid30 .slot{min-width:20px;width:20px;text-align:center}.grid30 .st-work{background:#dff2d8}.grid30 .st-break{background:#dbeafe}.grid30 .st-freela{background:#fff3bf}.grid30 .st-status{background:#ffe0e0}`;
}

function getWeatherForMonth(ano, mes) {
  return getWeatherForTiradentes_(Number(ano), Number(mes));
}

function fetchJsonWithLog_(url) {
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const text = resp.getContentText() || "";
    try { return JSON.parse(text || "{}"); }
    catch (e) { Logger.log("Weather JSON inválido: " + url + " / trecho: " + text.substring(0, 300)); return {}; }
  } catch (e) {
    Logger.log("Weather fetch falhou: " + url + " / erro: " + (e && e.message ? e.message : e));
    return {};
  }
}

function getWeatherForTiradentes_(ano, mes) {
  try {
    const key = `weather_${ano}_${mes}`;
    const cache = CacheService.getScriptCache();
    const cached = cache.get(key);
    if (cached) return JSON.parse(cached);

    const geo = fetchJsonWithLog_('https://geocoding-api.open-meteo.com/v1/search?name=Tiradentes&count=1&language=pt&format=json');
    const item = (geo.results || [])[0];
    if (!item) return {};

    const tz = CONFIG.TIMEZONE;
    const today = new Date(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd') + 'T00:00:00');
    const start = new Date(Number(ano), Number(mes) - 1, 1);
    const end = new Date(Number(ano), Number(mes), 0);

    const out = {};
    const fmt = d => Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    const addDaily = daily => {
      (daily.time || []).forEach((t, idx) => {
        out[t] = {
          min: daily.temperature_2m_min ? daily.temperature_2m_min[idx] : '—',
          max: daily.temperature_2m_max ? daily.temperature_2m_max[idx] : '—',
          rain: daily.precipitation_probability_max ? daily.precipitation_probability_max[idx] : '—',
        };
      });
    };

    const fetchArchive = (sDate, eDate) => {
      if (sDate > eDate) return;
      const url = `https://archive-api.open-meteo.com/v1/archive?latitude=${item.latitude}&longitude=${item.longitude}&daily=temperature_2m_min,temperature_2m_max,precipitation_probability_max&timezone=America%2FSao_Paulo&start_date=${fmt(sDate)}&end_date=${fmt(eDate)}`;
      const data = fetchJsonWithLog_(url);
      addDaily(data.daily || {});
    };
    const fetchForecast = (sDate, eDate) => {
      if (sDate > eDate) return;
      const url = `https://api.open-meteo.com/v1/forecast?latitude=${item.latitude}&longitude=${item.longitude}&daily=temperature_2m_min,temperature_2m_max,precipitation_probability_max&timezone=America%2FSao_Paulo&start_date=${fmt(sDate)}&end_date=${fmt(eDate)}`;
      const data = fetchJsonWithLog_(url);
      addDaily(data.daily || {});
    };

    const horizon = new Date(today);
    horizon.setDate(horizon.getDate() + 16);

    if (end < today) {
      fetchArchive(start, end);
    } else if (start <= today && end >= today) {
      const yesterday = new Date(today); yesterday.setDate(yesterday.getDate() - 1);
      fetchArchive(start, yesterday);
      const futureEnd = end < horizon ? end : horizon;
      fetchForecast(today, futureEnd);
    } else if (start > horizon) {
      cache.put(key, JSON.stringify({}), 21600);
      return {};
    } else {
      const futureEnd = end < horizon ? end : horizon;
      fetchForecast(start, futureEnd);
    }

    cache.put(key, JSON.stringify(out), 21600);
    return out;
  } catch (e) {
    return {};
  }
}

function formatDateFullPt_(ano, mes, dia) {
  const d = new Date(Number(ano), Number(mes) - 1, Number(dia));
  return `${String(dia).padStart(2, '0')}/${String(mes).padStart(2, '0')}/${ano} (${CONFIG.WEEKDAYS_PT[d.getDay()]})`;
}

function toHHMM_(v) { if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIMEZONE, 'HH:mm'); return String(v || '').trim(); }
function isHHMM_(s) { return /^\d{1,2}:\d{2}$/.test(String(s || '')) && hhmmToMin_(s) >= 0; }
function isHHMMAllowOver24_(s) { return /^\d{1,3}:\d{2}$/.test(String(s || '')) && hhmmToMin_(s) >= 0; }
function hhmmToMin_(s) { const m = String(s || '').match(/^(-?\d{1,3}):(\d{2})$/); if (!m) return -1; const h = Number(m[1]), mm = Number(m[2]); if (mm < 0 || mm > 59) return -1; return h * 60 + mm; }
function minToHHMM_(m) { const sign = m < 0 ? '-' : ''; const abs = Math.abs(m); return `${sign}${String(Math.floor(abs / 60)).padStart(2, '0')}:${String(abs % 60).padStart(2, '0')}`; }
function toMoney_(n) { return Number(n || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }); }
function dateKey_(d) { return Utilities.formatDate(new Date(d), CONFIG.TIMEZONE, 'yyyy-MM-dd'); }
function safeJsonParse_(s) { try { return JSON.parse(s); } catch (e) { return null; } }
function uniqueKeepOrder_(arr) { const seen = new Set(); return arr.filter(v => (seen.has(v) ? false : (seen.add(v), true))); }
function generateProtocol_() { return `PRT-${Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss')}-${Math.floor(Math.random() * 10000).toString().padStart(4, '0')}`; }
