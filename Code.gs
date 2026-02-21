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
};

const ESCAL_HEADERS = [
  'Protocolo', 'Empresa', 'Mês', 'Ano', 'EmailAnalise', 'ValorTaxi', 'ValorRefeicao',
  'ValorDobra', 'ValorHoraFreelancer', 'DuracaoIntervalo', 'MaxSemIntervalo',
  'PayloadJSON', 'LinkPdfAnalise', 'LinkPdfColaboradores', 'LinkPdfCopia', 'CriadoEm', 'AtualizadoEm',
];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('WebApp de Escalas')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function getBootstrapData() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const now = new Date();
  const ano = Number(getNamedSingleValue_(ss, 'ANO_AQUI')) || now.getFullYear();

  const empresas = getNamedRangeValues_(ss, 'EMP_COMP').flat().map(String).map(s => s.trim()).filter(Boolean);
  const colabData = getNamedRangeValues_(ss, 'COLAB');
  const funcData = getNamedRangeValues_(ss, 'FUNC');

  const colaboradores = colabData
    .filter(r => String(r[0] || '').trim())
    .map(r => ({
      nome: String(r[0] || '').trim(),
      funcao: String(r[1] || '').trim(),
      empresa: String(r[2] || '').trim(),
      escala: String(r[6] || '').trim(),
      jornada: toHHMM_(r[7]),
      historicoDobras: String(r[8] || '').trim(),
    }));

  const funcoes = funcData
    .filter(r => String(r[0] || '').trim())
    .map(r => ({
      nome: String(r[0] || '').trim(),
      empresasTokens: String(r[1] || '').split(',').map(s => s.trim()).filter(Boolean),
    }));

  const feriados = readDateDescMapByNamedRangeColumns_(ss, 'FERIAD', 24, 25);
  const eventos = readDateDescMapByNamedRangeColumns_(ss, 'EVENT', 0, 1);
  const prevFat = readDateValueMapByNamedRangeColumns_(ss, 'PREV_FAT', 0, 2);

  const weather = getWeatherForTiradentes_(ano);

  return {
    empresas,
    colaboradores,
    funcoes,
    feriados,
    eventos,
    prevFat,
    weather,
    config: {
      specialStatuses: CONFIG.SPECIAL_STATUSES,
      maxRows: CONFIG.MAX_ROWS,
      defaultRows: CONFIG.DEFAULT_ROWS,
      timezone: CONFIG.TIMEZONE,
      yearDefault: ano,
    },
  };
}

function loadProtocols(protocolList) {
  const parsed = String(protocolList || '').split(',').map(s => s.trim()).filter(Boolean);
  if (!parsed.length) return { merged: null, warnings: ['Nenhum protocolo informado'] };

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  ensureEscalHeader_(ss);
  const escal = getNamedRangeValues_(ss, 'ESCAL');
  if (!escal.length) return { merged: null, warnings: ['Nenhum registro em ESCAL'] };

  const byProtocol = new Map();
  escal.forEach(row => {
    const p = String(row[0] || '').trim();
    if (p) byProtocol.set(p, row);
  });

  const rows = [];
  const warnings = [];
  parsed.forEach(p => {
    if (!byProtocol.has(p)) warnings.push(`Protocolo não encontrado: ${p}`);
    else rows.push(byProtocol.get(p));
  });

  if (!rows.length) return { merged: null, warnings };

  let mergedPayload = null;
  rows.forEach((r, idx) => {
    const payload = safeJsonParse_(r[11]);
    if (!payload) return;
    if (!mergedPayload) {
      mergedPayload = payload;
      return;
    }
    mergedPayload = mergePayloads_(mergedPayload, payload, idx);
  });

  return { merged: mergedPayload, warnings };
}

function saveAndGeneratePdfAnalise(payload) {
  validatePayload_(payload, { strictForSubmit: true });
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
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

    appendEscal_(ss, {
      protocol,
      payload,
      linkAnalise: pdfLink,
    });

    appendDobrasHistory_(ss, payload, analysis);

    return { ok: true, protocol, pdfLink, emailError: emailErr || null };
  } finally {
    lock.releaseLock();
  }
}

function generatePdfCopia(payload) {
  validatePayload_(payload, { strictForSubmit: false });
  const html = buildWeeklyHtml_(payload, {
    title: 'Salvar cópia',
    showPrevFat: false,
    hideEscalaJornadaAfterName: false,
    hideDobraText: false,
    transformForColab: false,
  });
  const link = saveHtmlPdf_(html, `copia_${payload.protocol || generateProtocol_()}.pdf`);
  return { ok: true, pdfLink: link };
}

function generatePdfColaboradores(payload) {
  validatePayload_(payload, { strictForSubmit: false });
  const html = buildWeeklyHtml_(payload, {
    title: 'Envio colaboradores',
    showPrevFat: false,
    hideEscalaJornadaAfterName: true,
    hideDobraText: true,
    transformForColab: true,
  });
  const link = saveHtmlPdf_(html, `colaboradores_${payload.protocol || generateProtocol_()}.pdf`);
  return { ok: true, pdfLink: link };
}

/* ========================= Helpers de dados ========================= */

function getNamedRangeValues_(ss, rangeName) {
  const r = ss.getRangeByName(rangeName);
  if (!r) throw new Error(`Intervalo nomeado não encontrado: ${rangeName}`);
  return r.getValues();
}

function getNamedSingleValue_(ss, rangeName) {
  const values = getNamedRangeValues_(ss, rangeName);
  return values && values[0] ? values[0][0] : '';
}

function updateAnoAqui_(ss, ano) {
  const r = ss.getRangeByName('ANO_AQUI');
  if (!r) throw new Error('ANO_AQUI não encontrado');
  r.setValue(Number(ano));
}

function readDateDescMapByNamedRangeColumns_(ss, namedRange, dateColZeroBasedFromA, descColZeroBasedFromA) {
  const range = ss.getRangeByName(namedRange);
  if (!range) return {};
  const sheet = range.getSheet();
  const values = sheet.getRange(1, 1, sheet.getLastRow(), Math.max(dateColZeroBasedFromA, descColZeroBasedFromA) + 1).getValues();
  const out = {};
  values.forEach(r => {
    const date = r[dateColZeroBasedFromA];
    const desc = String(r[descColZeroBasedFromA] || '').trim();
    if (!date || !desc) return;
    const key = dateKey_(date);
    if (!out[key]) out[key] = [];
    out[key].push(desc);
  });
  return out;
}

function readDateValueMapByNamedRangeColumns_(ss, namedRange, dateColZeroBasedFromA, valueColZeroBasedFromA) {
  const range = ss.getRangeByName(namedRange);
  if (!range) return {};
  const values = range.getValues();
  const out = {};
  values.forEach(r => {
    const date = r[dateColZeroBasedFromA];
    if (!date) return;
    out[dateKey_(date)] = r[valueColZeroBasedFromA];
  });
  return out;
}

function ensureEscalHeader_(ss) {
  const range = ss.getRangeByName('ESCAL');
  if (!range) throw new Error('Intervalo nomeado ESCAL não encontrado');
  const sheet = range.getSheet();
  const first = sheet.getRange(range.getRow(), range.getColumn(), 1, ESCAL_HEADERS.length).getValues()[0];
  const hasHeader = first.some(Boolean);
  if (!hasHeader) {
    sheet.getRange(range.getRow(), range.getColumn(), 1, ESCAL_HEADERS.length).setValues([ESCAL_HEADERS]);
  }
}

function appendEscal_(ss, params) {
  const range = ss.getRangeByName('ESCAL');
  const sheet = range.getSheet();
  const now = new Date();
  const row = [
    params.protocol,
    params.payload.empresa,
    Number(params.payload.mes),
    Number(params.payload.ano),
    params.payload.emailAnalise,
    Number(params.payload.valorTaxi),
    Number(params.payload.valorRefeicao),
    Number(params.payload.valorDobra),
    Number(params.payload.valorHoraFreelancer),
    params.payload.duracaoIntervalo,
    params.payload.maxSemIntervalo,
    JSON.stringify(params.payload),
    params.linkAnalise || '',
    '',
    '',
    now,
    now,
  ];
  sheet.appendRow(row);
}

function appendDobrasHistory_(ss, payload, analysis) {
  const colabRange = ss.getRangeByName('COLAB');
  const colabSheet = colabRange.getSheet();
  const values = colabRange.getValues();
  const updateCol = 9;
  const updates = [];

  const mapDobras = analysis.totalDobrasByPerson || {};
  values.forEach((r, idx) => {
    const nome = String(r[0] || '').trim();
    if (!nome || payload.rows.some(x => x.isFreelancer && x.nome === nome)) return;
    if (!(nome in mapDobras)) return;

    const oldVal = String(r[8] || '').trim();
    const block = `{${payload.protocol || ''},${payload.mes},${payload.ano},${toMoney_(mapDobras[nome])}}`;
    const newVal = oldVal ? `${oldVal}, ${block}` : block;
    if (newVal !== oldVal) updates.push({ row: colabRange.getRow() + idx, value: newVal });
  });

  updates.forEach(u => colabSheet.getRange(u.row, colabRange.getColumn() + updateCol - 1).setValue(u.value));
}

function mergePayloads_(a, b) {
  const out = JSON.parse(JSON.stringify(a));
  const byName = new Map();
  (out.rows || []).forEach(r => byName.set(r.nome, r));
  (b.rows || []).forEach(r => byName.set(r.nome, r));
  out.rows = Array.from(byName.values());
  if (b.header) out.header = b.header;
  out.protocolMergedFrom = [...(a.protocolMergedFrom || []), b.protocol || ''];
  return out;
}

/* ========================= Regras / validação ========================= */

function validatePayload_(payload, opts) {
  if (!payload) throw new Error('Payload vazio');
  const required = ['empresa', 'valorTaxi', 'valorRefeicao', 'valorDobra', 'emailAnalise', 'duracaoIntervalo', 'maxSemIntervalo', 'mes', 'ano', 'valorHoraFreelancer'];
  required.forEach(k => {
    if (payload[k] === null || payload[k] === undefined || payload[k] === '') throw new Error(`Campo obrigatório: ${k}`);
  });
  if (!/^\S+@\S+\.\S+$/.test(payload.emailAnalise)) throw new Error('Email inválido');
  if (Number(payload.mes) < 1 || Number(payload.mes) > 12) throw new Error('Mês inválido');
  if (!/^\d{4}$/.test(String(payload.ano))) throw new Error('Ano inválido');
  ['duracaoIntervalo', 'maxSemIntervalo'].forEach(k => {
    if (!isHHMM_(payload[k])) throw new Error(`Formato inválido para ${k}, use HH:MM`);
  });
  if (!Array.isArray(payload.rows) || payload.rows.length < 2) throw new Error('É necessário no mínimo 2 linhas');
  if (payload.rows.length > CONFIG.MAX_ROWS) throw new Error(`Máximo de ${CONFIG.MAX_ROWS} linhas`);

  const names = new Set();
  payload.rows.forEach(row => {
    if (!row.nome) throw new Error('Linha sem nome');
    if (names.has(row.nome)) throw new Error(`Colaborador repetido no protocolo: ${row.nome}`);
    names.add(row.nome);

    (row.days || []).forEach(day => {
      const e = day.entrada || '';
      const i = day.intervalo || '';
      const s = day.saida || '';
      const isSpecial = CONFIG.SPECIAL_STATUSES.includes(e);
      if (isSpecial) {
        if (i !== e || s !== e) throw new Error(`Dia especial inconsistente em ${row.nome}`);
        return;
      }
      if (![e, i, s].every(x => isHHMMAllowOver24_(x) || x === '')) throw new Error(`Horário inválido em ${row.nome}`);
      if (e && i && s) validateTimeOrder_(e, i, s, payload.duracaoIntervalo);
    });
  });

  if (opts && opts.strictForSubmit) {
    payload.rows.forEach(row => {
      if (row.isFreelancer) return;
      (row.days || []).forEach((d, idx) => {
        if (!d.entrada || !d.intervalo || !d.saida) {
          throw new Error(`Preencha todos os dias para não freelancers: ${row.nome} dia ${idx + 1}`);
        }
      });
    });
  }
}

function validateTimeOrder_(entrada, intervalo, saida, duracaoIntervalo) {
  const en = hhmmToMin_(entrada);
  const it = hhmmToMin_(intervalo);
  const ex = hhmmToMin_(saida);
  const dur = hhmmToMin_(duracaoIntervalo);
  if (!(en < it && it < ex)) throw new Error('Obrigatório Entrada < Intervalo < Saída');
  if (it + dur > ex) throw new Error('Intervalo + duração deve ser <= Saída');
}

/* ========================= Cálculos ========================= */

function computeAnalysis_(payload) {
  const dur = hhmmToMin_(payload.duracaoIntervalo);
  const daysInMonth = new Date(payload.ano, payload.mes, 0).getDate();
  const result = {
    totalTaxi: 0,
    totalRefeicao: 0,
    totalDobrasMes: 0,
    totalDobrasByPerson: {},
    freelancersMinutes: 0,
    freelancersValue: 0,
    bancoHorasByPerson: {},
    days: [],
  };

  for (let d = 1; d <= daysInMonth; d++) {
    let taxiCount = 0;
    payload.rows.forEach(row => {
      const cell = row.days[d - 1] || {};
      const e = cell.entrada;
      const s = cell.saida;
      if (isHHMMAllowOver24_(e) && isHHMMAllowOver24_(s)) {
        const en = hhmmToMin_(e);
        const ex = hhmmToMin_(s);
        if (ex > 23 * 60) taxiCount += 1;

        if (en < 15 * 60) result.totalRefeicao += 1;
        if (ex >= 18 * 60) result.totalRefeicao += 1;

        if (row.isFreelancer) {
          const it = hhmmToMin_(cell.intervalo || e);
          result.freelancersMinutes += Math.max(0, ex - en - Math.max(0, dur));
        } else {
          const jornada = hhmmToMin_(row.jornada || '00:00');
          let extra = ex - en - dur - jornada;
          if (cell.isDobra) extra = 0;
          if (CONFIG.SPECIAL_STATUSES.includes(e) && e !== 'Abatimento') extra = 0;
          if (e === 'Abatimento') extra = -jornada;
          result.bancoHorasByPerson[row.nome] = (result.bancoHorasByPerson[row.nome] || 0) + extra;
        }
      }

      if (cell.isDobra) {
        result.totalDobrasMes += Number(payload.valorDobra);
        result.totalDobrasByPerson[row.nome] = (result.totalDobrasByPerson[row.nome] || 0) + Number(payload.valorDobra);
      }
    });

    result.totalTaxi += Math.ceil(taxiCount / 4) * Number(payload.valorTaxi);
    result.days.push({ day: d, taxiCount });
  }

  result.freelancersValue = (result.freelancersMinutes / 60) * Number(payload.valorHoraFreelancer);
  return result;
}

/* ========================= PDFs ========================= */

function saveHtmlPdf_(html, filename) {
  const blob = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(filename);
  const folder = DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID);
  const file = folder.createFile(blob);
  return file.getUrl();
}

function sendEmailWithAttachment_(to, subject, body, html, filename) {
  try {
    const attachment = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(filename);
    GmailApp.sendEmail(to, subject, body, { attachments: [attachment] });
    return null;
  } catch (e) {
    return String(e && e.message || e);
  }
}

function buildAnaliseHtml_(payload, analysis, opts) {
  const daysInMonth = new Date(payload.ano, payload.mes, 0).getDate();
  let blocks = '';
  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(payload.ano, payload.mes - 1, d);
    const dateLabel = Utilities.formatDate(date, CONFIG.TIMEZONE, 'dd/MM/yyyy (EEE)');
    const obs = buildHeaderObs_(payload, d, opts.showPrevFat);
    blocks += `<section class="day-block"><h3>${dateLabel} ${obs}</h3>${buildDayGridHtml_(payload, d, opts.showDobraText)}</section>`;
  }

  const bankRows = Object.keys(analysis.bancoHorasByPerson).sort().map(n => `<tr><td>${n}</td><td>${minToHHMM_(analysis.bancoHorasByPerson[n])}</td></tr>`).join('');
  const dobraRows = Object.keys(analysis.totalDobrasByPerson).sort().map(n => `<tr><td>${n}</td><td>${toMoney_(analysis.totalDobrasByPerson[n])}</td></tr>`).join('');

  return `
  <html><head><meta charset="utf-8" /><style>${pdfCss_()}</style></head><body>
  <section class="page first">
    <h1>Análise de Escala - ${payload.empresa}</h1>
    ${buildHeaderFieldsHtml_(payload)}
    <h2>Resumo do mês</h2>
    <ul>
      <li>Taxi total: ${toMoney_(analysis.totalTaxi)}</li>
      <li>Refeição total: ${toMoney_(analysis.totalRefeicao * Number(payload.valorRefeicao))} (${analysis.totalRefeicao} refeições)</li>
      <li>Dobras total: ${toMoney_(analysis.totalDobrasMes)}</li>
      <li>Freelancers: ${minToHHMM_(analysis.freelancersMinutes)} (${toMoney_(analysis.freelancersValue)})</li>
    </ul>
    <h3>Dobras por pessoa</h3>
    <table><tr><th>Nome</th><th>Valor</th></tr>${dobraRows}</table>
    <h3>Banco de horas</h3>
    <table><tr><th>Nome</th><th>HH:MM</th></tr>${bankRows}</table>
  </section>
  ${blocks}
  </body></html>`;
}

function buildWeeklyHtml_(payload, options) {
  const daysInMonth = new Date(payload.ano, payload.mes, 0).getDate();
  let out = `<html><head><meta charset="utf-8" /><style>${pdfCss_()}</style></head><body><h1>${options.title} - ${payload.empresa}</h1>`;

  for (let start = 1; start <= daysInMonth; start += 7) {
    const end = Math.min(daysInMonth, start + 6);
    out += `<section class="week-block"><h2>Dias ${start} a ${end}</h2><table><tr><th>Nome</th>`;
    for (let d = start; d <= end; d++) {
      const obs = buildHeaderObs_(payload, d, options.showPrevFat);
      out += `<th>${String(d).padStart(2, '0')} ${obs}</th>`;
    }
    out += '</tr>';

    payload.rows.forEach(row => {
      const displayName = options.hideEscalaJornadaAfterName ? row.nome : `${row.nome} (${row.escala || '-'} / ${row.jornada || '-'})`;
      out += `<tr><td>${displayName}</td>`;
      for (let d = start; d <= end; d++) {
        let cell = row.days[d - 1] || {};
        if (options.transformForColab) cell = transformCellForColabPdf_(row, cell, payload);
        let text = [cell.entrada || '-', cell.intervalo || '-', cell.saida || '-'].join(' / ');
        if (cell.isDobra && !options.hideDobraText) text += ' (dobra)';
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
  if (c.exceptionEscala) {
    c.entrada = 'Folga';
    c.intervalo = 'Folga';
    c.saida = 'Folga';
  }
  if (c.exceptionExtra && isHHMMAllowOver24_(c.entrada) && isHHMM_(row.jornada || '')) {
    const calc = hhmmToMin_(c.entrada) + hhmmToMin_(row.jornada) + hhmmToMin_(payload.duracaoIntervalo);
    c.saida = minToHHMM_(calc);
  }
  if (c.exceptionMaxSemInt && isHHMMAllowOver24_(c.entrada) && isHHMM_(row.jornada || '')) {
    const mid = hhmmToMin_(c.entrada) + Math.round((hhmmToMin_(row.jornada) / 2) / 30) * 30;
    c.intervalo = minToHHMM_(mid);
  }
  c.isDobra = false;
  return c;
}

function buildHeaderFieldsHtml_(payload) {
  const f = [
    ['Empresa', payload.empresa], ['Mês/Ano', `${payload.mes}/${payload.ano}`], ['Email análise', payload.emailAnalise],
    ['Valor Taxi', toMoney_(payload.valorTaxi)], ['Valor Refeição', toMoney_(payload.valorRefeicao)],
    ['Valor Dobra', toMoney_(payload.valorDobra)], ['Valor Hora Freelancer', toMoney_(payload.valorHoraFreelancer)],
    ['Intervalo', payload.duracaoIntervalo], ['Max sem intervalo', payload.maxSemIntervalo], ['Protocolo', payload.protocol || '-'],
  ];
  return `<table>${f.map(x => `<tr><th>${x[0]}</th><td>${x[1]}</td></tr>`).join('')}</table>`;
}

function buildHeaderObs_(payload, day, showPrevFat) {
  const date = new Date(payload.ano, payload.mes - 1, day);
  const key = dateKey_(date);
  const f = (payload.feriados && payload.feriados[key]) || [];
  const e = (payload.eventos && payload.eventos[key]) || [];
  const x = [...f, ...e];
  const obs = x.length ? `(${x.join(', ')})` : '';
  const p = showPrevFat && payload.prevFat ? payload.prevFat[key] : '';
  return `${obs} ${p ? `[${p}]` : ''}`.trim();
}

function buildDayGridHtml_(payload, day, showDobraText) {
  const rows = payload.rows.map(r => {
    const c = r.days[day - 1] || {};
    let txt = [c.entrada || '-', c.intervalo || '-', c.saida || '-'].join(' / ');
    if (c.isDobra && showDobraText) txt += ' / dobra';
    return `<tr><td>${r.nome}</td><td>${txt}</td></tr>`;
  }).join('');
  return `<table><tr><th>Nome</th><th>Horários</th></tr>${rows}</table>`;
}

function pdfCss_() {
  return `
    body { font-family: Arial, sans-serif; font-size: 10px; }
    h1,h2,h3 { margin: 6px 0; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 8px; }
    th, td { border: 1px solid #bbb; padding: 4px; vertical-align: top; }
    .day-block, .week-block { page-break-inside: avoid; margin-bottom: 10px; }
    .first { page-break-after: always; }
  `;
}

/* ========================= Weather ========================= */

function getWeatherForTiradentes_(ano) {
  try {
    const cache = CacheService.getScriptCache();
    const cKey = `weather_${ano}`;
    const cached = cache.get(cKey);
    if (cached) return JSON.parse(cached);

    const geoUrl = 'https://geocoding-api.open-meteo.com/v1/search?name=Tiradentes&count=1&language=pt&format=json';
    const geo = JSON.parse(UrlFetchApp.fetch(geoUrl, { muteHttpExceptions: true }).getContentText() || '{}');
    const item = (geo.results || [])[0];
    if (!item) return {};

    const start = `${ano}-01-01`;
    const end = `${ano}-12-31`;
    const forecastUrl = `https://api.open-meteo.com/v1/forecast?latitude=${item.latitude}&longitude=${item.longitude}&daily=temperature_2m_min,temperature_2m_max,precipitation_probability_max&timezone=America%2FSao_Paulo&start_date=${start}&end_date=${end}`;
    const data = JSON.parse(UrlFetchApp.fetch(forecastUrl, { muteHttpExceptions: true }).getContentText() || '{}');
    const daily = data.daily || {};
    const out = {};
    (daily.time || []).forEach((t, idx) => {
      out[t] = {
        min: daily.temperature_2m_min ? daily.temperature_2m_min[idx] : null,
        max: daily.temperature_2m_max ? daily.temperature_2m_max[idx] : null,
        rain: daily.precipitation_probability_max ? daily.precipitation_probability_max[idx] : null,
      };
    });
    cache.put(cKey, JSON.stringify(out), 21600);
    return out;
  } catch (e) {
    return {};
  }
}

/* ========================= Util ========================= */

function toHHMM_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIMEZONE, 'HH:mm');
  const s = String(v || '').trim();
  if (!s) return '';
  if (/^\d{1,2}:\d{2}$/.test(s)) return s;
  return s;
}

function isHHMM_(s) {
  return /^\d{1,2}:\d{2}$/.test(String(s || '')) && hhmmToMin_(s) >= 0;
}

function isHHMMAllowOver24_(s) {
  return /^\d{1,3}:\d{2}$/.test(String(s || '')) && hhmmToMin_(s) >= 0;
}

function hhmmToMin_(s) {
  const m = String(s || '').match(/^(\d{1,3}):(\d{2})$/);
  if (!m) return -1;
  const h = Number(m[1]);
  const mm = Number(m[2]);
  if (mm < 0 || mm > 59) return -1;
  return h * 60 + mm;
}

function minToHHMM_(m) {
  const sign = m < 0 ? '-' : '';
  const abs = Math.abs(m);
  const h = Math.floor(abs / 60);
  const mm = abs % 60;
  return `${sign}${String(h).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
}

function toMoney_(n) {
  return Number(n || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function dateKey_(d) {
  return Utilities.formatDate(new Date(d), CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function safeJsonParse_(s) {
  try { return JSON.parse(s); } catch (e) { return null; }
}

function generateProtocol_() {
  const stamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss');
  const rand = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
  return `PRT-${stamp}-${rand}`;
}
