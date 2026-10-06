import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import Ajv from 'ajv';
import { parse } from 'csv-parse/sync';
import ExcelJS from 'exceljs';

export const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const schema = JSON.parse(await fs.readFile(path.join(ROOT, 'schema/report.schema.json'), 'utf8'));
const validateShape = new Ajv({ allErrors: true, strict: false }).compile(schema);
const own = (obj, key) => Object.prototype.hasOwnProperty.call(obj, key);
export class ReportError extends Error {
  constructor(message, issues = []) { super(message); this.name = 'ReportError'; this.issues = issues; }
}
export const digest = bytes => createHash('sha256').update(bytes).digest('hex');
export const escapeHTML = value => String(value).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
export const textUnits = value => [...String(value)].reduce((n, c) => n + (/[^\u0000-\u00ff]/.test(c) ? 1 : 0.52), 0);

export function assertShape(report) {
  if (!validateShape(report)) throw new ReportError('원고 형식이 맞지 않습니다', validateShape.errors.map(e => `${e.instancePath || '/'} ${e.message}`));
}
function uniqueIds(items, kind) {
  const ids = items.map(v => v.id);
  if (new Set(ids).size !== ids.length) throw new ReportError(`${kind}의 id가 중복되었습니다`);
}
export async function safeInputPath(root, relative) {
  if (path.isAbsolute(relative)) throw new ReportError('데이터 경로는 원고 폴더에 대한 상대 경로여야 합니다');
  const base = await fs.realpath(root);
  const candidate = path.resolve(base, relative);
  const inside = target => { const rel = path.relative(base, target); return rel !== '' && rel !== '..' && !rel.startsWith(`..${path.sep}`) && !path.isAbsolute(rel); };
  if (!inside(candidate)) throw new ReportError('데이터 경로가 원고 폴더 밖을 가리킵니다');
  let real;
  try { real = await fs.realpath(candidate); } catch { throw new ReportError(`데이터 파일을 찾을 수 없습니다: ${relative}`); }
  if (!inside(real)) throw new ReportError('데이터 링크가 원고 폴더 밖을 가리킵니다');
  return real;
}
function checkHeaders(headers, name) {
  if (!headers.length || headers.some(h => !h) || new Set(headers).size !== headers.length) throw new ReportError(`${name}: 열 이름은 비어 있거나 중복될 수 없습니다`);
  if (headers.some(h => ['__proto__', 'constructor', 'prototype'].includes(h))) throw new ReportError(`${name}: 지원하지 않는 열 이름입니다`);
}
async function readDataset(spec, baseDir) {
  const filename = await safeInputPath(baseDir, spec.file);
  const stat = await fs.stat(filename);
  if (!stat.isFile() || stat.size > 20 * 1024 * 1024) throw new ReportError(`${spec.id}: 데이터 파일은 20MB 이하의 일반 파일이어야 합니다`);
  const bytes = await fs.readFile(filename);
  const extension = path.extname(filename).toLowerCase();
  let headers, rows;
  if (extension === '.csv') {
    rows = parse(bytes, { bom: true, skip_empty_lines: true, columns: raw => { headers = raw.map(h => h.trim()); checkHeaders(headers, spec.id); return headers; } });
  } else if (extension === '.xlsx') {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(bytes);
    const sheet = spec.sheet ? workbook.getWorksheet(spec.sheet) : workbook.worksheets[0];
    if (!sheet) throw new ReportError(`${spec.id}: 지정한 엑셀 시트를 찾을 수 없습니다`);
    headers = sheet.getRow(1).values.slice(1).map(v => String(v ?? '').trim());
    checkHeaders(headers, spec.id);
    rows = [];
    sheet.eachRow((row, number) => {
      if (number === 1) return;
      const record = {};
      headers.forEach((head, i) => {
        let value = row.getCell(i + 1).value;
        if (value && typeof value === 'object' && ('formula' in value || 'sharedFormula' in value)) {
          if (value.result === undefined || value.result === null) throw new ReportError(`${spec.id}: ${number}행 ${head}의 수식에 저장된 계산 결과가 없습니다. Excel에서 계산 후 저장해 주세요`);
          value = value.result;
        }
        if (value instanceof Date) value = value.toISOString().slice(0, 10);
        else if (value && typeof value === 'object' && value.richText) value = value.richText.map(r => r.text).join('');
        else if (value && typeof value === 'object' && value.text !== undefined) value = value.text;
        else if (value && typeof value === 'object') throw new ReportError(`${spec.id}: ${number}행 ${head}에 지원하지 않는 엑셀 셀 형식이 있습니다`);
        record[head] = value ?? '';
      });
      rows.push(record);
    });
  } else throw new ReportError(`${spec.id}: CSV 또는 XLSX 파일을 사용해 주세요`);
  if (!headers?.length || !rows.length || rows.length > 10000) throw new ReportError(`${spec.id}: 데이터는 헤더와 1~10,000개의 행을 포함해야 합니다`);
  return { ...spec, headers, rows, sha256: digest(bytes), filename };
}
export function numericValue(value, location) {
  if ((typeof value !== 'number' && typeof value !== 'string') || String(value).trim() === '') throw new ReportError(`${location}: 비어 있거나 숫자가 아닌 값입니다`);
  const result = Number(value);
  if (!Number.isFinite(result)) throw new ReportError(`${location}: 유한한 숫자가 필요합니다`);
  return result;
}
function requireColumn(dataset, column) {
  if (!dataset.headers.includes(column)) throw new ReportError(`${dataset.id}: ${column} 열을 찾을 수 없습니다`);
}
export function calculateMetric(metric, dataset) {
  requireColumn(dataset, metric.column);
  const values = metric.operation === 'count' ? [] : dataset.rows.map((r, i) => numericValue(r[metric.column], `${dataset.id} ${i + 2}행 ${metric.column}`));
  let value;
  switch (metric.operation) {
    case 'sum': value = values.reduce((a, b) => a + b, 0); break;
    case 'mean': value = values.reduce((a, b) => a + b, 0) / values.length; break;
    case 'first': value = values[0]; break;
    case 'last': value = values.at(-1); break;
    case 'count': value = dataset.rows.length; break;
    case 'growthPct':
      if (values.length < 2 || values[0] === 0) throw new ReportError(`${metric.id}: 변화율은 2개 이상의 값과 0이 아닌 첫 값이 필요합니다`);
      value = (values.at(-1) - values[0]) / values[0] * 100; break;
    default: throw new ReportError(`${metric.id}: 지원하지 않는 계산입니다`);
  }
  if (!Number.isFinite(value)) throw new ReportError(`${metric.id}: 계산 결과가 유한한 숫자가 아닙니다`);
  const decimals = metric.decimals ?? (metric.operation === 'growthPct' ? 1 : 0);
  const formattedNumber = new Intl.NumberFormat('ko-KR', { minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(value);
  const formatted = `${formattedNumber}${metric.unit === '%' ? '%' : metric.unit ? ` ${metric.unit}` : ''}`;
  return { ...metric, value, formatted, sourceIds: dataset.sourceIds };
}
function interpolate(value, metrics) {
  if (typeof value === 'string') return value.replace(/\{\{\s*([^{}]+?)\s*\}\}/g, (_, id) => {
    if (!own(metrics, id)) throw new ReportError(`원고에 알 수 없는 지표가 있습니다: {{${id}}}`);
    return metrics[id].formatted;
  });
  if (Array.isArray(value)) return value.map(v => interpolate(v, metrics));
  if (value && typeof value === 'object') return Object.fromEntries(Object.entries(value).map(([k, v]) => [k, interpolate(v, metrics)]));
  return value;
}
export async function loadTheme(name = 'corporate', customFile) {
  const filename = customFile ? path.resolve(customFile) : path.join(ROOT, 'themes', `${name}.json`);
  let theme;
  try { theme = JSON.parse(await fs.readFile(filename, 'utf8')); } catch { throw new ReportError('테마 파일을 읽을 수 없습니다'); }
  const colors = ['background', 'foreground', 'muted', 'accent', 'secondary', 'rule', 'tableHeader', 'tableStripe'];
  if (colors.some(k => !/^[a-fA-F0-9]{6}$/.test(theme[k])) || typeof theme.font !== 'string' || !theme.font || theme.font.length > 100 || typeof theme.name !== 'string') throw new ReportError('테마에 글꼴과 6자리 색상 코드가 필요합니다');
  return theme;
}
export async function resolveReport(input, baseDir, { themeFile } = {}) {
  assertShape(input);
  uniqueIds(input.sources, '출처'); uniqueIds(input.datasets, '데이터'); uniqueIds(input.metrics ?? [], '지표');
  const sources = Object.fromEntries(input.sources.map((s, i) => [s.id, { ...s, number: i + 1 }]));
  const checkSources = (ids, label) => ids.forEach(id => { if (!own(sources, id)) throw new ReportError(`${label}: 출처 ${id}를 찾을 수 없습니다`); });
  input.datasets.forEach(d => checkSources(d.sourceIds, d.id));
  const datasets = Object.create(null);
  for (const ds of input.datasets) datasets[ds.id] = await readDataset(ds, baseDir);
  const getDataset = id => { if (!own(datasets, id)) throw new ReportError(`데이터 ${id}를 찾을 수 없습니다`); return datasets[id]; };
  const metrics = Object.create(null);
  for (const m of input.metrics ?? []) metrics[m.id] = calculateMetric(m, getDataset(m.dataset));
  const doc = interpolate(structuredClone(input), metrics);
  if (textUnits(doc.title) > 58) throw new ReportError('표지 제목이 너무 깁니다. 제목과 부제를 나눠 주세요');
  if (doc.date && (!Number.isFinite(Date.parse(doc.date)) || new Date(doc.date).toISOString().slice(0, 10) !== doc.date)) throw new ReportError('날짜가 올바르지 않습니다');
  const warnings = [];
  doc.sections = doc.sections.map((s, i) => {
    const typeFields = { summary: [], text: [], conclusion: [], metrics: ['metricIds'], chart: ['dataset', 'chartType', 'labelColumn', 'series', 'unit'], table: ['dataset', 'columns'], comparison: ['options'], timeline: ['steps'] };
    const allowed = new Set(['type', 'title', 'lead', 'body', 'bullets', 'sourceIds', ...typeFields[s.type]]);
    const incompatible = Object.keys(s).filter(k => !allowed.has(k));
    if (incompatible.length) throw new ReportError(`블록 ${i + 1}: ${s.type} 유형에서 사용할 수 없는 항목입니다: ${incompatible.join(', ')}`);
    checkSources(s.sourceIds ?? [], `블록 ${i + 1}`);
    if (textUnits(s.title) > 47) throw new ReportError(`블록 ${i + 1}: 제목이 너무 깁니다. 47개 한글 글자 정도로 줄여 주세요`);
    if (textUnits(s.lead ?? '') > 95) throw new ReportError(`블록 ${i + 1}: 핵심 문장이 너무 깁니다. 상세 설명에 나눠 적어 주세요`);
    const ids = new Set(s.sourceIds ?? []);
    const section = { ...s };
    if (s.dataset) { const ds = getDataset(s.dataset); ds.sourceIds.forEach(id => ids.add(id)); }
    if (s.type === 'metrics') {
      section.resolvedMetrics = s.metricIds.map(id => { if (!own(metrics, id)) throw new ReportError(`지표 ${id}를 찾을 수 없습니다`); metrics[id].sourceIds.forEach(v => ids.add(v)); return metrics[id]; });
      for (const metric of section.resolvedMetrics) {
        if (textUnits(metric.label) > 24 || textUnits(metric.formatted) > 16 || textUnits(metric.description ?? '') > 50) throw new ReportError(`${metric.label}: 지표 이름이나 표시 값이 너무 깁니다. 단위를 조정하거나 표 블록을 사용해 주세요`);
      }
    }
    if (s.type === 'chart') {
      const ds = getDataset(s.dataset);
      if (ds.rows.length > 24) throw new ReportError(`${s.title}: 차트는 최대 24개의 범주를 지원합니다. 원본에서 집계해 주세요`);
      requireColumn(ds, s.labelColumn);
      const labels = ds.rows.map(r => String(r[s.labelColumn]));
      if (labels.some(l => !l.trim() || textUnits(l) > 16)) throw new ReportError(`${s.title}: 차트 범주 이름은 비어 있지 않은 짧은 문구여야 합니다`);
      section.chartData = s.series.map(series => {
        requireColumn(ds, series.column);
        return { name: series.name, labels, values: ds.rows.map((r, j) => numericValue(r[series.column], `${ds.id} ${j + 2}행 ${series.column}`)) };
      });
    }
    if (s.type === 'table') {
      const ds = getDataset(s.dataset); s.columns.forEach(c => requireColumn(ds, c));
      section.table = { headers: s.columns, rows: ds.rows.map(r => s.columns.map(c => String(r[c]))) };
    }
    if (s.type === 'comparison') section.table = { headers: ['선택지', '기대 효과', '검토할 점'], rows: s.options.map(o => [o.name, o.strength, o.tradeoff]) };
    if (s.type === 'timeline') section.table = { headers: ['시점', '실행 내용', '담당'], rows: s.steps.map(step => [step.when, step.title, step.owner ?? '']) };
    section.sourceIds = [...ids];
    if (!section.sourceIds.length) warnings.push(`${i + 1}. ${s.title}: 출처가 연결되어 있지 않습니다`);
    if (['summary', 'text', 'conclusion'].includes(s.type) && !(s.body?.length || s.bullets?.length || s.lead)) throw new ReportError(`${s.title}: 본문이나 핵심 문장을 입력해 주세요`);
    return section;
  });
  const theme = await loadTheme(doc.theme, themeFile);
  return { doc, datasets, metrics, sources, theme, warnings, baseDir: path.resolve(baseDir) };
}
export function sourceShort(model, ids) {
  const full = ids.map(id => { const s = model.sources[id]; return `[${s.number}] ${s.publisher || s.title}`; }).join(' / ');
  return textUnits(full) <= 90 ? full : `${ids.map(id => `[${model.sources[id].number}]`).join(' ')} 출처 상세는 참고 자료 참조`;
}
export function sourceFull(s) { return `[${s.number}] ${s.title}${s.publisher ? `, ${s.publisher}` : ''}${s.date ? `, ${s.date}` : ''}${s.locator ? `, ${s.locator}` : ''}${s.url ? `\n${s.url}` : ''}`; }

function flowPages(section) {
  const items = [...(section.body ?? []).map(text => ({ text, bullet: false })), ...(section.bullets ?? []).map(text => ({ text, bullet: true }))];
  const pages = []; let current = [], used = 0;
  for (const item of items) {
    const lines = Math.max(1, Math.ceil(textUnits(item.text) / 46));
    const h = lines * 0.34 + 0.22;
    if (h > 4.25) throw new ReportError(`${section.title}: 한 문단이 너무 깁니다. 여러 문단으로 나눠 주세요`);
    if (current.length && used + h > 4.25) { pages.push(current); current = []; used = 0; }
    current.push({ ...item, h }); used += h;
  }
  if (current.length || !pages.length) pages.push(current);
  return pages;
}
function tablePages(section) {
  const { headers, rows } = section.table;
  // Narrow number columns and wide narrative columns share a conservative width.
  const widths = headers.map(() => 12 / headers.length);
  if (section.type === 'comparison') widths.splice(0, widths.length, 2.8, 4.6, 4.6);
  if (section.type === 'timeline') widths.splice(0, widths.length, 2.1, 7.0, 2.9);
  const headerH = Math.max(0.5, ...headers.map((v, i) => Math.ceil(textUnits(v) / ((widths[i] - 0.3) * 72 / 16)) * 0.26 + 0.18));
  if (headerH > 1.4) throw new ReportError(`${section.title}: 표의 열 이름이 너무 깁니다. 짧은 이름을 사용해 주세요`);
  const pages = []; let current = [], heights = [headerH], used = headerH;
  for (const row of rows) {
    const h = Math.max(0.49, ...row.map((v, i) => Math.ceil(textUnits(v) / ((widths[i] - 0.3) * 72 / 16)) * 0.26 + 0.18));
    if (h > 3.6) throw new ReportError(`${section.title}: 표의 한 셀이 너무 깁니다. 열을 나누거나 설명을 본문으로 옮겨 주세요`);
    if (current.length && (used + h > 4.1 || current.length >= 8)) { pages.push({ rows: current, heights, widths }); current = []; heights = [headerH]; used = headerH; }
    current.push(row); heights.push(h); used += h;
  }
  if (current.length) pages.push({ rows: current, heights, widths });
  return pages;
}
export function planSlides(model) {
  const pages = [{ kind: 'cover' }]; let chartIndex = 0;
  model.doc.sections.forEach((section, sectionIndex) => {
    let blocks;
    if (section.table) blocks = tablePages(section).map(tablePage => ({ kind: 'table', tablePage }));
    else if (section.type === 'chart') blocks = [{ kind: 'chart', chartIndex: ++chartIndex }];
    else if (section.type === 'metrics') blocks = [{ kind: 'metrics' }];
    else blocks = flowPages(section).map(items => ({ kind: 'flow', items }));
    const evidenceType = ['chart', 'metrics', 'table', 'comparison', 'timeline'].includes(section.type);
    if (evidenceType && (section.body?.length || section.bullets?.length)) {
      const shortChartNote = section.type === 'chart' && section.body?.length === 1 && !section.bullets?.length && textUnits(section.body[0]) <= 90;
      if (shortChartNote) blocks[0].chartNote = section.body[0];
      else blocks.push(...flowPages(section).map(items => ({ kind: 'flow', items })));
    }
    blocks.forEach((p, i) => pages.push({ ...p, section, sectionIndex, part: i + 1, parts: blocks.length }));
  });
  let references = [], used = 0;
  for (const source of Object.values(model.sources)) {
    const h = Math.max(0.72, Math.ceil(textUnits(sourceFull({ ...source, url: undefined })) / 64) * 0.27 + (source.url ? 0.34 : 0.1) + 0.18);
    if (h > 4.3) throw new ReportError(`${source.title}: 발표 자료의 참고 자료 설명이 너무 깁니다`);
    if (references.length && used + h > 4.3) { pages.push({ kind: 'references', references }); references = []; used = 0; }
    references.push({ ...source, slideH: h }); used += h;
  }
  if (references.length) pages.push({ kind: 'references', references });
  return pages.map((p, i) => ({ ...p, number: i + 1 }));
}
