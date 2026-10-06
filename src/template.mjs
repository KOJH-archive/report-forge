import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { ReportError, safeInputPath, digest, textUnits, sourceFull, sourceShort, escapeHTML } from './core.mjs';

const NS = { p: 'http://schemas.openxmlformats.org/presentationml/2006/main', a: 'http://schemas.openxmlformats.org/drawingml/2006/main', c: 'http://schemas.openxmlformats.org/drawingml/2006/chart', r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', rel: 'http://schemas.openxmlformats.org/package/2006/relationships', ct: 'http://schemas.openxmlformats.org/package/2006/content-types' };
const EMU = 914400;
export const TEMPLATE_KINDS = ['cover', 'flow', 'metrics', 'chart', 'table', 'references'];
export const TEXT_ROLES = ['clear', 'keep', 'title', 'lead', 'body', 'footer', 'page', 'meta', 'disclosure', 'unit', 'note', ...Array.from({ length: 3 }, (_, i) => ['label', 'value', 'description'].map(k => `metric${i + 1}.${k}`)).flat()];
const children = (node, ns, local) => Array.from(node?.childNodes ?? []).filter(n => n.nodeType === 1 && (!ns || n.namespaceURI === ns) && (!local || n.localName === local));
const all = (node, ns, local) => Array.from(node?.getElementsByTagNameNS(ns, local) ?? []);
const first = (node, ns, local) => all(node, ns, local)[0];
const serialize = node => new XMLSerializer().serializeToString(node);
function xml(source) {
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw new ReportError('PPTX에 지원하지 않는 XML 선언이 있습니다');
  try { return new DOMParser({ onError: level => { if (level !== 'warning') throw new Error('invalid XML'); } }).parseFromString(source, 'application/xml'); }
  catch { throw new ReportError('PPTX 내부 XML을 읽을 수 없습니다'); }
}
function el(doc, prefix, local, attrs = {}) { const n = doc.createElementNS(NS[prefix], `${prefix === 'rel' || prefix === 'ct' ? '' : prefix + ':'}${local}`); for (const [k, v] of Object.entries(attrs)) n.setAttribute(k, String(v)); return n; }
function relFile(part) { return path.posix.join(path.posix.dirname(part), '_rels', path.posix.basename(part) + '.rels'); }
function targetPart(part, target) {
  const p = path.posix.normalize(target.startsWith('/') ? target.slice(1) : path.posix.join(path.posix.dirname(part), target));
  if (p.startsWith('../') || p.startsWith('/') || p.includes('\\')) throw new ReportError('PPTX 내부 연결 경로가 올바르지 않습니다');
  return p;
}
async function readXML(zip, name) { if (!zip.file(name)) throw new ReportError(`PPTX 내부 파일이 없습니다: ${name}`); return xml(await zip.file(name).async('string')); }
async function relationships(zip, part) { return zip.file(relFile(part)) ? children((await readXML(zip, relFile(part))).documentElement, NS.rel, 'Relationship').map(n => ({ id: n.getAttribute('Id'), type: n.getAttribute('Type'), target: n.getAttribute('Target'), external: n.getAttribute('TargetMode') === 'External' })) : []; }
function relXML(rels) { return xml(`<Relationships xmlns="${NS.rel}">${rels.map(r => `<Relationship Id="${escapeHTML(r.id)}" Type="${escapeHTML(r.type)}" Target="${escapeHTML(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`); }
async function loadPackage(bytes) {
  if (bytes.length > 40 * 1024 * 1024) throw new ReportError('회사 PPTX는 40MB 이하여야 합니다');
  let zip; try { zip = await JSZip.loadAsync(bytes); } catch { throw new ReportError('올바른 PPTX 파일을 선택해 주세요'); }
  const entries = Object.values(zip.files); let unpacked = 0;
  if (entries.length > 12000) throw new ReportError('PPTX 내부 파일이 너무 많습니다');
  for (const e of entries) { unpacked += e._data?.uncompressedSize ?? 0; if (unpacked > 240 * 1024 * 1024) throw new ReportError('PPTX 압축 해제 크기가 너무 큽니다'); }
  if (!zip.file('ppt/presentation.xml') || !zip.file('[Content_Types].xml')) throw new ReportError('PowerPoint PPTX 파일이 필요합니다');
  const types = await zip.file('[Content_Types].xml').async('string');
  if (/macroEnabled|vbaProject/i.test(types)) throw new ReportError('매크로 없는 PPTX 파일로 저장해 주세요');
  return zip;
}
function geometry(node) {
  const t = first(node, NS.a, 'xfrm') ?? first(node, NS.p, 'xfrm');
  const off = first(t, NS.a, 'off'), ext = first(t, NS.a, 'ext');
  return { x: Number(off?.getAttribute('x') ?? 0) / EMU, y: Number(off?.getAttribute('y') ?? 0) / EMU, w: Number(ext?.getAttribute('cx') ?? 0) / EMU, h: Number(ext?.getAttribute('cy') ?? 0) / EMU };
}
function shapeInfo(node) {
  const id = first(node, NS.p, 'cNvPr');
  const texts = all(node, NS.a, 't').map(n => n.textContent);
  const ph = first(node, NS.p, 'ph'), props = first(node, NS.a, 'rPr') ?? first(node, NS.a, 'defRPr');
  const table = first(node, NS.a, 'tbl'), chart = first(node, NS.c, 'chart');
  const unsupported = /\/diagram$/.test(first(node, NS.a, 'graphicData')?.getAttribute('uri') ?? '') || Boolean(first(node, NS.p, 'oleObj'));
  const type = unsupported ? 'unsupported' : table ? 'table' : chart ? 'chart' : node.localName === 'pic' ? 'image' : first(node, NS.p, 'txBody') ? 'text' : 'shape';
  const grouped = node.parentNode?.localName === 'grpSp';
  return { id: id?.getAttribute('id') ?? '', name: id?.getAttribute('name') ?? '', type, text: texts.join('\n'), placeholder: ph?.getAttribute('type') || (ph ? 'body' : ''), placeholderIndex: ph?.getAttribute('idx') || '0', ...geometry(node), grouped,
    fontSize: Number(props?.getAttribute('sz') || 1800) / 100,
    font: first(node, NS.a, 'latin')?.getAttribute('typeface') ?? '',
    color: first(props, NS.a, 'srgbClr')?.getAttribute('val') ?? '',
    columns: table ? children(first(table, NS.a, 'tblGrid'), NS.a, 'gridCol').length : undefined,
    rows: table ? children(table, NS.a, 'tr').length : undefined,
    tableLayout: table ? { widths: children(first(table, NS.a, 'tblGrid'), NS.a, 'gridCol').map(n => Number(n.getAttribute('w')) / EMU),
      rows: children(table, NS.a, 'tr').map(n => ({ h: Number(n.getAttribute('h')) / EMU, fonts: children(n, NS.a, 'tc').map(c => Number((first(c, NS.a, 'rPr') ?? first(c, NS.a, 'defRPr'))?.getAttribute('sz') || 1800) / 100) })) } : undefined,
    relationship: chart?.getAttributeNS(NS.r, 'id') ?? undefined };
}
function shapeNodes(doc) { return all(doc, NS.p, 'spTree').flatMap(tree => Array.from(tree.getElementsByTagName('*')).filter(n => n.namespaceURI === NS.p && ['sp', 'pic', 'graphicFrame'].includes(n.localName))); }
function inferRoles(slide, height, repeated) {
  const text = slide.shapes.filter(s => s.type === 'text' && !s.grouped);
  const roles = Object.fromEntries(text.map(s => [s.id, 'clear']));
  for (const s of text) {
    const named = s.name.match(/^rf:(.+)$/i)?.[1];
    if (TEXT_ROLES.includes(named)) roles[s.id] = named;
    else if (['title', 'ctrTitle'].includes(s.placeholder)) roles[s.id] = 'title';
    else if (s.placeholder === 'sldNum') roles[s.id] = 'page';
    else if (s.placeholder === 'ftr') roles[s.id] = 'footer';
    else if (s.placeholder === 'dt') roles[s.id] = 'meta';
    else if (s.placeholder === 'subTitle') roles[s.id] = 'lead';
    else if (s.placeholder === 'body') roles[s.id] = slide.index === 1 ? 'lead' : 'body';
    else if (/^\d+$/.test(s.text.trim()) && s.y > height * .84) roles[s.id] = 'page';
    else if (s.y > height * .84 && /출처|참고|source|reference/i.test(s.text)) roles[s.id] = 'footer';
    else if (repeated.has(`${s.text}|${s.x.toFixed(2)}|${s.y.toFixed(2)}`) && (s.y < height * .15 || s.y > height * .84)) roles[s.id] = 'keep';
  }
  if (!Object.values(roles).includes('title')) {
    const title = text.filter(s => roles[s.id] === 'clear' && s.y < height * .45).sort((a, b) => b.fontSize - a.fontSize || a.y - b.y)[0];
    if (title) roles[title.id] = 'title';
  }
  const visual = slide.shapes.find(s => ['chart', 'table'].includes(s.type));
  if (visual) {
    const options = text.filter(s => roles[s.id] === 'clear');
    const unit = options.find(s => /단위\s*[:：]|\bunit\s*:/i.test(s.text)); if (unit) roles[unit.id] = 'unit';
    const note = options.filter(s => roles[s.id] === 'clear' && s.y >= visual.y + visual.h - .08 && s.y < height * .85).sort((a, b) => b.w * b.h - a.w * a.h)[0];
    if (note && visual.type === 'chart') roles[note.id] = 'note';
    const lead = options.filter(s => roles[s.id] === 'clear' && s.y > height * .15 && s.y < visual.y).sort((a, b) => b.w - a.w || a.y - b.y)[0]; if (lead) roles[lead.id] = 'lead';
  }
  const numericCards = text.filter(s => roles[s.id] === 'clear' && s.y > height * .3 && s.y < height * .8 && s.fontSize >= 22 && /^[-+]?\d[\d,.%\s가-힣A-Za-z]*$/.test(s.text.trim())).sort((a, b) => a.x - b.x);
  if (numericCards.length >= 2 && numericCards.length <= 3) {
    numericCards.forEach((value, i) => {
      roles[value.id] = `metric${i + 1}.value`;
      const nearby = text.filter(s => roles[s.id] === 'clear' && Math.abs(s.x - value.x) < .35 && s.w <= value.w * 1.5);
      const label = nearby.filter(s => s.y < value.y && value.y - s.y < 1.5).sort((a, b) => b.y - a.y)[0];
      const description = nearby.filter(s => s.y > value.y && s.y - value.y < 1.5).sort((a, b) => a.y - b.y)[0];
      if (label) roles[label.id] = `metric${i + 1}.label`; if (description) roles[description.id] = `metric${i + 1}.description`;
    });
  }
  if (slide.index === 1) {
    const lower = text.filter(s => roles[s.id] === 'clear' && s.y > height * .65 && s.fontSize <= 16).sort((a, b) => a.y - b.y);
    if (lower[0]) roles[lower[0].id] = 'meta'; if (lower[1]) roles[lower[1].id] = 'disclosure';
  }
  const available = text.filter(s => roles[s.id] === 'clear' && s.w > 1 && s.y > height * .13 && s.y < height * .85);
  const body = !visual && !Object.values(roles).includes('metric1.value') && !Object.values(roles).includes('body') && !Object.values(roles).includes('lead') ? available.sort((a, b) => b.w * b.h - a.w * a.h)[0] : undefined;
  if (body) roles[body.id] = slide.index === 1 ? 'lead' : 'body';
  const lead = available.filter(s => roles[s.id] === 'clear' && s.y < (body?.y ?? height * .6)).sort((a, b) => a.y - b.y)[0];
  if (lead && !Object.values(roles).includes('lead')) roles[lead.id] = 'lead';
  return roles;
}
export async function inspectTemplate(bytes) {
  const zip = await loadPackage(bytes), presentation = await readXML(zip, 'ppt/presentation.xml');
  const size = first(presentation, NS.p, 'sldSz');
  const width = Number(size?.getAttribute('cx')) / EMU, height = Number(size?.getAttribute('cy')) / EMU;
  if (!(width > 0 && height > 0)) throw new ReportError('PPTX 슬라이드 크기를 읽을 수 없습니다');
  const rels = await relationships(zip, 'ppt/presentation.xml'), relMap = new Map(rels.map(r => [r.id, r]));
  const ids = all(presentation, NS.p, 'sldId');
  if (!ids.length || ids.length > 200) throw new ReportError('템플릿에는 1~200장의 슬라이드가 필요합니다');
  const slides = [], warnings = [];
  for (const [i, id] of ids.entries()) {
    const r = relMap.get(id.getAttributeNS(NS.r, 'id'));
    if (!r || r.external) throw new ReportError('슬라이드 연결이 올바르지 않습니다');
    const part = targetPart('ppt/presentation.xml', r.target), doc = await readXML(zip, part);
    const shapes = shapeNodes(doc).map(shapeInfo), slideRels = await relationships(zip, part);
    const layoutRel = slideRels.find(r => r.type.endsWith('/slideLayout'));
    if (layoutRel && !layoutRel.external) {
      const layoutPart = targetPart(part, layoutRel.target), layout = await readXML(zip, layoutPart);
      // PowerPoint often stores placeholder coordinates in the layout rather than the slide.
      const layoutShapes = shapeNodes(layout).map(shapeInfo);
      const masterRel = (await relationships(zip, layoutPart)).find(r => r.type.endsWith('/slideMaster'));
      const masterShapes = masterRel && !masterRel.external ? shapeNodes(await readXML(zip, targetPart(layoutPart, masterRel.target))).map(shapeInfo) : [];
      for (const s of shapes.filter(s => s.placeholder && !s.w)) {
        const inherited = layoutShapes.find(l => l.w && l.placeholder === s.placeholder && l.placeholderIndex === s.placeholderIndex) ?? masterShapes.find(l => l.w && l.placeholder === s.placeholder);
        if (inherited) for (const k of ['x', 'y', 'w', 'h', 'fontSize', 'font', 'color']) s[k] = inherited[k];
      }
    }
    for (const s of shapes.filter(s => s.type === 'chart')) {
      const cr = slideRels.find(r => r.id === s.relationship);
      if (!cr || cr.external) { s.chartType = 'unsupported'; continue; }
      const chart = await readXML(zip, targetPart(part, cr.target));
      const plot = first(chart, NS.c, 'plotArea');
      const chartTypes = children(plot, NS.c).filter(n => /Chart$/.test(n.localName));
      s.chartType = chartTypes.length === 1 && ['lineChart', 'barChart'].includes(chartTypes[0].localName) ? chartTypes[0].localName.replace('Chart', '') : 'unsupported';
    }
    if (shapes.some(s => s.grouped && s.type === 'text')) warnings.push(`${i + 1}번: 그룹 내부 텍스트는 디자인 요소로 보존됩니다. 원래 업무 내용이면 PowerPoint에서 그룹을 풀어 주세요`);
    if (shapes.some(s => s.type === 'unsupported')) warnings.push(`${i + 1}번: SmartArt/OLE가 포함되어 있어 생성용 원본으로 선택할 수 없습니다`);
    slides.push({ index: i + 1, part, shapes });
  }
  const counts = new Map();
  for (const slide of slides) for (const key of new Set(slide.shapes.filter(s => s.type === 'text' && s.text).map(s => `${s.text}|${s.x.toFixed(2)}|${s.y.toFixed(2)}`))) counts.set(key, (counts.get(key) ?? 0) + 1);
  const repeated = new Set([...counts].filter(([, n]) => n >= 2).map(([key]) => key));
  for (const slide of slides) slide.roles = inferRoles(slide, height, repeated);
  const bodySlides = slides.filter(s => Object.values(s.roles).includes('body') && !s.shapes.some(n => ['chart', 'table', 'unsupported'].includes(n.type)));
  const layouts = { cover: 1, flow: bodySlides[0]?.index ?? 1, references: bodySlides[0]?.index ?? 1, metrics: slides.find(s => Object.values(s.roles).includes('metric1.value'))?.index ?? bodySlides[0]?.index ?? 1,
    chart: slides.find(s => s.shapes.filter(n => n.type === 'chart').length === 1 && s.shapes.find(n => n.type === 'chart').chartType !== 'unsupported')?.index ?? null,
    table: slides.find(s => s.shapes.filter(n => n.type === 'table').length === 1)?.index ?? null };
  return { version: 1, sha256: digest(bytes), width, height, slides, layouts, warnings, roles: TEXT_ROLES };
}
export async function importTemplate(source, outDir) {
  if (path.extname(source).toLowerCase() !== '.pptx') throw new ReportError('.pptx 파일을 선택해 주세요');
  if ((await fs.stat(source)).size > 40 * 1024 * 1024) throw new ReportError('회사 PPTX는 40MB 이하여야 합니다');
  const bytes = await fs.readFile(source), inspection = await inspectTemplate(bytes), out = path.resolve(outDir);
  try { await fs.lstat(out); throw new ReportError('새 템플릿 폴더를 지정해 주세요'); } catch (e) { if (e.code !== 'ENOENT') throw e; }
  await fs.mkdir(out, { recursive: true });
  await fs.writeFile(path.join(out, 'source.pptx'), bytes);
  await fs.writeFile(path.join(out, 'design.json'), JSON.stringify(inspection, null, 2));
  return inspection;
}
export async function loadTemplate(config, baseDir) {
  if (!config) return null;
  const filename = await safeInputPath(baseDir, config.file);
  if (path.extname(filename).toLowerCase() !== '.pptx') throw new ReportError('템플릿 파일은 .pptx여야 합니다');
  if ((await fs.stat(filename)).size > 40 * 1024 * 1024) throw new ReportError('회사 PPTX는 40MB 이하여야 합니다');
  const bytes = await fs.readFile(filename), inspection = await inspectTemplate(bytes);
  const layouts = { ...inspection.layouts, ...config.layouts };
  for (const [kind, index] of Object.entries(layouts)) {
    if (index !== null && !inspection.slides[index - 1]) throw new ReportError(`${kind}: 존재하지 않는 템플릿 슬라이드입니다`);
  }
  for (const [index, overrides] of Object.entries(config.roles ?? {})) {
    const slide = inspection.slides[Number(index) - 1];
    if (!slide) throw new ReportError('텍스트 역할의 슬라이드 번호가 잘못되었습니다');
    for (const [id, role] of Object.entries(overrides)) {
      if (!slide.shapes.some(s => s.id === id && s.type === 'text' && !s.grouped) || !TEXT_ROLES.includes(role)) throw new ReportError(`${index}번 슬라이드 ${id}: 지원하지 않는 텍스트 역할입니다`);
      slide.roles[id] = role;
    }
  }
  for (const slide of inspection.slides) for (const role of TEXT_ROLES.filter(r => !['clear', 'keep'].includes(r))) {
    if (Object.values(slide.roles).filter(r => r === role).length > 1) throw new ReportError(`${slide.index}번: ${role} 역할은 하나의 텍스트 상자에만 지정해 주세요`);
  }
  return { bytes, filename, inspection, layouts, config };
}

function fitCapacity(shape) {
  const size = shape.fontSize || 18;
  const columns = Math.max(1, (shape.w * 72 - 16) / size);
  const lines = Math.max(1, Math.floor((shape.h * 72 - 8) / (size * 1.25)));
  return columns * lines;
}
function textCost(shape, value) {
  const columns = Math.max(1, (shape.w * 72 - 16) / (shape.fontSize || 18));
  return String(value).split('\n').reduce((sum, line) => sum + Math.max(1, Math.ceil(textUnits(line) / columns)) * columns, 0);
}
function content(page, model) {
  if (page.kind === 'cover') return { title: model.doc.title, lead: model.doc.subtitle ?? '', body: '', meta: [model.doc.author, model.doc.date].filter(Boolean).join('  '), disclosure: model.doc.disclosure ?? '' };
  const s = page.section;
  const values = { title: page.kind === 'references' ? '참고 자료' : `${s.title}${page.parts > 1 ? ` (${page.part}/${page.parts})` : ''}`, lead: s?.lead ?? '', footer: s ? sourceShort(model, s.sourceIds) : '', page: String(page.number), body: '' };
  if (page.kind === 'flow') values.body = page.items.map(item => `${item.bullet ? '• ' : ''}${item.text}`).join('\n');
  if (page.kind === 'references') values.body = page.references.map(sourceFull).join('\n\n');
  if (page.kind === 'chart') { values.unit = `단위: ${s.unit}`; values.note = page.chartNote ?? ''; }
  if (page.kind === 'metrics') {
    values.body = s.resolvedMetrics.map(m => `${m.label}  ${m.formatted}${m.description ? '\n' + m.description : ''}`).join('\n\n');
    s.resolvedMetrics.forEach((m, i) => { for (const k of ['label', 'value', 'description']) values[`metric${i + 1}.${k}`] = k === 'value' ? m.formatted : m[k] ?? ''; });
  }
  return values;
}
function chosenSlide(template, page) {
  let index = template.layouts[page.kind];
  if (page.kind === 'table' && !Object.hasOwn(template.config.layouts ?? {}, 'table')) {
    index = template.inspection.slides.find(s => s.shapes.filter(n => n.type === 'table').length === 1 && s.shapes.find(n => n.type === 'table').columns === page.section.table.headers.length)?.index;
  }
  const slide = template.inspection.slides[index - 1];
  if (!slide) throw new ReportError(`${page.kind}: 사용할 회사 슬라이드가 없습니다. 디자인 설정에서 해당 페이지를 지정해 주세요`);
  return slide;
}
export function planTemplateSlides(model, pages, template) {
  if (!template) return pages;
  const result = [];
  for (const page of pages) {
    const slide = chosenSlide(template, page), body = slide.shapes.find(s => slide.roles[s.id] === 'body');
    if (slide.shapes.some(s => s.type === 'unsupported')) throw new ReportError(`${slide.index}번: SmartArt/OLE의 내용을 교체할 수 없습니다. 일반 텍스트·표·차트가 있는 원본 슬라이드를 선택해 주세요`);
    if (page.kind === 'table') {
      const tables = slide.shapes.filter(s => s.type === 'table');
      if (tables.length !== 1 || tables[0].columns !== page.section.table.headers.length) throw new ReportError(`${page.section.title}: 원본 표의 열 개수가 맞지 않습니다. 같은 열 개수의 슬라이드를 선택해 주세요`);
      const capacity = Math.max(1, tables[0].rows - 1);
      if (tables[0].rows < 2) throw new ReportError(`${slide.index}번: 표에 헤더와 데이터 행이 필요합니다`);
      for (let i = 0; i < page.tablePage.rows.length; i += capacity) {
        const rows = page.tablePage.rows.slice(i, i + capacity), layout = tables[0].tableLayout;
        [page.section.table.headers, ...rows].forEach((row, ri) => row.forEach((v, ci) => {
          const seed = layout.rows[ri], cell = { w: layout.widths[ci], h: seed.h, fontSize: seed.fonts[ci] };
          if (textCost(cell, v) > fitCapacity(cell) * 1.05) throw new ReportError(`${page.section.title}: 표의 ${ci + 1}번째 열 내용이 원본 셀보다 깁니다. 더 넓은 표 슬라이드를 선택해 주세요`);
        }));
        result.push({ ...page, tablePage: { ...page.tablePage, rows }, templateSlide: slide.index });
      }
      continue;
    }
    if (['flow', 'references'].includes(page.kind)) {
      if (!body) throw new ReportError(`${slide.index}번: 본문 텍스트 상자를 지정해 주세요`);
      const key = page.kind === 'flow' ? 'items' : 'references', items = page[key]; let batch = [], used = 0;
      const capacity = fitCapacity(body) * .88;
      for (const item of items) {
        const value = page.kind === 'flow' ? item.text : sourceFull(item);
        const size = textCost(body, value) + (page.kind === 'references' ? (body.w * 72 - 16) / body.fontSize : 0);
        if (size > capacity) throw new ReportError(`${slide.index}번: 한 문단이 원본 본문 영역보다 깁니다. 원고를 나누거나 더 넓은 슬라이드를 선택해 주세요`);
        if (batch.length && used + size > capacity) { result.push({ ...page, [key]: batch, templateSlide: slide.index }); batch = []; used = 0; }
        batch.push(item); used += size;
      }
      result.push({ ...page, [key]: batch, templateSlide: slide.index }); continue;
    }
    result.push({ ...page, templateSlide: slide.index });
  }
  // Recalculate continuation labels after template-specific splitting.
  const counts = new Map(); for (const p of result) if (p.section) counts.set(p.sectionIndex, (counts.get(p.sectionIndex) ?? 0) + 1);
  const current = new Map();
  result.forEach((p, i) => { p.number = i + 1; if (p.section) { p.parts = counts.get(p.sectionIndex); p.part = (current.get(p.sectionIndex) ?? 0) + 1; current.set(p.sectionIndex, p.part); } });
  for (const page of result) {
    const slide = template.inspection.slides[page.templateSlide - 1], values = content(page, model);
    if (!Object.values(slide.roles).includes('title')) throw new ReportError(`${slide.index}번: 제목 텍스트 상자가 필요합니다`);
    if (page.kind === 'chart' && (slide.shapes.filter(s => s.type === 'chart').length !== 1 || slide.shapes.find(s => s.type === 'chart').chartType === 'unsupported')) throw new ReportError(`${slide.index}번: 편집 가능한 선 또는 막대 차트 하나가 필요합니다`);
    if (page.kind === 'metrics' && Object.values(slide.roles).includes('metric1.value')) values.body = '';
    // Missing slots must never silently drop content. Append ancillary text to the body slot.
    for (const role of ['lead', 'meta', 'disclosure', 'unit', 'note', 'footer']) if (values[role] && !Object.values(slide.roles).includes(role)) { values.body = [values.body, values[role]].filter(Boolean).join('\n'); values[role] = ''; }
    for (const [role, value] of Object.entries(values)) {
      if (!value || role === 'page') continue;
      const shape = slide.shapes.find(s => slide.roles[s.id] === role);
      if (!shape) throw new ReportError(`${slide.index}번: ${role} 내용을 담을 텍스트 상자가 없습니다. 역할 설정을 수정해 주세요`);
      if (textCost(shape, value) > fitCapacity(shape) * 1.05) throw new ReportError(`${slide.index}번 ${shape.name}: 새 내용이 원본 영역보다 깁니다. 내용을 줄이거나 영역이 큰 슬라이드를 선택해 주세요`);
    }
    page.templateContent = values;
  }
  return result;
}

function replaceText(node, value) {
  const tx = first(node, NS.p, 'txBody') ?? first(node, NS.a, 'txBody');
  if (!tx) return;
  const doc = node.ownerDocument, paragraphs = children(tx, NS.a, 'p');
  const seed = paragraphs[0]?.cloneNode(true) ?? el(doc, 'a', 'p');
  paragraphs.forEach(p => tx.removeChild(p));
  for (const line of String(value ?? '').split('\n')) {
    const p = seed.cloneNode(true), run = first(p, NS.a, 'r'), props = first(run, NS.a, 'rPr')?.cloneNode(true);
    for (const child of children(p)) if (child.localName !== 'pPr' && child.localName !== 'endParaRPr') p.removeChild(child);
    const r = el(doc, 'a', 'r'); if (props) { for (const k of ['hlinkClick', 'hlinkMouseOver']) all(props, NS.a, k).forEach(n => n.parentNode.removeChild(n)); r.appendChild(props); }
    const t = el(doc, 'a', 't'); t.appendChild(doc.createTextNode(line)); r.appendChild(t);
    const end = children(p, NS.a, 'endParaRPr')[0]; if (end) p.insertBefore(r, end); else p.appendChild(r); tx.appendChild(p);
  }
}
function updateTable(node, headers, rows) {
  const table = first(node, NS.a, 'tbl'), original = children(table, NS.a, 'tr');
  if (original.length < 2) throw new ReportError('원본 표에 헤더와 데이터 행이 필요합니다');
  if (all(table, NS.a, 'tc').some(c => ['gridSpan', 'rowSpan', 'hMerge', 'vMerge'].some(a => c.hasAttribute(a)))) throw new ReportError('병합되지 않은 표가 있는 원본 슬라이드를 선택해 주세요');
  for (const tr of original) table.removeChild(tr);
  [headers, ...rows].forEach((row, i) => {
    const tr = original[Math.min(i, original.length - 1)].cloneNode(true), cells = children(tr, NS.a, 'tc');
    row.forEach((v, c) => replaceText(cells[c], v)); table.appendChild(tr);
  });
  // Preserve original row heights; unused rows disappear rather than changing typography.
  const h = children(table, NS.a, 'tr').reduce((sum, r) => sum + Number(r.getAttribute('h')), 0);
  const ext = first(first(node, NS.p, 'xfrm'), NS.a, 'ext'); if (ext) ext.setAttribute('cy', String(h));
}
function mergeChart(source, generated, title) {
  const originalPlot = first(source, NS.c, 'plotArea'), newPlot = first(generated, NS.c, 'plotArea');
  const oldChart = children(originalPlot, NS.c).find(n => /Chart$/.test(n.localName)), newChart = children(newPlot, NS.c).find(n => /Chart$/.test(n.localName));
  if (!oldChart || !newChart || oldChart.localName !== newChart.localName) throw new ReportError('원본 차트 형태와 원고의 차트 형태가 다릅니다. 선/막대 형태를 원본에 맞춰 주세요');
  const oldSeries = children(oldChart, NS.c, 'ser'), newSeries = children(newChart, NS.c, 'ser');
  if (!oldSeries.length) throw new ReportError('원본 차트에 계열이 없습니다');
  if (newSeries.length > oldSeries.length) throw new ReportError('원본 차트보다 새 계열이 많습니다. 계열이 충분한 원본 차트를 선택해 주세요');
  const insertBefore = children(oldChart, NS.c).find(n => ['dLbls', 'dropLines', 'hiLowLines', 'upDownBars', 'marker', 'smooth', 'axId', 'gapWidth', 'overlap', 'serLines', 'extLst'].includes(n.localName));
  oldSeries.forEach(n => oldChart.removeChild(n));
  newSeries.forEach((s, i) => {
    const n = oldSeries[Math.min(i, oldSeries.length - 1)].cloneNode(true);
    for (const k of ['idx', 'order', 'tx', 'cat', 'val']) {
      const old = children(n, NS.c, k)[0], fresh = children(s, NS.c, k)[0];
      if (old && fresh) n.replaceChild(source.importNode(fresh, true), old);
      else if (fresh) n.appendChild(source.importNode(fresh, true));
    }
    // Point-specific formatting and labels refer to old categories.
    for (const k of ['dPt', 'dLbls']) children(n, NS.c, k).forEach(c => n.removeChild(c));
    if (insertBefore) oldChart.insertBefore(n, insertBefore); else oldChart.appendChild(n);
  });
  // Fixed old axis ranges must not clip new data; preserve all other axis styling.
  for (const scaling of all(source, NS.c, 'scaling')) for (const k of ['min', 'max']) children(scaling, NS.c, k).forEach(n => scaling.removeChild(n));
  for (const chartTitle of all(source, NS.c, 'title')) {
    const texts = all(chartTitle, NS.a, 't');
    if (texts.length) texts.forEach((n, i) => { n.textContent = i === 0 ? title : ''; });
    else chartTitle.parentNode.removeChild(chartTitle);
  }
  for (const autoUpdate of all(source, NS.c, 'autoUpdate')) autoUpdate.setAttribute('val', '0');
  return source;
}

export async function templatePptxBuffer(model, pages, template, generatedBytes) {
  const zip = await loadPackage(template.bytes), sourceZip = await loadPackage(template.bytes), generated = await JSZip.loadAsync(generatedBytes);
  const types = await readXML(zip, '[Content_Types].xml');
  const generatedTypes = await readXML(generated, '[Content_Types].xml');
  const contentType = (sourceTypes, name) => children(sourceTypes.documentElement, NS.ct, 'Override').find(n => n.getAttribute('PartName') === '/' + name)?.getAttribute('ContentType');
  function addType(name, type) { if (!type) return; const old = children(types.documentElement, NS.ct, 'Override').find(n => n.getAttribute('PartName') === '/' + name); if (old) old.setAttribute('ContentType', type); else types.documentElement.appendChild(el(types, 'ct', 'Override', { PartName: '/' + name, ContentType: type })); }
  for (const d of children(generatedTypes.documentElement, NS.ct, 'Default')) if (!children(types.documentElement, NS.ct, 'Default').some(n => n.getAttribute('Extension') === d.getAttribute('Extension'))) types.documentElement.appendChild(types.importNode(d, true));
  let serial = 0;
  async function copyGraph(sourceZip, sourceTypes, part, map = new Map()) {
    if (map.has(part)) return map.get(part);
    const name = `ppt/rfassets/part${++serial}${path.posix.extname(part)}`; map.set(part, name);
    if (!sourceZip.file(part)) throw new ReportError(`차트 연결 파일이 없습니다: ${part}`);
    zip.file(name, await sourceZip.file(part).async('nodebuffer')); addType(name, contentType(sourceTypes, part));
    const rels = await relationships(sourceZip, part);
    for (const r of rels) if (!r.external) r.target = '/' + await copyGraph(sourceZip, sourceTypes, targetPart(part, r.target), map);
    if (rels.length) zip.file(relFile(name), serialize(relXML(rels)));
    return name;
  }
  const presentation = await readXML(zip, 'ppt/presentation.xml'), presRels = await relationships(zip, 'ppt/presentation.xml');
  const slideList = first(presentation, NS.p, 'sldIdLst'); while (slideList.firstChild) slideList.removeChild(slideList.firstChild);
  // Remove old slide navigation and personal annotations, then prune unreachable parts.
  for (const k of ['custShowLst', 'sectionLst', 'tagLst', 'custDataLst']) all(presentation, NS.p, k).forEach(n => n.parentNode.removeChild(n));
  all(presentation, NS.p, 'extLst').forEach(n => n.parentNode.removeChild(n));
  const retainedRels = presRels.filter(r => !/\/(slide|commentAuthors|customXml|tags)$/.test(r.type));
  for (const name of Object.keys(zip.files)) if (/^ppt\/(slides|notesSlides)\//.test(name)) zip.remove(name);
  const notesMasterRel = retainedRels.find(r => r.type.endsWith('/notesMaster'));
  let notesMaster = notesMasterRel ? targetPart('ppt/presentation.xml', notesMasterRel.target) : null;
  if (!notesMaster) {
    const gm = (await relationships(generated, 'ppt/presentation.xml')).find(r => r.type.endsWith('/notesMaster'));
    if (gm) {
      notesMaster = await copyGraph(generated, generatedTypes, targetPart('ppt/presentation.xml', gm.target));
      retainedRels.push({ id: 'rIdRFNotesMaster', type: `${NS.r}/notesMaster`, target: '/' + notesMaster });
      let list = first(presentation, NS.p, 'notesMasterIdLst');
      if (!list) { list = el(presentation, 'p', 'notesMasterIdLst'); presentation.documentElement.insertBefore(list, slideList); }
      const id = el(presentation, 'p', 'notesMasterId'); id.setAttributeNS(NS.r, 'r:id', 'rIdRFNotesMaster'); list.appendChild(id);
    }
  }
  for (const [i, page] of pages.entries()) {
    const src = template.inspection.slides[page.templateSlide - 1], slideDoc = await readXML(sourceZip, src.part);
    const nodes = shapeNodes(slideDoc);
    for (const k of ['tagLst', 'custDataLst']) all(slideDoc, NS.p, k).forEach(n => n.parentNode.removeChild(n));
    const values = page.templateContent;
    for (const node of nodes) {
      const info = shapeInfo(node), role = src.roles[info.id];
      if (info.type === 'text' && !info.grouped && role !== 'keep') replaceText(node, values[role] ?? '');
      if (info.type === 'table') {
        if (page.kind === 'table') updateTable(node, page.section.table.headers, page.tablePage.rows);
        else node.parentNode.removeChild(node);
      }
      if (info.type === 'chart' && page.kind !== 'chart') node.parentNode.removeChild(node);
    }
    const name = `ppt/slides/slide${i + 1}.xml`;
    let rels = (await relationships(sourceZip, src.part)).filter(r => !/\/(notesSlide|comments|tags|slide|customXml)$/.test(r.type));
    for (const r of rels) if (!r.external) r.target = '/' + targetPart(src.part, r.target);
    rels = rels.filter(r => !r.type.endsWith('/chart') || page.kind === 'chart');
    if (page.kind === 'chart') {
      const gr = (await relationships(generated, `ppt/slides/slide${page.number}.xml`)).find(r => r.type.endsWith('/chart'));
      if (!gr) throw new ReportError('새 차트 패키지를 찾을 수 없습니다');
      const sourceRel = rels.find(r => r.type.endsWith('/chart'));
      const originalPart = sourceRel.target.slice(1), generatedPart = targetPart(`ppt/slides/slide${page.number}.xml`, gr.target);
      const clone = await copyGraph(sourceZip, types, originalPart);
      const merged = mergeChart(await readXML(sourceZip, originalPart), await readXML(generated, generatedPart), page.section.title);
      const oldRels = await relationships(zip, clone), newRels = await relationships(generated, generatedPart);
      const workbook = newRels.find(r => r.type.endsWith('/package'));
      if (!workbook || workbook.external) throw new ReportError('새 차트의 편집용 엑셀 데이터가 없습니다');
      const workbookName = await copyGraph(generated, generatedTypes, targetPart(generatedPart, workbook.target));
      const external = first(merged, NS.c, 'externalData');
      const workbookID = external?.getAttributeNS(NS.r, 'id') ?? 'rIdRFWorkbook';
      if (!external) { const e = el(merged, 'c', 'externalData'); e.setAttributeNS(NS.r, 'r:id', workbookID); merged.documentElement.appendChild(e); }
      const patched = oldRels.filter(r => !r.type.endsWith('/package')); patched.push({ id: workbookID, type: `${NS.r}/package`, target: '/' + workbookName });
      zip.file(clone, serialize(merged)); zip.file(relFile(clone), serialize(relXML(patched))); sourceRel.target = '/' + clone;
    }
    const notePart = `ppt/notesSlides/notesSlide${i + 1}.xml`;
    const noteText = [values.title, page.section?.lead, ...(page.section?.body ?? []), ...(page.section?.bullets ?? []), ...(page.section?.sourceIds ?? []).map(id => sourceFull(model.sources[id])), ...(page.references ?? []).map(sourceFull)].filter(Boolean).join('\n\n');
    zip.file(notePart, `<p:notes xmlns:p="${NS.p}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeHTML(noteText)}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`);
    const nr = [{ id: 'rIdSlide', type: `${NS.r}/slide`, target: '/' + name }]; if (notesMaster) nr.push({ id: 'rIdMaster', type: `${NS.r}/notesMaster`, target: '/' + notesMaster });
    zip.file(relFile(notePart), serialize(relXML(nr))); addType(notePart, 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');
    rels.push({ id: 'rIdRFNotes', type: `${NS.r}/notesSlide`, target: '/' + notePart });
    // Drop links into deleted original slides. Static brand images and layout relationships stay intact.
    for (const k of ['hlinkClick', 'hlinkMouseOver']) all(slideDoc, NS.a, k).filter(n => !rels.some(r => r.id === n.getAttributeNS(NS.r, 'id'))).forEach(n => n.parentNode.removeChild(n));
    zip.file(name, serialize(slideDoc)); zip.file(relFile(name), serialize(relXML(rels))); addType(name, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
    const rid = `rIdRFSlide${i + 1}`; retainedRels.push({ id: rid, type: `${NS.r}/slide`, target: `slides/slide${i + 1}.xml` });
    const id = el(presentation, 'p', 'sldId', { id: 256 + i }); id.setAttributeNS(NS.r, 'r:id', rid); slideList.appendChild(id);
  }
  zip.file('ppt/presentation.xml', serialize(presentation)); zip.file('ppt/_rels/presentation.xml.rels', serialize(relXML(retainedRels)));
  // Old chart workbooks, notes, source slides and thumbnails must not ride along as hidden content.
  const root = await readXML(zip, '_rels/.rels');
  children(root.documentElement, NS.rel, 'Relationship').filter(n => /\/(thumbnail|custom-properties)$/.test(n.getAttribute('Type'))).forEach(n => n.parentNode.removeChild(n)); zip.file('_rels/.rels', serialize(root));
  const seen = new Set(['[Content_Types].xml', '_rels/.rels']);
  async function visit(part) {
    if (seen.has(part)) return; if (!zip.file(part)) throw new ReportError(`PPTX 연결이 끊어졌습니다: ${part}`); seen.add(part);
    const rf = relFile(part); if (zip.file(rf)) seen.add(rf);
    for (const r of await relationships(zip, part)) if (!r.external) await visit(targetPart(part, r.target));
  }
  for (const r of children(root.documentElement, NS.rel, 'Relationship')) if (r.getAttribute('TargetMode') !== 'External') await visit(targetPart('', r.getAttribute('Target')));
  for (const name of Object.keys(zip.files)) if (!zip.files[name].dir && !seen.has(name)) zip.remove(name);
  children(types.documentElement, NS.ct, 'Override').filter(n => !zip.file(n.getAttribute('PartName').slice(1))).forEach(n => n.parentNode.removeChild(n));
  zip.file('[Content_Types].xml', serialize(types));
  const core = zip.file('docProps/core.xml'); if (core) { const d = xml(await core.async('string')); for (const k of ['title', 'subject', 'description']) Array.from(d.getElementsByTagNameNS('*', k)).forEach(n => { n.textContent = k === 'title' ? model.doc.title : model.doc.subtitle ?? ''; }); zip.file('docProps/core.xml', serialize(d)); }
  const app = zip.file('docProps/app.xml'); if (app) { const d = xml(await app.async('string')); for (const k of ['Slides', 'Notes']) Array.from(d.getElementsByTagNameNS('*', k)).forEach(n => { n.textContent = String(pages.length); }); for (const k of ['TitlesOfParts', 'HeadingPairs']) Array.from(d.getElementsByTagNameNS('*', k)).forEach(n => n.parentNode.removeChild(n)); zip.file('docProps/app.xml', serialize(d)); }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}
