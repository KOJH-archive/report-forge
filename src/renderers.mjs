import PptxGenJS from 'pptxgenjs';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, HeadingLevel, AlignmentType, BorderStyle, ExternalHyperlink, Footer, PageNumber } from 'docx';
import { escapeHTML as esc, planSlides, sourceShort, sourceFull } from './core.mjs';

const W = 13.333333, H = 7.5;
const chartFont = (theme, size = 13) => ({
  catAxisLabelFontFace: theme.font, catAxisLabelFontSize: size, catAxisLabelColor: theme.muted,
  valAxisLabelFontFace: theme.font, valAxisLabelFontSize: size, valAxisLabelColor: theme.muted,
  legendFontFace: theme.font, legendFontSize: size, legendColor: theme.foreground,
  dataLabelFormatCode: '#,##0.##', valAxisLabelFormatCode: '#,##0.##',
});
function slideText(slide, text, theme, x, y, w, h, size = 18, extra = {}) {
  slide.addText(String(text), { x, y, w, h, fontFace: theme.font, fontSize: size, color: theme.foreground, margin: 0, breakLine: false, valign: 'top', ...extra });
}
export async function pptxBuffer(model, pages = planSlides(model)) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE'; pptx.author = model.doc.author ?? 'Report Forge';
  pptx.subject = model.doc.subtitle ?? ''; pptx.title = model.doc.title; pptx.lang = 'ko-KR';
  const t = model.theme;
  pptx.theme = { headFontFace: t.font, bodyFontFace: t.font, lang: 'ko-KR' };
  for (const page of pages) {
    const slide = pptx.addSlide(); slide.background = { color: t.background };
    const section = page.section;
    if (page.kind === 'cover') {
      slideText(slide, model.doc.title, t, 0.75, 1.45, 11.85, 1.8, 42, { bold: true, color: t.accent });
      slideText(slide, model.doc.subtitle ?? '', t, 0.8, 3.7, 11.7, 0.95, 22);
      slideText(slide, [model.doc.author, model.doc.date].filter(Boolean).join('   '), t, 0.8, 5.35, 11.5, 0.4, 16, { color: t.muted });
      if (model.doc.disclosure) slideText(slide, model.doc.disclosure, t, 0.8, 6.3, 11.5, 0.6, 13, { color: t.muted });
      slide.addNotes([model.doc.title, model.doc.disclosure ?? ''].join('\n'));
      continue;
    }
    const title = page.kind === 'references' ? '참고 자료' : `${section.title}${page.parts > 1 ? ` (${page.part}/${page.parts})` : ''}`;
    slideText(slide, title, t, 0.65, 0.55, 12.0, 0.9, 32, { bold: true });
    if (section?.lead) slideText(slide, section.lead, t, 0.68, 1.52, 11.9, 0.72, 18, { color: t.accent });
    if (page.kind === 'flow') {
      let y = 2.4;
      page.items.forEach(item => {
        slideText(slide, item.text, t, item.bullet ? 0.87 : 0.68, y, item.bullet ? 11.72 : 11.95, item.h, 18,
          item.bullet ? { bullet: { indent: 16 }, hanging: 4, paraSpaceAfterPt: 8 } : {});
        y += item.h;
      });
    } else if (page.kind === 'metrics') {
      const metrics = section.resolvedMetrics;
      const width = 12 / metrics.length;
      metrics.forEach((metric, i) => {
        const x = 0.68 + width * i;
        slideText(slide, metric.label, t, x, 2.75, width - 0.35, 0.8, 20);
        slideText(slide, metric.formatted, t, x, 3.7, width - 0.35, 1.15, 30, { color: t.accent, bold: true });
        if (metric.description) slideText(slide, metric.description, t, x, 5.1, width - 0.35, 0.9, 16, { color: t.muted });
      });
    } else if (page.kind === 'chart') {
      slideText(slide, `단위: ${section.unit}`, t, 0.7, 2.24, 3, 0.3, 13, { color: t.muted });
      const allValues = section.chartData.flatMap(s => s.values);
      slide.addChart(section.chartType === 'bar' ? pptx.ChartType.bar : pptx.ChartType.line, structuredClone(section.chartData), {
        x: 0.65, y: 2.55, w: 12, h: page.chartNote ? 3.3 : 3.95,
        chartColors: [t.accent, t.secondary, t.muted], showLegend: section.chartData.length > 1, legendPos: 'b',
        showTitle: false, showValue: false, showBorder: false, showCatName: false,
        showMarker: section.chartType === 'line', markerSize: 5, lineSize: 2.5,
        barDir: 'col', barGrouping: 'clustered', gapSize: 80,
        catAxisLineShow: false, valAxisLineShow: false, catAxisMajorTickMark: 'none', valAxisMajorTickMark: 'none',
        valGridLine: { color: t.rule, width: 0.7 }, catGridLine: { style: 'none' },
        valAxisMinVal: Math.min(0, ...allValues),
        chartArea: { fill: { color: t.background }, border: { color: t.background, pt: 0 } },
        plotArea: { fill: { color: t.background }, border: { color: t.background, pt: 0 } },
        ...chartFont(t),
        valAxisLabelFormatCode: allValues.every(Number.isInteger) ? '#,##0' : '#,##0.00',
      });
      if (page.chartNote) slideText(slide, page.chartNote, t, 0.7, 6.0, 11.9, 0.65, 16);
    } else if (page.kind === 'table') {
      const rows = [section.table.headers.map(text => ({ text, options: { bold: true, color: 'FFFFFF', fill: t.tableHeader } })),
        ...page.tablePage.rows.map((row, i) => row.map(text => ({ text, options: { fill: i % 2 ? t.tableStripe : t.background } })))];
      slide.addTable(rows, {
        x: 0.65, y: 2.45, w: 12, h: page.tablePage.heights.reduce((a, b) => a + b, 0),
        colW: page.tablePage.widths, rowH: page.tablePage.heights,
        fontFace: t.font, fontSize: 16, color: t.foreground,
        border: { color: t.rule, pt: 0.6 }, margin: [0.08, 0.12, 0.08, 0.12],
        valign: 'middle', autoPage: false, paraSpaceAfterPt: 0,
      });
    } else if (page.kind === 'references') {
      let y = 2.15;
      page.references.forEach(source => {
        const short = sourceFull({ ...source, url: undefined });
        slideText(slide, short, t, 0.7, y, 11.9, source.slideH - (source.url ? 0.34 : 0), 15);
        if (source.url) slideText(slide, '원문 열기', t, 0.7, y + source.slideH - 0.35, 11.9, 0.28, 13, { color: t.accent, hyperlink: { url: source.url } });
        y += source.slideH;
      });
    }
    const note = section ? sourceShort(model, section.sourceIds) : '';
    slideText(slide, note, t, 0.68, 6.91, 11.2, 0.33, 10, { color: t.muted });
    slideText(slide, String(page.number), t, 12.05, 6.9, 0.55, 0.32, 11, { color: t.muted, align: 'right' });
    slide.addNotes(section ? [section.title, ...(section.body ?? []), ...(section.bullets ?? []), ...section.sourceIds.map(id => sourceFull(model.sources[id]))].join('\n\n') : page.references.map(sourceFull).join('\n\n'));
  }
  const bytes = Buffer.from(await pptx.write({ outputType: 'nodebuffer', compression: true }));
  const zip = await JSZip.loadAsync(bytes);
  const types = await zip.file('[Content_Types].xml').async('string');
  // PptxGenJS 4.0.1 can declare unused slide-master parts. Remove only
  // declarations whose target does not exist, preserving every actual part.
  const cleaned = types.replace(/<Override\s[^>]+\/>/g, entry => {
    const target = entry.match(/PartName="([^\"]+)"/)?.[1]?.replace(/^\//, '');
    return target && !zip.file(target) ? '' : entry;
  });
  zip.file('[Content_Types].xml', cleaned);
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

function wordTable(headers, rows, theme, width = 9360) {
  const border = { style: BorderStyle.SINGLE, size: 4, color: 'D9D9D9' };
  const widths = headers.map(() => Math.floor(width / headers.length));
  if (headers.join(',') === '시점,실행 내용,담당') widths.splice(0, widths.length, 1500, 5860, 2000);
  if (headers.join(',') === '선택지,기대 효과,검토할 점') widths.splice(0, widths.length, 2000, 3680, 3680);
  return new Table({ width: { size: width, type: WidthType.DXA }, columnWidths: widths,
    rows: [headers, ...rows].map((row, index) => new TableRow({ tableHeader: index === 0, cantSplit: true,
      children: row.map((text, i) => new TableCell({ width: { size: widths[i], type: WidthType.DXA },
        shading: { fill: index === 0 ? theme.tableHeader : index % 2 === 0 ? theme.tableStripe : 'FFFFFF' },
        borders: { top: border, bottom: border, left: border, right: border },
        margins: { top: 120, bottom: 120, left: 120, right: 120 }, verticalAlign: 'center',
        children: [new Paragraph({ spacing: { after: 0, line: 270 }, children: [new TextRun({ text: String(text), bold: index === 0, color: index === 0 ? 'FFFFFF' : '000000', size: 21 })] })],
      })),
    })),
  });
}
function sourceParagraph(model, ids) {
  return new Paragraph({ spacing: { before: 100, after: 180 }, children: [new TextRun({ text: `출처: ${sourceShort(model, ids)}`, size: 18, color: '595959' })] });
}
// Reuse the same data-backed Office chart package in Word and PowerPoint.
// This keeps Word charts editable without rasterizing them into screenshots.
export async function embedWordCharts(docxBytes, pptxBytes, markers) {
  if (!markers.length) return docxBytes;
  const wordZip = await JSZip.loadAsync(docxBytes), deckZip = await JSZip.loadAsync(pptxBytes);
  let documentXML = await wordZip.file('word/document.xml').async('string');
  let relationships = await wordZip.file('word/_rels/document.xml.rels').async('string');
  let contentTypes = await wordZip.file('[Content_Types].xml').async('string');
  // PptxGenJS chart filenames use a process-wide counter. Follow each slide's
  // relationships rather than assuming that every new deck starts at chart1.
  const slideRelFiles = Object.keys(deckZip.files).filter(n => /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(n))
    .sort((a, b) => Number(a.match(/slide(\d+)/)[1]) - Number(b.match(/slide(\d+)/)[1]));
  const chartFiles = [];
  for (const filename of slideRelFiles) {
    const xml = await deckZip.file(filename).async('string');
    for (const relationship of xml.matchAll(/<Relationship\s[^>]+\/>/g)) {
      if (!/Type="[^\"]+\/chart"/.test(relationship[0])) continue;
      const target = relationship[0].match(/Target="([^\"]+)"/)?.[1];
      if (target) chartFiles.push(pathForChart(target));
    }
  }
  for (const [filename, entry] of Object.entries(deckZip.files)) {
    if (entry.dir || !(filename.startsWith('ppt/charts/') || filename.startsWith('ppt/embeddings/'))) continue;
    wordZip.file(filename.replace(/^ppt\//, 'word/'), await entry.async('nodebuffer'));
  }
  for (const marker of markers) {
    const n = marker.index, rId = `rIdReportForgeChart${n}`;
    const chartFile = chartFiles[n - 1], chartName = chartFile?.split('/').at(-1);
    if (!chartFile || !deckZip.file(chartFile)) throw new Error(`Native chart ${n} is missing`);
    const drawing = `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="5943600" cy="3200400"/><wp:docPr id="${1000 + n}" name="${esc(marker.title)}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rId}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    const paragraphPattern = new RegExp(`<w:p(?:\\s[^>]*)?>(?:(?!<w:p(?:\\s|>))[\\s\\S])*?${marker.token}(?:(?!<w:p(?:\\s|>))[\\s\\S])*?</w:p>`);
    if (!paragraphPattern.test(documentXML)) throw new Error(`Chart placeholder ${n} is missing`);
    documentXML = documentXML.replace(paragraphPattern, drawing);
    relationships = relationships.replace('</Relationships>', `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/${chartName}"/></Relationships>`);
    contentTypes = contentTypes.replace('</Types>', `<Override PartName="/word/charts/${chartName}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`);
  }
  if (!contentTypes.includes('Extension="xlsx"')) contentTypes = contentTypes.replace('</Types>', '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/></Types>');
  wordZip.file('word/document.xml', documentXML); wordZip.file('word/_rels/document.xml.rels', relationships); wordZip.file('[Content_Types].xml', contentTypes);
  return wordZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}
function pathForChart(target) {
  if (target.startsWith('/')) return target.slice(1);
  const segments = ['ppt', 'slides'];
  for (const segment of target.split('/')) { if (segment === '..') segments.pop(); else if (segment && segment !== '.') segments.push(segment); }
  return segments.join('/');
}
export async function docxBuffer(model, deckBytes) {
  const { doc, theme } = model;
  const children = [new Paragraph({ text: doc.title, heading: HeadingLevel.TITLE }),
    new Paragraph({ text: doc.subtitle ?? '', spacing: { after: 180 } }),
    new Paragraph({ text: [doc.author, doc.date].filter(Boolean).join('   '), spacing: { after: 160 } })];
  if (doc.disclosure) children.push(new Paragraph({ text: doc.disclosure, spacing: { after: 220 } }));
  const markers = []; let chartIndex = 0;
  doc.sections.forEach(section => {
    children.push(new Paragraph({ text: section.title, heading: HeadingLevel.HEADING_1, keepNext: true }));
    if (section.lead) children.push(new Paragraph({ children: [new TextRun({ text: section.lead, bold: true })], keepNext: ['chart', 'table', 'metrics', 'comparison', 'timeline'].includes(section.type), spacing: { after: 160 } }));
    if (section.type === 'metrics') children.push(wordTable(['지표', '값', '계산 범위'], section.resolvedMetrics.map(m => [m.label, m.formatted, m.description ?? m.operation]), theme));
    if (section.table) children.push(wordTable(section.table.headers, section.table.rows, theme));
    if (section.type === 'chart') {
      const index = ++chartIndex, token = `REPORTFORGE_NATIVE_CHART_${index}_END`;
      markers.push({ index, token, title: section.title }); children.push(new Paragraph(token));
      children.push(new Paragraph({ text: `단위: ${section.unit}`, spacing: { before: 80, after: 160 } }));
    }
    (section.body ?? []).forEach(text => children.push(new Paragraph({ text, spacing: { after: 160, line: 320 } })));
    (section.bullets ?? []).forEach(text => children.push(new Paragraph({ text, bullet: { level: 0 }, spacing: { after: 100, line: 300 } })));
    if (section.sourceIds.length) children.push(sourceParagraph(model, section.sourceIds));
  });
  if (doc.sources.length) children.push(new Paragraph({ text: '참고 자료', heading: HeadingLevel.HEADING_1, keepNext: true }));
  Object.values(model.sources).forEach(source => {
    children.push(new Paragraph({ text: sourceFull({ ...source, url: undefined }), spacing: { after: source.url ? 60 : 180 } }));
    if (source.url) children.push(new Paragraph({ spacing: { after: 180 }, children: [new ExternalHyperlink({ link: source.url, children: [new TextRun({ text: source.url, style: 'Hyperlink' })] })] }));
  });
  const file = new Document({ creator: doc.author ?? 'Report Forge', title: doc.title, description: doc.subtitle ?? '',
    styles: { default: { document: { run: { font: theme.font, size: 22, color: '000000' }, paragraph: { spacing: { line: 300, after: 140 } } } }, paragraphStyles: [
      { id: 'Title', name: 'Title', basedOn: 'Normal', next: 'Normal', run: { color: '000000', size: 38, bold: true }, paragraph: { spacing: { after: 220 } } },
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', run: { color: '000000', size: 28, bold: true }, paragraph: { spacing: { before: 300, after: 180 }, keepNext: true } },
    ] },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1100, bottom: 1100, left: 1273, right: 1273 } } },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ children: [PageNumber.CURRENT], size: 18, color: '595959' })] })] }) }, children }],
  });
  return embedWordCharts(await Packer.toBuffer(file), deckBytes, markers);
}
export async function xlsxBuffer(model) {
  const workbook = new ExcelJS.Workbook(); workbook.creator = model.doc.author ?? 'Report Forge';
  const metadata = workbook.addWorksheet('보고서 정보');
  metadata.addRows([['항목', '내용'], ['제목', model.doc.title], ['작성일', model.doc.date ?? ''], ['고지', model.doc.disclosure ?? ''], ['설명', '데이터 스냅샷이며 수식은 저장된 계산 결과를 사용합니다']]);
  const metrics = workbook.addWorksheet('계산 지표');
  metrics.addRows([['지표 id', '이름', '원본 데이터', '열', '계산', '원시 값', '단위', '표시 값'], ...Object.values(model.metrics).map(m => [m.id, m.label, m.dataset, m.column, m.operation, m.value, m.unit ?? '', m.formatted])]);
  const refs = workbook.addWorksheet('참고 자료');
  refs.addRows([['번호', '자료명', '작성 기관', '날짜', '위치', 'URL'], ...Object.values(model.sources).map(s => [s.number, s.title, s.publisher ?? '', s.date ?? '', s.locator ?? '', s.url ?? ''])]);
  const used = new Set(workbook.worksheets.map(s => s.name));
  Object.values(model.datasets).forEach(ds => {
    let name = ds.id.replace(/[\\/*?:[\]]/g, '_').slice(0, 25) || 'data';
    const base = name; let n = 1; while (used.has(name)) name = `${base}_${n++}`;
    used.add(name);
    const sheet = workbook.addWorksheet(name); sheet.addRow(ds.headers);
    ds.rows.forEach(row => sheet.addRow(ds.headers.map(h => row[h])));
  });
  workbook.worksheets.forEach(sheet => {
    sheet.views = [{ state: 'frozen', ySplit: 1 }];
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${model.theme.tableHeader}` } };
    sheet.columns.forEach((column, i) => { column.width = i < 2 ? 26 : 22; });
    sheet.autoFilter = { from: 'A1', to: { row: 1, column: sheet.columnCount } };
  });
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

export function chartSVG(section, theme) {
  const x0 = 70, y0 = 24, width = 950, height = 310, bottom = y0 + height;
  const values = section.chartData.flatMap(s => s.values);
  const low = Math.min(0, ...values), high = Math.max(0, ...values);
  const span = high - low || 1;
  const y = value => bottom - (value - low) / span * height;
  const labels = section.chartData[0].labels;
  const x = i => x0 + (i + 0.5) * width / labels.length;
  const colors = [theme.accent, theme.secondary, theme.muted];
  const marks = [];
  for (let i = 0; i <= 4; i++) {
    const value = low + span * i / 4, yy = y(value);
    marks.push(`<line x1="${x0}" y1="${yy}" x2="${x0 + width}" y2="${yy}" stroke="#${theme.rule}"/><text x="${x0 - 12}" y="${yy + 5}" text-anchor="end">${esc(new Intl.NumberFormat('ko-KR', { maximumFractionDigits: 1 }).format(value))}</text>`);
  }
  labels.forEach((label, i) => marks.push(`<text x="${x(i)}" y="${bottom + 27}" text-anchor="middle">${esc(label)}</text>`));
  section.chartData.forEach((series, n) => {
    if (section.chartType === 'line') {
      marks.push(`<polyline points="${series.values.map((v, i) => `${x(i)},${y(v)}`).join(' ')}" fill="none" stroke="#${colors[n]}" stroke-width="3"/>`);
      series.values.forEach((v, i) => marks.push(`<circle cx="${x(i)}" cy="${y(v)}" r="4" fill="#${colors[n]}"><title>${esc(labels[i])}, ${esc(series.name)}: ${esc(v)} ${esc(section.unit)}</title></circle>`));
    } else {
      const group = width / labels.length * 0.75, bw = group / section.chartData.length;
      series.values.forEach((v, i) => marks.push(`<rect x="${x(i) - group / 2 + n * bw}" y="${Math.min(y(v), y(0))}" width="${bw * 0.88}" height="${Math.abs(y(v) - y(0))}" fill="#${colors[n]}"><title>${esc(labels[i])}, ${esc(series.name)}: ${esc(v)} ${esc(section.unit)}</title></rect>`));
    }
  });
  return `<svg viewBox="0 0 1080 395" role="img" aria-label="${esc(section.title)}" style="font-family:inherit;font-size:14px;fill:#${theme.muted}">${marks.join('')}</svg><div class="legend">${section.chartData.map((s, i) => `<span><i style="background:#${colors[i]}"></i>${esc(s.name)}</span>`).join('')}<span>단위: ${esc(section.unit)}</span></div>`;
}
function htmlTable(headers, rows) { return `<div class="table-wrap"><table><thead><tr>${headers.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.map(row => `<tr>${row.map(c => `<td>${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`; }
export function slideHTML(model, page) {
  const s = page.section;
  let body = '';
  if (page.kind === 'cover') return `<article class="slide cover" data-page="${page.number}"><div><h1>${esc(model.doc.title)}</h1><p class="subtitle">${esc(model.doc.subtitle ?? '')}</p><p>${esc([model.doc.author, model.doc.date].filter(Boolean).join('   '))}</p><small>${esc(model.doc.disclosure ?? '')}</small></div></article>`;
  if (page.kind === 'flow') body = `<div class="flow">${page.items.map(i => i.bullet ? `<p class="bullet">${esc(i.text)}</p>` : `<p>${esc(i.text)}</p>`).join('')}</div>`;
  if (page.kind === 'metrics') body = `<div class="metrics">${s.resolvedMetrics.map(m => `<div><h3>${esc(m.label)}</h3><strong>${esc(m.formatted)}</strong><p>${esc(m.description ?? '')}</p></div>`).join('')}</div>`;
  if (page.kind === 'chart') body = `<div class="chart">${chartSVG(s, model.theme)}</div>${page.chartNote ? `<p class="chart-note">${esc(page.chartNote)}</p>` : ''}`;
  if (page.kind === 'table') body = htmlTable(s.table.headers, page.tablePage.rows);
  if (page.kind === 'references') body = `<div class="references">${page.references.map(r => `<p>${esc(sourceFull(r)).replace(/\n/g, '<br>')}</p>`).join('')}</div>`;
  const title = s ? `${s.title}${page.parts > 1 ? ` (${page.part}/${page.parts})` : ''}` : '참고 자료';
  return `<article class="slide" data-page="${page.number}"><h2>${esc(title)}</h2>${s?.lead ? `<p class="lead">${esc(s.lead)}</p>` : '<div class="lead-spacer"></div>'}<div class="evidence">${body}</div><footer><span>${s ? esc(sourceShort(model, s.sourceIds)) : ''}</span><span>${page.number}</span></footer></article>`;
}
export function previewHTML(model, pages = planSlides(model), { companyDesign = false } = {}) {
  const t = model.theme;
  const refs = Object.values(model.sources).map(r => `<p>${esc(sourceFull(r)).replace(/\n/g, '<br>')}</p>`).join('');
  const report = model.doc.sections.map(s => `<section><h2>${esc(s.title)}</h2>${s.lead ? `<p><b>${esc(s.lead)}</b></p>` : ''}${s.type === 'chart' ? chartSVG(s, t) : ''}${s.type === 'metrics' ? htmlTable(['지표', '값', '범위'], s.resolvedMetrics.map(m => [m.label, m.formatted, m.description ?? ''])) : ''}${s.table ? htmlTable(s.table.headers, s.table.rows) : ''}${(s.body ?? []).map(p => `<p>${esc(p)}</p>`).join('')}${s.bullets?.length ? `<ul>${s.bullets.map(b => `<li>${esc(b)}</li>`).join('')}</ul>` : ''}${s.sourceIds.length ? `<small>출처: ${esc(sourceShort(model, s.sourceIds))}</small>` : ''}</section>`).join('');
  return `<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(model.doc.title)}</title><style>
  *{box-sizing:border-box}body{margin:0;background:#e9edf0;color:#${t.foreground};font-family:"${esc(t.font)}","Noto Sans KR",sans-serif}nav{position:sticky;top:0;z-index:2;display:flex;gap:8px;align-items:center;padding:12px 24px;background:#${t.background};border-bottom:1px solid #${t.rule}}button{font:inherit;padding:8px 16px;border:1px solid #${t.rule};background:#${t.background};color:#${t.foreground};cursor:pointer}button.active{background:#${t.accent};color:#fff}nav small{margin-left:auto;color:#${t.muted}}main{max-width:1160px;margin:24px auto;padding:0 16px}.slide{position:relative;background:#${t.background};aspect-ratio:16/9;margin-bottom:24px;padding:4.1% 5%;overflow:hidden}.slide h2{font-size:clamp(19px,2.45vw,32px);margin:0 0 18px;font-weight:700}.lead{color:#${t.accent};font-size:clamp(13px,1.5vw,19px);margin:0 0 20px;min-height:40px}.lead-spacer{height:56px}.cover{display:flex;align-items:center;padding:7%}.cover h1{font-size:clamp(25px,3.6vw,46px);max-width:95%;color:#${t.accent};line-height:1.5;margin:0 0 28px}.subtitle{font-size:22px;line-height:1.6}.cover small{display:block;margin-top:42px;color:#${t.muted}}.evidence{font-size:clamp(13px,1.5vw,19px);line-height:1.65}.flow p{margin:0 0 20px}.bullet{padding-left:20px;position:relative}.bullet:before{content:'•';position:absolute;left:0}.metrics{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:30px;padding-top:46px}.metrics h3{font-size:20px;font-weight:500}.metrics strong{font-size:32px;color:#${t.accent}}.metrics p{font-size:16px;color:#${t.muted}}table{width:100%;border-collapse:collapse;font-size:clamp(12px,1.35vw,17px)}th,td{padding:12px;border:1px solid #${t.rule};text-align:left;vertical-align:middle}th{background:#${t.tableHeader};color:#fff}tbody tr:nth-child(even){background:#${t.tableStripe}}.table-wrap{overflow-x:auto}.chart{max-width:100%}.chart svg{width:100%;max-height:315px;display:block}.legend{display:flex;justify-content:center;gap:24px;flex-wrap:wrap;font-size:14px}.legend i{display:inline-block;width:11px;height:11px;margin-right:7px}.chart-note{font-size:15px;margin-top:14px}.references{font-size:16px;overflow-wrap:anywhere}.references p{margin:0 0 23px}footer{position:absolute;left:5%;right:5%;bottom:3%;display:flex;justify-content:space-between;font-size:11px;color:#${t.muted}}.report{background:#fff;color:#000;padding:48px 56px;font-size:15px;line-height:1.8}.report h1{font-size:28px}.report h2{font-size:21px;margin-top:32px}.report section{margin-bottom:30px}.report small{color:#595959}.report .table-wrap{margin:20px 0}.report th,.report td{font-size:14px}.report svg{width:100%;height:auto}.report .legend{margin-bottom:20px}[hidden]{display:none!important}@media(max-width:650px){main{padding:0 6px;margin:10px auto}.slide{aspect-ratio:auto;min-height:400px;padding:24px 20px 55px}.metrics{gap:14px;padding-top:10px}.metrics strong{font-size:22px}.metrics h3{font-size:15px}.metrics p{font-size:12px}.report{padding:24px}.subtitle{font-size:16px}nav{padding:10px}nav small{display:none}.chart-note{font-size:13px}}@media print{nav{display:none}body{background:white}main{max-width:none;margin:0;padding:0}.slide{break-after:page;margin:0;box-shadow:none}.report{padding:0}}
  </style></head><body><nav><button class="active" data-mode="slides">발표 자료</button><button data-mode="report">보고서</button><small>${companyDesign ? '원고 확인용 기본 서식입니다. 회사 디자인은 내보낸 PowerPoint에 적용됩니다' : '내용 미리보기이며 Office 파일의 실제 렌더링과는 다를 수 있습니다'}</small></nav><main id="slides">${pages.map(p => slideHTML(model, p)).join('')}</main><main id="report" class="report" hidden><h1>${esc(model.doc.title)}</h1><p>${esc(model.doc.subtitle ?? '')}</p><p>${esc(model.doc.disclosure ?? '')}</p>${report}<h2>참고 자료</h2>${refs}</main><script>document.querySelectorAll('[data-mode]').forEach(button=>button.addEventListener('click',()=>{document.querySelectorAll('[data-mode]').forEach(b=>b.classList.toggle('active',b===button));document.getElementById('slides').hidden=button.dataset.mode!=='slides';document.getElementById('report').hidden=button.dataset.mode!=='report';}));</script></body></html>`;
}
