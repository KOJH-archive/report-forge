import fs from 'node:fs/promises';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { resolveReport, planSlides, ReportError, digest } from './core.mjs';
import { pptxBuffer, docxBuffer, xlsxBuffer, previewHTML } from './renderers.mjs';
import { loadTemplate, planTemplateSlides, templatePptxBuffer } from './template.mjs';

export const OUTPUT_FILES = ['presentation.pptx', 'report.docx', 'evidence.xlsx', 'preview.html', 'resolved-report.json', 'manifest.json'];
export async function readProject(filename) {
  const full = await fs.realpath(path.resolve(filename));
  if ((await fs.stat(full)).size > 1024 * 1024) throw new ReportError('원고 파일은 1MB 이하여야 합니다');
  let report;
  try { report = JSON.parse((await fs.readFile(full, 'utf8')).replace(/^\uFEFF/, '')); } catch { throw new ReportError('원고 파일에 올바른 JSON이 필요합니다'); }
  return { report, filename: full, baseDir: path.dirname(full) };
}
export async function prepareReport(report, baseDir, options = {}) {
  const model = await resolveReport(report, baseDir, options);
  const template = await loadTemplate(report.template, baseDir);
  const pages = planTemplateSlides(model, planSlides(model), template);
  if (template) model.warnings.push(...template.inspection.warnings, '회사 PPTX의 마스터와 이미지, 보존으로 지정한 문구는 그대로 사용됩니다. 남겨둘 디자인 요소인지 확인해 주세요');
  return { model, pages, template };
}
export async function buildReport(report, baseDir, outDir, { force = false, themeFile, inputFile } = {}) {
  const { model, pages, template } = await prepareReport(report, baseDir, { themeFile });
  const out = path.resolve(outDir);
  const protectedInputs = new Set(Object.values(model.datasets).map(d => path.resolve(d.filename)));
  if (template) protectedInputs.add(template.filename);
  if (inputFile) protectedInputs.add(path.resolve(inputFile));
  for (const name of OUTPUT_FILES) {
    const target = path.join(out, name);
    if (protectedInputs.has(target)) throw new ReportError('출력 위치가 원고 또는 데이터 파일과 겹칩니다');
    let existing;
    try { existing = await fs.lstat(target); } catch (e) { if (e.code !== 'ENOENT') throw e; }
    if (existing?.isSymbolicLink() || existing?.isDirectory()) throw new ReportError(`${name}: 출력 위치에 링크나 폴더가 있습니다`);
    if (existing && !force) throw new ReportError('출력 파일이 이미 있습니다. 다른 출력 폴더를 지정하거나 --force를 사용해 주세요');
  }
  // Validate and generate everything before touching final output files.
  const standardDeck = await pptxBuffer(model, pages);
  const deck = template ? await templatePptxBuffer(model, pages, template, standardDeck) : standardDeck;
  const word = await docxBuffer(model, standardDeck);
  const excel = await xlsxBuffer(model);
  const preview = previewHTML(model, pages, { companyDesign: Boolean(template) });
  const snapshot = { ...model.doc, metrics: Object.values(model.metrics), datasets: Object.values(model.datasets).map(({ filename, ...d }) => d) };
  const buffers = {
    'presentation.pptx': deck, 'report.docx': word, 'evidence.xlsx': excel,
    'preview.html': Buffer.from(preview, 'utf8'), 'resolved-report.json': Buffer.from(JSON.stringify(snapshot, null, 2), 'utf8'),
  };
  const manifest = {
    generator: 'Report Forge', version: '0.2.0', generatedAt: new Date().toISOString(),
    title: model.doc.title, theme: model.theme.name, slides: pages.length,
    charts: model.doc.sections.filter(s => s.type === 'chart').length,
    sourceCount: model.doc.sources.length, warnings: model.warnings,
    inputSha256: digest(JSON.stringify(report)),
    template: template ? { file: report.template.file, sha256: template.inspection.sha256, width: template.inspection.width, height: template.inspection.height, sourceSlides: pages.map(p => p.templateSlide), mode: 'native OOXML clone', wordDesign: 'standard report theme' } : null,
    datasets: Object.values(model.datasets).map(d => ({ id: d.id, file: d.file, rows: d.rows.length, sha256: d.sha256 })),
    files: Object.entries(buffers).map(([name, data]) => ({ name, bytes: data.length, sha256: digest(data) })),
    verification: { packageStructure: 'covered by automated tests', officeVisualRendering: 'not verified by this generator', preview: 'HTML approximation, not an Office render' },
  };
  buffers['manifest.json'] = Buffer.from(JSON.stringify(manifest, null, 2), 'utf8');
  await fs.mkdir(out, { recursive: true });
  const realOut = await fs.realpath(out);
  const stage = path.join(realOut, `.reportforge-${randomUUID()}`);
  await fs.mkdir(stage);
  try {
    for (const [name, bytes] of Object.entries(buffers)) await fs.writeFile(path.join(stage, name), bytes);
    for (const name of OUTPUT_FILES) await fs.copyFile(path.join(stage, name), path.join(realOut, name));
  } finally {
    const relative = path.relative(realOut, stage);
    if (relative.startsWith('.reportforge-') && !relative.includes(path.sep)) await fs.rm(stage, { recursive: true, force: true });
  }
  return { model, pages, manifest, outDir: realOut };
}
