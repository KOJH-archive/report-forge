import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { DOMParser } from '@xmldom/xmldom';
import { ROOT, digest } from '../src/core.mjs';
import { readProject, buildReport, prepareReport } from '../src/build.mjs';
import { inspectTemplate, importTemplate, loadTemplate } from '../src/template.mjs';
import { createStudioServer } from '../src/server.mjs';
import { companyFixture } from './company-fixture.mjs';
const sample = await readProject(path.join(ROOT, 'examples/demo/report.json'));
const bytes = await companyFixture();
async function sandbox(t) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'report-forge-template-'));
  t.after(async () => { assert.equal(path.dirname(path.resolve(dir)), path.resolve(os.tmpdir())); assert.ok(path.basename(dir).startsWith('report-forge-template-')); await fs.rm(dir, { recursive: true, force: true }); });
  await fs.cp(sample.baseDir, dir, { recursive: true });
  await fs.writeFile(path.join(dir, 'company.pptx'), bytes); return dir;
}
const fresh = () => ({ ...structuredClone(sample.report), template: { file: 'company.pptx' } });
test('company import discovers dimensions, layouts, slots, charts and tables', async t => {
  const dir = await sandbox(t), info = await importTemplate(path.join(dir, 'company.pptx'), path.join(dir, 'design'));
  assert.equal(info.width, 10); assert.equal(info.height, 7.5); assert.equal(info.slides.length, 6);
  assert.equal(info.layouts.flow, 2); assert.equal(info.layouts.metrics, 3); assert.equal(info.layouts.chart, 4);
  assert.equal(info.slides[4].shapes.find(s => s.type === 'table').columns, 4);
  assert.equal(info.slides[3].shapes.find(s => s.type === 'chart').chartType, 'line');
  await assert.rejects(importTemplate(path.join(dir, 'company.pptx'), path.join(dir, 'design')), /새 템플릿/);
  await assert.rejects(inspectTemplate(Buffer.from('not pptx')), /올바른 PPTX/);
});
test('native company export preserves brand geometry and assets, replaces old content and editable data', async t => {
  const dir = await sandbox(t), report = fresh(), result = await buildReport(report, dir, path.join(dir, 'output'));
  const before = await JSZip.loadAsync(bytes), after = await JSZip.loadAsync(await fs.readFile(path.join(result.outDir, 'presentation.pptx')));
  assert.equal(result.manifest.template.width, 10); assert.ok(result.pages.length > 10);
  for (const [name, entry] of Object.entries(before.files)) if (!entry.dir && /ppt\/(media|slideMasters|slideLayouts|theme)\//.test(name)) {
    assert.ok(after.file(name), `Missing brand asset: ${name}`);
    assert.equal(digest(await after.file(name).async('nodebuffer')), digest(await entry.async('nodebuffer')), name);
  }
  const originalCover = new DOMParser().parseFromString(await before.file('ppt/slides/slide1.xml').async('string'), 'application/xml');
  const cover = new DOMParser().parseFromString(await after.file('ppt/slides/slide1.xml').async('string'), 'application/xml');
  assert.equal(cover.getElementsByTagName('a:xfrm')[0].toString(), originalCover.getElementsByTagName('a:xfrm')[0].toString());
  const allText = (await Promise.all(Object.entries(after.files).filter(([n, e]) => !e.dir && n.endsWith('.xml')).map(([, e]) => e.async('string')))).join('\n');
  assert.ok(!allText.includes('OLD_PRIVATE')); assert.ok(!allText.includes('OLD_CELL')); assert.ok(!allText.includes('OLD_SERIES')); assert.ok(!allText.includes('OLD_CATEGORY')); assert.ok(!allText.includes('OLD_LABEL')); assert.ok(!allText.includes('999'));
  assert.match(allText, /873\s*백만원/); assert.ok(allText.includes('006D77')); assert.ok(allText.includes('E29578'));
  const chartFiles = Object.keys(after.files).filter(n => /rfassets\/.+\.xml$/.test(n));
  const chart = (await Promise.all(chartFiles.map(n => after.file(n).async('string')))).find(s => s.includes('c:lineChart'));
  assert.ok(chart.includes('<c:v>178</c:v>')); assert.doesNotMatch(chart, /<c:(min|max)(?:\s|\/|>)/);
  const workbookName = Object.keys(after.files).find(n => n.startsWith('ppt/rfassets/') && n.endsWith('.xlsx'));
  const workbook = new ExcelJS.Workbook(); await workbook.xlsx.load(await after.file(workbookName).async('nodebuffer'));
  assert.equal(workbook.worksheets[0].getCell('B7').value, 178);
  assert.equal((allText.match(/<p:notes\b/g) ?? []).length, result.pages.length);
  const tablePages = result.pages.filter(p => p.section?.type === 'table');
  assert.equal(tablePages.flatMap(p => p.tablePage.rows).length, 6);
  for (const p of tablePages) { const s = await after.file(`ppt/slides/slide${p.number}.xml`).async('string'); assert.ok(s.includes('006D77')); assert.ok(s.includes('월')); }
  const word = await JSZip.loadAsync(await fs.readFile(path.join(result.outDir, 'report.docx')));
  assert.ok((await word.file('word/document.xml').async('string')).includes('상반기 사업 성과'));
});
test('automatic inference works without role names and lets users override mistaken slots', async t => {
  const dir = await sandbox(t), zip = await JSZip.loadAsync(bytes);
  for (const name of Object.keys(zip.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))) zip.file(name, (await zip.file(name).async('string')).replace(/name="rf:[^"]+"/g, 'name="TextBox"'));
  const plain = await zip.generateAsync({ type: 'nodebuffer' }); await fs.writeFile(path.join(dir, 'plain.pptx'), plain);
  const info = await inspectTemplate(plain);
  assert.ok(Object.values(info.slides[1].roles).includes('title')); assert.ok(Object.values(info.slides[1].roles).includes('body'));
  assert.equal(info.layouts.metrics, 3); assert.equal(info.layouts.chart, 4);
  const automaticReport = fresh(); automaticReport.template.file = 'plain.pptx';
  const automatic = await buildReport(automaticReport, dir, path.join(dir, 'automatic'));
  assert.equal(automatic.manifest.template.sourceSlides[2], 3);
  const title = info.slides[1].shapes.find(s => info.slides[1].roles[s.id] === 'title');
  const loaded = await loadTemplate({ file: 'plain.pptx', roles: { 2: { [title.id]: 'keep' } } }, dir);
  assert.equal(loaded.inspection.slides[1].roles[title.id], 'keep');
  await assert.rejects(loadTemplate({ file: 'plain.pptx', roles: { 2: { 9999: 'body' } } }, dir), /텍스트 역할/);
});
test('reusing a chart slide creates independent caches and embedded workbooks', async t => {
  const dir = await sandbox(t), report = fresh();
  const original = report.sections.find(s => s.type === 'chart');
  report.sections.splice(3, 0, { ...structuredClone(original), title: '다른 데이터', series: [{ column: '오프라인', name: '독립 계열' }], body: [] });
  const result = await buildReport(report, dir, path.join(dir, 'two-charts'));
  const zip = await JSZip.loadAsync(await fs.readFile(path.join(result.outDir, 'presentation.pptx')));
  const charts = [];
  for (const name of Object.keys(zip.files).filter(n => /^ppt\/rfassets\/.*\.xml$/.test(n))) {
    const text = await zip.file(name).async('string'); if (text.includes('<c:lineChart>')) charts.push({ name, text });
  }
  assert.equal(charts.length, 2);
  assert.ok(charts[0].text.includes('온라인')); assert.ok(charts[1].text.includes('독립 계열')); assert.ok(!charts[1].text.includes('온라인'));
  const workbooks = [];
  for (const name of Object.keys(zip.files).filter(n => /^ppt\/rfassets\/.*\.xlsx$/.test(n))) {
    const w = new ExcelJS.Workbook(); await w.xlsx.load(await zip.file(name).async('nodebuffer')); workbooks.push(w.worksheets[0].getCell('B2').value);
  }
  assert.deepEqual(workbooks.sort((a, b) => a - b), [90, 120]);
});
test('missing layouts, incompatible chart type, crowded slots and template traversal fail before output writes', async t => {
  const dir = await sandbox(t);
  const noBody = fresh(); noBody.template.layouts = { flow: 1 }; await assert.rejects(prepareReport(noBody, dir), /본문 텍스트/);
  const wrong = fresh(); wrong.sections.find(s => s.type === 'chart').chartType = 'bar';
  await assert.rejects(buildReport(wrong, dir, path.join(dir, 'bad')), /차트 형태/);
  assert.equal(await fs.stat(path.join(dir, 'bad')).catch(() => null), null);
  const overflow = fresh(); overflow.template.layouts = { flow: 3 }; await assert.rejects(prepareReport(overflow, dir), /본문 텍스트/);
  const unsafe = fresh(); unsafe.template.file = '../company.pptx'; await assert.rejects(prepareReport(unsafe, dir), /밖/);
});
test('studio accepts local PPTX upload and exports after inspection with authentication', async t => {
  const dir = await sandbox(t), server = await createStudioServer(path.join(dir, 'report.json'), { outputRoot: path.join(dir, 'jobs') });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve)); t.after(() => new Promise(resolve => server.close(resolve)));
  const url = `http://127.0.0.1:${server.address().port}`, project = await (await fetch(url + '/api/project')).json();
  const send = (route, payload, token = project.token) => fetch(url + route, { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Report-Forge-Token': token }, body: JSON.stringify(payload) });
  assert.equal((await send('/api/template/import', { name: 'company.pptx', data: bytes.toString('base64') }, 'bad')).status, 403);
  const uploaded = await (await send('/api/template/import', { name: 'company.pptx', data: bytes.toString('base64') })).json();
  assert.equal(uploaded.design.slides.length, 6);
  const report = { ...project.report, template: uploaded.template };
  assert.equal((await send('/api/template/inspect', { report })).status, 200);
  const preview = await (await send('/api/preview', { report })).json(); assert.equal(preview.companyDesign, true);
  const response = await send('/api/build', { report }); const result = await response.json(); assert.equal(response.status, 200, JSON.stringify(result));
  const deck = await fetch(url + result.files.find(f => f.name === 'presentation.pptx').url); assert.equal(deck.status, 200);
  const generated = await inspectTemplate(Buffer.from(await deck.arrayBuffer())); assert.equal(generated.width, 10);
  assert.equal((await send('/api/template/import', { name: 'bad.pptx', data: '!!!' })).status, 400);
});
