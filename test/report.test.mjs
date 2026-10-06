import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { promisify } from 'node:util';
import { execFile } from 'node:child_process';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { ROOT, resolveReport, planSlides, safeInputPath, ReportError } from '../src/core.mjs';
import { readProject, buildReport } from '../src/build.mjs';
import { previewHTML } from '../src/renderers.mjs';
import { createStudioServer } from '../src/server.mjs';

const execute = promisify(execFile);
const sample = await readProject(path.join(ROOT, 'examples/demo/report.json'));
const fresh = () => structuredClone(sample.report);
async function sandbox(t) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'report-forge-test-'));
  t.after(async () => {
    const abs = path.resolve(root);
    assert.equal(path.dirname(abs), path.resolve(os.tmpdir()));
    assert.ok(path.basename(abs).startsWith('report-forge-test-'));
    await fs.rm(abs, { recursive: true, force: true });
  });
  return root;
}
test('same dataset drives metrics and narrative without mutating the input', async () => {
  const report = fresh(); const before = JSON.stringify(report);
  const model = await resolveReport(report, sample.baseDir);
  assert.equal(model.metrics.onlineTotal.value, 873);
  assert.equal(model.metrics.offlineTotal.value, 560);
  assert.equal(model.metrics.onlineGrowth.formatted, '48.3%');
  assert.ok(model.doc.sections[0].lead.includes('48.3%'));
  assert.deepEqual(model.doc.sections[1].sourceIds, ['sales-demo']);
  assert.equal(JSON.stringify(report), before);
});
test('unknown references and unknown interpolated metrics are rejected', async () => {
  const report = fresh(); report.sections[0].sourceIds.push('missing');
  await assert.rejects(resolveReport(report, sample.baseDir), /출처 missing/);
  const other = fresh(); other.sections[0].lead = '{{unavailable}}';
  await assert.rejects(resolveReport(other, sample.baseDir), /알 수 없는 지표/);
});
test('unsafe file paths and missing files are rejected before reading', async () => {
  await assert.rejects(safeInputPath(sample.baseDir, '../report.json'), /밖/);
  await assert.rejects(safeInputPath(sample.baseDir, path.join(ROOT, 'package.json')), /상대 경로/);
  await assert.rejects(safeInputPath(sample.baseDir, 'missing.csv'), /찾을 수 없습니다/);
});
test('blank numbers are rejected instead of being silently treated as zero', async t => {
  const dir = await sandbox(t); await fs.writeFile(path.join(dir, 'data.csv'), '월,값\n1월,10\n2월,\n');
  const report = { version: 1, title: '수치 검증', sources: [{ id: 's', title: '테스트' }], datasets: [{ id: 'd', file: 'data.csv', sourceIds: ['s'] }], metrics: [{ id: 'total', label: '합계', dataset: 'd', column: '값', operation: 'sum' }], sections: [{ type: 'metrics', title: '결과', metricIds: ['total'] }] };
  await assert.rejects(resolveReport(report, dir), /비어 있거나/);
});
test('growth rejects a zero baseline and tables preserve every row through pagination', async t => {
  const dir = await sandbox(t); await fs.writeFile(path.join(dir, 'rows.csv'), '순서,값\n' + Array.from({ length: 27 }, (_, i) => `${i},${i}`).join('\n'));
  const report = { version: 1, title: '페이지 분할', sources: [{ id: 's', title: '테스트' }], datasets: [{ id: 'd', file: 'rows.csv', sourceIds: ['s'] }], sections: [{ type: 'table', title: '전체 행', dataset: 'd', columns: ['순서', '값'] }] };
  const model = await resolveReport(report, dir), pages = planSlides(model).filter(p => p.kind === 'table');
  assert.equal(pages.length, 4);
  assert.deepEqual(pages.flatMap(p => p.tablePage.rows).map(r => r[0]), Array.from({ length: 27 }, (_, i) => String(i)));
  report.metrics = [{ id: 'growth', label: '변화율', dataset: 'd', column: '값', operation: 'growthPct' }];
  await assert.rejects(resolveReport(report, dir), /0이 아닌 첫 값/);
});
test('Excel uses cached formula values and rejects formulas without cached results', async t => {
  const dir = await sandbox(t); const wb = new ExcelJS.Workbook(); const sheet = wb.addWorksheet('실적');
  sheet.addRows([['월', '값'], ['1월', 10], ['2월', { formula: 'B2*2', result: 20 }]]); await wb.xlsx.writeFile(path.join(dir, 'data.xlsx'));
  const report = { version: 1, title: '엑셀 수치', sources: [{ id: 's', title: '테스트' }], datasets: [{ id: 'd', file: 'data.xlsx', sheet: '실적', sourceIds: ['s'] }], metrics: [{ id: 'total', label: '합계', dataset: 'd', column: '값', operation: 'sum' }], sections: [{ type: 'metrics', title: '결과', metricIds: ['total'] }] };
  assert.equal((await resolveReport(report, dir)).metrics.total.value, 30);
  sheet.getCell('B3').value = { formula: 'B2*2' }; await wb.xlsx.writeFile(path.join(dir, 'data.xlsx'));
  await assert.rejects(resolveReport(report, dir), /저장된 계산 결과/);
});
test('preview escapes user text', async () => {
  const report = fresh(); report.sections[6].body = ['<script>alert("unsafe")</script>'];
  const model = await resolveReport(report, sample.baseDir);
  const html = previewHTML(model);
  assert.ok(html.includes('&lt;script&gt;alert'));
  assert.ok(!html.includes('<script>alert("unsafe")'));
  assert.equal((html.match(/data-page=/g) ?? []).length, 10);
});
test('negative bar values remain negative and mismatched block properties are rejected', async t => {
  const dir = await sandbox(t); await fs.writeFile(path.join(dir, 'values.csv'), '범주,값\nA,-10\nB,20\n');
  const report = { version: 1, title: '음수 차트', sources: [{ id: 's', title: '테스트' }], datasets: [{ id: 'd', file: 'values.csv', sourceIds: ['s'] }], sections: [{ type: 'chart', title: '비교', dataset: 'd', chartType: 'bar', labelColumn: '범주', series: [{ column: '값', name: '값' }], unit: '개' }] };
  const model = await resolveReport(report, dir);
  assert.deepEqual(model.doc.sections[0].chartData[0].values, [-10, 20]);
  assert.ok(previewHTML(model).includes('값: -10 개'));
  report.sections[0].metricIds = ['unknown'];
  await assert.rejects(resolveReport(report, dir), /사용할 수 없는 항목/);
});
test('generated Office packages contain native editable charts, workbooks, tables, and source notes', async t => {
  const dir = await sandbox(t);
  const result = await buildReport(fresh(), sample.baseDir, path.join(dir, '출력 자료'));
  const deck = await JSZip.loadAsync(await fs.readFile(path.join(result.outDir, 'presentation.pptx')));
  const word = await JSZip.loadAsync(await fs.readFile(path.join(result.outDir, 'report.docx')));
  const chart = await deck.file('ppt/charts/chart1.xml').async('string');
  assert.ok(chart.includes('<c:lineChart>'));
  assert.ok(chart.includes('<c:v>178</c:v>'));
  assert.ok(Object.keys(deck.files).some(k => k.startsWith('ppt/embeddings/') && k.endsWith('.xlsx')));
  const types = await deck.file('[Content_Types].xml').async('string');
  for (const target of types.matchAll(/PartName="\/([^"]+)"/g)) assert.ok(deck.file(target[1]), `content type target ${target[1]} exists`);
  for (const filename of Object.keys(deck.files).filter(k => /^ppt\/slides\/slide\d+\.xml$/.test(k))) {
    const xml = await deck.file(filename).async('string');
    for (const anchor of xml.matchAll(/<a:tcPr[^>]*anchor="([^"]+)"/g)) assert.ok(['t', 'ctr', 'b', 'just', 'dist'].includes(anchor[1]), `valid native table anchor ${anchor[1]}`);
  }
  const chartInWord = await word.file('word/charts/chart1.xml').async('string');
  assert.equal(chartInWord, chart);
  const wordDocument = await word.file('word/document.xml').async('string');
  assert.ok(wordDocument.includes('rIdReportForgeChart1'));
  assert.ok(wordDocument.includes('<w:tbl>'));
  assert.ok(wordDocument.includes('참고 자료'));
  assert.ok(!wordDocument.includes('REPORTFORGE_NATIVE_CHART_'));
  assert.ok(!wordDocument.includes('{{'));
  assert.ok((await word.file('[Content_Types].xml').async('string')).includes('/word/charts/chart1.xml'));
  const rels = await word.file('word/charts/_rels/chart1.xml.rels').async('string');
  const targets = [...rels.matchAll(/Target="([^"]+)"/g)].map(m => m[1]);
  for (const target of targets) assert.ok(word.file(path.posix.normalize(path.posix.join('word/charts', target))), `chart target ${target} exists`);
  const notes = await deck.file('ppt/notesSlides/notesSlide4.xml').async('string');
  assert.ok(notes.includes('data/sales.csv'));
  const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(path.join(result.outDir, 'evidence.xlsx'));
  assert.equal(wb.getWorksheet('계산 지표').getCell('F2').value, 873);
  assert.equal(wb.getWorksheet('sales').rowCount, 7);
  assert.equal(result.manifest.slides, 10);
});
test('invalid builds leave existing outputs unchanged and overwrite requires an explicit option', async t => {
  const dir = await sandbox(t), out = path.join(dir, 'out'); await fs.mkdir(out); await fs.writeFile(path.join(out, 'presentation.pptx'), 'keep');
  await assert.rejects(buildReport(fresh(), sample.baseDir, out), /이미 있습니다/);
  const invalid = fresh(); invalid.sections[0].sourceIds = ['missing'];
  await assert.rejects(buildReport(invalid, sample.baseDir, out, { force: true }), /출처 missing/);
  assert.equal(await fs.readFile(path.join(out, 'presentation.pptx'), 'utf8'), 'keep');
});
test('the CLI runs outside the repository in a path containing Korean characters', async t => {
  const dir = await sandbox(t), projectDir = path.join(dir, '새 보고서');
  await execute(process.execPath, [path.join(ROOT, 'src/cli.mjs'), 'init', projectDir], { cwd: dir });
  const { stdout } = await execute(process.execPath, [path.join(ROOT, 'src/cli.mjs'), 'build', path.join(projectDir, 'report.json'), '--out', path.join(dir, '결과')], { cwd: dir });
  assert.ok(stdout.includes('자료 생성 완료'));
  assert.ok((await fs.stat(path.join(dir, '결과', 'report.docx'))).size > 1000);
});
test('theme switching retains the same numbers and section order', async () => {
  const report = fresh(); report.theme = 'editorial';
  const model = await resolveReport(report, sample.baseDir);
  assert.equal(model.theme.name, 'editorial'); assert.equal(model.metrics.onlineTotal.value, 873);
  assert.deepEqual(model.doc.sections.map(s => s.title), report.sections.map(s => s.title));
});
test('studio previews, exports files, rejects cross-site writes, and reports validation errors', async t => {
  const dir = await sandbox(t); const server = await createStudioServer(sample.filename, { outputRoot: path.join(dir, 'studio') });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  t.after(() => new Promise(resolve => server.close(resolve)));
  const base = `http://127.0.0.1:${server.address().port}`;
  const session = await (await fetch(`${base}/api/project`)).json();
  const post = (endpoint, report, overrides = {}) => fetch(`${base}${endpoint}`, { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Report-Forge-Token': session.token, ...overrides }, body: JSON.stringify({ report }) });
  assert.equal((await post('/api/build', session.report, { Origin: 'https://other.example' })).status, 403);
  const preview = await post('/api/preview', session.report); assert.equal(preview.status, 200); assert.ok((await preview.json()).html.includes('발표 자료'));
  const invalid = fresh(); invalid.sections[0].sourceIds.push('missing');
  assert.equal((await post('/api/build', invalid)).status, 400);
  const built = await post('/api/build', session.report); const builtBody = await built.json(); assert.equal(built.status, 200, JSON.stringify(builtBody));
  const files = builtBody.files;
  const downloaded = await fetch(base + files.find(f => f.name === 'presentation.pptx').url);
  assert.equal(downloaded.status, 200); assert.ok((await downloaded.arrayBuffer()).byteLength > 1000);
  assert.equal((await fetch(`${base}/artifacts/unknown/report.docx`)).status, 404);
});
