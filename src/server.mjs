import http from 'node:http';
import fs from 'node:fs/promises';
import path from 'node:path';
import { randomUUID, timingSafeEqual } from 'node:crypto';
import { ROOT, ReportError } from './core.mjs';
import { readProject, buildReport, prepareReport, OUTPUT_FILES } from './build.mjs';
import { previewHTML } from './renderers.mjs';
import { inspectTemplate, loadTemplate } from './template.mjs';

const TYPES = { '.html': 'text/html; charset=utf-8', '.json': 'application/json; charset=utf-8', '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation', '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' };
async function bodyJSON(request, limit = 1024 * 1024) {
  let size = 0; const chunks = [];
  for await (const chunk of request) { size += chunk.length; if (size > limit) throw new ReportError('업로드 가능한 파일 크기를 초과했습니다'); chunks.push(chunk); }
  try { return JSON.parse(Buffer.concat(chunks).toString('utf8')); } catch { throw new ReportError('올바른 원고가 필요합니다'); }
}
function json(response, status, data) { response.writeHead(status, { 'Content-Type': TYPES['.json'] }); response.end(JSON.stringify(data)); }
export async function createStudioServer(specFile, { outputRoot = path.join(ROOT, '.build/studio') } = {}) {
  const project = await readProject(specFile), token = randomUUID(), jobs = new Map();
  const studio = await fs.readFile(path.join(ROOT, 'src/studio.html'), 'utf8');
  const server = http.createServer(async (request, response) => {
    response.setHeader('X-Content-Type-Options', 'nosniff'); response.setHeader('Cache-Control', 'no-store');
    try {
      const address = server.address();
      const expectedHost = `127.0.0.1:${address.port}`;
      if (![expectedHost, `localhost:${address.port}`].includes(request.headers.host)) { json(response, 403, { error: '로컬 접속만 지원합니다' }); return; }
      if (request.headers.origin && ![`http://${expectedHost}`, `http://localhost:${address.port}`].includes(request.headers.origin)) { json(response, 403, { error: '다른 사이트에서의 요청은 허용하지 않습니다' }); return; }
      const url = new URL(request.url, `http://${expectedHost}`);
      if (request.method === 'GET' && url.pathname === '/') { response.writeHead(200, { 'Content-Type': TYPES['.html'] }); response.end(studio); return; }
      if (request.method === 'GET' && url.pathname === '/api/project') { json(response, 200, { report: project.report, token }); return; }
      if (request.method === 'GET' && url.pathname.startsWith('/artifacts/')) {
        const [id, filename, ...rest] = url.pathname.slice('/artifacts/'.length).split('/');
        if (rest.length || !jobs.has(id) || !OUTPUT_FILES.includes(filename)) { json(response, 404, { error: '파일을 찾을 수 없습니다' }); return; }
        const bytes = await fs.readFile(path.join(jobs.get(id), filename));
        response.writeHead(200, { 'Content-Type': TYPES[path.extname(filename)] ?? 'application/octet-stream', ...(filename.endsWith('.html') ? {} : { 'Content-Disposition': `attachment; filename="${filename}"` }) }); response.end(bytes); return;
      }
      if (request.method !== 'POST' || !['/api/preview', '/api/build', '/api/save', '/api/template/import', '/api/template/inspect'].includes(url.pathname)) { json(response, 404, { error: '작업을 찾을 수 없습니다' }); return; }
      const supplied = Buffer.from(String(request.headers['x-report-forge-token'] ?? ''));
      if (supplied.length !== Buffer.byteLength(token) || !timingSafeEqual(supplied, Buffer.from(token))) { json(response, 403, { error: '페이지를 다시 열어 주세요' }); return; }
      const payload = await bodyJSON(request, url.pathname === '/api/template/import' ? 56 * 1024 * 1024 : 1024 * 1024);
      if (url.pathname === '/api/template/import') {
        if (typeof payload.data !== 'string' || !/^[A-Za-z0-9+/]*={0,2}$/.test(payload.data) || !String(payload.name).toLowerCase().endsWith('.pptx')) throw new ReportError('PPTX 파일을 선택해 주세요');
        const bytes = Buffer.from(payload.data, 'base64'), design = await inspectTemplate(bytes);
        const parent = path.join(project.baseDir, 'templates');
        await fs.mkdir(parent, { recursive: true });
        const realParent = await fs.realpath(parent), relativeParent = path.relative(project.baseDir, realParent);
        if (relativeParent.startsWith('..') || path.isAbsolute(relativeParent)) throw new ReportError('템플릿 저장 폴더가 프로젝트 밖을 가리킵니다');
        const relative = `templates/company-${randomUUID()}`;
        const folder = path.join(project.baseDir, relative);
        await fs.mkdir(folder, { recursive: true });
        await fs.writeFile(path.join(folder, 'source.pptx'), bytes);
        await fs.writeFile(path.join(folder, 'design.json'), JSON.stringify(design, null, 2));
        json(response, 200, { template: { file: `${relative}/source.pptx` }, design }); return;
      }
      if (url.pathname === '/api/template/inspect') {
        const template = await loadTemplate(payload.report.template, project.baseDir);
        json(response, 200, { design: template?.inspection ?? null }); return;
      }
      if (url.pathname === '/api/preview') {
        const { model, pages, template } = await prepareReport(payload.report, project.baseDir);
        json(response, 200, { html: previewHTML(model, pages, { companyDesign: Boolean(template) }), slides: pages.length, warnings: model.warnings, companyDesign: Boolean(template) }); return;
      }
      if (url.pathname === '/api/save') {
        await prepareReport(payload.report, project.baseDir);
        const tmp = `${project.filename}.reportforge-${randomUUID()}.tmp`;
        try { await fs.writeFile(tmp, JSON.stringify(payload.report, null, 2), 'utf8'); await fs.copyFile(tmp, project.filename); } finally { await fs.rm(tmp, { force: true }); }
        project.report = payload.report; json(response, 200, { saved: true }); return;
      }
      const id = randomUUID(), out = path.join(outputRoot, id);
      const result = await buildReport(payload.report, project.baseDir, out, { inputFile: project.filename });
      jobs.set(id, result.outDir);
      if (jobs.size > 30) jobs.delete(jobs.keys().next().value);
      json(response, 200, { slides: result.pages.length, warnings: result.manifest.warnings, files: OUTPUT_FILES.map(name => ({ name, url: `/artifacts/${id}/${name}` })) });
    } catch (error) { json(response, error instanceof ReportError ? 400 : 500, { error: error.message, issues: error.issues ?? [] }); }
  });
  return server;
}
export async function startStudio(specFile, { port = 8765 } = {}) {
  const server = await createStudioServer(specFile);
  await new Promise((resolve, reject) => { server.once('error', reject); server.listen(port, '127.0.0.1', resolve); });
  console.log(`자료 제작 화면: http://127.0.0.1:${server.address().port}\n종료하려면 Ctrl+C를 누르세요`);
  return server;
}
