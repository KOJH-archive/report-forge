#!/usr/bin/env node
import path from 'node:path';
import fs from 'node:fs/promises';
import { ROOT, ReportError } from './core.mjs';
import { readProject, buildReport, prepareReport } from './build.mjs';
import { startStudio } from './server.mjs';
import { importTemplate } from './template.mjs';

const help = `Report Forge 0.2.0

같은 원고와 데이터로 발표 자료와 보고서를 만듭니다.

  node src/cli.mjs init <새 폴더>
  node src/cli.mjs import-template <회사.pptx> --out <새 템플릿 폴더>
  node src/cli.mjs validate <report.json>
  node src/cli.mjs build <report.json> --out <출력 폴더> [--force] [--theme <테마.json>]
  node src/cli.mjs studio <report.json> [--port 8765]

결과: presentation.pptx, report.docx, evidence.xlsx, preview.html
필수: Node.js 22 이상. AI 계정이나 Office 설치 없이 생성할 수 있습니다.
Office 파일의 실제 표시 확인에는 PowerPoint, Word 또는 LibreOffice가 필요합니다.
`;
function parseArgs(args) {
  const [command, input, ...rest] = args;
  const options = {};
  for (let i = 0; i < rest.length; i++) {
    const arg = rest[i];
    if (arg === '--force') options.force = true;
    else if (['--out', '--theme', '--port'].includes(arg)) {
      if (!rest[i + 1] || rest[i + 1].startsWith('--')) throw new ReportError(`${arg}의 값을 입력해 주세요`);
      options[arg.slice(2)] = rest[++i];
    } else throw new ReportError(`알 수 없는 옵션: ${arg}`);
  }
  return { command, input, options };
}
async function main() {
  if (process.argv.includes('--help') || process.argv.includes('-h') || process.argv.length < 3) { console.log(help); return; }
  const { command, input, options } = parseArgs(process.argv.slice(2));
  if (!input) throw new ReportError('원고 파일 또는 프로젝트 폴더를 입력해 주세요');
  if (command === 'import-template') {
    if (!options.out) throw new ReportError('--out 옵션으로 새 템플릿 폴더를 지정해 주세요');
    const design = await importTemplate(path.resolve(input), options.out);
    console.log(`회사 디자인 분석 완료: ${design.slides.length}장, ${design.width.toFixed(2)} × ${design.height.toFixed(2)} 인치\n${path.resolve(options.out)}`);
    design.warnings.forEach(w => console.log(`확인 사항: ${w}`)); return;
  }
  if (command === 'init') {
    const target = path.resolve(input);
    try { await fs.lstat(target); throw new ReportError('새 프로젝트 폴더 이름을 사용해 주세요'); } catch (e) { if (e.code !== 'ENOENT') throw e; }
    await fs.cp(path.join(ROOT, 'examples/demo'), target, { recursive: true, errorOnExist: true, force: false });
    const reportPath = path.join(target, 'report.json');
    const report = JSON.parse(await fs.readFile(reportPath, 'utf8')); delete report.$schema;
    await fs.writeFile(reportPath, JSON.stringify(report, null, 2), 'utf8');
    console.log(`새 프로젝트: ${target}\nreport.json의 원고와 data/sales.csv를 수정해 주세요`); return;
  }
  if (!['build', 'validate', 'studio'].includes(command)) throw new ReportError(`알 수 없는 작업: ${command}`);
  const project = await readProject(input);
  if (command === 'validate') {
    const { model, pages } = await prepareReport(project.report, project.baseDir, { themeFile: options.theme });
    console.log(`원고 확인 완료: ${project.report.sections.length}개 블록, ${pages.length}장, ${Object.keys(model.datasets).length}개 데이터`);
    model.warnings.forEach(w => console.log(`확인 사항: ${w}`)); return;
  }
  if (command === 'studio') {
    const port = options.port ? Number(options.port) : 8765;
    if (!Number.isInteger(port) || port < 1 || port > 65535) throw new ReportError('포트는 1~65535의 정수여야 합니다');
    await startStudio(project.filename, { port }); return;
  }
  if (!options.out) throw new ReportError('출력 폴더를 --out 옵션으로 지정해 주세요');
  const result = await buildReport(project.report, project.baseDir, options.out, { force: options.force, themeFile: options.theme, inputFile: project.filename });
  console.log(`자료 생성 완료: ${result.manifest.slides}장\n${result.outDir}`);
  result.manifest.warnings.forEach(w => console.log(`확인 사항: ${w}`));
}
main().catch(error => {
  console.error(`생성하지 못했습니다: ${error.message}`);
  for (const issue of error.issues ?? []) console.error(`  ${issue}`);
  if (process.env.REPORT_FORGE_DEBUG === '1') console.error(error.stack);
  process.exitCode = 1;
});
