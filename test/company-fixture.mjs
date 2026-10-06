import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';

// A deliberately different 4:3 brand deck. Original business text/data is fictional.
export async function companyFixture() {
  const p = new PptxGenJS(); p.defineLayout({ name: 'COMPANY_4_3', width: 10, height: 7.5 }); p.layout = 'COMPANY_4_3'; p.author = 'Fixture author';
  p.title = 'OLD_PRIVATE_TITLE'; p.theme = { headFontFace: 'Malgun Gothic', bodyFontFace: 'Malgun Gothic', lang: 'ko-KR' };
  const logo = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==';
  p.defineSlideMaster({ title: 'COMPANY', background: { color: 'FCF8EF' }, objects: [
    { rect: { x: 0, y: 0, w: .2, h: 7.5, fill: { color: '006D77' }, line: { color: '006D77' } } },
    { image: { data: logo, x: .5, y: .3, w: .25, h: .25 } },
    { text: { text: 'ACME / DESIGN TEST', options: { x: .85, y: .28, w: 5, h: .3, fontFace: 'Malgun Gothic', fontSize: 10, color: '006D77', margin: 0 } } },
  ] });
  function txt(slide, text, role, x, y, w, h, size, extra = {}) {
    slide.addText(text, { objectName: role ? `rf:${role}` : undefined, x, y, w, h, fontFace: 'Malgun Gothic', fontSize: size, color: '263D42', margin: 0, valign: 'top', ...extra });
  }
  function contentSlide(title) {
    const s = p.addSlide('COMPANY');
    s.addNotes('OLD_PRIVATE_SPEAKER_NOTES');
    txt(s, title, 'title', .6, .85, 8.8, .9, 27, { bold: true, color: '006D77' });
    txt(s, 'OLD_PRIVATE_LEAD', 'lead', .6, 1.8, 8.8, .7, 15);
    txt(s, 'OLD_PRIVATE_SOURCE', 'footer', .6, 7.0, 8.4, .32, 8);
    txt(s, '99', 'page', 9.1, 7.0, .4, .32, 9);
    return s;
  }
  let s = p.addSlide('COMPANY');
  s.addNotes('OLD_PRIVATE_COVER_NOTES');
  s.addShape(p.ShapeType.rect, { x: 7.8, y: 1.0, w: 2.2, h: 5.7, fill: { color: 'E29578' }, line: { color: 'E29578' } });
  txt(s, 'OLD_PRIVATE_TITLE', 'title', .6, 1.35, 6.8, 1.6, 30, { bold: true, color: '006D77' });
  txt(s, 'OLD_PRIVATE_SUBTITLE', 'lead', .6, 3.2, 6.8, 1.0, 18);
  txt(s, 'OLD_PRIVATE_AUTHOR', 'meta', .6, 5.4, 6.8, .6, 12);
  txt(s, 'OLD_PRIVATE_DISCLOSURE', 'disclosure', .6, 6.4, 6.8, .7, 11);
  s = contentSlide('OLD_PRIVATE_FLOW');
  txt(s, 'OLD_PRIVATE_BODY', 'body', .65, 2.65, 8.65, 3.9, 15);
  // Unmapped old prose must be erased, while geometric decoration is retained.
  txt(s, 'OLD_PRIVATE_EXTRA', null, 8.1, 6.6, 1.1, .2, 7);
  s = contentSlide('OLD_PRIVATE_METRICS');
  for (let i = 0; i < 3; i++) {
    const x = .65 + i * 3.0;
    s.addShape(p.ShapeType.rect, { x, y: 2.8, w: 2.7, h: 2.8, fill: { color: 'EDF3EF' }, line: { color: 'EDF3EF' } });
    txt(s, 'OLD_LABEL', `metric${i + 1}.label`, x + .12, 3.1, 2.45, .7, 15);
    txt(s, '999', `metric${i + 1}.value`, x + .12, 4.0, 2.45, .8, 24, { color: '006D77', bold: true });
    txt(s, 'OLD_RANGE', `metric${i + 1}.description`, x + .12, 5.0, 2.45, .5, 10);
  }
  s = contentSlide('OLD_PRIVATE_CHART');
  txt(s, '단위: OLD', 'unit', .65, 2.55, 3, .3, 10);
  s.addChart(p.ChartType.line, [{ name: 'OLD_SERIES_A', labels: ['OLD_CATEGORY_1', 'OLD_CATEGORY_2'], values: [999, 888] }, { name: 'OLD_SERIES_B', labels: ['OLD_CATEGORY_1', 'OLD_CATEGORY_2'], values: [777, 666] }], {
    x: .65, y: 2.95, w: 8.7, h: 2.85, chartColors: ['006D77', 'E29578'], showLegend: true, legendPos: 'b', showTitle: false, showValue: false,
    catAxisLabelFontFace: 'Malgun Gothic', catAxisLabelFontSize: 10, valAxisLabelFontFace: 'Malgun Gothic', valAxisLabelFontSize: 10,
    legendFontFace: 'Malgun Gothic', legendFontSize: 10, valAxisMaxVal: 1000, valAxisMinVal: 500,
    chartArea: { fill: { color: 'FCF8EF' }, border: { color: 'FCF8EF', pt: 0 } }, plotArea: { fill: { color: 'FCF8EF' } },
  });
  txt(s, 'OLD_PRIVATE_NOTE', 'note', .65, 6.05, 8.7, .7, 11);
  for (const count of [4, 3]) {
    s = contentSlide(`OLD_PRIVATE_TABLE_${count}`);
    const header = Array.from({ length: count }, (_, i) => ({ text: `OLD_HEADER_${i}`, options: { fill: '006D77', color: 'FFFFFF', bold: true } }));
    const rows = Array.from({ length: 4 }, (_, r) => Array.from({ length: count }, (_, c) => ({ text: `OLD_CELL_${r}_${c}`, options: { fill: r % 2 ? 'FCF8EF' : 'EDF3EF' } })));
    s.addTable([header, ...rows], { x: .65, y: 2.65, w: 8.7, colW: count === 3 ? [2.0, 4.2, 2.5] : [1.3, 2.3, 2.3, 2.8], rowH: [.5, .82, .82, .82, .82], fontFace: 'Malgun Gothic', fontSize: 11, color: '263D42', margin: .08, border: { color: 'B9CEC9', pt: .6 }, valign: 'middle', autoPage: false });
  }
  const bytes = Buffer.from(await p.write({ outputType: 'nodebuffer', compression: true }));
  const zip = await JSZip.loadAsync(bytes);
  const types = await zip.file('[Content_Types].xml').async('string');
  zip.file('[Content_Types].xml', types.replace(/<Override\s[^>]+\/>/g, entry => { const target = entry.match(/PartName="([^"]+)"/)?.[1]?.slice(1); return target && !zip.file(target) ? '' : entry; }));
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}
