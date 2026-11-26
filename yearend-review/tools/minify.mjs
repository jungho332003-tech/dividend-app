// dist/App.html 압축: <style> 안은 csso, 앱 <script> 안은 terser. HTML 구조와 템플릿(<?!= boot ?>)은 건드리지 않는다.
// 쓰는 법: node tools/minify.mjs <입력 html> <출력 html>   (패키지: npm install, package.json 참고)
import fs from 'fs';
import { minify as minifyJs } from 'terser';
import { minify as minifyCss } from 'csso';

const [, , src, dst] = process.argv;
let html = fs.readFileSync(src, 'utf8');

html = html.replace(/<style>([\s\S]*?)<\/style>/, (m, css) => `<style>${minifyCss(css).css}</style>`);

// 앱 스크립트: 속성 없는 <script> 중 가장 큰 것. 맨 위 이름(화면의 onclick에서 부름)은 바꾸지 않는다.
const scripts = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)];
const main = scripts.sort((a, b) => b[1].length - a[1].length)[0];
const out = await minifyJs(main[1], { toplevel: false, compress: { passes: 2 }, mangle: true, format: { comments: false } });
if (/<\/script/i.test(out.code)) throw new Error('압축 결과에 </script>가 생겼습니다');
// 압축기가 '\uFFFD'(깨진 글자 표시)를 실제 글자로 바꾸므로 다시 이스케이프로 되돌린다 (문자열·정규식 안에서만 쓰임)
out.code = out.code.replace(/\uFFFD/g, '\\ufffd');
html = html.replace(main[0], () => `<script>${out.code}</script>`);

fs.writeFileSync(dst, html);
console.log(`${src} ${fs.statSync(src).size} → ${dst} ${Buffer.byteLength(html)} bytes`);
