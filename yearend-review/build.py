"""
한 파일 설치본 만들기: python3 build.py
  dist/Code.gs          스크립트 9개를 하나로 합친 파일
  dist/App.html         웹앱 화면 (그대로 복사)
  dist/appsscript.json
  dist/demo.html        예시 데이터로 바로 열어보는 화면 (App.html + demo/Mock.js)
  dist/install.html     설치 코드 복사 페이지 (3개 파일을 차례로 복사)
"""
import json, os, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
ORDER = ['Config', 'Setup', 'People', 'Events', 'Notice', 'Rules', 'Files', 'Admin', 'WebApp']

def read(name):
    with open(os.path.join(HERE, name), encoding='utf-8') as f:
        return f.read()

def write(name, text):
    with open(os.path.join(HERE, 'dist', name), 'w', encoding='utf-8') as f:
        f.write(text)

parts = [
    '/**\n'
    ' * 연말정산 검토 관리 - 한 파일 설치본 (build.py로 자동 생성, 직접 고치지 말고 원본 .gs 파일을 고친 뒤 다시 생성)\n'
    ' * Apps Script의 Code.gs 내용을 전부 지우고 이 파일 전체를 붙여넣으세요.\n'
    ' */\n'
]
for name in ORDER:
    parts.append(f'\n// ===================================================================\n// {name}.gs\n// ===================================================================\n\n')
    parts.append(read(name + '.gs').rstrip() + '\n')
code = ''.join(parts)

os.makedirs(os.path.join(HERE, 'dist'), exist_ok=True)
write('Code.gs', code)
for name in ['App.html', 'appsscript.json']:
    shutil.copyfile(os.path.join(HERE, name), os.path.join(HERE, 'dist', name))

app = read('App.html')
assert '<!-- MOCK -->' in app
write('demo.html', app.replace('<!-- MOCK -->', '<script>\n' + read('demo/Mock.js') + '</script>'))

files = []
for file, name, ext, kind, text in [
    ('Code.gs', 'Code', '.gs', 'script', code),
    ('App.html', 'App', '.html', 'html', app),
    ('appsscript.json', 'appsscript', '.json', 'json', read('appsscript.json')),
]:
    files.append({'name': name, 'ext': ext, 'file': file, 'kind': kind, 'code': text,
                  'lines': text.count('\n') + 1, 'size': len(text.encode('utf-8'))})
# </script> 가 JSON 안에 있으면 페이지가 깨지므로 < 를 이스케이프
payload = json.dumps(files, ensure_ascii=False).replace('<', '\\u003c')
write('install.html', read('demo/install_template.html').replace('__FILES__', payload))
print('dist/ 생성 완료')
