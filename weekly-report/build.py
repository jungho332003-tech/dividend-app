"""
한 파일 설치본 만들기: python3 build.py
  dist/Code.gs        스크립트 11개 + 예산전용 신청 화면(BudgetForm)을 하나로 합친 파일
  dist/App.html       웹앱 화면 (그대로 복사)
  dist/appsscript.json
"""
import json, os, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
ORDER = ['Config', 'Setup', 'Report', 'Automation', 'Budget', 'Calendar', 'Board', 'Rules', 'Files', 'Admin', 'WebApp']

def read(name):
    with open(os.path.join(HERE, name), encoding='utf-8') as f:
        return f.read()

parts = [
    '/**\n'
    ' * 인사팀 업무관리 - 한 파일 설치본 (build.py로 자동 생성, 직접 고치지 말고 원본 .gs 파일을 고친 뒤 다시 생성)\n'
    ' * Apps Script의 Code.gs 내용을 전부 지우고 이 파일 전체를 붙여넣으세요.\n'
    ' */\n'
]
for name in ORDER:
    parts.append(f'\n// ===================================================================\n// {name}.gs\n// ===================================================================\n\n')
    parts.append(read(name + '.gs').rstrip() + '\n')

form = json.dumps(read('BudgetForm.html'), ensure_ascii=False)
parts.append('\n// ===================================================================\n// BudgetForm.html (스프레드시트 메뉴의 예산전용 신청 화면)\n// ===================================================================\n\n')
parts.append(f'const BUDGET_FORM_HTML = {form};\n')

os.makedirs(os.path.join(HERE, 'dist'), exist_ok=True)
with open(os.path.join(HERE, 'dist', 'Code.gs'), 'w', encoding='utf-8') as f:
    f.write(''.join(parts))
for name in ['App.html', 'appsscript.json']:
    shutil.copyfile(os.path.join(HERE, name), os.path.join(HERE, 'dist', name))
print('dist/ 생성 완료')
