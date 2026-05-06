#!/bin/zsh
set -euo pipefail

cd "${0:A:h}"

if [[ ! -x ".venv/bin/python" ]]; then
  echo "Не найдено виртуальное окружение .venv. Создаю..."
  /usr/bin/python3 -m venv .venv
fi

".venv/bin/python" -m pip install --upgrade pip >/dev/null
".venv/bin/python" -m pip install -q pyinstaller pandas openpyxl xlrd xlwt

# Чистим предыдущие артефакты сборки
rm -rf dist/merge_pallets build/merge_pallets
rm -rf dist/merge_pallets*.app(N)
".venv/bin/python" -m PyInstaller --clean -y merge_pallets.spec

# Оставляем только .app как основной артефакт сборки
rm -rf dist/merge_pallets dist/merge_pallets.exe 2>/dev/null || true

# Иногда Finder/распаковка создаёт имя вида "merge_pallets 2.app".
# Берём самый новый *.app и приводим к ожидаемому имени.
apps=(dist/merge_pallets*.app(Nom))
latest_app=${apps[1]:-}
if [[ -z "$latest_app" ]]; then
  echo "Ошибка: не найдено .app в dist/" >&2
  exit 2
fi

if [[ "$latest_app" != "dist/merge_pallets.app" ]]; then
  rm -rf "dist/merge_pallets.app" 2>/dev/null || true
  mv "$latest_app" "dist/merge_pallets.app"
fi

# Ad-hoc подпись всего бандла (уменьшает шанс ошибок Gatekeeper вида
# "повреждено" после переноса/распаковки).
/usr/bin/codesign --force --deep --sign - "dist/merge_pallets.app" >/dev/null

# Предпроверка Gatekeeper (не гарантирует запуск на другом Mac без предупреждений,
# но помогает ловить очевидные проблемы сразу).
/usr/sbin/spctl --assess --type execute -vv "dist/merge_pallets.app" >/dev/null 2>&1 || true

# Упаковка в один файл для передачи (рекомендуемый вариант для macOS)
rm -f "dist/merge_pallets_mac.zip" 2>/dev/null || true
ditto -c -k --sequesterRsrc --keepParent "dist/merge_pallets.app" "dist/merge_pallets_mac.zip"

echo "Готово: dist/merge_pallets.app"
echo "Один файл для передачи: dist/merge_pallets_mac.zip"
