# merge_pallets — запуск на macOS

## Готовая программа

После сборки она лежит здесь:

- `dist/merge_pallets.app`

Для передачи на другой Mac (одним файлом):

- `dist/merge_pallets_mac.zip`

Запуск:

1. Откройте папку `dist/`
2. Дважды кликните `merge_pallets.app`

Если macOS блокирует запуск (Gatekeeper):

- Правый клик по `merge_pallets.app` → **Открыть** → **Открыть**

Если на другом Mac пишет «Приложение повреждено» или «не удаётся проверить разработчика» и не запускается:

1. Переместите `merge_pallets.app` в папку `Applications` (Программы)
2. В терминале выполните:

```bash
xattr -dr com.apple.quarantine "/Applications/merge_pallets.app"
```

После этого попробуйте открыть снова.

Если вообще «ничего не происходит» при запуске:

- Проверьте лог падения:

```bash
cat "$HOME/Library/Logs/merge_pallets/crash.log" 2>/dev/null || echo "crash.log не найден"
```

- И проверку Gatekeeper:

```bash
spctl --assess --type execute -vv "/Applications/merge_pallets.app" 2>&1
```

## Сборка на этом Mac (если нужно пересобрать)

В терминале из папки проекта:

```bash
./build_mac.command
```

Результат: `dist/merge_pallets.app`

И один файл для передачи: `dist/merge_pallets_mac.zip`

Логи при аварийном падении (если вдруг):

- `~/Library/Logs/merge_pallets/crash.log`
