# ExcelPrototype

## Настройка Codespace (алиасы + логи)

Если работаешь через расширение и оно работает нормально, этот шаг можно пропустить.
Используй его как fallback для CLI-сценария (когда расширение недоступно/нестабильно), чтобы писать локальные логи сессии в терминале.
```bash
bash .env/Codex/setup_codexlog.sh
```
Использование:
```bash
codexlog
```
Это запускает Codex с логированием терминала в `.env/Codex/logs/codex-YYYY-MM-DD_HH-MM.log`.

## Git-алиасы для репозитория

```bash
bash .env/git/setup_git_aliases.sh
bash .env/git/remove_git_aliases.sh
```

Проверить, какие алиасы активны и откуда они взялись:
```bash
git config --show-origin --get-regexp '^alias\.'
```
