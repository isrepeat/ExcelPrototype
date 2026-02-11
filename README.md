# ExcelPrototype

## Настройка Codespace (алиасы + логи)

В каждом новом Codespace подключи алиас `codexlog`, чтобы терминал логировался и перезапусти shell
```bash
echo 'source /workspaces/ExcelPrototype/Codex/aliases.sh' >> ~/.bashrc
source /workspaces/ExcelPrototype/Codex/aliases.sh
```
Использование:
```bash
codexlog
```
Это запускает Codex с логированием терминала в `Codex/logs/codex-YYYY-MM-DD_HH-MM.log`.
