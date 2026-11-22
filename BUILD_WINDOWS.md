# Reindeer Ledger v0.6.0 RC — Windows 11 installer

## Для бухгалтера

После сборки разработчик передаёт **ReindeerLedger_Setup_v0.6.0_RC.exe**. Дважды щёлкнуть по установщику, выбрать «Установить». Опция «Создать ярлык на рабочем столе» включена по умолчанию. После установки включён запуск приложения. Python, pip, интернет и отдельные библиотеки не требуются.

## Для разработчика — GitHub Actions

1. Создать репозиторий и поместить **содержимое этой папки** в корень репозитория, включая `.github/workflows/build-windows.yml`.
2. Actions → `Build Windows installer (v0.6.0 RC)` → `Run workflow`.
3. Дождаться успешного выполнения всех шагов, скачать артефакт `ReindeerLedger-v0.6.0-RC-Windows-Setup`.
4. Извлечь EXE из артефакта и **проверить на настоящей Windows 11**, прежде чем передавать бухгалтеру.

## Ручная сборка на Windows 11

Установить Python 3.12 и Inno Setup 6 (только на машине разработчика). В PowerShell из корня проекта:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-build.txt
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --onedir --windowed --name ReindeerLedger --icon assets/app.ico --add-data "assets/kvramis_logo.jpg;assets" --add-data "assets/accounting_template.xlsx;assets" main.py
& "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe" installer\ReindeerLedger.iss
```

Итог: `release\ReindeerLedger_Setup_v0.6.0_RC.exe`.

## Контроль перед релизом

- Установить на чистой Windows 11 без Python, проверить ярлык рабочего стола и запуск после установки.
- Открыть XLSX 17 ноября; подтвердить 102 головы, 2542 кг и закупные 26 голов, 867 кг.
- Проверить экспорт Excel (четыре листа) и PDF, резервное сохранение и восстановление.
- Убрать выбранный XLSX, убедиться, что исходный файл на диске не удалён.
- Проверить разделы «Обучение», «О программе», ссылки и обе темы.
- Проверить деинсталляцию и сохранность пользовательских данных.

**Ограничения:** этот ZIP содержит проект и сценарий сборки, а не уже скомпилированный Windows EXE. Неподписанный установщик может вызвать SmartScreen. Приложение работает офлайн; открытие сайта и почтового клиента требует соответствующих внешних программ и сети.
