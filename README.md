# TechnoPan

Цель: автоматизировать получение спецификации из проекта раскладки панелей.

На текущем этапе реализован MVP пайплайн:

1. На вход подаётся файл проекта в DXF.
2. Система извлекает блоки панелей и их атрибуты и собирает спецификацию в Excel (`.xlsx`).

## Графический интерфейс (GUI)

Для запуска удобного графического интерфейса выполните команду:

```powershell
python run_gui.py
```

Или через модуль:

```powershell
python -m technopan_spec gui
```

В интерфейсе можно выбрать файл (DWG/DXF), конфигурацию и папку для сохранения результата.

## Приложение для Windows (EXE)

Вы можете собрать автономное приложение для проектировщиков, не требующее установки Python.

1. Установите зависимости для сборки:
```powershell
pip install pyinstaller
```

2. Запустите сборку:
```powershell
pyinstaller --noconfirm --onedir --windowed --name "TechnoPanSpec" --add-data "path\to\site-packages\customtkinter;customtkinter/" --add-data "configs;configs/" run_gui.py
```
*(Замените путь к customtkinter на актуальный, см. `pip show customtkinter`)*.

3. Готовое приложение будет в папке `dist/TechnoPanSpec`. Скопируйте эту папку проектировщику.
   - Внутри папки есть `TechnoPanSpec.exe` для запуска.
   - Папка `configs` рядом с exe позволяет редактировать настройки.

## Установка

```powershell
python -m pip install -r requirements.txt
```

## Быстрый старт

Проверить, какие блоки и какие теги атрибутов есть в DXF:

```powershell
python -m technopan_spec inspect path\to\project.dxf
```

Настроить соответствие блоков и атрибутов в `configs/default.yml`, затем сгенерировать спецификацию:

```powershell
python -m technopan_spec generate path\to\project.dxf -c configs\default.yml -o out\spec.xlsx
```

## Режим по размерам (DIMENSION)

Если в проекте **нет блоков с атрибутами**, а количество панелей считается по объектам `DIMENSION` на заданных слоях и есть маркеры в `TEXT` (например, `в 346`), можно включить режим `dimension_extraction` в YAML-конфиге.

Пример конфига: `configs/abk_dimensions.yml`

## DWG

DWG можно обрабатывать двумя способами:

1. Экспортировать DWG в DXF (AutoCAD / ODA File Converter) и подавать DXF.
2. Установить ODA File Converter и подавать DWG напрямую (через `ezdxf.addons.odafc`).

Проверка доступности конвертера:

```powershell
python -m technopan_spec odafc-check
```

Чтобы включить чтение DWG напрямую, установите ODA File Converter и задайте переменную окружения `ODA_FILE_CONVERTER_EXE` (путь к `ODAFileConverter.exe`), например:

```powershell
$env:ODA_FILE_CONVERTER_EXE = "C:\Program Files\ODA\ODAFileConverter\ODAFileConverter.exe"
```

Чтобы сохранить переменную в системе (для новых терминалов), используйте:

```powershell
setx ODA_FILE_CONVERTER_EXE "C:\Program Files\ODA\ODAFileConverter\ODAFileConverter.exe"
```

