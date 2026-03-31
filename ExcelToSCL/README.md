# Excel to TIA Portal SCL Code Generator

Генератор кода для TIA Portal (SCL) на основе таблиц Excel с системой ввода-вывода.

## Описание

Программа парсит Excel файлы с разметкой системы ввода-вывода (листы ШС, ШСАУ) и генерирует готовый код SCL для TIA Portal в формате, аналогичном приведенным примерам.

## Возможности

- **Парсинг Excel файлов** с формулами и ссылками на ячейки
- **Автоматическое определение типов устройств** по названиям сигналов
- **Генерация регионов (REGION)** для различных типов оборудования:
  - Doliv (доливы)
  - Temperature (нагревы/охлаждения)
  - Cover (крышки ванн)
  - Jr (жироуловители)
  - Mixer (мешалки/барботажи)
  - Filtr (фильтры)
  - SafetyBar (барьеры безопасности)
  - И другие...

- **Генерация кода для модулей ввода/вывода**:
  - AI (аналоговые входы)
  - AO (аналоговые выходы)
  - DI (цифровые входы)
  - DO (цифровые выходы)

- **Поддержка формул Excel** вида `=$A$7`

## Установка

Требуется Python 3.8+ и библиотека openpyxl:

```bash
pip install openpyxl
```

## Использование

### Базовое использование

```bash
python excel_to_scl.py "Система ввода-вывода.xlsx" -o generated_code.txt
```

### Параметры командной строки

```
positional arguments:
  excel_file           Путь к Excel файлу с системой ввода-вывода

optional arguments:
  -h, --help           показать справку
  -o OUTPUT, --output OUTPUT
                       Путь к выходному файлу (по умолчанию: generated_code.txt)
```

### Примеры

```bash
# Генерация кода с именем файла по умолчанию
python excel_to_scl.py "Система ввода-вывода.xlsx"

# Генерация кода с указанием выходного файла
python excel_to_scl.py "Система ввода-вывода.xlsx" -o my_project_code.txt

# Генерация кода из другого листа
python excel_to_scl.py "Спецификация.xlsx" -o spec_code.txt
```

## Структура входного файла

Программа ожидает Excel файл со следующими листами:

### Лист "ШС" (шкаф силовой)

| шкаф | № п/п | обозн. | тип | адрес | № вх. | Наименование сигнала | Тех. поз. | № п/п основ. поз. | № п/п вспомогат. поз. | ... |
|------|-------|--------|-----|-------|-------|---------------------|-----------|------------------|---------------------|-----|
| ШС1 | 1 | А14 | DO | 181 | 1 | Температура | Нагрев | 2 | 2 | ... |

### Лист "ШСАУ" (шкаф автоматики)

Аналогичная структура для сигналов автоматики.

## Формат выходного файла

Программа генерирует текстовый файл в формате SCL для TIA Portal:

```scl
// Сгенерировано из Система ввода-вывода.xlsx
// Дата: 31.03.2026 11:47

// Установленное количество устройств
"Options".Count.MaxDoliv := 12;
"Options".Count.MaxTemperature := 11;
"Options".Count.MaxCover := 17;
...

REGION Doliv
(* Доливы
CfgPlace        номер ванны
CfgType         тип
CfgCascadeVann  для каскадных доливов указываем номер второй ванны каскада
*)
    #iCnt := 1;
    WHILE #iCnt <= "Options".Count.MaxDoliv DO
        "Doliv".Dev[#iCnt].CfgAutoLevel := TRUE;
        #iCnt := #iCnt + 1;
    END_WHILE;
    
    "Doliv".Dev[1].CfgPlace := 2;
    "Doliv".Dev[1].CfgType := 23;
    
    "Doliv".Dev[2].CfgPlace := 4;
    "Doliv".Dev[2].CfgType := 15;
    "Doliv".Dev[2].CfgCascadeVann := 3;
    ...

END_REGION

REGION Module1
    // А24. Адрес 141
    "Vann".Dev[2].Temperature := "MapAin".Module[1].Channel[1];
    "Vann".Dev[3].Temperature := "MapAin".Module[1].Channel[2];
    ...
END_REGION
```

## Расширение функциональности

### Добавление новых типов устройств

Для добавления поддержки нового типа оборудования отредактируйте словарь `EQUIPMENT_TYPE_MAP` в классе `ExcelToSCLConverter`:

```python
EQUIPMENT_TYPE_MAP = {
    'Температура': 'Temperature',
    'НовыйТип': 'NewRegion',
    ...
}
```

Затем добавьте метод генерации кода для нового региона:

```python
elif region_name == 'NewRegion':
    lines.extend([
        '(* Новый регион *)',
    ])
    for dev in devices:
        lines.append(f'    "NewRegion".Dev[{dev.device_num}].CfgPlace := {dev.place};')
```

### Настройка параметров по умолчанию

Измените словарь `DEFAULT_PARAMS`:

```python
DEFAULT_PARAMS = {
    'Doliv': {'CfgAutoLevel': True},
    'Mixer': {'CfgPnevmo': True},
    'NewRegion': {'SomeParam': 'DefaultValue'},
    ...
}
```

## Структура проекта

```
ExcelToSCL/
├── excel_to_scl.py      # Основной скрипт генератора
├── README.md            # Документация
└── output/              # Папка для выходных файлов (опционально)
```

## Требования

- Python 3.8+
- openpyxl >= 3.0.0

## Лицензия

Свободное использование без ограничений.

## Контакты

Для вопросов и предложений обращайтесь к разработчику.
