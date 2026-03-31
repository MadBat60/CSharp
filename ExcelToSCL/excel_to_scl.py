#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel to TIA Portal SCL Code Generator

Программа для парсинга таблиц Excel с разметкой системы ввода-вывода
и генерации готового кода для TIA Portal (SCL).
"""

import openpyxl
import re
import os
from typing import Dict, List, Any, Optional
from collections import defaultdict


# Константы для типов оборудования
EQUIPMENT_TYPES = {
    'DOZING': 'Дозирование',
    'TEMP': 'Температура',
    'COVER': 'Крышка',
    'MIXER': 'Мешалка',
    'FILTER': 'Фильтр',
    'LINE_DEV': 'Линейное устройство',
    'JR': 'Jr (Реактор)',
    'CART': 'Cart (Тележка)',
    'VALVE': 'Клапан',
    'PUMP': 'Насос',
    'SENSOR': 'Датчик'
}


class ExcelToSCLConverter:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.workbook = None
        self.devices = {}  # Словарь устройств: ключ - имя устройства
        self.modules = []  # Список всех модулей (каналов)
        self.errors = []
        
    def load_workbook(self):
        """Загружает книгу и инициирует парсинг"""
        try:
            self.workbook = openpyxl.load_workbook(self.file_path, data_only=True)
            self._parse_all_sheets()
            return True
        except Exception as e:
            self.errors.append(f"Ошибка загрузки файла: {str(e)}")
            return False

    def _parse_all_sheets(self):
        """Проходит по всем листам и ищет таблицы с оборудованием"""
        if not self.workbook:
            return

        # Нас интересуют только листы ШС и ШСАУ
        target_sheets = ['ШС', 'ШСАУ']
        
        for sheet_name in target_sheets:
            if sheet_name in self.workbook.sheetnames:
                sheet = self.workbook[sheet_name]
                self._parse_sheet_data(sheet, sheet_name)

    def _parse_sheet_data(self, sheet, sheet_name: str):
        """
        Парсер листа ШС/ШСАУ.
        Структура таблицы (начиная с строки 6-7):
        A: шкаф | B: № п/п | C: обозн. | D: тип | E: адрес | F: № вх. | G: Device | H: Наименование сигнала | ...
        """
        max_row = sheet.max_row
        
        # Проходим по строкам данных (пропускаем заголовки до строки 6)
        for row_idx in range(6, max_row + 1):
            row_values = [cell.value for cell in sheet[row_idx]]
            
            # Пропускаем пустые строки
            if not any(row_values):
                continue
            
            # Извлекаем ключевые поля
            # Индексы колонок (0-based): A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9
            cabinet = row_values[0] if len(row_values) > 0 else None  # Шкаф
            device_name_raw = row_values[6] if len(row_values) > 7 else None  # Device (колонка G)
            signal_name = row_values[7] if len(row_values) > 7 else None  # Наименование сигнала (колонка H)
            tech_pos = row_values[8] if len(row_values) > 8 else None  # Тех. поз. (колонка I)
            main_pos = row_values[9] if len(row_values) > 9 else None  # № п/п основ. поз. (колонка J)
            
            sig_type = row_values[3] if len(row_values) > 3 else None  # Тип сигнала (DI, DO, AI, AO)
            address = row_values[4] if len(row_values) > 4 else None  # Адрес
            channel_num = row_values[5] if len(row_values) > 5 else None  # № входа
            
            # Пропускаем если нет устройства или сигнала
            if not device_name_raw and not signal_name:
                continue
            
            # Определяем устройство по комбинации Device + Тех. поз.
            if device_name_raw:
                device_key = self._create_device_key(device_name_raw, tech_pos, sheet_name)
            else:
                # Пытаемся использовать последнее известное устройство или создаем новое
                device_key = self._infer_device_key(signal_name, tech_pos, sheet_name)
            
            # Создаем или получаем устройство
            if device_key not in self.devices:
                dev_type = self._detect_device_type(device_name_raw, signal_name, tech_pos)
                dev_name = self._format_device_name(dev_type, tech_pos, main_pos)
                
                self.devices[device_key] = {
                    'type': dev_type,
                    'name': dev_name,
                    'sheet': sheet_name,
                    'cabinet': cabinet,
                    'tech_pos': tech_pos,
                    'main_pos': main_pos,
                    'signals': {}
                }
            
            # Добавляем сигнал к устройству
            if signal_name:
                signal_key = f"{sig_type}_{channel_num}" if channel_num else signal_name
                self.devices[device_key]['signals'][signal_key] = {
                    'name': signal_name,
                    'type': sig_type,
                    'address': address,
                    'channel': channel_num,
                    'cabinet': cabinet
                }

    def _create_device_key(self, device_raw: Any, tech_pos: Any, sheet: str) -> str:
        """Создает уникальный ключ для устройства"""
        d_name = str(device_raw).strip() if device_raw else "Unknown"
        t_pos = str(tech_pos).strip() if tech_pos else ""
        return f"{sheet}_{d_name}_{t_pos}"

    def _infer_device_key(self, signal_name: Any, tech_pos: Any, sheet: str) -> str:
        """Пытается определить устройство по сигналу"""
        t_pos = str(tech_pos).strip() if tech_pos else "0"
        return f"{sheet}_Group_{t_pos}"

    def _detect_device_type(self, device_raw: Any, signal_name: Any, tech_pos: Any) -> str:
        """Определяет тип устройства по названию и сигналам"""
        text = ""
        if device_raw:
            text += str(device_raw).lower()
        if signal_name:
            text += " " + str(signal_name).lower()
        
        if re.search(r'долив|дозир|dosing', text):
            return 'DOZING'
        elif re.search(r'температур|temp|нагрев|охлажд', text):
            return 'TEMP'
        elif re.search(r'крышк|cover|cap|лоток', text):
            return 'COVER'
        elif re.search(r'мешалк|mixer|agit|перемешив', text):
            return 'MIXER'
        elif re.search(r'фильтр|filter', text):
            return 'FILTER'
        elif re.search(r'насос|pump', text):
            return 'PUMP'
        elif re.search(r'клапан|valve|кран', text):
            return 'VALVE'
        elif re.search(r'тележк|cart|каретк|перемещени', text):
            return 'CART'
        elif re.search(r'jr\s*|^j\s*\d+', text):
            return 'JR'
        else:
            return 'SENSOR'

    def _format_device_name(self, dev_type: str, tech_pos: Any, main_pos: Any) -> str:
        """Формирует читаемое имя устройства"""
        type_names = {
            'DOZING': 'Дополив',
            'TEMP': 'Температура',
            'COVER': 'Крышка',
            'MIXER': 'Мешалка',
            'FILTER': 'Фильтр',
            'JR': 'Jr',
            'CART': 'Cart',
            'PUMP': 'Насос',
            'VALVE': 'Клапан',
            'SENSOR': 'Датчик'
        }
        
        base_name = type_names.get(dev_type, dev_type)
        num = str(main_pos).strip() if main_pos else (str(tech_pos).strip() if tech_pos else "?")
        return f"{base_name} {num}"

    def analyze_workbook(self) -> Dict:
        """Возвращает аналитику по книге"""
        stats = {
            'total_devices': len(self.devices),
            'total_signals': sum(len(d['signals']) for d in self.devices.values()),
            'by_type': defaultdict(int),
            'sheets': list(set(d['sheet'] for d in self.devices.values())),
            'errors': self.errors
        }
        
        for dev in self.devices.values():
            stats['by_type'][dev['type']] += 1
            
        return stats

    def generate_full_code(self) -> str:
        """Генерирует полный SCL код"""
        output = []
        output.append("// === ГЕНЕРАЦИЯ КОДА TIA PORTAL (SCL) ===")
        output.append(f"// Источник: {os.path.basename(self.file_path)}")
        output.append(f"// Всего устройств: {len(self.devices)}")
        output.append(f"// Всего сигналов: {sum(len(d['signals']) for d in self.devices.values())}")
        output.append("")
        
        # Группируем по типам
        grouped = defaultdict(list)
        for dev in self.devices.values():
            grouped[dev['type']].append(dev)
            
        for dev_type, devs in grouped.items():
            output.append(f"REGION \"{EQUIPMENT_TYPES.get(dev_type, dev_type)}\"")
            output.append(f"// Количество: {len(devs)}")
            output.append("")
            
            for dev in sorted(devs, key=lambda x: x['name']):
                output.append(f"    // ============================================")
                output.append(f"    // {dev['name']} (Тех. поз.: {dev['tech_pos']}, Шкаф: {dev['cabinet']})")
                output.append(f"    // ============================================")
                
                signals = sorted(dev['signals'].items(), key=lambda x: (x[1]['type'] or '', str(x[1]['channel'] or '')))
                
                for sig_key, sig_data in signals:
                    sig_name = sig_data['name']
                    sig_type = sig_data['type']
                    addr = sig_data['address']
                    channel = sig_data['channel']
                    
                    output.append(f"    // {sig_type} Channel {channel}: {sig_name}")
                    if addr:
                        output.append(f"    //   Address: {addr}")
                
                output.append("")
            
            output.append("END_REGION")
            output.append("")
            
        return "\n".join(output)

    def get_statistics(self) -> Dict[str, int]:
        """Возвращает детальную статистику для UI"""
        stats = {
            'ДОЛИВЫ': 0,
            'ТЕМПЕРАТУРЫ': 0,
            'КРЫШКИ': 0,
            'МЕШАЛКИ': 0,
            'ФИЛЬТРЫ': 0,
            'НАСОСЫ': 0,
            'КЛАПАНЫ': 0,
            'ТЕЛЕЖКИ': 0,
            'ДРУГОЕ': 0
        }
        
        for dev in self.devices.values():
            t = dev['type']
            if t == 'DOZING': stats['ДОЛИВЫ'] += 1
            elif t == 'TEMP': stats['ТЕМПЕРАТУРЫ'] += 1
            elif t == 'COVER': stats['КРЫШКИ'] += 1
            elif t == 'MIXER': stats['МЕШАЛКИ'] += 1
            elif t == 'FILTER': stats['ФИЛЬТРЫ'] += 1
            elif t == 'PUMP': stats['НАСОСЫ'] += 1
            elif t == 'VALVE': stats['КЛАПАНЫ'] += 1
            elif t == 'CART': stats['ТЕЛЕЖКИ'] += 1
            else: stats['ДРУГОЕ'] += 1
            
        stats['ВСЕГО'] = sum(stats.values())
        return stats


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Использование: python excel_to_scl.py <путь_к_файлу.xlsx> [выходной_файл.txt]")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else "generated_code.txt"
    
    converter = ExcelToSCLConverter(input_file)
    
    if converter.load_workbook():
        code = converter.generate_full_code()
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(code)
        
        stats = converter.get_statistics()
        print(f"Код сгенерирован и сохранен в {output_file}")
        print(f"\nСтатистика:")
        for k, v in stats.items():
            print(f"  {k}: {v}")
    else:
        print(f"Ошибка загрузки файла: {converter.errors}")
        sys.exit(1)
