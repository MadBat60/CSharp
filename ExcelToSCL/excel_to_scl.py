#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel to TIA Portal SCL Code Generator v2.0

Программа для парсинга ТРЕХ таблиц Excel:
1. Спецификация.xlsx - технологические данные (Config_Line, Config_Transport)
2. Система ввода-вывода.xlsx - аппаратная привязка (ШС, ШСАУ)
3. Список переменных от системы ввода-вывода.xlsx - имена переменных

И генерации готового кода для TIA Portal (SCL).
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
    def __init__(self, spec_file: str = None, io_file: str = None, vars_file: str = None):
        self.spec_file = spec_file
        self.io_file = io_file
        self.vars_file = vars_file
        
        self.workbook_spec = None
        self.workbook_io = None
        self.workbook_vars = None
        
        # Данные из спецификации
        self.spec_devices = {}  # Ключ: tech_pos -> {name, type, row, ...}
        self.transport_devices = {}  # Ключ: name_dev -> {devices list}
        
        # Данные из ШС/ШСАУ
        self.io_signals = []  # Список всех сигналов с привязкой
        self.devices = {}  # Итоговый словарь устройств
        self.modules = []  # Список всех модулей (каналов)
        
        # Переменные
        self.variables_map = {}  # Маппинг сигналов на переменные
        
        self.errors = []

    def load_all_workbooks(self) -> bool:
        """Загружает все три книги"""
        success = True
        
        # Загружаем Спецификацию
        if self.spec_file and os.path.exists(self.spec_file):
            try:
                self.workbook_spec = openpyxl.load_workbook(self.spec_file, data_only=True)
                self._parse_specification()
            except Exception as e:
                self.errors.append(f"Ошибка загрузки Спецификации: {str(e)}")
                success = False
        else:
            self.errors.append("Файл Спецификации не найден")
            success = False
        
        # Загружаем Систему ввода-вывода
        if self.io_file and os.path.exists(self.io_file):
            try:
                self.workbook_io = openpyxl.load_workbook(self.io_file, data_only=True)
                self._parse_io_system()
            except Exception as e:
                self.errors.append(f"Ошибка загрузки Системы ввода-вывода: {str(e)}")
                success = False
        else:
            self.errors.append("Файл Системы ввода-вывода не найден")
            success = False
        
        # Загружаем Список переменных
        if self.vars_file and os.path.exists(self.vars_file):
            try:
                self.workbook_vars = openpyxl.load_workbook(self.vars_file, data_only=True)
                self._parse_variables()
            except Exception as e:
                self.errors.append(f"Ошибка загрузки Списка переменных: {str(e)}")
                success = False
        else:
            self.errors.append("Файл Списка переменных не найден")
            success = False
        
        # Объединяем данные
        if success:
            self._merge_data()
        
        return success

    def _parse_specification(self):
        """Парсит Спецификацию (Config_Line и Config_Transport)"""
        if not self.workbook_spec:
            return
        
        # Парсим Config_Transport - основные устройства
        if 'Config_Transport' in self.workbook_spec.sheetnames:
            ws = self.workbook_spec['Config_Transport']
            for row_idx in range(3, ws.max_row + 1):
                row = [cell.value for cell in ws[row_idx]]
                
                # Пропускаем пустые строки и разделители
                if not any(row) or row[0] == 'Общая информация:':
                    continue
                
                name = row[2] if len(row) > 2 else None  # Наименование
                name_dev = row[3] if len(row) > 3 else None  # NameDev (Cart, AO, etc.)
                tech_type = row[4] if len(row) > 4 else None  # Type
                dev_id = row[5] if len(row) > 5 else None  # Id устройства
                position = row[6] if len(row) > 6 else None  # Позиция
                row_num = row[7] if len(row) > 7 else None  # Ряд
                
                if name and name_dev:
                    key = f"{name_dev}_{position}"
                    self.transport_devices[key] = {
                        'name': name,
                        'name_dev': name_dev,
                        'type': tech_type,
                        'id': dev_id,
                        'position': position,
                        'row': row_num
                    }
        
        # Парсим Config_Line - оснащение ванн
        if 'Config_Line' in self.workbook_spec.sheetnames:
            ws = self.workbook_spec['Config_Line']
            for row_idx in range(7, ws.max_row + 1):
                row = [cell.value for cell in ws[row_idx]]
                
                num = row[1] if len(row) > 1 else None  # Num (технологический номер)
                name = row[2] if len(row) > 2 else None  # Name
                dopop = row[5] if len(row) > 5 else None  # DopOp (оснащение)
                
                if num and dopop:
                    self.spec_devices[str(num)] = {
                        'name': name,
                        'dopop': dopop
                    }

    def _parse_io_system(self):
        """Парсит Систему ввода-вывода (листы ШС и ШСАУ)"""
        if not self.workbook_io:
            return
        
        target_sheets = ['ШС', 'ШСАУ']
        
        for sheet_name in target_sheets:
            if sheet_name not in self.workbook_io.sheetnames:
                continue
            
            ws = self.workbook_io[sheet_name]
            
            # Структура: A=шкаф, B=№п/п, C=обозн., D=тип, E=адрес, F=№вх., G=Device, H=сигнал, I=тех.поз., J=осн.поз.
            for row_idx in range(7, ws.max_row + 1):
                row = [cell.value for cell in ws[row_idx]]
                
                if not any(row):
                    continue
                
                cabinet = row[0] if len(row) > 0 else None
                sig_type = row[3] if len(row) > 3 else None  # DO, DI, AI, AO
                address = row[4] if len(row) > 4 else None
                channel = row[5] if len(row) > 5 else None
                device_raw = row[6] if len(row) > 6 else None  # Device column
                signal_name = row[7] if len(row) > 7 else None
                tech_pos = row[8] if len(row) > 8 else None
                main_pos = row[9] if len(row) > 9 else None
                
                # Разрешаем формулы вида =$A$7
                if isinstance(cabinet, str) and cabinet.startswith('='):
                    cabinet = self._resolve_formula(ws, row_idx, 0)
                
                self.io_signals.append({
                    'sheet': sheet_name,
                    'cabinet': cabinet,
                    'sig_type': sig_type,
                    'address': address,
                    'channel': channel,
                    'device_raw': device_raw,
                    'signal_name': signal_name,
                    'tech_pos': tech_pos,
                    'main_pos': main_pos
                })

    def _resolve_formula(self, ws, row_idx, col_idx):
        """Пытается разрешить простую формулу Excel"""
        # Для упрощения возвращаем значение из первой строки данных
        return ws.cell(row=7, column=col_idx+1).value

    def _parse_variables(self):
        """Парсит Список переменных"""
        if not self.workbook_vars:
            return
        
        if 'Лист1' in self.workbook_vars.sheetnames:
            ws = self.workbook_vars['Лист1']
            
            # Структура: A=DI описание, B=DI переменная, D=DO описание, E=DO переменная, G=AI описание, H=AI переменная
            for row_idx in range(2, ws.max_row + 1):
                row = [cell.value for cell in ws[row_idx]]
                
                # DI
                di_desc = row[0] if len(row) > 0 else None
                di_var = row[1] if len(row) > 1 else None
                
                # DO
                do_desc = row[3] if len(row) > 3 else None
                do_var = row[4] if len(row) > 4 else None
                
                # AI
                ai_desc = row[6] if len(row) > 6 else None
                ai_var = row[7] if len(row) > 7 else None
                
                if di_desc and di_var:
                    self.variables_map[di_desc] = {'var': di_var, 'type': 'DI'}
                if do_desc and do_var:
                    self.variables_map[do_desc] = {'var': do_var, 'type': 'DO'}
                if ai_desc and ai_var:
                    self.variables_map[ai_desc] = {'var': ai_var, 'type': 'AI'}

    def _merge_data(self):
        """Объединяет данные из всех источников в итоговые устройства"""
        # Группируем сигналы по tech_pos + device_raw
        signals_by_device = defaultdict(list)
        
        for sig in self.io_signals:
            tech_pos = str(sig['tech_pos']).strip() if sig['tech_pos'] else ''
            device_raw = str(sig['device_raw']).strip() if sig['device_raw'] else ''
            key = f"{device_raw}_{tech_pos}"
            signals_by_device[key].append(sig)
        
        # Создаем устройства на основе сигналов
        for device_key, signals in signals_by_device.items():
            if not signals:
                continue
            
            first_sig = signals[0]
            device_raw = first_sig['device_raw']
            tech_pos = first_sig['tech_pos']
            main_pos = first_sig['main_pos']
            
            # Определяем тип устройства
            dev_type = self._detect_device_type(device_raw, signals)
            
            # Формируем имя
            dev_name = self._format_device_name(dev_type, tech_pos, main_pos)
            
            # Создаем устройство
            self.devices[device_key] = {
                'type': dev_type,
                'name': dev_name,
                'tech_pos': tech_pos,
                'main_pos': main_pos,
                'device_raw': device_raw,
                'signals': [],
                'cabinet': first_sig['cabinet']
            }
            
            # Добавляем сигналы
            for sig in signals:
                var_info = self.variables_map.get(sig['signal_name'])
                self.devices[device_key]['signals'].append({
                    'name': sig['signal_name'],
                    'type': sig['sig_type'],
                    'address': sig['address'],
                    'channel': sig['channel'],
                    'variable': var_info['var'] if var_info else None
                })
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
