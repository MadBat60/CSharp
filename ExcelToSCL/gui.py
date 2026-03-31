#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TIA Portal SCL Code Generator - GUI Interface
Генератор кода для TIA Portal из Excel файлов с графическим интерфейсом
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import sys

# Добавляем путь к основному скрипту
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_to_scl import ExcelToSCLConverter


class ExcelToSCLGUI:
    """Графический интерфейс для генератора кода TIA Portal"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("TIA Portal SCL Code Generator")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 600)
        
        # Переменные
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.generator = None
        self.is_processing = False
        
        # Настройка стилей
        self.setup_styles()
        
        # Создание интерфейса
        self.create_widgets()
        
    def setup_styles(self):
        """Настройка стилей интерфейса"""
        style = ttk.Style()
        
        # Доступные темы
        available_themes = style.theme_names()
        if 'vista' in available_themes:
            style.theme_use('vista')
        elif 'clam' in available_themes:
            style.theme_use('clam')
        
        # Настройка цветов
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'))
        style.configure('Subtitle.TLabel', font=('Arial', 12))
        style.configure('Success.TLabel', foreground='green')
        style.configure('Error.TLabel', foreground='red')
        style.configure('Info.TLabel', foreground='blue')
        
    def create_widgets(self):
        """Создание всех элементов интерфейса"""
        # Главный контейнер
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настройка растягивания
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        # Заголовок
        title_label = ttk.Label(
            main_frame, 
            text="Генератор кода TIA Portal (SCL)", 
            style='Title.TLabel'
        )
        title_label.grid(row=0, column=0, pady=(0, 20))
        
        subtitle_label = ttk.Label(
            main_frame,
            text="Из Excel файлов системы ввода-вывода в код для SIMATIC",
            style='Subtitle.TLabel'
        )
        subtitle_label.grid(row=1, column=0, pady=(0, 20))
        
        # Фрейм для выбора файлов
        file_frame = ttk.LabelFrame(main_frame, text="Выбор файлов", padding="10")
        file_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        # Входной файл
        ttk.Label(file_frame, text="Excel файл:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(
            file_frame, 
            textvariable=self.input_file, 
            width=50
        ).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Button(
            file_frame, 
            text="Обзор...", 
            command=self.browse_input_file
        ).grid(row=0, column=2, pady=5)
        
        # Выходной файл
        ttk.Label(file_frame, text="Выходной файл:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(
            file_frame, 
            textvariable=self.output_file, 
            width=50
        ).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Button(
            file_frame, 
            text="Обзор...", 
            command=self.browse_output_file
        ).grid(row=1, column=2, pady=5)
        
        # Кнопки управления
        button_frame = ttk.Frame(file_frame)
        button_frame.grid(row=2, column=0, columnspan=3, pady=10)
        
        self.generate_btn = ttk.Button(
            button_frame, 
            text="🚀 Генерировать код", 
            command=self.start_generation,
            width=20
        )
        self.generate_btn.grid(row=0, column=0, padx=5)
        
        self.clear_btn = ttk.Button(
            button_frame,
            text="🗑️ Очистить",
            command=self.clear_output,
            width=15
        )
        self.clear_btn.grid(row=0, column=1, padx=5)
        
        self.save_btn = ttk.Button(
            button_frame,
            text="💾 Сохранить результат",
            command=self.save_result,
            width=20,
            state=tk.DISABLED
        )
        self.save_btn.grid(row=0, column=2, padx=5)
        
        # Индикатор прогресса
        progress_frame = ttk.Frame(file_frame)
        progress_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        progress_frame.columnconfigure(1, weight=1)
        
        ttk.Label(progress_frame, text="Прогресс:").grid(row=0, column=0, sticky=tk.W)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_var, 
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        
        self.status_label = ttk.Label(progress_frame, text="Готов к работе", style='Info.TLabel')
        self.status_label.grid(row=0, column=2, padx=5)
        
        # Фрейм для вывода результата
        output_frame = ttk.LabelFrame(main_frame, text="Результат генерации", padding="10")
        output_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1)
        
        # Текстовое поле с прокруткой
        self.output_text = scrolledtext.ScrolledText(
            output_frame,
            wrap=tk.WORD,
            width=120,
            height=30,
            font=('Consolas', 9),
            bg='#f8f8f8',
            fg='#333'
        )
        self.output_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настройка тегов для подсветки
        self.output_text.tag_configure('comment', foreground='gray', font=('Consolas', 9, 'italic'))
        self.output_text.tag_configure('keyword', foreground='blue', font=('Consolas', 9, 'bold'))
        self.output_text.tag_configure('string', foreground='green')
        self.output_text.tag_configure('number', foreground='red')
        self.output_text.tag_configure('header', foreground='purple', font=('Consolas', 9, 'bold'))
        self.output_text.tag_configure('success', foreground='green', font=('Consolas', 9, 'bold'))
        self.output_text.tag_configure('error', foreground='red', font=('Consolas', 9, 'bold'))
        
        # Статус бар
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=4, column=0, sticky=(tk.W, tk.E))
        
        self.stats_label = ttk.Label(status_frame, text="")
        self.stats_label.grid(row=0, column=0, sticky=tk.W)
        
    def browse_input_file(self):
        """Выбор входного Excel файла"""
        filename = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.input_file.set(filename)
            # Автоматически предложить имя выходного файла
            base_name = os.path.splitext(os.path.basename(filename))[0]
            default_output = os.path.join(
                os.path.dirname(filename),
                f"{base_name}_SCL.txt"
            )
            self.output_file.set(default_output)
            self.update_status(f"Выбран файл: {os.path.basename(filename)}", 'info')
    
    def browse_output_file(self):
        """Выбор выходного файла"""
        filename = filedialog.asksaveasfilename(
            title="Сохранить результат как",
            defaultextension=".txt",
            filetypes=[
                ("Text files", "*.txt"),
                ("SCL files", "*.scl"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.output_file.set(filename)
    
    def update_status(self, message, status_type='info'):
        """Обновление статуса"""
        self.status_label.config(text=message)
        if status_type == 'success':
            self.status_label.configure(style='Success.TLabel')
        elif status_type == 'error':
            self.status_label.configure(style='Error.TLabel')
        else:
            self.status_label.configure(style='Info.TLabel')
        self.root.update_idletasks()
    
    def update_progress(self, value):
        """Обновление прогресс-бара"""
        self.progress_var.set(value)
        self.root.update_idletasks()
    
    def append_output(self, text, tag=None):
        """Добавление текста в окно вывода"""
        self.output_text.insert(tk.END, text, tag)
        self.output_text.see(tk.END)
        self.root.update_idletasks()
    
    def clear_output(self):
        """Очистка окна вывода"""
        self.output_text.delete(1.0, tk.END)
        self.stats_label.config(text="")
        self.update_status("Очищено", 'info')
        self.progress_var.set(0)
    
    def start_generation(self):
        """Запуск генерации в отдельном потоке"""
        if not self.input_file.get():
            messagebox.showerror("Ошибка", "Пожалуйста, выберите Excel файл!")
            return
        
        if not os.path.exists(self.input_file.get()):
            messagebox.showerror("Ошибка", "Указанный файл не существует!")
            return
        
        if self.is_processing:
            messagebox.showwarning("Предупреждение", "Обработка уже выполняется!")
            return
        
        # Блокировка кнопок
        self.is_processing = True
        self.generate_btn.config(state=tk.DISABLED)
        self.save_btn.config(state=tk.DISABLED)
        self.clear_output()
        
        # Запуск в отдельном потоке
        thread = threading.Thread(target=self.run_generation, daemon=True)
        thread.start()
    
    def run_generation(self):
        """Основная логика генерации (выполняется в потоке)"""
        try:
            input_path = self.input_file.get()
            output_path = self.output_file.get()
            
            self.append_output("=" * 80 + "\n", 'header')
            self.append_output("ГЕНЕРАЦИЯ КОДА TIA PORTAL (SCL)\n", 'header')
            self.append_output("=" * 80 + "\n\n", 'header')
            
            self.update_status("Инициализация генератора...", 'info')
            self.update_progress(10)
            
            # Создание генератора
            self.generator = ExcelToSCLConverter(input_path)
            
            self.update_status("Анализ структуры Excel файла...", 'info')
            self.update_progress(20)
            
            # Анализ workbook
            workbook_info = self.generator.analyze_workbook()
            
            self.append_output(f"📊 Входной файл: {os.path.basename(input_path)}\n", 'info')
            self.append_output(f"📋 Листы найдено: {len(workbook_info.get('sheets', []))}\n")
            
            for sheet_name in workbook_info.get('sheets', []):
                self.append_output(f"   • {sheet_name}\n")
            
            self.update_progress(30)
            
            # Генерация кода
            self.update_status("Генерация кода SCL...", 'info')
            self.append_output("\n" + "=" * 80 + "\n", 'header')
            self.append_output("ПРОЦЕСС ГЕНЕРАЦИИ\n", 'header')
            self.append_output("=" * 80 + "\n\n")
            
            generated_code = self.generator.generate_full_code(
                callback=lambda msg, prog: self.process_callback(msg, prog)
            )
            
            self.update_progress(90)
            
            # Вывод результата
            self.append_output("\n" + "=" * 80 + "\n", 'header')
            self.append_output("РЕЗУЛЬТАТ ГЕНЕРАЦИИ\n", 'header')
            self.append_output("=" * 80 + "\n\n")
            
            self.append_output(generated_code)
            
            # Сохранение в файл
            if output_path:
                self.update_status("Сохранение результата...", 'info')
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(generated_code)
                
                self.append_output(f"\n\n✅ Файл сохранён: {output_path}\n", 'success')
            
            # Статистика
            stats = self.generator.get_statistics()
            stats_text = (
                f"\n{'='*80}\n"
                f"СТАТИСТИКА:\n"
                f"{'='*80}\n"
                f"• Типов оборудования: {stats['equipment_types']}\n"
                f"• Всего устройств: {stats['total_devices']}\n"
                f"• Каналов модулей: {stats['module_channels']}\n"
                f"{'='*80}\n"
            )
            self.append_output(stats_text, 'success')
            
            self.stats_label.config(
                text=f"Типов: {stats['equipment_types']} | Устройств: {stats['total_devices']} | "
                     f"Каналов: {stats['module_channels']}"
            )
            
            self.update_progress(100)
            self.update_status("Генерация завершена успешно!", 'success')
            self.save_btn.config(state=tk.NORMAL)
            
            messagebox.showinfo("Успех", "Код успешно сгенерирован!")
            
        except Exception as e:
            error_msg = f"❌ Ошибка: {str(e)}"
            self.append_output(f"\n{error_msg}\n", 'error')
            self.update_status("Ошибка генерации!", 'error')
            messagebox.showerror("Ошибка", str(e))
        
        finally:
            self.is_processing = False
            self.generate_btn.config(state=tk.NORMAL)
    
    def process_callback(self, message, progress):
        """Callback для обновления прогресса во время генерации"""
        self.update_status(message, 'info')
        self.update_progress(30 + (progress * 0.6))  # 30-90% диапазон
        self.append_output(f"→ {message}\n")
    
    def save_result(self):
        """Сохранение результата в файл"""
        if not self.output_file.get():
            self.browse_output_file()
        
        if self.output_file.get():
            try:
                content = self.output_text.get(1.0, tk.END)
                with open(self.output_file.get(), 'w', encoding='utf-8') as f:
                    f.write(content)
                messagebox.showinfo("Успех", f"Результат сохранён в:\n{self.output_file.get()}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{str(e)}")


def main():
    """Точка входа приложения"""
    root = tk.Tk()
    
    # Установка иконки (если доступна)
    try:
        # Можно добавить иконку при наличии
        pass
    except:
        pass
    
    app = ExcelToSCLGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
