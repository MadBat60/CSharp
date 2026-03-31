#!/bin/bash
# Запуск GUI интерфейса генератора кода TIA Portal

echo "=================================================="
echo "  TIA Portal SCL Generator - GUI Launcher"
echo "=================================================="
echo ""

# Проверка наличия tkinter
python3 -c "import tkinter" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "❌ Tkinter не установлен!"
    echo ""
    echo "Для установки выполните:"
    echo "  Ubuntu/Debian: sudo apt-get install python3-tk"
    echo "  Fedora: sudo dnf install python3-tkinter"
    echo "  macOS: brew install python-tk"
    echo ""
    echo "Или используйте веб-интерфейс:"
    echo "  ./run_web_gui.sh"
    echo ""
    exit 1
fi

# Запуск GUI
echo "🚀 Запуск графического интерфейса..."
python3 "$(dirname "$0")/gui.py"
