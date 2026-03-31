#!/bin/bash
# Запуск веб-интерфейса генератора кода TIA Portal

echo "=================================================="
echo "  TIA Portal SCL Generator - Web Interface"
echo "=================================================="
echo ""

# Проверка наличия Flask
python3 -c "import flask" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "⚠️  Flask не установлен. Установка..."
    pip3 install flask --quiet
fi

echo ""
echo "🌐 Веб-сервер запускается..."
echo ""
echo "📍 Откройте в браузере: http://localhost:5000"
echo "   или http://127.0.0.1:5000"
echo ""
echo "⚠️  Для остановки нажмите Ctrl+C"
echo "=================================================="
echo ""

# Запуск веб-сервера
python3 "$(dirname "$0")/web_gui.py"
