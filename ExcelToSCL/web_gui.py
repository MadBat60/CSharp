#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TIA Portal SCL Code Generator - Web Interface
Генератор кода для TIA Portal из Excel файлов с веб-интерфейсом
"""

import os
import sys
import tempfile
from pathlib import Path

# Проверка наличия flask
try:
    from flask import Flask, render_template_string, request, send_file, jsonify
    from werkzeug.utils import secure_filename
except ImportError:
    print("❌ Flask не установлен. Установите командой: pip install flask")
    sys.exit(1)

# Добавляем путь к основному скрипту
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_to_scl import ExcelToSCLConverter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

# HTML шаблон интерфейса
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>TIA Portal SCL Generator</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
        }
        
        .header p {
            opacity: 0.9;
            font-size: 1.1em;
        }
        
        .content {
            padding: 40px;
        }
        
        .upload-section {
            border: 3px dashed #3498db;
            border-radius: 10px;
            padding: 40px;
            text-align: center;
            background: #f8f9fa;
            transition: all 0.3s ease;
            margin-bottom: 30px;
        }
        
        .upload-section:hover {
            border-color: #2ecc71;
            background: #e8f8f0;
        }
        
        .upload-section.dragover {
            border-color: #e74c3c;
            background: #fdedec;
        }
        
        .upload-icon {
            font-size: 4em;
            margin-bottom: 20px;
        }
        
        .btn {
            background: linear-gradient(135deg, #3498db 0%, #2980b9 100%);
            color: white;
            border: none;
            padding: 15px 40px;
            font-size: 1.1em;
            border-radius: 8px;
            cursor: pointer;
            transition: all 0.3s ease;
            margin: 10px 5px;
        }
        
        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(52, 152, 219, 0.4);
        }
        
        .btn:disabled {
            background: #bdc3c7;
            cursor: not-allowed;
            transform: none;
        }
        
        .btn-success {
            background: linear-gradient(135deg, #2ecc71 0%, #27ae60 100%);
        }
        
        .btn-download {
            background: linear-gradient(135deg, #9b59b6 0%, #8e44ad 100%);
        }
        
        .progress-container {
            margin: 30px 0;
            display: none;
        }
        
        .progress-bar {
            width: 100%;
            height: 30px;
            background: #ecf0f1;
            border-radius: 15px;
            overflow: hidden;
        }
        
        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, #3498db, #2ecc71);
            width: 0%;
            transition: width 0.3s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: bold;
        }
        
        .status-message {
            margin-top: 15px;
            padding: 15px;
            border-radius: 8px;
            display: none;
        }
        
        .status-success {
            background: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }
        
        .status-error {
            background: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
        
        .output-section {
            margin-top: 30px;
            display: none;
        }
        
        .output-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
        }
        
        .code-preview {
            background: #282c34;
            color: #abb2bf;
            padding: 20px;
            border-radius: 8px;
            max-height: 500px;
            overflow-y: auto;
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 13px;
            line-height: 1.5;
            white-space: pre-wrap;
            word-wrap: break-word;
        }
        
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin: 20px 0;
        }
        
        .stat-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            text-align: center;
        }
        
        .stat-number {
            font-size: 2.5em;
            font-weight: bold;
            margin-bottom: 5px;
        }
        
        .stat-label {
            opacity: 0.9;
        }
        
        input[type="file"] {
            display: none;
        }
        
        .file-info {
            margin-top: 15px;
            padding: 10px;
            background: #e3f2fd;
            border-radius: 5px;
            display: none;
        }
        
        .loading-spinner {
            display: inline-block;
            width: 20px;
            height: 20px;
            border: 3px solid rgba(255,255,255,0.3);
            border-radius: 50%;
            border-top-color: white;
            animation: spin 1s ease-in-out infinite;
            margin-right: 10px;
        }
        
        @keyframes spin {
            to { transform: rotate(360deg); }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🏭 TIA Portal SCL Generator</h1>
            <p>Генерация кода для SIMATIC из Excel файлов системы ввода-вывода</p>
        </div>
        
        <div class="content">
            <div class="upload-section" id="dropZone">
                <div class="upload-icon">📊</div>
                <h2>Загрузите Excel файл</h2>
                <p>Перетащите файл сюда или нажмите кнопку ниже</p>
                <p style="color: #7f8c8d; margin-top: 10px;">
                    Поддерживаются файлы .xlsx и .xls с листами ШС и ШСАУ
                </p>
                <button class="btn" onclick="document.getElementById('fileInput').click()">
                    📁 Выбрать файл
                </button>
                <input type="file" id="fileInput" accept=".xlsx,.xls" onchange="handleFileSelect(this)">
                
                <div class="file-info" id="fileInfo"></div>
            </div>
            
            <div style="text-align: center; margin: 30px 0;">
                <button class="btn btn-success" id="generateBtn" onclick="generateCode()" disabled>
                    🚀 Сгенерировать код
                </button>
            </div>
            
            <div class="progress-container" id="progressContainer">
                <div class="progress-bar">
                    <div class="progress-fill" id="progressFill">0%</div>
                </div>
                <p id="statusMessage" style="text-align: center; margin-top: 10px; color: #7f8c8d;"></p>
            </div>
            
            <div class="output-section" id="outputSection">
                <div class="output-header">
                    <h2>✅ Результат генерации</h2>
                    <button class="btn btn-download" onclick="downloadResult()">
                        💾 Скачать файл
                    </button>
                </div>
                
                <div class="stats-grid" id="statsGrid"></div>
                
                <h3 style="margin: 20px 0 10px;">Предпросмотр кода:</h3>
                <div class="code-preview" id="codePreview"></div>
            </div>
        </div>
    </div>
    
    <script>
        let selectedFile = null;
        let generatedFileName = null;
        
        // Drag and drop
        const dropZone = document.getElementById('dropZone');
        
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, preventDefaults, false);
        });
        
        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }
        
        ['dragenter', 'dragover'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => dropZone.classList.add('dragover'), false);
        });
        
        ['dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => dropZone.classList.remove('dragover'), false);
        });
        
        dropZone.addEventListener('drop', handleDrop, false);
        
        function handleDrop(e) {
            const dt = e.dataTransfer;
            const files = dt.files;
            if (files.length > 0) {
                handleFile(files[0]);
            }
        }
        
        function handleFileSelect(input) {
            if (input.files.length > 0) {
                handleFile(input.files[0]);
            }
        }
        
        function handleFile(file) {
            selectedFile = file;
            document.getElementById('fileInfo').style.display = 'block';
            document.getElementById('fileInfo').innerHTML = `
                <strong>Выбран файл:</strong> ${file.name}<br>
                <strong>Размер:</strong> ${(file.size / 1024 / 1024).toFixed(2)} MB
            `;
            document.getElementById('generateBtn').disabled = false;
        }
        
        async function generateCode() {
            if (!selectedFile) {
                alert('Пожалуйста, выберите файл!');
                return;
            }
            
            const formData = new FormData();
            formData.append('file', selectedFile);
            
            document.getElementById('progressContainer').style.display = 'block';
            document.getElementById('generateBtn').disabled = true;
            document.getElementById('outputSection').style.display = 'none';
            
            try {
                // Шаг 1: Загрузка файла
                updateProgress(10, 'Загрузка файла...');
                
                const uploadResponse = await fetch('/upload', {
                    method: 'POST',
                    body: formData
                });
                
                if (!uploadResponse.ok) {
                    throw new Error('Ошибка загрузки файла');
                }
                
                const uploadResult = await uploadResponse.json();
                const fileId = uploadResult.file_id;
                
                // Шаг 2: Генерация кода
                updateProgress(30, 'Анализ структуры Excel...');
                
                const generateResponse = await fetch('/generate', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({ file_id: fileId })
                });
                
                if (!generateResponse.ok) {
                    throw new Error('Ошибка генерации кода');
                }
                
                const result = await generateResponse.json();
                
                if (result.success) {
                    updateProgress(100, 'Генерация завершена!');
                    showResult(result);
                } else {
                    throw new Error(result.error || 'Неизвестная ошибка');
                }
                
            } catch (error) {
                updateProgress(0, '');
                document.getElementById('progressContainer').style.display = 'none';
                document.getElementById('generateBtn').disabled = false;
                alert('Ошибка: ' + error.message);
            }
        }
        
        function updateProgress(percent, message) {
            const progressFill = document.getElementById('progressFill');
            const statusMessage = document.getElementById('statusMessage');
            
            progressFill.style.width = percent + '%';
            progressFill.textContent = Math.round(percent) + '%';
            statusMessage.textContent = message;
        }
        
        function showResult(result) {
            document.getElementById('outputSection').style.display = 'block';
            generatedFileName = result.filename;
            
            // Показываем статистику
            const stats = result.statistics;
            document.getElementById('statsGrid').innerHTML = `
                <div class="stat-card">
                    <div class="stat-number">${stats.equipment_types}</div>
                    <div class="stat-label">Типов оборудования</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">${stats.total_devices}</div>
                    <div class="stat-label">Всего устройств</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">${stats.module_channels}</div>
                    <div class="stat-label">Каналов модулей</div>
                </div>
            `;
            
            // Показываем предпросмотр кода (первые 2000 символов)
            const preview = result.code_preview.substring(0, 3000);
            document.getElementById('codePreview').textContent = preview + (result.code_preview.length > 3000 ? '\\n\\n... (полный код в скачанном файле)' : '');
            
            setTimeout(() => {
                document.getElementById('generateBtn').disabled = false;
            }, 1000);
        }
        
        function downloadResult() {
            if (generatedFileName) {
                window.location.href = '/download/' + generatedFileName;
            }
        }
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    """Главная страница"""
    return render_template_string(HTML_TEMPLATE)

@app.route('/upload', methods=['POST'])
def upload_file():
    """Загрузка файла"""
    if 'file' not in request.files:
        return jsonify({'error': 'Файл не найден'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Файл не выбран'}), 400
    
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Поддерживаются только Excel файлы (.xlsx, .xls)'}), 400
    
    # Сохраняем файл
    filename = secure_filename(file.filename)
    file_id = f"{os.getpid()}_{filename}"
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file_id)
    file.save(filepath)
    
    return jsonify({
        'success': True,
        'file_id': file_id,
        'filename': filename
    })

@app.route('/generate', methods=['POST'])
def generate():
    """Генерация кода"""
    data = request.json
    file_id = data.get('file_id')
    
    if not file_id:
        return jsonify({'error': 'ID файла не указан'}), 400
    
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file_id)
    
    if not os.path.exists(filepath):
        return jsonify({'error': 'Файл не найден'}), 404
    
    try:
        # Создаем генератор
        generator = ExcelToSCLConverter(filepath)
        
        # Анализируем workbook
        workbook_info = generator.analyze_workbook()
        
        # Генерируем код
        generated_code = generator.generate_full_code()
        
        # Получаем статистику
        stats = generator.get_statistics()
        
        # Сохраняем результат
        output_filename = f"{os.path.splitext(os.path.basename(filepath))[0]}_SCL.txt"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(generated_code)
        
        return jsonify({
            'success': True,
            'filename': output_filename,
            'code_preview': generated_code[:5000],
            'statistics': stats,
            'workbook_info': workbook_info
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download(filename):
    """Скачивание результата"""
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(filename))
    
    if not os.path.exists(filepath):
        return jsonify({'error': 'Файл не найден'}), 404
    
    return send_file(
        filepath,
        as_attachment=True,
        download_name=filename,
        mimetype='text/plain'
    )

def main():
    """Запуск веб-сервера"""
    print("=" * 60)
    print("🏭 TIA Portal SCL Generator - Web Interface")
    print("=" * 60)
    print()
    print("🌐 Откройте в браузере: http://localhost:5000")
    print()
    print("⚠️  Для остановки нажмите Ctrl+C")
    print("=" * 60)
    
    app.run(debug=True, host='0.0.0.0', port=5000)

if __name__ == '__main__':
    main()
