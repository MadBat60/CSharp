#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TIA Portal SCL Code Generator - Web Interface (v2.0)
Генератор кода для TIA Portal из ТРЕХ Excel файлов:
1. Спецификация.xlsx - технологические данные (Config_Line, Config_Transport)
2. Система ввода-вывода.xlsx - аппаратная привязка (ШС, ШСАУ)
3. Список переменных от системы ввода-вывода.xlsx - имена переменных
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
    <title>TIA Portal SCL Generator v2.0</title>
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
            max-width: 1400px;
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
        
        .upload-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(350px, 1fr));
            gap: 25px;
            margin-bottom: 30px;
        }
        
        .upload-card {
            border: 3px dashed #3498db;
            border-radius: 10px;
            padding: 25px;
            text-align: center;
            background: #f8f9fa;
            transition: all 0.3s ease;
        }
        
        .upload-card:hover {
            border-color: #2ecc71;
            background: #e8f8f0;
        }
        
        .upload-card.dragover {
            border-color: #e74c3c;
            background: #fdedec;
        }
        
        .upload-card.required {
            border-color: #e74c3c;
            background: #fdedec;
        }
        
        .upload-icon {
            font-size: 3em;
            margin-bottom: 15px;
        }
        
        .upload-card h3 {
            color: #2c3e50;
            margin-bottom: 10px;
        }
        
        .upload-card p {
            color: #7f8c8d;
            font-size: 0.9em;
            margin-bottom: 15px;
        }
        
        .btn {
            background: linear-gradient(135deg, #3498db 0%, #2980b9 100%);
            color: white;
            border: none;
            padding: 12px 30px;
            font-size: 1em;
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
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
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
            font-size: 0.9em;
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
            font-size: 0.9em;
        }
        
        .file-status {
            display: inline-block;
            padding: 3px 10px;
            border-radius: 15px;
            font-size: 0.8em;
            margin-left: 10px;
        }
        
        .file-status.loaded {
            background: #2ecc71;
            color: white;
        }
        
        .file-status.missing {
            background: #e74c3c;
            color: white;
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
        
        .files-required-note {
            background: #fff3cd;
            border: 1px solid #ffc107;
            color: #856404;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🏭 TIA Portal SCL Generator v2.0</h1>
            <p>Генерация кода для SIMATIC из ТРЕХ Excel файлов системы ввода-вывода</p>
        </div>
        
        <div class="content">
            <div class="files-required-note">
                <strong>⚠️ Важно:</strong> Для корректной генерации кода необходимо загрузить ВСЕ три файла.
                Программа объединяет данные из всех источников для создания полной карты устройств.
            </div>
            
            <div class="upload-grid">
                <!-- Файл 1: Спецификация -->
                <div class="upload-card required" id="card-spec" ondrop="handleDrop(event, 'spec')" ondragover="handleDragOver(event)" ondragleave="handleDragLeave(event)">
                    <div class="upload-icon">📋</div>
                    <h3>1. Спецификация.xlsx</h3>
                    <p>Технологические данные (Config_Line, Config_Transport)</p>
                    <button class="btn" onclick="document.getElementById('file-spec').click()">📁 Выбрать файл</button>
                    <input type="file" id="file-spec" accept=".xlsx,.xls" onchange="handleFileSelect(event, 'spec')">
                    <div class="file-info" id="info-spec"></div>
                </div>
                
                <!-- Файл 2: Система ввода-вывода -->
                <div class="upload-card required" id="card-io" ondrop="handleDrop(event, 'io')" ondragover="handleDragOver(event)" ondragleave="handleDragLeave(event)">
                    <div class="upload-icon">🔌</div>
                    <h3>2. Система ввода-вывода.xlsx</h3>
                    <p>Аппаратная привязка (листы ШС, ШСАУ)</p>
                    <button class="btn" onclick="document.getElementById('file-io').click()">📁 Выбрать файл</button>
                    <input type="file" id="file-io" accept=".xlsx,.xls" onchange="handleFileSelect(event, 'io')">
                    <div class="file-info" id="info-io"></div>
                </div>
                
                <!-- Файл 3: Список переменных -->
                <div class="upload-card required" id="card-vars" ondrop="handleDrop(event, 'vars')" ondragover="handleDragOver(event)" ondragleave="handleDragLeave(event)">
                    <div class="upload-icon">📝</div>
                    <h3>3. Список переменных.xlsx</h3>
                    <p>Имена переменных для привязки</p>
                    <button class="btn" onclick="document.getElementById('file-vars').click()">📁 Выбрать файл</button>
                    <input type="file" id="file-vars" accept=".xlsx,.xls" onchange="handleFileSelect(event, 'vars')">
                    <div class="file-info" id="info-vars"></div>
                </div>
            </div>
            
            <div style="text-align: center; margin: 30px 0;">
                <button class="btn btn-success" id="generateBtn" onclick="generateCode()" disabled>
                    🚀 Сгенерировать код
                </button>
                <span id="filesStatus" style="margin-left: 15px; color: #7f8c8d;"></span>
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
        let files = {
            spec: null,
            io: null,
            vars: null
        };
        let generatedFileName = null;
        
        function updateFilesStatus() {
            const loaded = Object.values(files).filter(f => f !== null).length;
            const statusEl = document.getElementById('filesStatus');
            const generateBtn = document.getElementById('generateBtn');
            
            if (loaded === 3) {
                statusEl.textContent = '✅ Все файлы загружены';
                statusEl.style.color = '#27ae60';
                generateBtn.disabled = false;
            } else {
                statusEl.textContent = `⏳ Загружено файлов: ${loaded} из 3`;
                statusEl.style.color = '#e67e22';
                generateBtn.disabled = true;
            }
        }
        
        // Drag and drop
        function handleDragOver(e) {
            e.preventDefault();
            e.currentTarget.classList.add('dragover');
        }
        
        function handleDragLeave(e) {
            e.currentTarget.classList.remove('dragover');
        }
        
        function handleDrop(e, fileType) {
            e.preventDefault();
            e.currentTarget.classList.remove('dragover');
            const dt = e.dataTransfer;
            const fileList = dt.files;
            if (fileList.length > 0) {
                handleFile(fileList[0], fileType);
            }
        }
        
        function handleFileSelect(event, fileType) {
            const input = event.target;
            if (input.files.length > 0) {
                handleFile(input.files[0], fileType);
            }
        }
        
        function handleFile(file, fileType) {
            files[fileType] = file;
            const infoEl = document.getElementById('info-' + fileType);
            const cardEl = document.getElementById('card-' + fileType);
            
            infoEl.style.display = 'block';
            infoEl.innerHTML = `
                <strong>✅ Файл загружен:</strong> ${file.name}<br>
                <strong>Размер:</strong> ${(file.size / 1024 / 1024).toFixed(2)} MB
                <span class="file-status loaded">Готов</span>
            `;
            cardEl.classList.remove('required');
            
            updateFilesStatus();
        }
        
        async function generateCode() {
            // Проверка наличия всех файлов
            if (!files.spec || !files.io || !files.vars) {
                alert('Пожалуйста, загрузите ВСЕ три файла!');
                return;
            }
            
            const formData = new FormData();
            formData.append('spec_file', files.spec);
            formData.append('io_file', files.io);
            formData.append('vars_file', files.vars);
            
            document.getElementById('progressContainer').style.display = 'block';
            document.getElementById('generateBtn').disabled = true;
            document.getElementById('outputSection').style.display = 'none';
            
            try {
                // Шаг 1: Загрузка файлов
                updateProgress(10, 'Загрузка файлов...');
                
                const uploadResponse = await fetch('/upload', {
                    method: 'POST',
                    body: formData
                });
                
                if (!uploadResponse.ok) {
                    throw new Error('Ошибка загрузки файлов');
                }
                
                const uploadResult = await uploadResponse.json();
                const sessionId = uploadResult.session_id;
                
                // Шаг 2: Генерация кода
                updateProgress(30, 'Анализ структуры Excel файлов...');
                
                const generateResponse = await fetch('/generate', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({ session_id: sessionId })
                });
                
                if (!generateResponse.ok) {
                    const errorData = await generateResponse.json();
                    throw new Error(errorData.error || 'Ошибка генерации кода');
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
                <div class="stat-card">
                    <div class="stat-number">${stats.ai_count}</div>
                    <div class="stat-label">AI каналы</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">${stats.ao_count}</div>
                    <div class="stat-label">AO каналы</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">${stats.di_count}</div>
                    <div class="stat-label">DI каналы</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">${stats.do_count}</div>
                    <div class="stat-label">DO каналы</div>
                </div>
            `;
            
            // Показываем предпросмотр кода (первые 3000 символов)
            const preview = result.code_preview.substring(0, 3000);
            document.getElementById('codePreview').textContent = preview + (result.code_preview.length > 3000 ? '\n\n... (полный код в скачанном файле)' : '');
            
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
    """Загрузка ТРЕХ файлов"""
    required_files = ['spec_file', 'io_file', 'vars_file']
    file_names = {
        'spec_file': 'Спецификация',
        'io_file': 'Система ввода-вывода',
        'vars_file': 'Список переменных'
    }
    
    uploaded_files = {}
    
    for field_name in required_files:
        if field_name not in request.files:
            return jsonify({'error': f'Файл не найден: {file_names[field_name]}'}), 400
        
        file = request.files[field_name]
        if file.filename == '':
            return jsonify({'error': f'Файл не выбран: {file_names[field_name]}'}), 400
        
        if not file.filename.endswith(('.xlsx', '.xls')):
            return jsonify({'error': f'Поддерживаются только Excel файлы: {file_names[field_name]}'}), 400
        
        # Сохраняем файл с уникальным именем
        filename = secure_filename(file.filename)
        file_id = f"{os.getpid()}_{field_name}_{filename}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file_id)
        file.save(filepath)
        uploaded_files[field_name] = filepath
    
    # Создаем session_id из комбинации всех file_id
    session_id = f"{os.getpid()}_{uploaded_files['spec_file']}_{uploaded_files['io_file']}_{uploaded_files['vars_file']}"
    
    # Сохраняем информацию о сессии
    session_data = {
        'spec_file': uploaded_files['spec_file'],
        'io_file': uploaded_files['io_file'],
        'vars_file': uploaded_files['vars_file']
    }
    session_path = os.path.join(app.config['UPLOAD_FOLDER'], f"session_{session_id}.json")
    
    import json
    with open(session_path, 'w', encoding='utf-8') as f:
        json.dump(session_data, f, ensure_ascii=False)
    
    return jsonify({
        'success': True,
        'session_id': session_id,
        'files': {k: os.path.basename(v) for k, v in uploaded_files.items()}
    })

@app.route('/generate', methods=['POST'])
def generate():
    """Генерация кода из ТРЕХ файлов"""
    data = request.json
    session_id = data.get('session_id')
    
    if not session_id:
        return jsonify({'error': 'ID сессии не указан'}), 400
    
    # Загружаем информацию о сессии
    session_path = os.path.join(app.config['UPLOAD_FOLDER'], f"session_{session_id}.json")
    
    if not os.path.exists(session_path):
        return jsonify({'error': 'Сессия не найдена. Загрузите файлы заново.'}), 404
    
    import json
    with open(session_path, 'r', encoding='utf-8') as f:
        session_data = json.load(f)
    
    spec_file = session_data.get('spec_file')
    io_file = session_data.get('io_file')
    vars_file = session_data.get('vars_file')
    
    # Проверяем наличие всех файлов
    for fpath, fname in [(spec_file, 'Спецификация'), (io_file, 'Система ввода-вывода'), (vars_file, 'Список переменных')]:
        if not os.path.exists(fpath):
            return jsonify({'error': f'Файл не найден: {fname}'}), 404
    
    try:
        # Создаем генератор с тремя файлами
        generator = ExcelToSCLConverter(
            spec_file=spec_file,
            io_file=io_file,
            vars_file=vars_file
        )
        
        # Анализируем workbook
        workbook_info = generator.analyze_workbook()
        
        # Генерируем код
        generated_code = generator.generate_full_code()
        
        # Получаем статистику
        stats = generator.get_statistics()
        
        # Сохраняем результат
        output_filename = f"SCL_Generated_{session_id[:8]}.txt"
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
        import traceback
        error_details = traceback.format_exc()
        print("❌ ПОЛНАЯ ОШИБКА:")
        print(error_details)
        print("❌ КРАТКАЯ ОШИБКА:", str(e))
        return jsonify({
            'error': str(e),
            'details': error_details
        }), 500

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
