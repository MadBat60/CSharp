using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace MiniTest
{
    // Главная точка входа в программу
    static class Program
    {
        [STAThread] // Это нужно для работы с формами Windows
        static void Main()
        {
            // Включаем красивое оформление кнопок
            Application.EnableVisualStyles();
            // Настраиваем совместимость отображения текста
            Application.SetCompatibleTextRenderingDefault(false);
            // Запускаем главное окно программы
            Application.Run(new MainForm());
        }
    }

    // Главное окно программы
    public class MainForm : Form
    {
        // ========== ПОЛЯ КЛАССА ==========
        // Здесь хранятся все элементы управления (кнопки, поля ввода и т.д.)
        
        // Вкладки
        private TabControl tabControl;
        private TabPage tabSCL;
        private TabPage tabOther;
        
        // Поля для ввода путей к файлам (вкладка SCL)
        private TextBox txtExcelPath;   // Поле для адреса Excel файла
        private TextBox txtTxtPath;     // Поле для адреса TXT файла
        
        // Счётчики для выбора строк
        private NumericUpDown numStartRow;  // Начальная строка
        private NumericUpDown numEndRow;    // Конечная строка
        
        // Кнопки
        private Button btnBrowseExcel;  // Кнопка выбора Excel файла
        private Button btnBrowseTxt;    // Кнопка выбора TXT файла
        private Button btnGenerate;     // Кнопка генерации
        private Button btnExit;         // Кнопка выхода

        // Поле для вывода лога (сообщений о работе)
        private RichTextBox rtbLog;
        
        // Строка статуса и полоска прогресса
        private Label lblStatus;
        private ProgressBar progressBar;
        
        // Список всех устройств, которые нужно обработать
        private List<Device> devices;
        
        // Поля для второй вкладки
        private TextBox txtExcelPath2;
        private TextBox txtTxtPath2;
        private Button btnBrowseExcel2;
        private Button btnBrowseTxt2;
        private NumericUpDown numStartRow2;
        private NumericUpDown numEndRow2;
        private Button btnGenerate2;
        private RichTextBox rtbLog2;

        // ========== КОНСТРУКТОР ==========
        // Это метод, который вызывается при создании окна
        public MainForm()
        {
            // Создаём все кнопки, поля и настраиваем внешний вид
            InitializeComponent();
            // Заполняем список устройств
            InitializeDevices();
        }

        // ========== НАСТРОЙКА ВНЕШНЕГО ВИДА ==========
        private void InitializeComponent()
        {
            // ----- Настройки самого окна -----
            this.Text = "Excel to SCL Конвертер";  // Заголовок окна
            this.Size = new Size(700, 650);         // Размер окна
            this.StartPosition = FormStartPosition.CenterScreen;  // По центру экрана
            this.FormBorderStyle = FormBorderStyle.FixedDialog;   // Нельзя менять размер
            this.MaximizeBox = false;                // Отключаем кнопку "Развернуть"

            // ----- Создаём вкладки -----
            tabControl = new TabControl();
            tabControl.Location = new Point(10, 10);
            tabControl.Size = new Size(660, 520);
            
            tabSCL = new TabPage();
            tabSCL.Text = "Спецификация";
            
            tabOther = new TabPage();
            tabOther.Text = "Ввод-вывод";
            
            tabControl.Controls.Add(tabSCL);
            tabControl.Controls.Add(tabOther);
            
            // ----- Настройка первой вкладки (SCL) -----
            // Поле для Excel файла
            Label lblExcelPath = new Label();
            lblExcelPath.Text = "Excel файл (XLSX):";
            lblExcelPath.Location = new Point(10, 15);
            lblExcelPath.Size = new Size(120, 25);
            
            txtExcelPath = new TextBox();
            txtExcelPath.Location = new Point(140, 15);
            txtExcelPath.Size = new Size(400, 25);
            
            btnBrowseExcel = new Button();
            btnBrowseExcel.Text = "...";
            btnBrowseExcel.Location = new Point(550, 15);
            btnBrowseExcel.Size = new Size(35, 25);
            btnBrowseExcel.Click += BtnBrowseExcel_Click;

            // Поле для TXT файла
            Label lblTxtPath = new Label();
            lblTxtPath.Text = "TXT файл:";
            lblTxtPath.Location = new Point(10, 50);
            lblTxtPath.Size = new Size(120, 25);
            
            txtTxtPath = new TextBox();
            txtTxtPath.Location = new Point(140, 50);
            txtTxtPath.Size = new Size(400, 25);
            
            btnBrowseTxt = new Button();
            btnBrowseTxt.Text = "...";
            btnBrowseTxt.Location = new Point(550, 50);
            btnBrowseTxt.Size = new Size(35, 25);
            btnBrowseTxt.Click += BtnBrowseTxt_Click;

            // Выбор диапазона строк
            Label lblStartRow = new Label();
            lblStartRow.Text = "Начальная строка:";
            lblStartRow.Location = new Point(10, 85);
            lblStartRow.Size = new Size(110, 25);
            
            numStartRow = new NumericUpDown();
            numStartRow.Location = new Point(130, 85);
            numStartRow.Size = new Size(60, 25);
            numStartRow.Minimum = 1;
            numStartRow.Maximum = 200;
            numStartRow.Value = 8;

            Label lblEndRow = new Label();
            lblEndRow.Text = "Конечная строка:";
            lblEndRow.Location = new Point(210, 85);
            lblEndRow.Size = new Size(100, 25);
            
            numEndRow = new NumericUpDown();
            numEndRow.Location = new Point(320, 85);
            numEndRow.Size = new Size(60, 25);
            numEndRow.Minimum = 1;
            numEndRow.Maximum = 200;
            numEndRow.Value = 46;

            // Кнопки управления
            btnGenerate = new Button();
            btnGenerate.Text = "Сгенерировать";
            btnGenerate.Location = new Point(265, 125);
            btnGenerate.Size = new Size(130, 35);
            btnGenerate.BackColor = Color.FromArgb(76, 175, 80);
            btnGenerate.ForeColor = Color.White;
            btnGenerate.Click += BtnGenerate_Click;
            
            btnExit = new Button();
            btnExit.Text = "Выход";
            btnExit.Location = new Point(545, 490);
            btnExit.Size = new Size(100, 35);
            btnExit.BackColor = Color.FromArgb(244, 67, 54);
            btnExit.ForeColor = Color.White;
            btnExit.Click += (s, e) => Application.Exit();

            // Лог выполнения
            Label lblLog = new Label();
            lblLog.Text = "Лог выполнения:";
            lblLog.Location = new Point(10, 175);
            lblLog.Size = new Size(120, 20);
            
            rtbLog = new RichTextBox();
            rtbLog.Location = new Point(10, 195);
            rtbLog.Size = new Size(640, 280);
            rtbLog.ReadOnly = true;
            rtbLog.BackColor = Color.Black;
            rtbLog.ForeColor = Color.LightGreen;
            rtbLog.Font = new Font("Consolas", 9);

            // Добавляем элементы на первую вкладку
            tabSCL.Controls.AddRange(new Control[] {
                lblExcelPath, txtExcelPath, btnBrowseExcel,
                lblTxtPath, txtTxtPath, btnBrowseTxt,
                lblStartRow, numStartRow, lblEndRow, numEndRow,
                btnGenerate,
                lblLog, rtbLog
            });
            
            // ----- Настройка второй вкладки -----
            Label lblExcelPath2 = new Label();
            lblExcelPath2.Text = "Excel файл (XLSX):";
            lblExcelPath2.Location = new Point(10, 15);
            lblExcelPath2.Size = new Size(120, 25);
            
            txtExcelPath2 = new TextBox();
            txtExcelPath2.Location = new Point(140, 15);
            txtExcelPath2.Size = new Size(400, 25);
            
            btnBrowseExcel2 = new Button();
            btnBrowseExcel2.Text = "...";
            btnBrowseExcel2.Location = new Point(550, 15);
            btnBrowseExcel2.Size = new Size(35, 25);
            btnBrowseExcel2.Click += BtnBrowseExcel2_Click;
            
            Label lblTxtPath2 = new Label();
            lblTxtPath2.Text = "TXT файл:";
            lblTxtPath2.Location = new Point(10, 50);
            lblTxtPath2.Size = new Size(120, 25);
            
            txtTxtPath2 = new TextBox();
            txtTxtPath2.Location = new Point(140, 50);
            txtTxtPath2.Size = new Size(400, 25);
            
            btnBrowseTxt2 = new Button();
            btnBrowseTxt2.Text = "...";
            btnBrowseTxt2.Location = new Point(550, 50);
            btnBrowseTxt2.Size = new Size(35, 25);
            btnBrowseTxt2.Click += BtnBrowseTxt2_Click;

            // Выбор диапазона строк для второй вкладки
            Label lblStartRow2 = new Label();
            lblStartRow2.Text = "Начальная строка:";
            lblStartRow2.Location = new Point(10, 85);
            lblStartRow2.Size = new Size(110, 25);
            
            numStartRow2 = new NumericUpDown();
            numStartRow2.Location = new Point(130, 85);
            numStartRow2.Size = new Size(60, 25);
            numStartRow2.Minimum = 1;
            numStartRow2.Maximum = 200;
            numStartRow2.Value = 8;

            Label lblEndRow2 = new Label();
            lblEndRow2.Text = "Конечная строка:";
            lblEndRow2.Location = new Point(210, 85);
            lblEndRow2.Size = new Size(100, 25);
            
            numEndRow2 = new NumericUpDown();
            numEndRow2.Location = new Point(320, 85);
            numEndRow2.Size = new Size(60, 25);
            numEndRow2.Minimum = 1;
            numEndRow2.Maximum = 200;
            numEndRow2.Value = 46;

            // Кнопка генерации для второй вкладки
            btnGenerate2 = new Button();
            btnGenerate2.Text = "Сгенерировать";
            btnGenerate2.Location = new Point(265, 125);
            btnGenerate2.Size = new Size(130, 35);
            btnGenerate2.BackColor = Color.FromArgb(76, 175, 80);
            btnGenerate2.ForeColor = Color.White;
            btnGenerate2.Click += BtnGenerate2_Click;

            // Лог выполнения для второй вкладки
            Label lblLog2 = new Label();
            lblLog2.Text = "Лог выполнения:";
            lblLog2.Location = new Point(10, 175);
            lblLog2.Size = new Size(120, 20);
            
            rtbLog2 = new RichTextBox();
            rtbLog2.Location = new Point(10, 195);
            rtbLog2.Size = new Size(640, 280);
            rtbLog2.ReadOnly = true;
            rtbLog2.BackColor = Color.Black;
            rtbLog2.ForeColor = Color.LightGreen;
            rtbLog2.Font = new Font("Consolas", 9);
            
            // Добавляем элементы на вторую вкладку
            tabOther.Controls.AddRange(new Control[] {
                lblExcelPath2, txtExcelPath2, btnBrowseExcel2,
                lblTxtPath2, txtTxtPath2, btnBrowseTxt2,
                lblStartRow2, numStartRow2, lblEndRow2, numEndRow2,
                btnGenerate2,
                lblLog2, rtbLog2
            });
            
            // ----- Статус бар -----
            lblStatus = new Label();
            lblStatus.Text = "Готов к работе";
            lblStatus.Location = new Point(10, 540);
            lblStatus.Size = new Size(400, 25);
            
            progressBar = new ProgressBar();
            progressBar.Location = new Point(420, 540);
            progressBar.Size = new Size(250, 20);
            progressBar.Visible = false;
            progressBar.Style = ProgressBarStyle.Marquee;

            // ----- Добавляем все элементы на форму -----
            this.Controls.AddRange(new Control[] {
                tabControl,
                btnExit,
                lblStatus, progressBar
            });
        }

        // ========== ЗАПОЛНЯЕМ СПИСОК УСТРОЙСТВ ==========
        private void InitializeDevices()
        {
            // Создаём список устройств
            // Каждое устройство имеет: имя, комментарий, номер колонки для типа, номер колонки для имени
            devices = new List<Device>();
            
            // Добавляем устройства одно за другим
            devices.Add(new Device("Doliv", "Долив", 13, 14));
            devices.Add(new Device("Tmpr", "Температура", 15, 16));
            devices.Add(new Device("Cover", "Крышка", 17, 18));
            devices.Add(new Device("Jr", "Жироуловитель", 19, 20));
            devices.Add(new Device("Mix", "Перемешивание", 21, 22));
            devices.Add(new Device("Vip", "Выпрямитель", 23, 24));
            devices.Add(new Device("Filtr", "Фильтрование", 25, 26));
            devices.Add(new Device("Doser", "Дозирование", 27, 28));
            devices.Add(new Device("Shower", "Душирование", 29, 30));
            devices.Add(new Device("Pok", "Качание", 31, 32));
            devices.Add(new Device("Dry", "Сушилка", 33, 34));
            devices.Add(new Device("SafetyBar", "Барьер безопасности", 35, 36));
            devices.Add(new Device("Sink", "Слив", 37, 38));
            devices.Add(new Device("Blower", "Воздуходувка", 39, 40));
            devices.Add(new Device("BarRot", "Вращение барабанов", 41, 42));
            devices.Add(new Device("Chiller", "Чиллер", 43, 44));
            devices.Add(new Device("Lifter", "Подъемник", 45, 46));
            devices.Add(new Device("Vent", "Вентиляция", 47, 48));
        }

        // ========== ОБРАБОТЧИКИ КНОПОК ==========
        
        // Кнопка выбора Excel файла
        private void BtnBrowseExcel_Click(object sender, EventArgs e)
        {
            // Создаём диалог выбора файла
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                // Настраиваем фильтр: показываем только xlsx файлы
                dlg.Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                dlg.FilterIndex = 1;  // Выбираем первый фильтр по умолчанию
                
                // Если пользователь выбрал файл и нажал "ОК"
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    // Запоминаем путь к файлу
                    txtExcelPath.Text = dlg.FileName;
                    // Пишем в лог
                    Log($"Выбран файл: {dlg.FileName}");
                }
            }
        }
        
        // Кнопка выбора TXT файла (куда сохранять результат)
        private void BtnBrowseTxt_Click(object sender, EventArgs e)
        {
            // Создаём диалог сохранения файла
            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Filter = "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*";
                dlg.FilterIndex = 1;
                dlg.DefaultExt = "txt";  // Расширение по умолчанию
                
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    txtTxtPath.Text = dlg.FileName;
                    Log($"Файл будет сохранен: {dlg.FileName}");
                }
            }
        }
        
        // Кнопка выбора Excel файла для второй вкладки
        private void BtnBrowseExcel2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                dlg.FilterIndex = 1;
                
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    txtExcelPath2.Text = dlg.FileName;
                    Log($"Выбран файл для вкладки 2: {dlg.FileName}");
                }
            }
        }
        
        // Кнопка выбора TXT файла для второй вкладки
        private void BtnBrowseTxt2_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Filter = "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*";
                dlg.FilterIndex = 1;
                dlg.DefaultExt = "txt";
                
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    txtTxtPath2.Text = dlg.FileName;
                    Log($"Файл для вкладки 2 будет сохранен: {dlg.FileName}");
                }
            }
        }
        
        // Главная кнопка "Сгенерировать"
        private void BtnGenerate_Click(object sender, EventArgs e)
        {
            // ===== ПРОВЕРКИ =====
            // Проверяем, выбран ли Excel файл
            if (string.IsNullOrEmpty(txtExcelPath.Text))
            {
                MessageBox.Show("Выберите Excel файл!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;  // Выходим из метода
            }
            
            // Проверяем, выбран ли TXT файл
            if (string.IsNullOrEmpty(txtTxtPath.Text))
            {
                MessageBox.Show("Выберите путь для сохранения!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            // Проверяем, существует ли Excel файл
            if (!File.Exists(txtExcelPath.Text))
            {
                MessageBox.Show("Excel файл не найден!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            // Проверяем, что начальная строка не больше конечной
            if (numStartRow.Value > numEndRow.Value)
            {
                MessageBox.Show("Начальная строка не может быть больше конечной!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            // ===== НАЧИНАЕМ ГЕНЕРАЦИЮ =====
            // Блокируем кнопку генерации, чтобы не нажали повторно
            btnGenerate.Enabled = false;
            // Показываем полоску прогресса
            progressBar.Visible = true;
            // Пишем в лог
            Log("Начало генерации...");
            
            try
            {
                // Генерируем SCL код (это строка с результатом)
                string result = GenerateSCL();
                
                // Сохраняем результат в TXT файл
                File.WriteAllText(txtTxtPath.Text, result, Encoding.UTF8);
                
                // Прячем полоску прогресса
                progressBar.Visible = false;
                // Пишем об успешном завершении
                Log("✅ Генерация успешно завершена!");
                Log($"📄 Файл сохранен: {txtTxtPath.Text}");
                Log($"📊 Всего записей: {CountRecords(result)}");
                
                // Спрашиваем, открыть ли файл
                if (MessageBox.Show("Открыть полученный файл?", "Готово", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    // Открываем файл в блокноте
                    System.Diagnostics.Process.Start("notepad.exe", txtTxtPath.Text);
                }
            }
            catch (Exception ex)
            {
                // Если произошла ошибка, пишем в лог и показываем сообщение
                Log($"❌ Ошибка: {ex.Message}");
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Этот блок выполнится в любом случае (и при успехе, и при ошибке)
                // Разблокируем кнопку и прячем прогресс
                btnGenerate.Enabled = true;
                progressBar.Visible = false;
            }
        }
        
        // Главная кнопка "Сгенерировать" для второй вкладки
        private void BtnGenerate2_Click(object sender, EventArgs e)
        {
            // ===== ПРОВЕРКИ =====
            // Проверяем, выбран ли Excel файл
            if (string.IsNullOrEmpty(txtExcelPath2.Text))
            {
                MessageBox.Show("Выберите Excel файл!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            // Проверяем, выбран ли TXT файл
            if (string.IsNullOrEmpty(txtTxtPath2.Text))
            {
                MessageBox.Show("Выберите путь для сохранения!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            // Проверяем, существует ли Excel файл
            if (!File.Exists(txtExcelPath2.Text))
            {
                MessageBox.Show("Excel файл не найден!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            // Проверяем, что начальная строка не больше конечной
            if (numStartRow2.Value > numEndRow2.Value)
            {
                MessageBox.Show("Начальная строка не может быть больше конечной!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            // ===== НАЧИНАЕМ ГЕНЕРАЦИЮ =====
            // Блокируем кнопку генерации, чтобы не нажали повторно
            btnGenerate2.Enabled = false;
            // Показываем полоску прогресса
            progressBar.Visible = true;
            // Пишем в лог
            Log("🚀 Запуск генерации...");
            
            try
            {
                // Генерируем SCL код
                string sclCode = GenerateSCL2();
                
                // Сохраняем результат в файл
                File.WriteAllText(txtTxtPath2.Text, sclCode, Encoding.UTF8);
                Log($"✅ Файл сохранен: {txtTxtPath2.Text}");
                
                // Считаем количество записей
                int recordCount = CountRecords(sclCode);
                Log($"📊 Всего сгенерировано записей: {recordCount}");
                
                // Предлагаем открыть файл
                if (MessageBox.Show("Генерация завершена! Открыть файл?", "Успех",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start("notepad.exe", txtTxtPath2.Text);
                }
            }
            catch (Exception ex)
            {
                Log($"❌ Ошибка: {ex.Message}");
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnGenerate2.Enabled = true;
                progressBar.Visible = false;
            }
        }
        
        // ========== ГЛАВНАЯ ЛОГИКА ГЕНЕРАЦИИ ==========
        private string GenerateSCL()
        {
            // StringBuilder - специальный класс для эффективного создания строк
            StringBuilder result = new StringBuilder();
            
            // Получаем номера строк из полей ввода
            int startRow = (int)numStartRow.Value;
            int endRow = (int)numEndRow.Value;
            
            // Добавляем заголовок в результат
            result.AppendLine("// SCL код сгенерирован автоматически");
            result.AppendLine($"// Дата генерации: {DateTime.Now}");
            result.AppendLine($"// Диапазон строк: {startRow} - {endRow}");
            result.AppendLine($"// Файл источник: {Path.GetFileName(txtExcelPath.Text)}");
            result.AppendLine();  // Пустая строка для красоты
            
            // Открываем Excel файл
            // using гарантирует, что файл будет закрыт даже при ошибке
            using (FileStream fs = new FileStream(txtExcelPath.Text, FileMode.Open, FileAccess.Read))
            using (XSSFWorkbook workbook = new XSSFWorkbook(fs))
            {
                // Берем первый лист в книге (индекс 0)
                ISheet sheet = workbook.GetSheetAt(0);
                Log($"📋 Работаем с листом: {sheet.SheetName}");
                
                int totalRecords = 0;  // Счётчик обработанных записей
                
                // Проходим по всем устройствам из списка
                foreach (var device in devices)
                {
                    // Добавляем комментарий к устройству
                    result.AppendLine($"// {device.Comment}");
                    
                    // Для каждого устройства перебираем строки от startRow до endRow
                    for (int rowNum = startRow; rowNum <= endRow; rowNum++)
                    {
                        // Получаем строку из Excel по номеру
                        IRow row = sheet.GetRow(rowNum);
                        if (row == null) continue;  // Если строки нет, пропускаем
                        
                        // Получаем значение из колонки 0 (место)
                        CellValueInfo placeInfo = GetCellValueInfo(row.GetCell(0));
                        // Если место пустое, пропускаем эту строку
                        if (string.IsNullOrEmpty(placeInfo.Value)) continue;
                        
                        // Получаем тип из колонки device.FirstCol
                        CellValueInfo typeInfo = GetCellValueInfo(row.GetCell(device.FirstCol));
                        // Получаем имя из колонки device.SecondCol
                        CellValueInfo nameInfo = GetCellValueInfo(row.GetCell(device.SecondCol));
                        
                        // Если тип или имя пустые, пропускаем
                        if (string.IsNullOrEmpty(typeInfo.Value) || string.IsNullOrEmpty(nameInfo.Value))
                            continue;
                        
                        // Форматируем значения:
                        // Если значение числовое, пишем без кавычек, иначе в кавычках
                        string placeFormatted = placeInfo.IsNumeric ? placeInfo.Value : $"\"{placeInfo.Value}\"";
                        string typeFormatted = typeInfo.IsNumeric ? typeInfo.Value : $"\"{typeInfo.Value}\"";
                        string nameFormatted = nameInfo.IsNumeric ? nameInfo.Value : $"\"{nameInfo.Value}\"";
                        
                        // Добавляем строки в результат
                        result.AppendLine($"\"{device.Name}\".Dev[{nameFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"{device.Name}\".Dev[{nameFormatted}].CfgType := {typeFormatted};");
                        
                        totalRecords++;  // Увеличиваем счётчик
                    }
                    
                    result.AppendLine();  // Пустая строка между устройствами
                }
                
                Log($"✅ Обработано записей: {totalRecords}");
            }
            
            // Возвращаем полученный текст
            return result.ToString();
        }
        
        // Генерация для второй вкладки (Ввод-вывод)
        private string GenerateSCL2()
        {
            StringBuilder result = new StringBuilder();
            int startRow = (int)numStartRow2.Value;
            int endRow = (int)numEndRow2.Value;
            
            result.AppendLine("// SCL код сгенерирован автоматически (вкладка 2)");
            result.AppendLine($"// Дата генерации: {DateTime.Now}");
            result.AppendLine($"// Диапазон строк: {startRow} - {endRow}");
            result.AppendLine($"// Файл источник: {Path.GetFileName(txtExcelPath2.Text)}");
            result.AppendLine();
            
            using (FileStream fs = new FileStream(txtExcelPath2.Text, FileMode.Open, FileAccess.Read))
            using (XSSFWorkbook workbook = new XSSFWorkbook(fs))
            {
                ISheet sheet = workbook.GetSheetAt(0);
                Log2($"📋 Работаем с листом: {sheet.SheetName}");
                
                int totalRecords = 0;
                
                // Проходим по строкам от startRow до endRow
                for (int rowNum = startRow; rowNum <= endRow; rowNum++)
                {
                    IRow row = sheet.GetRow(rowNum);
                    if (row == null) continue;
                    
                    // Получаем значение из столбца A (индекс 0) - CfgPlace
                    CellValueInfo placeInfo = GetCellValueInfo(row.GetCell(0));
                    // Пропускаем пустые строки
                    if (string.IsNullOrEmpty(placeInfo.Value)) continue;
                    
                    string placeFormatted = placeInfo.IsNumeric ? placeInfo.Value : $"\"{placeInfo.Value}\"";
                    
                    // Doliv: Dev[столбец N(13)], CfgType из столбца M(12)
                    CellValueInfo dolivIndexInfo = GetCellValueInfo(row.GetCell(13));
                    CellValueInfo dolivTypeInfo = GetCellValueInfo(row.GetCell(12));
                    if (!string.IsNullOrEmpty(dolivIndexInfo.Value) && !string.IsNullOrEmpty(dolivTypeInfo.Value))
                    {
                        string dolivIndexFormatted = dolivIndexInfo.IsNumeric ? dolivIndexInfo.Value : $"\"{dolivIndexInfo.Value}\"";
                        string dolivTypeFormatted = dolivTypeInfo.IsNumeric ? dolivTypeInfo.Value : $"\"{dolivTypeInfo.Value}\"";
                        result.AppendLine($"\"Doliv\".Dev[{dolivIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"Doliv\".Dev[{dolivIndexFormatted}].CfgType := {dolivTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    // Tmpr: Dev[столбец P(15)], CfgType из столбца O(14)
                    CellValueInfo tmprIndexInfo = GetCellValueInfo(row.GetCell(15));
                    CellValueInfo tmprTypeInfo = GetCellValueInfo(row.GetCell(14));
                    if (!string.IsNullOrEmpty(tmprIndexInfo.Value) && !string.IsNullOrEmpty(tmprTypeInfo.Value))
                    {
                        string tmprIndexFormatted = tmprIndexInfo.IsNumeric ? tmprIndexInfo.Value : $"\"{tmprIndexInfo.Value}\"";
                        string tmprTypeFormatted = tmprTypeInfo.IsNumeric ? tmprTypeInfo.Value : $"\"{tmprTypeInfo.Value}\"";
                        result.AppendLine($"\"Tmpr\".Dev[{tmprIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"Tmpr\".Dev[{tmprIndexFormatted}].CfgType := {tmprTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    // Cover: Dev[столбец R(17)], CfgType из столбца Q(16)
                    CellValueInfo coverIndexInfo = GetCellValueInfo(row.GetCell(17));
                    CellValueInfo coverTypeInfo = GetCellValueInfo(row.GetCell(16));
                    if (!string.IsNullOrEmpty(coverIndexInfo.Value) && !string.IsNullOrEmpty(coverTypeInfo.Value))
                    {
                        string coverIndexFormatted = coverIndexInfo.IsNumeric ? coverIndexInfo.Value : $"\"{coverIndexInfo.Value}\"";
                        string coverTypeFormatted = coverTypeInfo.IsNumeric ? coverTypeInfo.Value : $"\"{coverTypeInfo.Value}\"";
                        result.AppendLine($"\"Cover\".Dev[{coverIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"Cover\".Dev[{coverIndexFormatted}].CfgType := {coverTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    // Jr: Dev[столбец T(19)], CfgType из столбца S(18)
                    CellValueInfo jrIndexInfo = GetCellValueInfo(row.GetCell(19));
                    CellValueInfo jrTypeInfo = GetCellValueInfo(row.GetCell(18));
                    if (!string.IsNullOrEmpty(jrIndexInfo.Value) && !string.IsNullOrEmpty(jrTypeInfo.Value))
                    {
                        string jrIndexFormatted = jrIndexInfo.IsNumeric ? jrIndexInfo.Value : $"\"{jrIndexInfo.Value}\"";
                        string jrTypeFormatted = jrTypeInfo.IsNumeric ? jrTypeInfo.Value : $"\"{jrTypeInfo.Value}\"";
                        result.AppendLine($"\"Jr\".Dev[{jrIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"Jr\".Dev[{jrIndexFormatted}].CfgType := {jrTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    // Mixer: Dev[столбец V(21)], CfgType из столбца U(20)
                    CellValueInfo mixerIndexInfo = GetCellValueInfo(row.GetCell(21));
                    CellValueInfo mixerTypeInfo = GetCellValueInfo(row.GetCell(20));
                    if (!string.IsNullOrEmpty(mixerIndexInfo.Value) && !string.IsNullOrEmpty(mixerTypeInfo.Value))
                    {
                        string mixerIndexFormatted = mixerIndexInfo.IsNumeric ? mixerIndexInfo.Value : $"\"{mixerIndexInfo.Value}\"";
                        string mixerTypeFormatted = mixerTypeInfo.IsNumeric ? mixerTypeInfo.Value : $"\"{mixerTypeInfo.Value}\"";
                        result.AppendLine($"\"Mixer\".Dev[{mixerIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"Mixer\".Dev[{mixerIndexFormatted}].CfgType := {mixerTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    // Vip: Dev[столбец X(23)], CfgType из столбца W(22)
                    CellValueInfo vipIndexInfo = GetCellValueInfo(row.GetCell(23));
                    CellValueInfo vipTypeInfo = GetCellValueInfo(row.GetCell(22));
                    if (!string.IsNullOrEmpty(vipIndexInfo.Value) && !string.IsNullOrEmpty(vipTypeInfo.Value))
                    {
                        string vipIndexFormatted = vipIndexInfo.IsNumeric ? vipIndexInfo.Value : $"\"{vipIndexInfo.Value}\"";
                        string vipTypeFormatted = vipTypeInfo.IsNumeric ? vipTypeInfo.Value : $"\"{vipTypeInfo.Value}\"";
                        result.AppendLine($"\"Vip\".Dev[{vipIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"Vip\".Dev[{vipIndexFormatted}].CfgType := {vipTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    // Filtr: Dev[столбец Z(25)], CfgType из столбца Y(24)
                    CellValueInfo filtrIndexInfo = GetCellValueInfo(row.GetCell(25));
                    CellValueInfo filtrTypeInfo = GetCellValueInfo(row.GetCell(24));
                    if (!string.IsNullOrEmpty(filtrIndexInfo.Value) && !string.IsNullOrEmpty(filtrTypeInfo.Value))
                    {
                        string filtrIndexFormatted = filtrIndexInfo.IsNumeric ? filtrIndexInfo.Value : $"\"{filtrIndexInfo.Value}\"";
                        string filtrTypeFormatted = filtrTypeInfo.IsNumeric ? filtrTypeInfo.Value : $"\"{filtrTypeInfo.Value}\"";
                        result.AppendLine($"\"Filtr\".Dev[{filtrIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"Filtr\".Dev[{filtrIndexFormatted}].CfgType := {filtrTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    // Doser: Dev[столбец AB(27)], CfgType из столбца AA(26)
                    CellValueInfo doserIndexInfo = GetCellValueInfo(row.GetCell(27));
                    CellValueInfo doserTypeInfo = GetCellValueInfo(row.GetCell(26));
                    if (!string.IsNullOrEmpty(doserIndexInfo.Value) && !string.IsNullOrEmpty(doserTypeInfo.Value))
                    {
                        string doserIndexFormatted = doserIndexInfo.IsNumeric ? doserIndexInfo.Value : $"\"{doserIndexInfo.Value}\"";
                        string doserTypeFormatted = doserTypeInfo.IsNumeric ? doserTypeInfo.Value : $"\"{doserTypeInfo.Value}\"";
                        result.AppendLine($"\"Doser\".Dev[{doserIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"Doser\".Dev[{doserIndexFormatted}].CfgType := {doserTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    // Shower: Dev[столбец AD(29)], CfgType из столбца AC(28)
                    CellValueInfo showerIndexInfo = GetCellValueInfo(row.GetCell(29));
                    CellValueInfo showerTypeInfo = GetCellValueInfo(row.GetCell(28));
                    if (!string.IsNullOrEmpty(showerIndexInfo.Value) && !string.IsNullOrEmpty(showerTypeInfo.Value))
                    {
                        string showerIndexFormatted = showerIndexInfo.IsNumeric ? showerIndexInfo.Value : $"\"{showerIndexInfo.Value}\"";
                        string showerTypeFormatted = showerTypeInfo.IsNumeric ? showerTypeInfo.Value : $"\"{showerTypeInfo.Value}\"";
                        result.AppendLine($"\"Shower\".Dev[{showerIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"Shower\".Dev[{showerIndexFormatted}].CfgType := {showerTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    // Pok: Dev[столбец AF(31)], CfgType из столбца AE(30)
                    CellValueInfo pokIndexInfo = GetCellValueInfo(row.GetCell(31));
                    CellValueInfo pokTypeInfo = GetCellValueInfo(row.GetCell(30));
                    if (!string.IsNullOrEmpty(pokIndexInfo.Value) && !string.IsNullOrEmpty(pokTypeInfo.Value))
                    {
                        string pokIndexFormatted = pokIndexInfo.IsNumeric ? pokIndexInfo.Value : $"\"{pokIndexInfo.Value}\"";
                        string pokTypeFormatted = pokTypeInfo.IsNumeric ? pokTypeInfo.Value : $"\"{pokTypeInfo.Value}\"";
                        result.AppendLine($"\"Pok\".Dev[{pokIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"Pok\".Dev[{pokIndexFormatted}].CfgType := {pokTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    // Dry: Dev[столбец AH(33)], CfgType из столбца AG(32)
                    CellValueInfo dryIndexInfo = GetCellValueInfo(row.GetCell(33));
                    CellValueInfo dryTypeInfo = GetCellValueInfo(row.GetCell(32));
                    if (!string.IsNullOrEmpty(dryIndexInfo.Value) && !string.IsNullOrEmpty(dryTypeInfo.Value))
                    {
                        string dryIndexFormatted = dryIndexInfo.IsNumeric ? dryIndexInfo.Value : $"\"{dryIndexInfo.Value}\"";
                        string dryTypeFormatted = dryTypeInfo.IsNumeric ? dryTypeInfo.Value : $"\"{dryTypeInfo.Value}\"";
                        result.AppendLine($"\"Dry\".Dev[{dryIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"Dry\".Dev[{dryIndexFormatted}].CfgType := {dryTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    // SafetyBar: Dev[столбец AJ(35)], CfgType из столбца AI(34)
                    CellValueInfo safetyBarIndexInfo = GetCellValueInfo(row.GetCell(35));
                    CellValueInfo safetyBarTypeInfo = GetCellValueInfo(row.GetCell(34));
                    if (!string.IsNullOrEmpty(safetyBarIndexInfo.Value) && !string.IsNullOrEmpty(safetyBarTypeInfo.Value))
                    {
                        string safetyBarIndexFormatted = safetyBarIndexInfo.IsNumeric ? safetyBarIndexInfo.Value : $"\"{safetyBarIndexInfo.Value}\"";
                        string safetyBarTypeFormatted = safetyBarTypeInfo.IsNumeric ? safetyBarTypeInfo.Value : $"\"{safetyBarTypeInfo.Value}\"";
                        result.AppendLine($"\"SafetyBar\".Dev[{safetyBarIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"SafetyBar\".Dev[{safetyBarIndexFormatted}].CfgType := {safetyBarTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    // Sink: Dev[столбец AL(37)], CfgType из столбца AK(36)
                    CellValueInfo sinkIndexInfo = GetCellValueInfo(row.GetCell(37));
                    CellValueInfo sinkTypeInfo = GetCellValueInfo(row.GetCell(36));
                    if (!string.IsNullOrEmpty(sinkIndexInfo.Value) && !string.IsNullOrEmpty(sinkTypeInfo.Value))
                    {
                        string sinkIndexFormatted = sinkIndexInfo.IsNumeric ? sinkIndexInfo.Value : $"\"{sinkIndexInfo.Value}\"";
                        string sinkTypeFormatted = sinkTypeInfo.IsNumeric ? sinkTypeInfo.Value : $"\"{sinkTypeInfo.Value}\"";
                        result.AppendLine($"\"Sink\".Dev[{sinkIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"Sink\".Dev[{sinkIndexFormatted}].CfgType := {sinkTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    // Blower: Dev[столбец AN(39)], CfgType из столбца AM(38)
                    CellValueInfo blowerIndexInfo = GetCellValueInfo(row.GetCell(39));
                    CellValueInfo blowerTypeInfo = GetCellValueInfo(row.GetCell(38));
                    if (!string.IsNullOrEmpty(blowerIndexInfo.Value) && !string.IsNullOrEmpty(blowerTypeInfo.Value))
                    {
                        string blowerIndexFormatted = blowerIndexInfo.IsNumeric ? blowerIndexInfo.Value : $"\"{blowerIndexInfo.Value}\"";
                        string blowerTypeFormatted = blowerTypeInfo.IsNumeric ? blowerTypeInfo.Value : $"\"{blowerTypeInfo.Value}\"";
                        result.AppendLine($"\"Blower\".Dev[{blowerIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"Blower\".Dev[{blowerIndexFormatted}].CfgType := {blowerTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    // BarRot: Dev[столбец AP(41)], CfgType из столбца AO(40)
                    CellValueInfo barRotIndexInfo = GetCellValueInfo(row.GetCell(41));
                    CellValueInfo barRotTypeInfo = GetCellValueInfo(row.GetCell(40));
                    if (!string.IsNullOrEmpty(barRotIndexInfo.Value) && !string.IsNullOrEmpty(barRotTypeInfo.Value))
                    {
                        string barRotIndexFormatted = barRotIndexInfo.IsNumeric ? barRotIndexInfo.Value : $"\"{barRotIndexInfo.Value}\"";
                        string barRotTypeFormatted = barRotTypeInfo.IsNumeric ? barRotTypeInfo.Value : $"\"{barRotTypeInfo.Value}\"";
                        result.AppendLine($"\"BarRot\".Dev[{barRotIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"BarRot\".Dev[{barRotIndexFormatted}].CfgType := {barRotTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    // Chiller: Dev[столбец AR(43)], CfgType из столбца AQ(42)
                    CellValueInfo chillerIndexInfo = GetCellValueInfo(row.GetCell(43));
                    CellValueInfo chillerTypeInfo = GetCellValueInfo(row.GetCell(42));
                    if (!string.IsNullOrEmpty(chillerIndexInfo.Value) && !string.IsNullOrEmpty(chillerTypeInfo.Value))
                    {
                        string chillerIndexFormatted = chillerIndexInfo.IsNumeric ? chillerIndexInfo.Value : $"\"{chillerIndexInfo.Value}\"";
                        string chillerTypeFormatted = chillerTypeInfo.IsNumeric ? chillerTypeInfo.Value : $"\"{chillerTypeInfo.Value}\"";
                        result.AppendLine($"\"Chiller\".Dev[{chillerIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"Chiller\".Dev[{chillerIndexFormatted}].CfgType := {chillerTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    // Lifter: Dev[столбец AT(45)], CfgType из столбца AS(44)
                    CellValueInfo lifterIndexInfo = GetCellValueInfo(row.GetCell(45));
                    CellValueInfo lifterTypeInfo = GetCellValueInfo(row.GetCell(44));
                    if (!string.IsNullOrEmpty(lifterIndexInfo.Value) && !string.IsNullOrEmpty(lifterTypeInfo.Value))
                    {
                        string lifterIndexFormatted = lifterIndexInfo.IsNumeric ? lifterIndexInfo.Value : $"\"{lifterIndexInfo.Value}\"";
                        string lifterTypeFormatted = lifterTypeInfo.IsNumeric ? lifterTypeInfo.Value : $"\"{lifterTypeInfo.Value}\"";
                        result.AppendLine($"\"Lifter\".Dev[{lifterIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"Lifter\".Dev[{lifterIndexFormatted}].CfgType := {lifterTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    // Vent: Dev[столбец AV(47)], CfgType из столбца AU(46)
                    CellValueInfo ventIndexInfo = GetCellValueInfo(row.GetCell(47));
                    CellValueInfo ventTypeInfo = GetCellValueInfo(row.GetCell(46));
                    if (!string.IsNullOrEmpty(ventIndexInfo.Value) && !string.IsNullOrEmpty(ventTypeInfo.Value))
                    {
                        string ventIndexFormatted = ventIndexInfo.IsNumeric ? ventIndexInfo.Value : $"\"{ventIndexInfo.Value}\"";
                        string ventTypeFormatted = ventTypeInfo.IsNumeric ? ventTypeInfo.Value : $"\"{ventTypeInfo.Value}\"";
                        result.AppendLine($"\"Vent\".Dev[{ventIndexFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"Vent\".Dev[{ventIndexFormatted}].CfgType := {ventTypeFormatted};");
                        totalRecords += 2;
                    }
                    
                    result.AppendLine();
                }
                
                Log2($"✅ Обработано записей: {totalRecords}");
            }
            
            return result.ToString();
        }
        
        // ========== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ==========
        
        // Подсчитывает количество записей в сгенерированном тексте
        private int CountRecords(string text)
        {
            // Разбиваем текст на строки, считаем те, которые содержат ".CfgPlace :="
            string[] lines = text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            int count = 0;
            foreach (string line in lines)
            {
                if (line.Contains(".CfgPlace :="))
                {
                    count++;
                }
            }
            return count;
        }
        
        // Получает значение ячейки Excel и определяет, число это или текст
        private CellValueInfo GetCellValueInfo(ICell cell)
        {
            // Если ячейка пустая, возвращаем пустую строку
            if (cell == null)
                return new CellValueInfo("", false);
            
            try
            {
                // Смотрим, какой тип данных в ячейке
                switch (cell.CellType)
                {
                    case CellType.String:  // Текст
                        string text = cell.StringCellValue;
                        if (text == null) text = "";
                        return new CellValueInfo(text.Trim(), false);
                        
                    case CellType.Numeric:  // Число
                        double numericValue = cell.NumericCellValue;
                        // Если число целое, показываем без .0
                        string stringValue;
                        if (numericValue == Math.Floor(numericValue))
                        {
                            stringValue = numericValue.ToString("0");
                        }
                        else
                        {
                            // Для дробных используем инвариантную культуру (точка, а не запятая)
                            stringValue = numericValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        }
                        return new CellValueInfo(stringValue, true);
                        
                    case CellType.Boolean:  // Логическое значение
                        return new CellValueInfo(cell.BooleanCellValue.ToString(), false);
                        
                    case CellType.Formula:  // Формула
                        try
                        {
                            // Пытаемся получить результат формулы
                            switch (cell.CachedFormulaResultType)
                            {
                                case CellType.String:
                                    string formulaText = cell.StringCellValue;
                                    if (formulaText == null) formulaText = "";
                                    return new CellValueInfo(formulaText.Trim(), false);
                                case CellType.Numeric:
                                    double formulaValue = cell.NumericCellValue;
                                    string formulaString;
                                    if (formulaValue == Math.Floor(formulaValue))
                                    {
                                        formulaString = formulaValue.ToString("0");
                                    }
                                    else
                                    {
                                        formulaString = formulaValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                                    }
                                    return new CellValueInfo(formulaString, true);
                                case CellType.Boolean:
                                    return new CellValueInfo(cell.BooleanCellValue.ToString(), false);
                                default:
                                    return new CellValueInfo("", false);
                            }
                        }
                        catch
                        {
                            return new CellValueInfo("", false);
                        }
                        
                    default:  // Остальные типы
                        return new CellValueInfo("", false);
                }
            }
            catch
            {
                // Если произошла ошибка, возвращаем пустую строку
                return new CellValueInfo("", false);
            }
        }
        
        // Записывает сообщение в лог и в строку статуса
        private void Log(string message)
        {
            if (rtbLog.InvokeRequired)
            {
                rtbLog.Invoke(new Action<string>(Log), message);
            }
            else
            {
                rtbLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
                rtbLog.ScrollToCaret();
                lblStatus.Text = message;
            }
        }
        
        // Записывает сообщение в лог второй вкладки
        private void Log2(string message)
        {
            if (rtbLog2.InvokeRequired)
            {
                rtbLog2.Invoke(new Action<string>(Log2), message);
            }
            else
            {
                rtbLog2.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
                rtbLog2.ScrollToCaret();
                lblStatus.Text = message;
            }
        }
        
        // ========== ВСПОМОГАТЕЛЬНЫЕ КЛАССЫ ==========
        
        // Класс, описывающий устройство
        class Device
        {
            public string Name { get; set; }      // Имя устройства
            public string Comment { get; set; }   // Комментарий (по-русски)
            public int FirstCol { get; set; }     // Номер колонки для типа
            public int SecondCol { get; set; }    // Номер колонки для имени
            
            // Конструктор - вызывается при создании устройства
            public Device(string name, string comment, int firstCol, int secondCol)
            {
                Name = name;
                Comment = comment;
                FirstCol = firstCol;
                SecondCol = secondCol;
            }
        }
        
        // Класс для хранения значения ячейки и информации, число это или нет
        class CellValueInfo
        {
            public string Value { get; set; }      // Значение в виде строки
            public bool IsNumeric { get; set; }    // Является ли значение числом
            
            public CellValueInfo(string value, bool isNumeric)
            {
                Value = value;
                IsNumeric = isNumeric;
            }
        }
    }
}