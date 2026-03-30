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
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    public class MainForm : Form
    {
        // Поля класса
        private TabControl tabControl;
        private TabPage tabSpec;
        private TabPage tabIO;

        // Элементы вкладки "Спецификация"
        private TextBox txtExcelPath;
        private TextBox txtTxtPath;
        private NumericUpDown numStartRow;
        private NumericUpDown numEndRow;
        private Button btnBrowseExcel;
        private Button btnBrowseTxt;
        private Button btnGenerate;
        private RichTextBox rtbLog;

        // Элементы вкладки "Ввод-вывод" (заготовка)
        private TextBox txtExcelPath2;
        private TextBox txtTxtPath2;
        private NumericUpDown numStartRow2;
        private NumericUpDown numEndRow2;
        private Button btnBrowseExcel2;
        private Button btnBrowseTxt2;
        private Button btnGenerate2;
        private RichTextBox rtbLog2;

        // Общие элементы
        private Button btnExit;
        private Label lblStatus;
        private ProgressBar progressBar;
        private List<Device> devices;

        public MainForm()
        {
            InitializeComponent();
            InitializeDevices();
        }

        private void InitializeComponent()
        {
            this.Text = "Excel to SCL Конвертер";
            this.Size = new Size(900, 750);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;

            tabControl = new TabControl();
            tabControl.Location = new Point(10, 10);
            tabControl.Size = new Size(860, 600);

            tabSpec = new TabPage();
            tabSpec.Text = "Спецификация";
            
            tabIO = new TabPage();
            tabIO.Text = "Ввод-вывод";

            tabControl.Controls.Add(tabSpec);
            tabControl.Controls.Add(tabIO);

            SetupSpecTab();
            SetupIOTab();

            lblStatus = new Label();
            lblStatus.Text = "Готов к работе";
            lblStatus.Location = new Point(10, 620);
            lblStatus.Size = new Size(600, 25);

            progressBar = new ProgressBar();
            progressBar.Location = new Point(620, 620);
            progressBar.Size = new Size(230, 20);
            progressBar.Visible = false;
            progressBar.Style = ProgressBarStyle.Marquee;

            btnExit = new Button();
            btnExit.Text = "Выход";
            btnExit.Location = new Point(740, 615);
            btnExit.Size = new Size(100, 35);
            btnExit.BackColor = Color.FromArgb(244, 67, 54);
            btnExit.ForeColor = Color.White;
            btnExit.Click += (s, e) => Application.Exit();

            this.Controls.AddRange(new Control[] { tabControl, btnExit, lblStatus, progressBar });
        }

        private void SetupSpecTab()
        {
            Label lblExcelPath = new Label();
            lblExcelPath.Text = "Excel файл (XLSX):";
            lblExcelPath.Location = new Point(10, 15);
            lblExcelPath.Size = new Size(120, 25);

            txtExcelPath = new TextBox();
            txtExcelPath.Location = new Point(140, 15);
            txtExcelPath.Size = new Size(580, 25);

            btnBrowseExcel = new Button();
            btnBrowseExcel.Text = "...";
            btnBrowseExcel.Location = new Point(730, 15);
            btnBrowseExcel.Size = new Size(35, 25);
            btnBrowseExcel.Click += BtnBrowseExcel_Click;

            Label lblTxtPath = new Label();
            lblTxtPath.Text = "TXT файл:";
            lblTxtPath.Location = new Point(10, 50);
            lblTxtPath.Size = new Size(120, 25);

            txtTxtPath = new TextBox();
            txtTxtPath.Location = new Point(140, 50);
            txtTxtPath.Size = new Size(580, 25);

            btnBrowseTxt = new Button();
            btnBrowseTxt.Text = "...";
            btnBrowseTxt.Location = new Point(730, 50);
            btnBrowseTxt.Size = new Size(35, 25);
            btnBrowseTxt.Click += BtnBrowseTxt_Click;

            Label lblStartRow = new Label();
            lblStartRow.Text = "Начальная строка:";
            lblStartRow.Location = new Point(10, 85);
            lblStartRow.Size = new Size(110, 25);

            numStartRow = new NumericUpDown();
            numStartRow.Location = new Point(130, 85);
            numStartRow.Size = new Size(60, 25);
            numStartRow.Minimum = 1;
            numStartRow.Maximum = 1000;
            numStartRow.Value = 8;

            Label lblEndRow = new Label();
            lblEndRow.Text = "Конечная строка:";
            lblEndRow.Location = new Point(210, 85);
            lblEndRow.Size = new Size(100, 25);

            numEndRow = new NumericUpDown();
            numEndRow.Location = new Point(320, 85);
            numEndRow.Size = new Size(60, 25);
            numEndRow.Minimum = 1;
            numEndRow.Maximum = 1000;
            numEndRow.Value = 50;

            btnGenerate = new Button();
            btnGenerate.Text = "Сгенерировать";
            btnGenerate.Location = new Point(350, 125);
            btnGenerate.Size = new Size(160, 35);
            btnGenerate.BackColor = Color.FromArgb(76, 175, 80);
            btnGenerate.ForeColor = Color.White;
            btnGenerate.Click += BtnGenerate_Click;

            Label lblLog = new Label();
            lblLog.Text = "Лог выполнения:";
            lblLog.Location = new Point(10, 175);
            lblLog.Size = new Size(120, 20);

            rtbLog = new RichTextBox();
            rtbLog.Location = new Point(10, 195);
            rtbLog.Size = new Size(830, 360);
            rtbLog.ReadOnly = true;
            rtbLog.BackColor = Color.Black;
            rtbLog.ForeColor = Color.LightGreen;
            rtbLog.Font = new Font("Consolas", 9);

            tabSpec.Controls.AddRange(new Control[] {
                lblExcelPath, txtExcelPath, btnBrowseExcel,
                lblTxtPath, txtTxtPath, btnBrowseTxt,
                lblStartRow, numStartRow, lblEndRow, numEndRow,
                btnGenerate, lblLog, rtbLog
            });
        }

        private void SetupIOTab()
        {
            Label lblExcelPath2 = new Label();
            lblExcelPath2.Text = "Excel файл (XLSX):";
            lblExcelPath2.Location = new Point(10, 15);
            lblExcelPath2.Size = new Size(120, 25);

            txtExcelPath2 = new TextBox();
            txtExcelPath2.Location = new Point(140, 15);
            txtExcelPath2.Size = new Size(580, 25);

            btnBrowseExcel2 = new Button();
            btnBrowseExcel2.Text = "...";
            btnBrowseExcel2.Location = new Point(730, 15);
            btnBrowseExcel2.Size = new Size(35, 25);
            btnBrowseExcel2.Click += BtnBrowseExcel2_Click;

            Label lblTxtPath2 = new Label();
            lblTxtPath2.Text = "TXT файл:";
            lblTxtPath2.Location = new Point(10, 50);
            lblTxtPath2.Size = new Size(120, 25);

            txtTxtPath2 = new TextBox();
            txtTxtPath2.Location = new Point(140, 50);
            txtTxtPath2.Size = new Size(580, 25);

            btnBrowseTxt2 = new Button();
            btnBrowseTxt2.Text = "...";
            btnBrowseTxt2.Location = new Point(730, 50);
            btnBrowseTxt2.Size = new Size(35, 25);
            btnBrowseTxt2.Click += BtnBrowseTxt2_Click;

            Label lblStartRow2 = new Label();
            lblStartRow2.Text = "Начальная строка:";
            lblStartRow2.Location = new Point(10, 85);
            lblStartRow2.Size = new Size(110, 25);

            numStartRow2 = new NumericUpDown();
            numStartRow2.Location = new Point(130, 85);
            numStartRow2.Size = new Size(60, 25);
            numStartRow2.Minimum = 1;
            numStartRow2.Maximum = 1000;
            numStartRow2.Value = 8;

            Label lblEndRow2 = new Label();
            lblEndRow2.Text = "Конечная строка:";
            lblEndRow2.Location = new Point(210, 85);
            lblEndRow2.Size = new Size(100, 25);

            numEndRow2 = new NumericUpDown();
            numEndRow2.Location = new Point(320, 85);
            numEndRow2.Size = new Size(60, 25);
            numEndRow2.Minimum = 1;
            numEndRow2.Maximum = 1000;
            numEndRow2.Value = 50;

            btnGenerate2 = new Button();
            btnGenerate2.Text = "Сгенерировать";
            btnGenerate2.Location = new Point(350, 125);
            btnGenerate2.Size = new Size(160, 35);
            btnGenerate2.BackColor = Color.FromArgb(76, 175, 80);
            btnGenerate2.ForeColor = Color.White;
            btnGenerate2.Click += BtnGenerate2_Click;

            Label lblLog2 = new Label();
            lblLog2.Text = "Лог выполнения:";
            lblLog2.Location = new Point(10, 175);
            lblLog2.Size = new Size(120, 20);

            rtbLog2 = new RichTextBox();
            rtbLog2.Location = new Point(10, 195);
            rtbLog2.Size = new Size(830, 360);
            rtbLog2.ReadOnly = true;
            rtbLog2.BackColor = Color.Black;
            rtbLog2.ForeColor = Color.LightGreen;
            rtbLog2.Font = new Font("Consolas", 9);

            tabIO.Controls.AddRange(new Control[] {
                lblExcelPath2, txtExcelPath2, btnBrowseExcel2,
                lblTxtPath2, txtTxtPath2, btnBrowseTxt2,
                lblStartRow2, numStartRow2, lblEndRow2, numEndRow2,
                btnGenerate2, lblLog2, rtbLog2
            });
        }

        private void InitializeDevices()
        {
            devices = new List<Device>();
            // Маппинг: Имя, Комментарий, Индекс колонки Типа, Индекс колонки Dev (Индекса)
            // Столбцы: A(0), ..., M(12), N(13), O(14), P(15) ...
            devices.Add(new Device("Doliv", "Долив", 12, 13));       // M, N
            devices.Add(new Device("Tmpr", "Температура", 14, 15));  // O, P
            devices.Add(new Device("Cover", "Крышка", 16, 17));      // Q, R
            devices.Add(new Device("Jr", "Жироуловитель", 18, 19));  // S, T
            devices.Add(new Device("Mixer", "Перемешивание", 20, 21)); // U, V
            devices.Add(new Device("Vip", "Выпрямитель", 22, 23));   // W, X
            devices.Add(new Device("Filtr", "Фильтрование", 24, 25)); // Y, Z
            devices.Add(new Device("Doser", "Дозирование", 26, 27)); // AA, AB
            devices.Add(new Device("Shower", "Душирование", 28, 29)); // AC, AD
            devices.Add(new Device("Pok", "Качание", 30, 31));       // AE, AF
            devices.Add(new Device("Dry", "Сушилка", 32, 33));       // AG, AH
            devices.Add(new Device("SafetyBar", "Барьер безопасности", 34, 35)); // AI, AJ
            devices.Add(new Device("Sink", "Слив", 36, 37));         // AK, AL
            devices.Add(new Device("Blower", "Воздуходувка", 38, 39)); // AM, AN
            devices.Add(new Device("BarRot", "Вращение барабанов", 40, 41)); // AO, AP
            devices.Add(new Device("Chiller", "Чиллер", 42, 43));    // AQ, AR
            devices.Add(new Device("Lifter", "Подъемник", 44, 45));  // AS, AT
            devices.Add(new Device("Vent", "Вентиляция", 46, 47));   // AU, AV
        }

        private void BtnBrowseExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    txtExcelPath.Text = dlg.FileName;
                    Log($"Выбран файл: {dlg.FileName}");
                }
            }
        }

        private void BtnBrowseTxt_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Filter = "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*";
                dlg.DefaultExt = "txt";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    txtTxtPath.Text = dlg.FileName;
                    Log($"Файл будет сохранен: {dlg.FileName}");
                }
            }
        }

        private void BtnBrowseExcel2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    txtExcelPath2.Text = dlg.FileName;
                    Log($"Выбран файл (вкладка 2): {dlg.FileName}");
                }
            }
        }

        private void BtnBrowseTxt2_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Filter = "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*";
                dlg.DefaultExt = "txt";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    txtTxtPath2.Text = dlg.FileName;
                    Log($"Файл (вкладка 2) будет сохранен: {dlg.FileName}");
                }
            }
        }

        private void BtnGenerate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtExcelPath.Text))
            {
                MessageBox.Show("Выберите Excel файл!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrEmpty(txtTxtPath.Text))
            {
                MessageBox.Show("Выберите путь для сохранения!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!File.Exists(txtExcelPath.Text))
            {
                MessageBox.Show("Excel файл не найден!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (numStartRow.Value > numEndRow.Value)
            {
                MessageBox.Show("Начальная строка не может быть больше конечной!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnGenerate.Enabled = false;
            progressBar.Visible = true;
            Log("Начало генерации (Спецификация)...");

            try
            {
                string result = GenerateSCL();
                File.WriteAllText(txtTxtPath.Text, result, Encoding.UTF8);
                Log("✅ Генерация успешно завершена!");
                Log($"📄 Файл сохранен: {txtTxtPath.Text}");

                if (MessageBox.Show("Открыть полученный файл?", "Готово",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start("notepad.exe", txtTxtPath.Text);
                }
            }
            catch (Exception ex)
            {
                Log($"❌ Ошибка: {ex.Message}");
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnGenerate.Enabled = true;
                progressBar.Visible = false;
            }
        }

        private void BtnGenerate2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Функционал вкладки 'Ввод-вывод' будет добавлен позже.", "Информация",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private string GenerateSCL()
        {
            StringBuilder result = new StringBuilder();
            int startRow = (int)numStartRow.Value;
            int endRow = (int)numEndRow.Value;

            result.AppendLine("// SCL код сгенерирован автоматически");
            result.AppendLine($"// Дата генерации: {DateTime.Now}");
            result.AppendLine($"// Диапазон строк: {startRow} - {endRow}");
            result.AppendLine($"// Файл источник: {Path.GetFileName(txtExcelPath.Text)}");
            result.AppendLine();

            // Словарь для хранения максимальных значений индексов по устройствам
            Dictionary<string, int> maxIndices = new Dictionary<string, int>();
            foreach (var dev in devices)
            {
                maxIndices[dev.Name] = 0;
            }

            using (FileStream fs = new FileStream(txtExcelPath.Text, FileMode.Open, FileAccess.Read))
            using (XSSFWorkbook workbook = new XSSFWorkbook(fs))
            {
                ISheet sheet = workbook.GetSheetAt(0);
                Log($"📋 Работаем с листом: {sheet.SheetName}");

                int totalRecords = 0;

                // Первый проход: сбор статистики максимумов и генерация кода
                foreach (var device in devices)
                {
                    result.AppendLine($"// {device.Comment}");

                    for (int rowNum = startRow; rowNum <= endRow; rowNum++)
                    {
                        IRow row = sheet.GetRow(rowNum);
                        if (row == null) continue;

                        // Столбец A (индекс 0) - место (CfgPlace)
                        CellValueInfo placeInfo = GetCellValueInfo(row.GetCell(0));
                        if (string.IsNullOrEmpty(placeInfo.Value)) continue;

                        // Столбец типа (FirstCol)
                        CellValueInfo typeInfo = GetCellValueInfo(row.GetCell(device.FirstCol));
                        
                        // Столбец индекса/имени (SecondCol) - именно отсюда берем максимум
                        CellValueInfo nameInfo = GetCellValueInfo(row.GetCell(device.SecondCol));

                        if (string.IsNullOrEmpty(typeInfo.Value) || string.IsNullOrEmpty(nameInfo.Value))
                            continue;

                        // Обновляем максимум, если текущее значение больше
                        if (nameInfo.IsNumeric)
                        {
                            if (double.TryParse(nameInfo.Value, out double val))
                            {
                                int intVal = (int)val;
                                if (intVal > maxIndices[device.Name])
                                {
                                    maxIndices[device.Name] = intVal;
                                }
                            }
                        }

                        string placeFormatted = placeInfo.IsNumeric ? placeInfo.Value : $"\"{placeInfo.Value}\"";
                        string typeFormatted = typeInfo.IsNumeric ? typeInfo.Value : $"\"{typeInfo.Value}\"";
                        string nameFormatted = nameInfo.IsNumeric ? nameInfo.Value : $"\"{nameInfo.Value}\"";

                        result.AppendLine($"\"{device.Name}\".Dev[{nameFormatted}].CfgPlace := {placeFormatted};");
                        result.AppendLine($"\"{device.Name}\".Dev[{nameFormatted}].CfgType := {typeFormatted};");

                        totalRecords++;
                    }
                    result.AppendLine();
                }

                Log($"✅ Обработано записей: {totalRecords}");

                // Вставка блока Options.Count в начало результата
                StringBuilder header = new StringBuilder();
                header.AppendLine("\"Options\".Count := RECORD");
                
                foreach (var dev in devices)
                {
                    // Формируем имя переменной: MaxDoliv, MaxTmpr и т.д.
                    string varName = $"Max{dev.Name}";
                    header.AppendLine($"\"Options\".Count.{varName} := {maxIndices[dev.Name]};");
                }
                
                header.AppendLine("END_RECORD;");
                header.AppendLine();

                // Объединяем заголовок и основной код
                return header.ToString() + result.ToString();
            }
        }

        private CellValueInfo GetCellValueInfo(ICell cell)
        {
            if (cell == null) return new CellValueInfo("", false);

            try
            {
                switch (cell.CellType)
                {
                    case CellType.String:
                        string text = cell.StringCellValue;
                        return new CellValueInfo(text == null ? "" : text.Trim(), false);
                    case CellType.Numeric:
                        double val = cell.NumericCellValue;
                        string strVal = (val == Math.Floor(val)) ? val.ToString("0") : val.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        return new CellValueInfo(strVal, true);
                    case CellType.Boolean:
                        return new CellValueInfo(cell.BooleanCellValue.ToString(), false);
                    case CellType.Formula:
                        try
                        {
                            switch (cell.CachedFormulaResultType)
                            {
                                case CellType.String:
                                    return new CellValueInfo(cell.StringCellValue?.Trim() ?? "", false);
                                case CellType.Numeric:
                                    double fVal = cell.NumericCellValue;
                                    string fStr = (fVal == Math.Floor(fVal)) ? fVal.ToString("0") : fVal.ToString(System.Globalization.CultureInfo.InvariantCulture);
                                    return new CellValueInfo(fStr, true);
                                case CellType.Boolean:
                                    return new CellValueInfo(cell.BooleanCellValue.ToString(), false);
                                default:
                                    return new CellValueInfo("", false);
                            }
                        }
                        catch { return new CellValueInfo("", false); }
                    default:
                        return new CellValueInfo("", false);
                }
            }
            catch
            {
                return new CellValueInfo("", false);
            }
        }

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

        class Device
        {
            public string Name { get; set; }
            public string Comment { get; set; }
            public int FirstCol { get; set; }
            public int SecondCol { get; set; }

            public Device(string name, string comment, int firstCol, int secondCol)
            {
                Name = name;
                Comment = comment;
                FirstCol = firstCol;
                SecondCol = secondCol;
            }
        }

        class CellValueInfo
        {
            public string Value { get; set; }
            public bool IsNumeric { get; set; }

            public CellValueInfo(string value, bool isNumeric)
            {
                Value = value;
                IsNumeric = isNumeric;
            }
        }
    }
}