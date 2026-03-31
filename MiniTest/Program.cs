using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
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
        private TabPage tabManual;
        private TabPage tabConfig;

        // Элементы вкладки "Спецификация"
        private TextBox txtExcelPath;
        private TextBox txtTxtPath;
        private NumericUpDown numStartRow;
        private NumericUpDown numEndRow;
        private Button btnBrowseExcel;
        private Button btnBrowseTxt;
        private Button btnGenerate;
        private RichTextBox rtbLog;

        // Элементы вкладки "Ввод-вывод"
        private TextBox txtIO_SystemPath;      // Система ввода-вывода.xlsx
        private TextBox txtIO_SpecPath;        // Спецификация.xlsx
        private TextBox txtIO_VarPath;         // Список переменных от системы ввода-вывода.xlsx
        private TextBox txtIOTxtPath;          // Выходной TXT файл
        private Button btnBrowseIO_System;
        private Button btnBrowseIO_Spec;
        private Button btnBrowseIO_Var;
        private Button btnBrowseIOTxt;
        private Button btnGenerateIO;
        private RichTextBox rtbLogIO;

        // Элементы вкладки "Ручная конфигурация"
        private Dictionary<string, NumericUpDown> manualInputs = new Dictionary<string, NumericUpDown>();
        private Button btnGenerateManual;
        private RichTextBox rtbLogManual;

        // Элементы вкладки "Конфигуратор устройств"
        private NumericUpDown numConfigPlace;
        private ComboBox cmbConfigDevice;
        private NumericUpDown numConfigIndex;
        private NumericUpDown numConfigType;
        private Button btnSaveConfig;
        private Button btnResetConfig;
        private Button btnGenerateConfig;
        private Label lblConfigCount;
        private RichTextBox rtbLogConfig;
        private List<DeviceConfig> savedConfigs = new List<DeviceConfig>();

        // Общие элементы
        private Button btnExit;
        private Label lblStatus;
        private ProgressBar progressBar;
        private List<Device> devices;

        public MainForm()
        {
            InitializeComponent();
            InitializeDevices();
            InitializeManualTab();
            InitializeConfigTab();
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

            tabManual = new TabPage();
            tabManual.Text = "Ручная конфигурация";

            tabConfig = new TabPage();
            tabConfig.Text = "Конфигуратор устройств";

            tabControl.Controls.Add(tabSpec);
            tabControl.Controls.Add(tabIO);
            tabControl.Controls.Add(tabManual);
            tabControl.Controls.Add(tabConfig);

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
            int y = 10;
            int labelWidth = 280;
            int textBoxWidth = 450;
            int btnWidth = 35;
            int startX = 10;
            
            // Система ввода-вывода
            Label lblIO_System = new Label();
            lblIO_System.Text = "Система ввода-вывода.xlsx:";
            lblIO_System.Location = new Point(startX, y);
            lblIO_System.Size = new Size(labelWidth, 25);
            
            txtIO_SystemPath = new TextBox();
            txtIO_SystemPath.Location = new Point(startX + labelWidth, y);
            txtIO_SystemPath.Size = new Size(textBoxWidth, 25);
            
            btnBrowseIO_System = new Button();
            btnBrowseIO_System.Text = "...";
            btnBrowseIO_System.Location = new Point(startX + labelWidth + textBoxWidth + 5, y);
            btnBrowseIO_System.Size = new Size(btnWidth, 25);
            btnBrowseIO_System.Click += BtnBrowseIO_System_Click;
            
            y += 35;
            
            // Спецификация
            Label lblIO_Spec = new Label();
            lblIO_Spec.Text = "Спецификация.xlsx:";
            lblIO_Spec.Location = new Point(startX, y);
            lblIO_Spec.Size = new Size(labelWidth, 25);
            
            txtIO_SpecPath = new TextBox();
            txtIO_SpecPath.Location = new Point(startX + labelWidth, y);
            txtIO_SpecPath.Size = new Size(textBoxWidth, 25);
            
            btnBrowseIO_Spec = new Button();
            btnBrowseIO_Spec.Text = "...";
            btnBrowseIO_Spec.Location = new Point(startX + labelWidth + textBoxWidth + 5, y);
            btnBrowseIO_Spec.Size = new Size(btnWidth, 25);
            btnBrowseIO_Spec.Click += BtnBrowseIO_Spec_Click;
            
            y += 35;
            
            // Список переменных
            Label lblIO_Var = new Label();
            lblIO_Var.Text = "Список переменных.xlsx:";
            lblIO_Var.Location = new Point(startX, y);
            lblIO_Var.Size = new Size(labelWidth, 25);
            
            txtIO_VarPath = new TextBox();
            txtIO_VarPath.Location = new Point(startX + labelWidth, y);
            txtIO_VarPath.Size = new Size(textBoxWidth, 25);
            
            btnBrowseIO_Var = new Button();
            btnBrowseIO_Var.Text = "...";
            btnBrowseIO_Var.Location = new Point(startX + labelWidth + textBoxWidth + 5, y);
            btnBrowseIO_Var.Size = new Size(btnWidth, 25);
            btnBrowseIO_Var.Click += BtnBrowseIO_Var_Click;
            
            y += 35;
            
            // Выходной TXT файл
            Label lblIOTxt = new Label();
            lblIOTxt.Text = "Выходной TXT файл:";
            lblIOTxt.Location = new Point(startX, y);
            lblIOTxt.Size = new Size(labelWidth, 25);
            
            txtIOTxtPath = new TextBox();
            txtIOTxtPath.Location = new Point(startX + labelWidth, y);
            txtIOTxtPath.Size = new Size(textBoxWidth, 25);
            
            btnBrowseIOTxt = new Button();
            btnBrowseIOTxt.Text = "...";
            btnBrowseIOTxt.Location = new Point(startX + labelWidth + textBoxWidth + 5, y);
            btnBrowseIOTxt.Size = new Size(btnWidth, 25);
            btnBrowseIOTxt.Click += BtnBrowseIOTxt_Click;
            
            y += 45;
            
            // Кнопка генерации
            btnGenerateIO = new Button();
            btnGenerateIO.Text = "Сгенерировать SCL";
            btnGenerateIO.Location = new Point(350, y);
            btnGenerateIO.Size = new Size(160, 35);
            btnGenerateIO.BackColor = Color.FromArgb(76, 175, 80);
            btnGenerateIO.ForeColor = Color.White;
            btnGenerateIO.Click += BtnGenerateIO_Click;
            
            y += 50;
            
            // Лог
            Label lblLogIO = new Label();
            lblLogIO.Text = "Лог выполнения:";
            lblLogIO.Location = new Point(startX, y);
            lblLogIO.Size = new Size(120, 20);
            
            rtbLogIO = new RichTextBox();
            rtbLogIO.Location = new Point(startX, y + 20);
            rtbLogIO.Size = new Size(830, 480);
            rtbLogIO.ReadOnly = true;
            rtbLogIO.BackColor = Color.Black;
            rtbLogIO.ForeColor = Color.LightGreen;
            rtbLogIO.Font = new Font("Consolas", 9);
            
            tabIO.Controls.AddRange(new Control[] {
                lblIO_System, txtIO_SystemPath, btnBrowseIO_System,
                lblIO_Spec, txtIO_SpecPath, btnBrowseIO_Spec,
                lblIO_Var, txtIO_VarPath, btnBrowseIO_Var,
                lblIOTxt, txtIOTxtPath, btnBrowseIOTxt,
                btnGenerateIO, lblLogIO, rtbLogIO
            });
        }

        private void InitializeDevices()
        {
            devices = new List<Device>();
            // Маппинг: Имя, Комментарий, Индекс колонки Типа, Индекс колонки Dev (Индекса)
            // Столбцы: A(0), ..., N(13), O(14), P(15), Q(16) ...
            devices.Add(new Device("Doliv", "Долив", 13, 14));       // N, O
            devices.Add(new Device("Tmpr", "Температура", 15, 16));  // O, P -> P, Q
            devices.Add(new Device("Cover", "Крышка", 17, 18));      // Q, R -> R, S
            devices.Add(new Device("Jr", "Жироуловитель", 19, 20));  // S, T -> T, U
            devices.Add(new Device("Mixer", "Перемешивание", 21, 22)); // U, V -> V, W
            devices.Add(new Device("Vip", "Выпрямитель", 23, 24));   // W, X -> X, Y
            devices.Add(new Device("Filtr", "Фильтрование", 25, 26)); // Y, Z -> Z, AA
            devices.Add(new Device("Doser", "Дозирование", 27, 28)); // AA, AB -> AB, AC
            devices.Add(new Device("Shower", "Душирование", 29, 30)); // AC, AD -> AD, AE
            devices.Add(new Device("Pok", "Качание", 31, 32));       // AE, AF -> AF, AG
            devices.Add(new Device("Dry", "Сушилка", 33, 34));       // AG, AH -> AH, AI
            devices.Add(new Device("SafetyBar", "Барьер безопасности", 35, 36)); // AI, AJ -> AJ, AK
            devices.Add(new Device("Sink", "Слив", 37, 38));         // AK, AL -> AL, AM
            devices.Add(new Device("Blower", "Воздуходувка", 39, 40)); // AM, AN -> AN, AO
            devices.Add(new Device("BarRot", "Вращение барабанов", 41, 42)); // AO, AP -> AP, AQ
            devices.Add(new Device("Chiller", "Чиллер", 43, 44));    // AQ, AR -> AR, AS
            devices.Add(new Device("Lifter", "Подъемник", 45, 46));  // AS, AT -> AT, AU
            devices.Add(new Device("Vent", "Вентиляция", 47, 48));   // AU, AV -> AV, AW
        }

        private void InitializeManualTab()
        {
            int y = 15;
            int colWidth = 280;
            int labelWidth = 200;
            int inputWidth = 70;
            
            // Список параметров для ручной конфигурации
            var manualParams = new List<(string Label, string Code)>
            {
                ("Число панелей оператора:", "MaxOP"),
                ("Число рядов ванн:", "MaxRow"),
                ("Число автооператоров:", "MaxAO"),
                ("Число тележек:", "MaxCart"),
                ("Число ванн:", "MaxVann"),
                ("Число доливов:", "MaxDoliv"),
                ("Число нагревов/охлаждений:", "MaxTemperature"),
                ("Число крышек:", "MaxCover"),
                ("Число жироуловителей:", "MaxJr"),
                ("Число перемешиваний:", "MaxMixer"),
                ("Число выпрямителей:", "MaxVip"),
                ("Число фильтрований:", "MaxFiltr"),
                ("Число дозаторов:", "MaxDoser"),
                ("Число душирований:", "MaxShower"),
                ("Число качалок:", "MaxPok"),
                ("Число сушилок:", "MaxDry"),
                ("Число сливов:", "MaxSink"),
                ("Число ПИД-регуляций:", "MaxPID"),
                ("Число воздуходувок:", "MaxBlower"),
                ("Число чиллеров:", "MaxChiller"),
                ("Число барьеров безопасности:", "MaxSafetyBar"),
                ("Число подъемников:", "MaxLifter")
            };

            foreach (var param in manualParams)
            {
                int col = (manualParams.IndexOf(param) / 11) * 430;
                int rowY = 15 + ((manualParams.IndexOf(param) % 11) * 35);
                
                Label lbl = new Label();
                lbl.Text = param.Label;
                lbl.Location = new Point(10 + col, rowY);
                lbl.Size = new Size(labelWidth, 25);
                
                NumericUpDown num = new NumericUpDown();
                num.Location = new Point(10 + labelWidth + col, rowY);
                num.Size = new Size(inputWidth, 25);
                num.Minimum = 0;
                num.Maximum = 1000;
                num.Value = 0;
                
                manualInputs[param.Code] = num;
                
                tabManual.Controls.Add(lbl);
                tabManual.Controls.Add(num);
            }

            btnGenerateManual = new Button();
            btnGenerateManual.Text = "Сгенерировать";
            btnGenerateManual.Location = new Point(350, 420);
            btnGenerateManual.Size = new Size(160, 35);
            btnGenerateManual.BackColor = Color.FromArgb(76, 175, 80);
            btnGenerateManual.ForeColor = Color.White;
            btnGenerateManual.Click += BtnGenerateManual_Click;

            Label lblLogManual = new Label();
            lblLogManual.Text = "Лог:";
            lblLogManual.Location = new Point(10, 465);
            lblLogManual.Size = new Size(50, 20);

            rtbLogManual = new RichTextBox();
            rtbLogManual.Location = new Point(10, 485);
            rtbLogManual.Size = new Size(830, 80);
            rtbLogManual.ReadOnly = true;
            rtbLogManual.BackColor = Color.Black;
            rtbLogManual.ForeColor = Color.LightGreen;
            rtbLogManual.Font = new Font("Consolas", 9);

            tabManual.Controls.AddRange(new Control[] { btnGenerateManual, lblLogManual, rtbLogManual });
        }

        private void InitializeConfigTab()
        {
            Label lblPlace = new Label();
            lblPlace.Text = "Номер позиции:";
            lblPlace.Location = new Point(10, 15);
            lblPlace.Size = new Size(120, 25);

            numConfigPlace = new NumericUpDown();
            numConfigPlace.Location = new Point(140, 15);
            numConfigPlace.Size = new Size(80, 25);
            numConfigPlace.Minimum = 0;
            numConfigPlace.Maximum = 1000;
            numConfigPlace.Value = 0;

            Label lblDevice = new Label();
            lblDevice.Text = "Устройство:";
            lblDevice.Location = new Point(240, 15);
            lblDevice.Size = new Size(100, 25);

            cmbConfigDevice = new ComboBox();
            cmbConfigDevice.Location = new Point(350, 15);
            cmbConfigDevice.Size = new Size(200, 25);
            cmbConfigDevice.DropDownStyle = ComboBoxStyle.DropDownList;
            foreach (var dev in devices)
            {
                cmbConfigDevice.Items.Add(dev.Comment);
            }

            Label lblIndex = new Label();
            lblIndex.Text = "Порядковый номер:";
            lblIndex.Location = new Point(10, 50);
            lblIndex.Size = new Size(120, 25);

            numConfigIndex = new NumericUpDown();
            numConfigIndex.Location = new Point(140, 50);
            numConfigIndex.Size = new Size(80, 25);
            numConfigIndex.Minimum = 0;
            numConfigIndex.Maximum = 1000;
            numConfigIndex.Value = 0;

            Label lblType = new Label();
            lblType.Text = "Тип устройства:";
            lblType.Location = new Point(240, 50);
            lblType.Size = new Size(100, 25);

            numConfigType = new NumericUpDown();
            numConfigType.Location = new Point(350, 50);
            numConfigType.Size = new Size(80, 25);
            numConfigType.Minimum = 0;
            numConfigType.Maximum = 100;
            numConfigType.Value = 0;

            btnSaveConfig = new Button();
            btnSaveConfig.Text = "Сохранить";
            btnSaveConfig.Location = new Point(560, 15);
            btnSaveConfig.Size = new Size(100, 35);
            btnSaveConfig.BackColor = Color.FromArgb(33, 150, 243);
            btnSaveConfig.ForeColor = Color.White;
            btnSaveConfig.Click += BtnSaveConfig_Click;

            btnResetConfig = new Button();
            btnResetConfig.Text = "Сброс";
            btnResetConfig.Location = new Point(670, 15);
            btnResetConfig.Size = new Size(80, 35);
            btnResetConfig.BackColor = Color.FromArgb(244, 67, 54);
            btnResetConfig.ForeColor = Color.White;
            btnResetConfig.Click += BtnResetConfig_Click;

            btnGenerateConfig = new Button();
            btnGenerateConfig.Text = "Сгенерировать код устройств";
            btnGenerateConfig.Location = new Point(350, 90);
            btnGenerateConfig.Size = new Size(200, 35);
            btnGenerateConfig.BackColor = Color.FromArgb(76, 175, 80);
            btnGenerateConfig.ForeColor = Color.White;
            btnGenerateConfig.Click += BtnGenerateConfig_Click;

            lblConfigCount = new Label();
            lblConfigCount.Text = "Сохранено устройств: 0";
            lblConfigCount.Location = new Point(10, 90);
            lblConfigCount.Size = new Size(200, 25);

            Label lblLogConfig = new Label();
            lblLogConfig.Text = "Лог:";
            lblLogConfig.Location = new Point(10, 130);
            lblLogConfig.Size = new Size(50, 20);

            rtbLogConfig = new RichTextBox();
            rtbLogConfig.Location = new Point(10, 150);
            rtbLogConfig.Size = new Size(830, 415);
            rtbLogConfig.ReadOnly = true;
            rtbLogConfig.BackColor = Color.Black;
            rtbLogConfig.ForeColor = Color.LightGreen;
            rtbLogConfig.Font = new Font("Consolas", 9);

            tabConfig.Controls.AddRange(new Control[] {
                lblPlace, numConfigPlace, lblDevice, cmbConfigDevice,
                lblIndex, numConfigIndex, lblType, numConfigType,
                btnSaveConfig, btnResetConfig, btnGenerateConfig,
                lblConfigCount, lblLogConfig, rtbLogConfig
            });
        }

        private void BtnGenerateManual_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var kvp in manualInputs)
            {
                sb.AppendLine($"\"Options\".Count.{kvp.Key} := {kvp.Value.Value};");
            }
            
            LogManual("✅ Код сгенерирован!");
            
            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Filter = "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*";
                dlg.DefaultExt = "txt";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
                    LogManual($"📄 Файл сохранен: {dlg.FileName}");
                }
            }
        }

        private void BtnSaveConfig_Click(object sender, EventArgs e)
        {
            if (cmbConfigDevice.SelectedIndex < 0)
            {
                MessageBox.Show("Выберите устройство!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string deviceName = devices[cmbConfigDevice.SelectedIndex].Name;
            int place = (int)numConfigPlace.Value;
            int index = (int)numConfigIndex.Value;
            int type = (int)numConfigType.Value;

            // Проверка на дубликат
            foreach (var cfg in savedConfigs)
            {
                if (cfg.DeviceName == deviceName && cfg.Index == index)
                {
                    MessageBox.Show($"Устройство {deviceName}[{index}] уже сохранено!", "Дубликат", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            savedConfigs.Add(new DeviceConfig(deviceName, place, index, type));
            lblConfigCount.Text = $"Сохранено устройств: {savedConfigs.Count}";
            LogConfig($"✅ Сохранено: {deviceName}.Dev[{index}].CfgPlace := {place}; {deviceName}.Dev[{index}].CfgType := {type};");
        }

        private void BtnResetConfig_Click(object sender, EventArgs e)
        {
            savedConfigs.Clear();
            lblConfigCount.Text = "Сохранено устройств: 0";
            rtbLogConfig.Clear();
            LogConfig("🔄 Сброс выполнен.");
        }

        private void BtnGenerateConfig_Click(object sender, EventArgs e)
        {
            if (savedConfigs.Count == 0)
            {
                MessageBox.Show("Нет сохраненных устройств!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            StringBuilder sb = new StringBuilder();
            foreach (var cfg in savedConfigs)
            {
                sb.AppendLine($"\"{cfg.DeviceName}\".Dev[{cfg.Index}].CfgPlace := {cfg.Place};");
                sb.AppendLine($"\"{cfg.DeviceName}\".Dev[{cfg.Index}].CfgType := {cfg.Type};");
            }

            LogConfig("✅ Код устройств сгенерирован!");

            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Filter = "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*";
                dlg.DefaultExt = "txt";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
                    LogConfig($"📄 Файл сохранен: {dlg.FileName}");
                }
            }
        }

        private void LogManual(string message)
        {
            rtbLogManual.AppendText(message + Environment.NewLine);
        }

        private void LogConfig(string message)
        {
            rtbLogConfig.AppendText(message + Environment.NewLine);
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

        private void BtnBrowseIO_System_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                dlg.Title = "Выберите файл 'Система ввода-вывода.xlsx'";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    txtIO_SystemPath.Text = dlg.FileName;
                    LogIO($"Выбран файл системы ввода-вывода: {dlg.FileName}");
                }
            }
        }

        private void BtnBrowseIO_Spec_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                dlg.Title = "Выберите файл 'Спецификация.xlsx'";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    txtIO_SpecPath.Text = dlg.FileName;
                    LogIO($"Выбран файл спецификации: {dlg.FileName}");
                }
            }
        }

        private void BtnBrowseIO_Var_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                dlg.Title = "Выберите файл 'Список переменных от системы ввода-вывода.xlsx'";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    txtIO_VarPath.Text = dlg.FileName;
                    LogIO($"Выбран файл списка переменных: {dlg.FileName}");
                }
            }
        }

        private void BtnBrowseIOTxt_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Filter = "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*";
                dlg.DefaultExt = "txt";
                dlg.Title = "Сохранить SCL файл";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    txtIOTxtPath.Text = dlg.FileName;
                    LogIO($"Файл будет сохранен: {dlg.FileName}");
                }
            }
        }

        private void BtnGenerateIO_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtIO_SystemPath.Text))
            {
                MessageBox.Show("Выберите файл 'Система ввода-вывода.xlsx'!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrEmpty(txtIO_SpecPath.Text))
            {
                MessageBox.Show("Выберите файл 'Спецификация.xlsx'!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrEmpty(txtIO_VarPath.Text))
            {
                MessageBox.Show("Выберите файл 'Список переменных от системы ввода-вывода.xlsx'!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrEmpty(txtIOTxtPath.Text))
            {
                MessageBox.Show("Выберите путь для сохранения!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!File.Exists(txtIO_SystemPath.Text))
            {
                MessageBox.Show("Файл 'Система ввода-вывода.xlsx' не найден!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!File.Exists(txtIO_SpecPath.Text))
            {
                MessageBox.Show("Файл 'Спецификация.xlsx' не найден!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!File.Exists(txtIO_VarPath.Text))
            {
                MessageBox.Show("Файл 'Список переменных от системы ввода-вывода.xlsx' не найден!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            btnGenerateIO.Enabled = false;
            progressBar.Visible = true;
            LogIO("Начало генерации SCL (Ввод-вывод)...");

            try
            {
                string result = GenerateIOSCL();
                File.WriteAllText(txtIOTxtPath.Text, result, Encoding.UTF8);
                LogIO("✅ Генерация SCL успешно завершена!");
                LogIO($"📄 Файл сохранен: {txtIOTxtPath.Text}");

                if (MessageBox.Show("Открыть полученный файл?", "Готово",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start("notepad.exe", txtIOTxtPath.Text);
                }
            }
            catch (Exception ex)
            {
                LogIO($"❌ Ошибка: {ex.Message}");
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnGenerateIO.Enabled = true;
                progressBar.Visible = false;
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

        private void LogIO(string message)
        {
            if (rtbLogIO.InvokeRequired)
            {
                rtbLogIO.Invoke(new Action<string>(LogIO), message);
            }
            else
            {
                rtbLogIO.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
                rtbLogIO.ScrollToCaret();
            }
        }

        private string GenerateIOSCL()
        {
            StringBuilder result = new StringBuilder();
            result.AppendLine("// SCL код карты ввода-вывода сгенерирован автоматически");
            result.AppendLine($"// Дата генерации: {DateTime.Now}");
            result.AppendLine();

            // Словари для маппинга
            Dictionary<string, string> diVariables = new Dictionary<string, string>();
            Dictionary<string, string> doVariables = new Dictionary<string, string>();
            Dictionary<int, string> specPositions = new Dictionary<int, string>();
            Dictionary<string, DeviceSpecInfo> specDeviceMap = new Dictionary<string, DeviceSpecInfo>();

            // 1. Читаем "Список переменных" (DI и DO таблицы)
            LogIO("Чтение файла 'Список переменных от системы ввода-вывода.xlsx'...");
            using (FileStream fs = new FileStream(txtIO_VarPath.Text, FileMode.Open, FileAccess.Read))
            using (XSSFWorkbook workbook = new XSSFWorkbook(fs))
            {
                ISheet sheet = workbook.GetSheetAt(0);
                
                // Читаем DI таблицу (колонки A и B)
                int rowIdx = 0;
                while (rowIdx < sheet.LastRowNum)
                {
                    IRow row = sheet.GetRow(rowIdx);
                    if (row == null) { rowIdx++; continue; }
                    
                    ICell diCell = row.GetCell(0);
                    ICell varCell = row.GetCell(1);
                    
                    if (diCell != null && varCell != null)
                    {
                        string diName = GetCellValue(diCell).Trim();
                        string varName = GetCellValue(varCell).Trim();
                        
                        if (!string.IsNullOrEmpty(diName) && !string.IsNullOrEmpty(varName) && 
                            diName != "DI" && diName != "Переменная")
                        {
                            if (!diVariables.ContainsKey(diName))
                                diVariables[diName] = varName;
                        }
                    }
                    
                    // Проверяем, не началась ли таблица DO (колонка D содержит "DO")
                    ICell doHeaderCell = row.GetCell(3);
                    if (doHeaderCell != null && GetCellValue(doHeaderCell).Trim() == "DO")
                        break;
                    
                    rowIdx++;
                }
                
                // Читаем DO таблицу (колонки D и E)
                rowIdx = 0;
                bool inDoTable = false;
                while (rowIdx < sheet.LastRowNum)
                {
                    IRow row = sheet.GetRow(rowIdx);
                    if (row == null) { rowIdx++; continue; }
                    
                    ICell doHeaderCell = row.GetCell(3);
                    if (doHeaderCell != null && GetCellValue(doHeaderCell).Trim() == "DO")
                    {
                        inDoTable = true;
                        rowIdx++;
                        continue;
                    }
                    
                    if (inDoTable)
                    {
                        ICell doCell = row.GetCell(3);
                        ICell varCell = row.GetCell(4);
                        
                        if (doCell != null && varCell != null)
                        {
                            string doName = GetCellValue(doCell).Trim();
                            string varName = GetCellValue(varCell).Trim();
                            
                            if (!string.IsNullOrEmpty(doName) && !string.IsNullOrEmpty(varName) && 
                                doName != "DO" && doName != "Переменная")
                            {
                                if (!doVariables.ContainsKey(doName))
                                    doVariables[doName] = varName;
                            }
                        }
                        
                        // Проверяем конец таблицы DO
                        ICell aiHeaderCell = row.GetCell(6);
                        if (aiHeaderCell != null && GetCellValue(aiHeaderCell).Trim() == "AI")
                            break;
                    }
                    
                    rowIdx++;
                }
            }
            LogIO($"Загружено DI переменных: {diVariables.Count}, DO переменных: {doVariables.Count}");

            // 2. Читаем "Спецификация" (лист Config_Line)
            LogIO("Чтение файла 'Спецификация.xlsx'...");
            using (FileStream fs = new FileStream(txtIO_SpecPath.Text, FileMode.Open, FileAccess.Read))
            using (XSSFWorkbook workbook = new XSSFWorkbook(fs))
            {
                ISheet sheet = null;
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    if (workbook.GetSheetName(i).Contains("Config_Line"))
                    {
                        sheet = workbook.GetSheetAt(i);
                        break;
                    }
                }
                
                if (sheet == null)
                {
                    throw new Exception("Лист 'Config_Line' не найден в файле Спецификации!");
                }
                
                // Находим заголовок с колонками устройств
                int headerRow = 0;
                Dictionary<string, int> deviceColIndex = new Dictionary<string, int>();
                
                IRow headerRowData = sheet.GetRow(headerRow);
                if (headerRowData != null)
                {
                    for (int col = 0; col < headerRowData.LastCellNum; col++)
                    {
                        string cellValue = GetCellValue(headerRowData.GetCell(col)).Trim();
                        // Сопоставляем русские названия с именами устройств
                        if (cellValue == "Долив") deviceColIndex["Doliv"] = col;
                        else if (cellValue == "Температура") deviceColIndex["Tmpr"] = col;
                        else if (cellValue == "Крышки" || cellValue == "Крышка") deviceColIndex["Cover"] = col;
                        else if (cellValue == "Жироуловитель") deviceColIndex["Jr"] = col;
                        else if (cellValue == "Перемешивание") deviceColIndex["Mixer"] = col;
                        else if (cellValue == "Выпрямитель") deviceColIndex["Vip"] = col;
                        else if (cellValue == "Фильтрование") deviceColIndex["Filtr"] = col;
                        else if (cellValue == "Дозирование") deviceColIndex["Doser"] = col;
                        else if (cellValue == "Душирование") deviceColIndex["Shower"] = col;
                        else if (cellValue == "Качание") deviceColIndex["Pok"] = col;
                        else if (cellValue == "Сушилка") deviceColIndex["Dry"] = col;
                        else if (cellValue == "Барьер безопасности") deviceColIndex["SafetyBar"] = col;
                        else if (cellValue == "Слив") deviceColIndex["Sink"] = col;
                        else if (cellValue == "Воздуходувка") deviceColIndex["Blower"] = col;
                        else if (cellValue == "Чиллер") deviceColIndex["Chiller"] = col;
                        else if (cellValue == "Подъемник") deviceColIndex["Lifter"] = col;
                    }
                }
                
                // Читаем позиции
                for (int rowNum = 1; rowNum <= sheet.LastRowNum; rowNum++)
                {
                    IRow row = sheet.GetRow(rowNum);
                    if (row == null) continue;
                    
                    CellValueInfo posInfo = GetCellValueInfo(row.GetCell(0));
                    if (!int.TryParse(posInfo.Value, out int positionNum))
                        continue;
                    
                    specPositions[positionNum] = $"Позиция {positionNum}";
                    
                    // Для каждого устройства сохраняем индекс
                    foreach (var devPair in deviceColIndex)
                    {
                        ICell cell = row.GetCell(devPair.Value);
                        if (cell != null)
                        {
                            CellValueInfo valInfo = GetCellValueInfo(cell);
                            if (int.TryParse(valInfo.Value, out int devIndex) && devIndex > 0)
                            {
                                string key = $"{devPair.Key}_{positionNum}";
                                specDeviceMap[key] = new DeviceSpecInfo { DeviceName = devPair.Key, Index = devIndex };
                            }
                        }
                    }
                }
            }
            LogIO($"Загружено позиций: {specPositions.Count}, привязок устройств: {specDeviceMap.Count}");

            // 3. Читаем "Система ввода-вывода" (лист ШС) и генерируем код
            LogIO("Чтение файла 'Система ввода-вывода.xlsx' и генерация SCL...");
            
            // Группы модулей: ключ = тип+адрес+обозначение
            Dictionary<string, IOModule> modules = new Dictionary<string, IOModule>();
            int moduleCounter = 0;
            
            using (FileStream fs = new FileStream(txtIO_SystemPath.Text, FileMode.Open, FileAccess.Read))
            using (XSSFWorkbook workbook = new XSSFWorkbook(fs))
            {
                ISheet sheet = null;
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    if (workbook.GetSheetName(i).Contains("ШС"))
                    {
                        sheet = workbook.GetSheetAt(i);
                        break;
                    }
                }
                
                if (sheet == null)
                {
                    throw new Exception("Лист 'ШС' не найден в файле Системы ввода-вывода!");
                }
                
                // Определяем индексы колонок по заголовку
                int colType = -1, colAddress = -1, colDesignation = -1, colSignalNum = -1;
                int colDevice = -1, colSignalName = -1, colTechPos = -1;
                
                IRow headerRow = sheet.GetRow(0);
                if (headerRow != null)
                {
                    for (int col = 0; col < headerRow.LastCellNum; col++)
                    {
                        string cellVal = GetCellValue(headerRow.GetCell(col)).Trim().ToLower();
                        if (cellVal == "тип") colType = col;
                        else if (cellVal == "адрес") colAddress = col;
                        else if (cellVal == "обозн.") colDesignation = col;
                        else if (cellVal == "№ вх." || cellVal == "номер входа") colSignalNum = col;
                        else if (cellVal == "устройство") colDevice = col;
                        else if (cellVal == "наименование сигнала") colSignalName = col;
                        else if (cellVal == "тех. поз." || cellVal == "технологическая позиция") colTechPos = col;
                    }
                }
                
                if (colType < 0 || colAddress < 0 || colDesignation < 0 || colSignalNum < 0 || 
                    colDevice < 0 || colSignalName < 0 || colTechPos < 0)
                {
                    throw new Exception("Не все требуемые колонки найдены в листе 'ШС'!");
                }
                
                // Читаем строки
                for (int rowNum = 1; rowNum <= sheet.LastRowNum; rowNum++)
                {
                    IRow row = sheet.GetRow(rowNum);
                    if (row == null) continue;
                    
                    string type = GetCellValue(row.GetCell(colType)).Trim();
                    string address = GetCellValue(row.GetCell(colAddress)).Trim();
                    string designation = GetCellValue(row.GetCell(colDesignation)).Trim();
                    string signalNumStr = GetCellValue(row.GetCell(colSignalNum)).Trim();
                    string deviceName = GetCellValue(row.GetCell(colDevice)).Trim();
                    string signalName = GetCellValue(row.GetCell(colSignalName)).Trim();
                    string techPosStr = GetCellValue(row.GetCell(colTechPos)).Trim();
                    
                    // Пропускаем некорректные строки
                    if (type != "DI" && type != "DO") continue;
                    if (string.IsNullOrEmpty(address) || address.ToLower() == "no data") continue;
                    if (string.IsNullOrEmpty(designation)) continue;
                    if (!int.TryParse(signalNumStr, out int signalNum)) continue;
                    if (string.IsNullOrEmpty(signalName)) continue;
                    if (string.IsNullOrEmpty(techPosStr)) continue;
                    
                    // Создаем ключ модуля
                    string moduleKey = $"{type}_{address}_{designation}";
                    
                    if (!modules.ContainsKey(moduleKey))
                    {
                        moduleCounter++;
                        modules[moduleKey] = new IOModule
                        {
                            Number = moduleCounter,
                            Type = type,
                            Address = address,
                            Designation = designation,
                            Signals = new List<IOSignal>()
                        };
                    }
                    
                    // Ищем устройство в спецификации
                    if (!int.TryParse(techPosStr, out int techPos))
                        continue;
                    
                    string specKey = "";
                    // Сопоставляем русское название устройства с именем
                    if (deviceName.Contains("Долив")) specKey = "Doliv";
                    else if (deviceName.Contains("Температур")) specKey = "Tmpr";
                    else if (deviceName.Contains("Крышк")) specKey = "Cover";
                    else if (deviceName.Contains("Жироуловит")) specKey = "Jr";
                    else if (deviceName.Contains("Перемешиван")) specKey = "Mixer";
                    else if (deviceName.Contains("Выпрямит")) specKey = "Vip";
                    else if (deviceName.Contains("Фильтрован")) specKey = "Filtr";
                    else if (deviceName.Contains("Дозир")) specKey = "Doser";
                    else if (deviceName.Contains("Душирован")) specKey = "Shower";
                    else if (deviceName.Contains("Качан")) specKey = "Pok";
                    else if (deviceName.Contains("Сушил")) specKey = "Dry";
                    else if (deviceName.Contains("Барьер")) specKey = "SafetyBar";
                    else if (deviceName.Contains("Слив")) specKey = "Sink";
                    else if (deviceName.Contains("Воздуходув")) specKey = "Blower";
                    else if (deviceName.Contains("Чиллер")) specKey = "Chiller";
                    else if (deviceName.Contains("Подъемник")) specKey = "Lifter";
                    else continue; // Устройство не найдено
                    
                    string deviceSpecKey = $"{specKey}_{techPos}";
                    if (!specDeviceMap.ContainsKey(deviceSpecKey))
                        continue; // Позиция не найдена в спецификации
                    
                    DeviceSpecInfo specInfo = specDeviceMap[deviceSpecKey];
                    
                    // Ищем переменную в списке
                    string variableName = "";
                    if (type == "DI")
                    {
                        if (diVariables.ContainsKey(signalName))
                            variableName = diVariables[signalName];
                    }
                    else if (type == "DO")
                    {
                        if (doVariables.ContainsKey(signalName))
                            variableName = doVariables[signalName];
                    }
                    
                    if (string.IsNullOrEmpty(variableName))
                        continue; // Переменная не найдена
                    
                    // Добавляем сигнал в модуль
                    modules[moduleKey].Signals.Add(new IOSignal
                    {
                        SignalNumber = signalNum,
                        DeviceName = specInfo.DeviceName,
                        DeviceIndex = specInfo.Index,
                        VariableName = variableName
                    });
                }
            }
            
            LogIO($"Найдено модулей: {modules.Count}");
            
            // Генерируем SCL код по модулям
            foreach (var modulePair in modules.OrderBy(k => k.Value.Number))
            {
                IOModule module = modulePair.Value;
                
                result.AppendLine($"REGION Module {module.Number}");
                result.AppendLine($"// {module.Designation}. Адрес {module.Address}");
                
                if (module.Type == "DO")
                {
                    result.AppendLine("#dwModuleBitMask := 0; // Обнулим маску выходов");
                }
                else
                {
                    result.AppendLine("#dwModuleBitMask := 0; // Обнулим маску входов");
                }
                
                // Сортируем сигналы по номеру
                foreach (var signal in module.Signals.OrderBy(s => s.SignalNumber))
                {
                    result.AppendLine($"#xBit.b{signal.SignalNumber} := \"{signal.DeviceName}\".Dev[{signal.DeviceIndex}].{signal.VariableName};");
                }
                
                // Генерируем OR операции
                foreach (var signal in module.Signals.OrderBy(s => s.SignalNumber))
                {
                    result.AppendLine($"#dwModuleBitMask := #dwModuleBitMask OR #xBit.b{signal.SignalNumber};");
                }
                
                if (module.Type == "DO")
                {
                    result.AppendLine($"\"MapDout\".Module[{module.Number}].BitMask := #dwModuleBitMask;");
                }
                else
                {
                    result.AppendLine($"\"MapDin\".Module[{module.Number}].BitMask := #dwModuleBitMask;");
                }
                
                result.AppendLine("END_REGION");
                result.AppendLine();
            }
            
            return result.ToString();
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
                
                foreach (var dev in devices)
                {
                    // Формируем имя переменной: MaxDoliv, MaxTmpr и т.д.
                    string varName = $"Max{dev.Name}";
                    header.AppendLine($"\"Options\".Count.{varName} := {maxIndices[dev.Name]};");
                }
                
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
                                    string fStr = (fVal == Math.Floor(fVal)) ? fVal.ToString("0") : fStr = fVal.ToString(System.Globalization.CultureInfo.InvariantCulture);
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

        private string GetCellValue(ICell cell)
        {
            CellValueInfo info = GetCellValueInfo(cell);
            return info.Value;
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

        class DeviceConfig
        {
            public string DeviceName { get; set; }
            public int Place { get; set; }
            public int Index { get; set; }
            public int Type { get; set; }

            public DeviceConfig(string deviceName, int place, int index, int type)
            {
                DeviceName = deviceName;
                Place = place;
                Index = index;
                Type = type;
            }
        }

        class DeviceSpecInfo
        {
            public string DeviceName { get; set; }
            public int Index { get; set; }
        }

        class IOModule
        {
            public int Number { get; set; }
            public string Type { get; set; }
            public string Address { get; set; }
            public string Designation { get; set; }
            public List<IOSignal> Signals { get; set; }
        }

        class IOSignal
        {
            public int SignalNumber { get; set; }
            public string DeviceName { get; set; }
            public int DeviceIndex { get; set; }
            public string VariableName { get; set; }
        }
    }
}