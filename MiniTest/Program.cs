using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelDataReader;
using System.Drawing;

namespace MiniTest
{
    public static class Program
    {
        [STAThread]
        public static void Main()
        {
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    public partial class MainForm : Form
    {
        // Глобальные переменные для путей к файлам
        private string ioSystemPath = "";
        private string specPath = "";
        private string variablesPath = "";

        // Маппинг русских названий устройств на английские (для генерации кода)
        private readonly Dictionary<string, string> deviceNames = new Dictionary<string, string>
        {
            { "Панель оператора", "OP" },
            { "Ряд ванн", "Row" },
            { "Автооператор", "AO" },
            { "Тележка", "Cart" },
            { "Ванна", "Vann" },
            { "Долив", "Doliv" },
            { "Температура", "Tmpr" },
            { "Крышки", "Cover" },
            { "Жироуловитель", "Jr" },
            { "Перемешивание", "Mixer" },
            { "Выпрямитель", "Vip" },
            { "Фильтрование", "Filtr" },
            { "Дозатор", "Doser" },
            { "Душирование", "Shower" },
            { "Качалка", "Pok" },
            { "Сушилка", "Dry" },
            { "Слив", "Sink" },
            { "ПИД-регуляция", "PID" },
            { "Воздуходувка", "Blower" },
            { "Чиллер", "Chiller" },
            { "Барьер безопасности", "SafetyBar" },
            { "Подъемник", "Lifter" }
        };

        // Список устройств для конфигуратора (только те, что указаны в задании)
        private readonly string[] deviceTypes = new string[] 
        { 
            "Панель оператора", "Ряд ванн", "Автооператор", "Тележка", "Ванна",
            "Долив", "Температура", "Крышки", "Жироуловитель", "Перемешивание", 
            "Выпрямитель", "Фильтрование", "Дозатор", "Душирование", "Качалка", 
            "Сушилка", "Слив", "ПИД-регуляция", "Воздуходувка", "Чиллер", 
            "Барьер безопасности", "Подъемник"
        };

        // Структуры для вкладки Ввод-Вывод
        public struct IoSignalInfo
        {
            public int SignalNumber; 
            public string DeviceRu; 
            public string SignalNameRu; 
            public int TechPos;     
            public string VarNameEn; 
            public int DevIndex;    
        }

        public struct IoModuleInfo
        {
            public int Id;
            public string Type; 
            public string Address;
            public string Label;
            public List<IoSignalInfo> Signals;
        }

        private Dictionary<string, IoModuleInfo> ioModules = new Dictionary<string, IoModuleInfo>();

        // Список устройств для конфигуратора
        private List<DeviceConfig> savedDevices = new List<DeviceConfig>();

        public struct DeviceConfig
        {
            public int Position;
            public string DeviceType;
            public int Index;
            public int TypeCode;
        }

        public MainForm()
        {
            InitializeCustomComponents();
        }

        private void InitializeCustomComponents()
        {
            this.Text = "Excel to SCL Converter";
            this.Size = new Size(1000, 700);
            this.StartPosition = FormStartPosition.CenterScreen;

            TabControl tabControl = new TabControl
            {
                Dock = DockStyle.Fill
            };

            // --- Вкладка 1: Спецификация ---
            TabPage tabSpec = new TabPage("Спецификация");
            SetupSpecTab(tabSpec);
            tabControl.TabPages.Add(tabSpec);

            // --- Вкладка 2: Ручная конфигурация ---
            TabPage tabManual = new TabPage("Ручная конфигурация");
            SetupManualTab(tabManual);
            tabControl.TabPages.Add(tabManual);

            // --- Вкладка 3: Конфигуратор устройств ---
            TabPage tabConfigurator = new TabPage("Конфигуратор устройств");
            SetupConfiguratorTab(tabConfigurator);
            tabControl.TabPages.Add(tabConfigurator);

            // --- Вкладка 4: Ввод-вывод ---
            TabPage tabIO = new TabPage("Ввод-вывод");
            SetupIOTab(tabIO);
            tabControl.TabPages.Add(tabIO);

            this.Controls.Add(tabControl);
        }

        #region Вкладка Спецификация

        private void SetupSpecTab(TabPage tab)
        {
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(10)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var btnSelect = new Button { Text = "Выбрать файл Спецификация.xlsx", Dock = DockStyle.Fill };
            btnSelect.Click += (s, e) => SelectFile("spec", ref specPath);

            var btnGenerate = new Button { Text = "Сгенерировать код", Dock = DockStyle.Fill };
            btnGenerate.Click += (s, e) => GenerateSpecCode();

            var lblStatus = new Label { Text = "Файл не выбран", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            lblStatus.Name = "lblSpecStatus";

            var txtOutput = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, ReadOnly = true };
            txtOutput.Name = "txtSpecOutput";

            layout.Controls.Add(btnSelect, 0, 0);
            layout.Controls.Add(btnGenerate, 0, 1);
            layout.Controls.Add(lblStatus, 0, 2);
            layout.Controls.Add(txtOutput, 0, 3);

            tab.Controls.Add(layout);
        }

        private void SelectFile(string type, ref string pathVar)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel Files|*.xlsx";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    pathVar = dlg.FileName;
                    UpdateStatusLabel(type, Path.GetFileName(pathVar));
                }
            }
        }

        private void UpdateStatusLabel(string type, string fileName)
        {
            foreach (TabPage tab in ((TabControl)this.Controls[0]).TabPages)
            {
                foreach (Control ctrl in tab.Controls)
                {
                    if (ctrl is TableLayoutPanel tlp)
                    {
                        foreach (Control c in tlp.Controls)
                        {
                            if (c is Label lbl && lbl.Name == $"lbl{type}Status")
                            {
                                lbl.Text = $"Выбран: {fileName}";
                                lbl.ForeColor = Color.Green;
                            }
                        }
                    }
                }
            }
        }

        private void GenerateSpecCode()
        {
            if (!File.Exists(specPath))
            {
                MessageBox.Show("Файл спецификации не выбран!");
                return;
            }

            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (var stream = File.Open(specPath, FileMode.Open, FileAccess.Read))
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var sb = new StringBuilder();
                    bool found = false;
                    
                    while (reader.Read())
                    {
                        if (reader.Name == "Config_Line")
                        {
                            found = true;
                            
                            // Читаем заголовок для определения индексов колонок
                            if (!reader.Read()) break; 
                            
                            var headers = new Dictionary<string, int>();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                var h = reader.GetValue(i)?.ToString()?.Trim() ?? "";
                                if (!string.IsNullOrEmpty(h) && !headers.ContainsKey(h))
                                    headers[h] = i;
                            }

                            // Проходим по строкам данных
                            while (reader.Read())
                            {
                                // Пропускаем строки без номера позиции
                                var posVal = reader.GetValue(0);
                                if (posVal == null || !int.TryParse(posVal.ToString(), out _)) continue;

                                // Генерируем код для каждого известного устройства
                                foreach (var kvp in deviceNames)
                                {
                                    string rusName = kvp.Key;
                                    string engName = kvp.Value;
                                    string propName = GetPropNameByDevice(rusName);

                                    if (headers.ContainsKey(rusName))
                                    {
                                        var val = reader.GetValue(headers[rusName]);
                                        if (val != null && int.TryParse(val.ToString(), out int count) && count > 0)
                                        {
                                            sb.AppendLine($"\"Options\".Count.{propName} := {count};");
                                        }
                                    }
                                }
                            }
                            break;
                        }
                    }
                    
                    if (!found) 
                        sb.AppendLine("Лист Config_Line не найден!");
                    else if (sb.Length == 0)
                        sb.AppendLine("// Данные не найдены или все значения равны 0");
                    
                    ShowOutput("txtSpecOutput", sb.ToString());
                    SaveToFile("Spec_Generation.txt", sb.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}\n{ex.StackTrace}");
            }
        }

        // Helper для маппинга Русское имя -> Свойство (Max...)
        private string GetPropNameByDevice(string rusName)
        {
            switch (rusName)
            {
                case "Панель оператора": return "MaxOP";
                case "Ряд ванн": return "MaxRow";
                case "Автооператор": return "MaxAO";
                case "Тележка": return "MaxCart";
                case "Ванна": return "MaxVann";
                case "Долив": return "MaxDoliv";
                case "Температура": return "MaxTemperature";
                case "Крышки": return "MaxCover";
                case "Жироуловитель": return "MaxJr";
                case "Перемешивание": return "MaxMixer";
                case "Выпрямитель": return "MaxVip";
                case "Фильтрование": return "MaxFiltr";
                case "Дозатор": return "MaxDoser";
                case "Душирование": return "MaxShower";
                case "Качалка": return "MaxPok";
                case "Сушилка": return "MaxDry";
                case "Слив": return "MaxSink";
                case "ПИД-регуляция": return "MaxPID";
                case "Воздуходувка": return "MaxBlower";
                case "Чиллер": return "MaxChiller";
                case "Барьер безопасности": return "MaxSafetyBar";
                case "Подъемник": return "MaxLifter";
                default: return "Unknown";
            }
        }

        #endregion

        #region Вкладка Ручная конфигурация

        private void SetupManualTab(TabPage tab)
        {
            var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };
            
            var parameters = new (string Label, string Prop)[]
            {
                ("Число панелей оператора", "MaxOP"), ("Число рядов ванн", "MaxRow"),
                ("Число автооператоров", "MaxAO"), ("Число тележек", "MaxCart"),
                ("Число ванн", "MaxVann"), ("Число доливов", "MaxDoliv"),
                ("Число нагревов/охлаждений", "MaxTemperature"), ("Число крышек", "MaxCover"),
                ("Число жироуловителей", "MaxJr"), ("Число перемешиваний", "MaxMixer"),
                ("Число выпрямителей", "MaxVip"), ("Число фильтрований", "MaxFiltr"),
                ("Число дозаторов", "MaxDoser"), ("Число душирований", "MaxShower"),
                ("Число качалок", "MaxPok"), ("Число сушилок", "MaxDry"),
                ("Число сливов", "MaxSink"), ("Число ПИД-регуляций", "MaxPID"),
                ("Число воздуходувок", "MaxBlower"), ("Число чиллеров", "MaxChiller"),
                ("Число барьеров безопасности", "MaxSafetyBar"), ("Число подъемников", "MaxLifter")
            };

            var inputs = new Dictionary<string, NumericUpDown>();

            foreach (var p in parameters)
            {
                var panel = new Panel { Height = 35, Width = 850 };
                var lbl = new Label { Text = p.Label + ":", Left = 10, Top = 8, Width = 220 };
                var num = new NumericUpDown { Left = 240, Top = 6, Width = 100, Minimum = 0, Maximum = 1000 };
                inputs[p.Prop] = num;
                panel.Controls.AddRange(new Control[] { lbl, num });
                flow.Controls.Add(panel);
            }

            var btnGen = new Button { Text = "Сгенерировать", Width = 200, Margin = new Padding(10), Height = 40 };
            btnGen.Click += (s, e) => GenerateManualCode(inputs);
            flow.Controls.Add(btnGen);

            tab.Controls.Add(flow);
        }

        private void GenerateManualCode(Dictionary<string, NumericUpDown> inputs)
        {
            var sb = new StringBuilder();
            foreach (var kvp in inputs)
            {
                sb.AppendLine($"\"Options\".Count.{kvp.Key} := {kvp.Value.Value};");
            }
            SaveToFile("Manual_Config.txt", sb.ToString());
            MessageBox.Show("Код ручной конфигурации сгенерирован!");
        }

        #endregion

        #region Вкладка Конфигуратор устройств

        private void SetupConfiguratorTab(TabPage tab)
        {
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 6,
                RowCount = 5,
                Padding = new Padding(20)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            
            for(int i=0; i<5; i++) layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var numPos = new NumericUpDown { Width = 100, Minimum = 1, Maximum = 1000 };
            var cmbDevice = new ComboBox { Width = 200, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbDevice.Items.AddRange(deviceTypes);
            var numIdx = new NumericUpDown { Width = 100, Minimum = 1, Maximum = 1000 };
            
            // Тип устройства - Input Field (NumericUpDown) от 1 до 100
            var numType = new NumericUpDown { Width = 100, Minimum = 1, Maximum = 100 };

            layout.Controls.Add(new Label { Text = "Номер позиции:", TextAlign = ContentAlignment.MiddleRight, AutoSize = true }, 0, 0);
            layout.Controls.Add(numPos, 1, 0);
            layout.Controls.Add(new Label { Text = "Устройство:", TextAlign = ContentAlignment.MiddleRight, AutoSize = true }, 2, 0);
            layout.Controls.Add(cmbDevice, 3, 0);
            
            layout.Controls.Add(new Label { Text = "Порядковый номер:", TextAlign = ContentAlignment.MiddleRight, AutoSize = true }, 0, 1);
            layout.Controls.Add(numIdx, 1, 1);
            layout.Controls.Add(new Label { Text = "Тип устройства:", TextAlign = ContentAlignment.MiddleRight, AutoSize = true }, 2, 1);
            layout.Controls.Add(numType, 3, 1);

            var btnSave = new Button { Text = "Сохранить", Width = 100, Height = 30 };
            btnSave.Click += (s, e) => SaveDeviceConfig((int)numPos.Value, cmbDevice.Text, (int)numIdx.Value, (int)numType.Value);
            
            var btnReset = new Button { Text = "Сброс", Width = 100, Height = 30 };
            btnReset.Click += (s, e) => { savedDevices.Clear(); UpdateConfigStatus(); };

            layout.Controls.Add(btnSave, 4, 0); 
            layout.Controls.Add(btnReset, 4, 1); 

            var lblStatus = new Label { Text = "Сохранено устройств: 0", AutoSize = true };
            lblStatus.Name = "lblConfigStatus";
            layout.Controls.Add(lblStatus, 0, 2);

            var btnGenDev = new Button { Text = "Сгенерировать код устройств", Width = 200, Height = 40 };
            btnGenDev.Click += (s, e) => GenerateDeviceCode();
            layout.Controls.Add(btnGenDev, 0, 3);

            tab.Controls.Add(layout);
        }

        private void SaveDeviceConfig(int pos, string devType, int idx, int typeCode)
        {
            if (string.IsNullOrEmpty(devType))
            {
                MessageBox.Show("Выберите устройство!");
                return;
            }

            if (savedDevices.Any(d => d.DeviceType == devType && d.Index == idx))
            {
                MessageBox.Show($"Устройство {devType} с индексом {idx} уже сохранено!");
                return;
            }

            savedDevices.Add(new DeviceConfig { Position = pos, DeviceType = devType, Index = idx, TypeCode = typeCode });
            UpdateConfigStatus();
            MessageBox.Show("Устройство сохранено!");
        }

        private void UpdateConfigStatus()
        {
            var lbl = this.Controls.Find("lblConfigStatus", true).FirstOrDefault() as Label;
            if (lbl != null) lbl.Text = $"Сохранено устройств: {savedDevices.Count}";
        }

        private void GenerateDeviceCode()
        {
            var sb = new StringBuilder();
            foreach (var dev in savedDevices)
            {
                string engName = GetEngName(dev.DeviceType);
                sb.AppendLine($"\"{engName}\".Dev[{dev.Index}].CfgPlace := {dev.Position};");
                sb.AppendLine($"\"{engName}\".Dev[{dev.Index}].CfgType := {dev.TypeCode};");
            }
            SaveToFile("Device_Config_Code.txt", sb.ToString());
            MessageBox.Show("Код устройств сгенерирован!");
        }

        private string GetEngName(string rus)
        {
            if (deviceNames.ContainsKey(rus))
                return deviceNames[rus];
            return "Unknown";
        }

        #endregion

        #region Вкладка Ввод-Вывод

        private void SetupIOTab(TabPage tab)
        {
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(20)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for(int i=0; i<6; i++) layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var btnIoSystem = new Button { Text = "1. Выбрать 'Система ввода-вывода.xlsx'", Dock = DockStyle.Fill, Font = new Font(this.Font, FontStyle.Bold) };
            btnIoSystem.Click += (s, e) => SelectFile("ioSystem", ref ioSystemPath);

            var btnSpecIo = new Button { Text = "2. Выбрать 'Спецификация.xlsx'", Dock = DockStyle.Fill };
            btnSpecIo.Click += (s, e) => SelectFile("specIo", ref specPath);

            var btnVars = new Button { Text = "3. Выбрать 'Список переменных...xlsx'", Dock = DockStyle.Fill };
            btnVars.Click += (s, e) => SelectFile("vars", ref variablesPath);

            var btnGenIo = new Button { Text = "Сгенерировать SCL (Ввод-Вывод)", Dock = DockStyle.Fill, Font = new Font(this.Font, FontStyle.Bold), BackColor = Color.LightGreen };
            btnGenIo.Click += (s, e) => GenerateIOCode();

            var lblStatus = new Label { Text = "Файлы не выбраны", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            lblStatus.Name = "lblIOStatus";

            layout.Controls.Add(btnIoSystem, 0, 0);
            layout.Controls.Add(btnSpecIo, 0, 1);
            layout.Controls.Add(btnVars, 0, 2);
            layout.Controls.Add(btnGenIo, 0, 3);
            layout.Controls.Add(lblStatus, 0, 4);

            tab.Controls.Add(layout);
        }

        private void GenerateIOCode()
        {
            if (!File.Exists(ioSystemPath) || !File.Exists(specPath) || !File.Exists(variablesPath))
            {
                MessageBox.Show("Не все файлы выбраны! Пожалуйста, выберите все три файла.");
                return;
            }

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            try
            {
                var variablesMap = ReadVariablesFile(variablesPath);
                var specMap = ReadSpecFile(specPath);

                var sb = new StringBuilder();
                ReadAndGenerateIO(ioSystemPath, variablesMap, specMap, sb);

                if (sb.Length == 0)
                {
                    MessageBox.Show("Данные не найдены или не сгенерированы. Проверьте файлы.");
                    return;
                }

                SaveToFile("IO_Map_Generated.scl", sb.ToString());
                MessageBox.Show("SCL код успешно сгенерирован!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка генерации: {ex.Message}\n{ex.StackTrace}");
            }
        }

        private Dictionary<string, string> ReadVariablesFile(string path)
        {
            var map = new Dictionary<string, string>();
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                do
                {
                    while (reader.Read())
                    {
                        for (int i = 0; i < reader.FieldCount - 1; i++)
                        {
                            var col1 = reader.GetValue(i)?.ToString()?.Trim();
                            var col2 = reader.GetValue(i + 1)?.ToString()?.Trim();

                            if (!string.IsNullOrEmpty(col1) && !string.IsNullOrEmpty(col2) && 
                                col1 != "DI" && col1 != "DO" && col1 != "AI" && col1 != "AO" &&
                                col1 != "Переменная" && col1 != "no data")
                            {
                                if (!map.ContainsKey(col1))
                                    map[col1] = col2;
                            }
                        }
                    }
                } while (reader.NextResult());
            }
            return map;
        }

        private Dictionary<int, Dictionary<string, int>> ReadSpecFile(string path)
        {
            var map = new Dictionary<int, Dictionary<string, int>>();
            
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                while (reader.Read())
                {
                    if (reader.Name == "Config_Line")
                    {
                        reader.Read(); // Заголовок
                        var headers = new Dictionary<string, int>();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            var h = reader.GetValue(i)?.ToString()?.Trim() ?? "";
                            if (!string.IsNullOrEmpty(h) && !headers.ContainsKey(h)) headers[h] = i;
                        }

                        while (reader.Read())
                        {
                            var posIdVal = reader.GetValue(0);
                            if (posIdVal == null || !int.TryParse(posIdVal.ToString(), out int posId)) continue;

                            var devices = new Dictionary<string, int>();
                            foreach (var devType in deviceTypes)
                            {
                                if (headers.ContainsKey(devType))
                                {
                                    var val = reader.GetValue(headers[devType]);
                                    if (val != null && int.TryParse(val.ToString(), out int idx) && idx > 0)
                                    {
                                        devices[devType] = idx;
                                    }
                                }
                            }
                            if (devices.Count > 0)
                                map[posId] = devices;
                        }
                        break;
                    }
                }
            }
            return map;
        }

        private void ReadAndGenerateIO(string path, Dictionary<string, string> varMap, Dictionary<int, Dictionary<string, int>> specMap, StringBuilder sb)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                while (reader.Read())
                {
                    if (reader.Name == "ШС")
                    {
                        // Пропуск служебных строк
                        reader.Read(); reader.Read(); reader.Read(); reader.Read();
                        if (!reader.Read()) break; // Заголовок

                        var colIndices = new Dictionary<string, int>();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            var rawHeader = reader.GetValue(i)?.ToString() ?? "";
                            var cleanHeader = rawHeader.Replace("\n", "").Replace("\r", "").Trim();
                            
                            if (cleanHeader.Contains("тип")) colIndices["type"] = i;
                            else if (cleanHeader.Contains("адрес")) colIndices["address"] = i;
                            else if (cleanHeader.Contains("обозн")) colIndices["label"] = i;
                            else if (cleanHeader.Contains("№ вх")) colIndices["signalNum"] = i;
                            else if (cleanHeader.Contains("Устройство")) colIndices["device"] = i;
                            else if (cleanHeader.Contains("Наименование сигнала")) colIndices["signalName"] = i;
                            else if (cleanHeader.Contains("Тех. поз")) colIndices["techPos"] = i;
                        }

                        if (!colIndices.ContainsKey("type") || !colIndices.ContainsKey("address"))
                        {
                            sb.AppendLine("// Ошибка: Не найдены обязательные колонки");
                            return;
                        }

                        var modulesDict = new Dictionary<string, IoModuleInfo>();
                        int currentModuleId = 0;

                        while (reader.Read())
                        {
                            string GetVal(string key)
                            {
                                if (!colIndices.ContainsKey(key)) return "";
                                var idx = colIndices[key];
                                if (idx >= reader.FieldCount) return "";
                                return reader.GetValue(idx)?.ToString()?.Trim() ?? "";
                            }

                            int GetIntVal(string key)
                            {
                                var s = GetVal(key);
                                return int.TryParse(s, out var res) ? res : 0;
                            }

                            string typeRaw = GetVal("type").Trim();
                            string type = typeRaw.StartsWith("DI") ? "DI" : (typeRaw.StartsWith("DO") ? "DO" : "");
                            
                            if (string.IsNullOrEmpty(type)) continue;

                            string address = GetVal("address");
                            string label = GetVal("label");
                            int signalNum = GetIntVal("signalNum");
                            string deviceName = GetVal("device");
                            string signalName = GetVal("signalName");
                            int techPos = GetIntVal("techPos");

                            if (string.IsNullOrEmpty(address) || signalNum == 0) continue;
                            if (deviceName == "Резерв" || string.IsNullOrEmpty(signalName) || signalName == "no data") continue;

                            string moduleKey = $"{type}_{address}_{label}";
                            
                            if (!modulesDict.ContainsKey(moduleKey))
                            {
                                currentModuleId++;
                                modulesDict[moduleKey] = new IoModuleInfo
                                {
                                    Id = currentModuleId,
                                    Type = type,
                                    Address = address,
                                    Label = label,
                                    Signals = new List<IoSignalInfo>(),
                                };
                            }

                            if (!varMap.ContainsKey(signalName)) continue;
                            
                            int devIndex = 0;
                            if (specMap.ContainsKey(techPos) && specMap[techPos].ContainsKey(deviceName))
                            {
                                devIndex = specMap[techPos][deviceName];
                            }
                            else
                            {
                                continue;
                            }

                            var signalInfo = new IoSignalInfo
                            {
                                SignalNumber = signalNum,
                                DeviceRu = deviceName,
                                SignalNameRu = signalName,
                                TechPos = techPos,
                                VarNameEn = varMap[signalName],
                                DevIndex = devIndex
                            };
                            
                            modulesDict[moduleKey].Signals.Add(signalInfo);
                        }

                        // Генерация кода
                        foreach (var mod in modulesDict.Values.OrderBy(m => m.Id))
                        {
                            var sortedSignals = mod.Signals.OrderBy(s => s.SignalNumber).ToList();
                            if (sortedSignals.Count == 0) continue;

                            string regionName = mod.Type == "DO" ? "MapDout" : "MapDin";
                            string maskVar = "#dwModuleBitMask";
                            string bitVar = "#xBit";

                            sb.AppendLine($"REGION Module {mod.Id}");
                            sb.AppendLine($"// {mod.Label}. Адрес {mod.Address}");
                            sb.AppendLine($"{maskVar} := 0; // Обнулим маску {(mod.Type == "DO" ? "выходов" : "входов")}");

                            foreach (var sig in sortedSignals)
                            {
                                string engDev = GetEngName(sig.DeviceRu);
                                sb.AppendLine($"{bitVar}.b{sig.SignalNumber} := \"{engDev}\".Dev[{sig.DevIndex}].{sig.VarNameEn};");
                            }

                            foreach (var sig in sortedSignals)
                            {
                                sb.AppendLine($"{maskVar} := {maskVar} OR {bitVar}.b{sig.SignalNumber};");
                            }

                            sb.AppendLine($"\"{regionName}\".Module[{mod.Id}].BitMask := {maskVar};");
                            sb.AppendLine("END_REGION");
                            sb.AppendLine();
                        }

                        break;
                    }
                }
            }
        }

        #endregion

        private void ShowOutput(string txtName, string text)
        {
            var txt = this.Controls.Find(txtName, true).FirstOrDefault() as TextBox;
            if (txt != null) txt.Text = text;
        }

        private void SaveToFile(string fileName, string content)
        {
            using (var dlg = new SaveFileDialog())
            {
                dlg.FileName = fileName;
                dlg.Filter = "Text Files|*.txt;*.scl|All Files|*.*";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(dlg.FileName, content, Encoding.UTF8);
                }
            }
        }
    }
}