using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Threading;
using System.IO;
using System.IO.Ports;
using System.Configuration;

using System.Collections.Specialized;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

using System;
using System.Collections.Generic;
using System.Text;
using System.Management;
using System.ComponentModel;
using System.Diagnostics;

using System.Runtime.InteropServices;

namespace ArduinoSerialLogger
{
    public partial class Form1 : Form
    {
        static SerialPort _serialPort;
        bool waitingForData = false;
        public Form1()
        {
            InitializeComponent();
            fillBaudRateBox(SerialPort.GetPortNames());
            _serialPort = new SerialPort();

            //Properties.Settings.Default.Reset();

            int baudRate = Properties.Settings.Default.baudRate;
            if (baudRate == 1)
            {
                baudRate = 9600;
                Properties.Settings.Default.baudRate = baudRate;
                Properties.Settings.Default.Save();
            }
            baudrateNumericUpDown.Value = baudRate;

            bool saveDataToFile = Properties.Settings.Default.saveDataToFile;
            saveDataCheckBox.Checked = saveDataToFile;

            string saveLocation = Properties.Settings.Default.saveLocation;
            if (saveLocation == "")
            {
                saveLocation = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Log.txt");
                Properties.Settings.Default.saveLocation = saveLocation;
                Properties.Settings.Default.Save();
            }
            saveLocationLabel.Text = saveLocation;

            StringCollection delimiterCollection = Properties.Settings.Default.delimiterCollection;
            delimiterComboBox.Items.Clear();
            bool delimiterFound = false;
            foreach (string delimiter in delimiterCollection)
            {
                delimiterComboBox.Items.Add(delimiter);
                if (delimiter == Properties.Settings.Default.delimiter)
                {
                    delimiterComboBox.SelectedItem = delimiter;
                    delimiterFound = true;
                }
            }
            if (portBox.Items.Count > 0 && !delimiterFound)
            {
                delimiterComboBox.SelectedIndex = 0;
            }

            writeToExcelCheckbox.Checked = Properties.Settings.Default.writeToExcel;

            if (Properties.Settings.Default.afterWritingToCell == 0)
            {
                noActionRadioButton.Checked = true;
            }
            if (Properties.Settings.Default.afterWritingToCell == 1)
            {
                nextRowRadioButton.Checked = true;
            }
            if (Properties.Settings.Default.afterWritingToCell == 2)
            {
                nextColumnRadioButton.Checked = true;
            }
        }


        private void refreshPortsButton_Click(object sender, EventArgs e)
        {
            fillBaudRateBox(SerialPort.GetPortNames());
        }


        private void fillBaudRateBox(string[] ports)
        {
            bool portFound = false;
            foreach (string port in ports)
            {
                Console.WriteLine(port);
                portBox.Items.Add(port);
                if (port == Properties.Settings.Default.port)
                {
                    portBox.SetSelected(portBox.Items.Count - 1, true);
                    portFound = true;
                }
            }
            if (portBox.Items.Count > 0 && !portFound)
            {
                portBox.SetSelected(0, true);
            }
        }

        private void writeLineToFile(String s)
        {
            //string appPath = Path.GetDirectoryName(AppContext.BaseDirectory);
            string appPath = saveLocationLabel.Text;
            //string appPath = "C:\\Users\\FH\\Desktop\\file.txt";
            using (StreamWriter writer = File.AppendText(appPath))
            {
                writer.Write(s);
            }
            // Read a file
            //string readText = File.ReadAllText(appPath);
            //Console.WriteLine(readText);
        }

        public async Task requestContinuousData()
        {
            string delimiter = Properties.Settings.Default.delimiter;
            if (delimiter == "newline")
            {
                delimiter = "\n";
            }
            if (delimiter == "carriage return")
            {
                delimiter = "\r";
            }

            _serialPort.ReadExisting();
            string s = "";
            waitingForData = true;
            while (waitingForData)
            {
                String r = _serialPort.ReadExisting();
                if (r == "")
                {
                    continue;
                }
                foreach (char c in r)
                {
                    if (c.ToString() == delimiter)
                    {
                        Console.WriteLine(s);
                        if (saveDataCheckBox.Checked)
                        {
                            writeLineToFile(s);
                        }
                        manageExcel(s);
                        s = "";
                    }
                    else
                    {
                        s += c;
                    }
                }
                //writeLineToFile();
                //Console.WriteLine(s);
                Thread.Sleep(200);
            }
            
        }

        private async void requestContinuousDataButton_Click(object sender, EventArgs e)
        {
            requestContinuousDataButton.Enabled = false;
            requestDataButton.Enabled = false;
            closeConnectionButton.Enabled = true;
            openSerialPort();
            _serialPort.ReadExisting();
            await Task.Run(() => requestContinuousData());
            requestContinuousDataButton.Enabled = true;
            requestDataButton.Enabled = true;
            closeConnectionButton.Enabled = false;
            //await requestContinuousData();
            //requestContinuousData();
            //Task obj = new Task(requestContinuousData);
            //obj.Start();
        }

        
        private async void requestDataButton_Click(object sender, EventArgs e)
        {
            requestContinuousDataButton.Enabled = false;
            requestDataButton.Enabled = false;
            closeConnectionButton.Enabled = true;
            openSerialPort();
            _serialPort.ReadExisting();
            //requestData();
            await Task.Run(() => requestData());
            //await requestData();
            requestContinuousDataButton.Enabled = true;
            requestDataButton.Enabled = true;
            closeConnectionButton.Enabled = false;
        }


        private void openSerialPort()
        {
            if (!_serialPort.IsOpen)
            {
                _serialPort.PortName = portBox.SelectedItem.ToString();
                _serialPort.BaudRate = (int)baudrateNumericUpDown.Value;
                _serialPort.DtrEnable = true;
                _serialPort.Open();
            }
        }

        public async Task requestData()
        {
            string delimiter = Properties.Settings.Default.delimiter;
            if (delimiter == "newline")
            {
                delimiter = "\n";
            }
            if (delimiter == "carriage return")
            {
                delimiter = "\r";
            }

            _serialPort.ReadExisting();
            _serialPort.Write(new byte[] { 1 }, 0, 1);
            string s = "";
            waitingForData = true;
            long time = DateTime.Now.Ticks;
            while (waitingForData)
            {
                long elapsedTicks = DateTime.Now.Ticks - time;
                TimeSpan elapsedSpan = new TimeSpan(elapsedTicks);
                if (elapsedSpan.TotalSeconds > 3)
                {
                    waitingForData = false;
                }
                String r = _serialPort.ReadExisting();
                if (r == "")
                {
                    continue;
                }
                foreach (char c in r)
                {
                    if (c.ToString() == delimiter)
                    {
                        Console.WriteLine(s);
                        if (saveDataCheckBox.Checked)
                        {
                            writeLineToFile(s);
                        }
                        manageExcel(s);
                        s = "";
                        waitingForData = false;
                        break;
                    }
                    else
                    {
                        s += c;
                    }
                }
                Thread.Sleep(200);
            }
        }

        private Excel.Application getExcelApp()
        {
            if (Process.GetProcessesByName("excel").Length > 0)
            {
                Excel.Application app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                return app;
            }
            else
            {
                return null;
            }
        }

        private void manageExcel(string value)
        {
            if (writeToExcelCheckbox.Checked)
            {
                Excel.Application app = getExcelApp();
                if (app != null)
                {
                    writeToExcelCell(app, value);
                }
                if (nextRowRadioButton.Checked)
                {
                    moveActiveCellRow(app);
                }
                if (nextColumnRadioButton.Checked)
                {
                    moveActiveCellColumn(app);
                }
            }
        }

        private void writeToExcelCell(Excel.Application app, string value)
        {
            Excel.Workbook book = app.ActiveWorkbook;
            Excel.Worksheet sheet = book.ActiveSheet;
            Excel.Range range = app.ActiveCell;

            int row = range.Row;
            int column = range.Column;

            sheet.Cells[row, column] = value;
        }

        private void moveActiveCellRow(Excel.Application app)
        {
            Excel.Workbook book = app.ActiveWorkbook;
            Excel.Worksheet sheet = book.ActiveSheet;
            Excel.Range range = app.ActiveCell;

            int row = range.Row;
            int column = range.Column;

            Excel.Range newRange = (Excel.Range)sheet.Cells[row + 1, column];
            newRange.Select();
        }

        private void moveActiveCellColumn(Excel.Application app)
        {
            Excel.Workbook book = app.ActiveWorkbook;
            Excel.Worksheet sheet = book.ActiveSheet;
            Excel.Range range = app.ActiveCell;

            int row = range.Row;
            int column = range.Column;
            
            Excel.Range newRange = (Excel.Range)sheet.Cells[row, column + 1];
            newRange.Select();
        }


        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (_serialPort.IsOpen)
            {
                _serialPort.Close();
            }
        }

        private void closeConnectionButton_Click(object sender, EventArgs e)
        {
            waitingForData = false;
            if (_serialPort.IsOpen)
            {
                _serialPort.Close();
            }
        }

        private void saveLocationLabel_Click(object sender, EventArgs e)
        {
            SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
            SaveFileDialog1.InitialDirectory = Environment.SpecialFolder.Desktop.ToString();
            //SaveFileDialog1.CheckFileExists = true;
            SaveFileDialog1.CheckPathExists = true;
            SaveFileDialog1.RestoreDirectory = true;
            SaveFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            SaveFileDialog1.Title = "Choose Save Location";
            SaveFileDialog1.DefaultExt = "txt";
            if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                saveLocationLabel.Text = SaveFileDialog1.FileName;
                Properties.Settings.Default.saveLocation = saveLocationLabel.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void portBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.port = portBox.SelectedItem.ToString();
            Properties.Settings.Default.Save();
        }

        private void baudrateNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.baudRate = (int)baudrateNumericUpDown.Value;
            Properties.Settings.Default.Save();
        }

        private void saveDataCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.saveDataToFile = saveDataCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        private void delimiterComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.delimiter = delimiterComboBox.Text;
            Properties.Settings.Default.Save();
        }

        public static string ListAllApplications()
        {
            StringBuilder sb = new StringBuilder();

            foreach (Process p in Process.GetProcesses("."))
            {
                try
                {
                    if (p.MainWindowTitle.Length > 0)
                    {
                        //sb.Append("Window Title:\t" +
                        //    p.MainWindowTitle.ToString()
                        //    + Environment.NewLine);

                        //sb.Append("Process Name:\t" +
                        //    p.ProcessName.ToString()
                        //    + Environment.NewLine);

                        //sb.Append("Window Handle:\t" +
                        //    p.MainWindowHandle.ToString()
                        //    + Environment.NewLine);

                        //sb.Append("Memory Allocation:\t" +
                        //    p.PrivateMemorySize64.ToString()
                        //    + Environment.NewLine);

                        //sb.Append(Environment.NewLine);

                        Console.WriteLine(p.MainWindowTitle.ToString());
                        if (p.MainWindowTitle.ToString().Contains("Excel"))
                        {
                            Console.WriteLine("???");
                            Excel.Application instance = (Excel.Application)p;
                            Console.WriteLine(instance.ActiveSheet);
                            //Excel.Worksheet cell = instance.Sheets;
                            //Console.WriteLine("excel");
                            //Console.WriteLine(cell);
                        }
                        else
                        {
                            Console.WriteLine("not found");
                        }
                    }
                }
                catch { }
            }

            return sb.ToString();
        }

        private void noActionRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.afterWritingToCell = 0;
            Properties.Settings.Default.Save();
        }

        private void nextRowRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.afterWritingToCell = 1;
            Properties.Settings.Default.Save();
        }

        private void nextColumnRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.afterWritingToCell = 2;
            Properties.Settings.Default.Save();
        }

        private void writeToExcelCheckbox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.writeToExcel = writeToExcelCheckbox.Checked;
            Properties.Settings.Default.Save();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
