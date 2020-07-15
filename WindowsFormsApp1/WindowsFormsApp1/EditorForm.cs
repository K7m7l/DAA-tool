using ClosedXML.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class EditorForm : Form
    {
        public EditorForm()
        {
            InitializeComponent();

            this.WindowState = FormWindowState.Maximized;

        }

        DataTable perfTable = new DataTable();

        private void ribbonButton3_Click(object sender, EventArgs e)
        {
            //add
            addStepItem(string.Empty, string.Empty);
        }

        Thread runner, perfThread;
        DateTime startTime;
        private void ribbonButton4_Click(object sender, EventArgs e)
        {
            //run
            if(flowLayoutPanel1.Controls.Count > 0)
            {
                

                List<Stepifier> steps = makeFlow();
                Type type = Type.GetType("WindowsFormsApp1.Util");

                runner = new Thread(() =>
                {
                    Thread.CurrentThread.IsBackground = true;
                    /* run your code here */

                    startTime = DateTime.Now;
                    perfThread = new Thread(() =>
                    {
                        Thread.CurrentThread.IsBackground = true;
                        if (perfTable.Columns.Count > 0)
                        {
                            perfTable.Rows.Clear();
                        }
                        else
                        {
                            perfTable.Columns.Add("Time");
                            perfTable.Columns.Add("RAM");
                            perfTable.Columns.Add("CPU");
                        }

                        while (true && runner.ThreadState == System.Threading.ThreadState.Running)
                        {
                            LogPeformance();
                        }

                    });
                    perfThread.Start();

                    foreach (Stepifier step in steps)
                    {
                        object instance = Activator.CreateInstance(type);
                        MethodInfo method = type.GetMethod(step.Step);
                        object[] paramss = new object[1] { step.Values };

                        method.Invoke(instance, paramss);
                    }

                });
                runner.Start();
            }
        }

        private void ribbonButton2_Click(object sender, EventArgs e)
        {
            //save
            if (flowLayoutPanel1.Controls.Count > 0)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.DefaultExt = "json";
                saveFileDialog.Filter = "json files (*.json)|*.json|All files (*.*)|*.*";
                saveFileDialog.FilterIndex = 1;
                DialogResult dr = saveFileDialog.ShowDialog();
                if (saveFileDialog.FileName != "")
                {
                    List<Stepifier> steps = makeFlow();
                    string json = JsonConvert.SerializeObject(steps);
                    //string filePath = System.IO.Path.Combine(Syroot.Windows.IO.KnownFolders.Desktop.Path, "flow" + DateTime.Now.ToString("MMddyyyyhhmmss"));
                    string filePath = saveFileDialog.FileName;
                    System.IO.File.WriteAllText(filePath, json);
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No Steps");
            }
        }
        private void ribbonButton1_Click(object sender, EventArgs e)
        {
            //open
            if (flowLayoutPanel1.Controls.Count > 0)
            {
                MessageBoxResult result = System.Windows.MessageBox.Show("Clear Existing Steps ?","Warning", MessageBoxButton.OKCancel);

                if (result == MessageBoxResult.OK)
                {
                    flowLayoutPanel1.Controls.Clear();
                }
                else
                {
                    return;
                }
            }
            else
            {

            }

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "json files (*.json)|*.json|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;
            DialogResult dr = openFileDialog.ShowDialog();
            if (openFileDialog.FileName != "" && System.IO.Path.GetExtension(openFileDialog.FileName).ToLower().IndexOf("json") > -1)
            {
                List<Stepifier> items = new List<Stepifier>();
                using (StreamReader r = new StreamReader(openFileDialog.FileName))
                {
                    string json = r.ReadToEnd();
                    items = JsonConvert.DeserializeObject<List<Stepifier>>(json);
                }
                foreach(Stepifier item in items)
                {
                    addStepItem(item.Step, item.Values);
                }
            }
        }
        private List<Stepifier> makeFlow()
        {
            List<Stepifier> stepifiers = new List<Stepifier>();

            Control.ControlCollection controls = flowLayoutPanel1.Controls;
            int counter = 1;

            foreach(Control control in controls)
            {
                var cBoxn = ((Panel)control).Controls.OfType<ComboBox>().First();
                string stSelect = cBoxn.SelectedItem.ToString();
                var rBoxn = ((Panel)control).Controls.OfType<RichTextBox>().First();
                string stValue = rBoxn.Text.ToString();
                Stepifier stepifier = new Stepifier();
                stepifier.ID = counter;
                stepifier.Step = stSelect;
                stepifier.Values = stValue;
                stepifiers.Add(stepifier);
                counter++;
            }

            return stepifiers;
        }
        private void addStepItem(string stepName, string stepValue)
        {
            Panel panel = new Panel();
            panel.Width = flowLayoutPanel1.Width;
            ComboBox cBox = new ComboBox();
            cBox.Width = (panel.Width * 3) / 4;
            cBox.DropDownStyle = ComboBoxStyle.DropDownList;
            cBox.Items.Add("Open");
            cBox.Items.Add("Close");
            cBox.Items.Add("Switch");
            cBox.Items.Add("CheckControlExists");
            cBox.Items.Add("Click");
            cBox.Items.Add("SetText");
            cBox.Items.Add("GetText");
            cBox.Items.Add("SetTextToClipBoard");
            cBox.Items.Add("GetTextFromClipBoard");
            cBox.Items.Add("Sleep");
            cBox.Items.Add("Scroll");
            cBox.Items.Add("FindControl");
            cBox.MouseWheel += (l, p) =>
            {
                ((HandledMouseEventArgs)p).Handled = true;
            };

            Button remover = new Button();
            remover.Width = panel.Width / 7;
            remover.Location = new System.Drawing.Point(((panel.Width * 4) / 5), 0);
            remover.Text = "Remove";
            remover.Click += (k, o) =>
            {
                ((FlowLayoutPanel)((Panel)((Button)k).Parent).Parent).Controls.Remove((Panel)((Button)k).Parent);
            };
            RichTextBox richer = new RichTextBox();
            richer.Width = panel.Width;
            richer.Location = new System.Drawing.Point(0, cBox.Height);

            if(stepName != string.Empty)
            {
                cBox.SelectedItem = stepName;
            }
            if (stepValue != string.Empty)
            {
                richer.Text = stepValue;
            }

            panel.Controls.Add(cBox);
            panel.Controls.Add(remover);
            panel.Controls.Add(richer);
            flowLayoutPanel1.Controls.Add(panel);
        }
        
        class Stepifier
        {
            public int ID { get; set; }
            public string Step { get; set; }
            public string Values { get; set; }
        }

        private void ribbonButton5_Click(object sender, EventArgs e)
        {
            MessageBoxResult result = System.Windows.MessageBox.Show("Clear Existing Steps ?", "Warning", MessageBoxButton.OKCancel);

            if (result == MessageBoxResult.OK)
            {
                flowLayoutPanel1.Controls.Clear();
            }
            else
            {
                return;
            }
        }

        private void ribbonButton6_Click(object sender, EventArgs e)
        {
            if(runner != null)
            {
                runner.Abort();

                DateTime stopTime = DateTime.Now;
                TimeSpan ttlTime = stopTime - startTime;
                perfThread.Abort();

                string logPath = @"C:\DAA-Files\";
                if (!Directory.Exists(logPath))
                {
                    Directory.CreateDirectory(logPath);
                }

                logPath = Path.Combine(logPath, "perfLog" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx");
                DataSet ds = new DataSet();
                ds.Tables.Add(perfTable);

                ExportDataSetToExcel(ds, logPath);

                System.Windows.Forms.MessageBox.Show("Completed. Total Time Taken : " + ttlTime.TotalSeconds.ToString());
            }  
        }
        public void LogPeformance()
        {
            string logPath = @"C:\DAA-Files\";
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            logPath = Path.Combine(logPath, "perfLog" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".txt");
            string logPath1 = Path.Combine(logPath, "perfLog" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx");

            string rammer = new Performance().getAvailableRAM();
            Thread.Sleep(500);
            //string cpuer = new Performance().getCurrentCpuUsage();
            string cpuer = new Performance().GetCpuUsage().ToString() + "%";

            perfTable.Rows.Add(DateTime.Now.ToString("yyyyMMddhhmmss"), rammer, cpuer);


            try
            {
                // Check if file already exists. If yes, delete i
                if (File.Exists(logPath))
                {
                    using (var tw = new StreamWriter(logPath, true))
                    {
                        tw.WriteLine("=====================================================================================================================");
                        tw.WriteLine("CPU :" + cpuer + "" + Environment.NewLine + "Date :" + DateTime.Now.ToString());
                        tw.WriteLine("RAM :" + rammer + "" + Environment.NewLine + "Date :" + DateTime.Now.ToString());
                        tw.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
                        tw.Close();
                    }
                }
                else
                {
                    // Create a new file 
                    using (StreamWriter sw = File.CreateText(logPath))
                    {
                        sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                        sw.WriteLine(Environment.MachineName);
                        sw.WriteLine("****//\\\\*****");
                        sw.WriteLine("CPU :" + cpuer + "" + Environment.NewLine + "Date :" + DateTime.Now.ToString());
                        sw.WriteLine("RAM :" + rammer + "" + Environment.NewLine + "Date :" + DateTime.Now.ToString());
                        sw.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
                        sw.WriteLine("=====================================================================================================================");
                        sw.Close();
                    }
                }
            }
            catch (Exception Ex)
            {
                LogException(Ex);
            }
        }
        public static void ExportDataSetToExcel(DataSet ds, string filepath)
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    wb.Worksheets.Add(ds.Tables[i], ds.Tables[i].TableName);
                }
                wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Style.Font.Bold = true;
                wb.SaveAs(filepath);
            }
        }
        public void LogException(Exception ex, [CallerMemberName] string memberName = "")
        {
            string logPath = @"C:\DAA-Files\";
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            logPath = Path.Combine(logPath, "dayLog" + DateTime.Now.ToString("yyyyMMdd") + ".txt");

            try
            {
                // Check if file already exists. If yes, delete i
                if (File.Exists(logPath))
                {
                    using (var tw = new StreamWriter(logPath, true))
                    {
                        tw.WriteLine("=====================================================================================================================");
                        tw.WriteLine("Message :" + ex.Message + "<br/>" + Environment.NewLine + "StackTrace :" + ex.StackTrace + "" + Environment.NewLine + "Date :" + DateTime.Now.ToString());
                        tw.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
                        tw.Close();
                    }
                }
                else
                {
                    // Create a new file 
                    using (StreamWriter sw = File.CreateText(logPath))
                    {
                        sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                        sw.WriteLine(Environment.MachineName);
                        sw.WriteLine("****//\\\\*****");
                        sw.WriteLine("Message :" + ex.Message + "<br/>" + Environment.NewLine + "StackTrace :" + ex.StackTrace + "" + Environment.NewLine + "Date :" + DateTime.Now.ToString());
                        sw.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
                        sw.WriteLine("=====================================================================================================================");
                        sw.Close();
                    }
                }
            }
            catch (Exception Ex)
            {
                string ec = Ex.Message;
            }
        }
    }
    class Performance
    {
        public string getCurrentCpuUsage()
        {
            PerformanceCounter cpuCounter;
            cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");
            return cpuCounter.NextValue() + "%";
        }
        public string getAvailableRAM()
        {
            PerformanceCounter ramCounter;
            ramCounter = new PerformanceCounter("Memory", "Available MBytes");
            return ramCounter.NextValue() + "MB";
        }
        public int GetCpuUsage()
        {
            var cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total", Environment.MachineName.ToString());
            cpuCounter.NextValue();
            System.Threading.Thread.Sleep(1000);
            return (int)cpuCounter.NextValue();
        }
    }
}
