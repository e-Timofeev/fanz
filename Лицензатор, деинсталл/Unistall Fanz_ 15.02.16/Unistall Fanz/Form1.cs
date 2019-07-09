using Microsoft.Win32;

using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace Unistall_Fanz
{
    public partial class Form1 : Form
    {
        public string path = string.Empty;


        public Form1()
        {
            InitializeComponent();
        }
        private void Error()
        {
            try
            {
                var keys = Registry.LocalMachine.OpenSubKey("Software\\Fanz", true);
                if (keys != null)
                {
                    Registry.LocalMachine.DeleteSubKey("Software\\Fanz");
                }
            }
            catch { }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (radioButton1.Checked)
                {
                    if (Directory.Exists(@"C:\Program Files (x86)\ФАНЗ"))
                    {
                        path = @"C:\Program Files (x86)\ФАНЗ\Uninstall.exe";
                        if (File.Exists(path))
                        {
                            System.Diagnostics.Process.Start(path);
                            Error();
                        }
                        Thread.Sleep(5000);
                        if (Directory.Exists(@"C:\Program Files (x86)\ФАНЗ"))
                        {
                            Directory.Delete(@"C:\Program Files (x86)\ФАНЗ", true);
                        }
                        Thread.Sleep(1000);
                        MessageBox.Show("Удаление выполнено успешно.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        if (Directory.Exists(@"C:\Program Files\ФАНЗ"))
                        {
                            path = @"C:\Program Files (x86)\ФАНЗ\Uninstall.exe";
                            if (File.Exists(path))
                            {
                                System.Diagnostics.Process.Start(path);
                                Error();
                            }
                            Thread.Sleep(5000);
                            if (Directory.Exists(@"C:\Program Files\ФАНЗ"))
                            {
                                Directory.Delete(@"C:\Program Files\ФАНЗ", true);
                            }
                            Thread.Sleep(1000);
                            MessageBox.Show("Удаление выполнено успешно.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else if (radioButton2.Checked)
                {
                    path = folderBrowserDialog1.SelectedPath + @"\Uninstall.exe";
                    if (File.Exists(path))
                    {
                        System.Diagnostics.Process.Start(path);
                        Error();
                    }
                    Thread.Sleep(5000);
                    if (Directory.Exists(folderBrowserDialog1.SelectedPath + @"\ФАНЗ"))
                    {
                        Directory.Delete(folderBrowserDialog1.SelectedPath + @"\ФАНЗ", true);
                    }
                    Thread.Sleep(1000);
                    MessageBox.Show("Удаление выполнено успешно.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
              
            }
            catch (Exception error)
            {
                DialogResult result;
                result = MessageBox.Show(error.ToString(), "ФАНЗ", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Stop);
                if (result == System.Windows.Forms.DialogResult.Abort)
                {
                    Application.Exit();
                }
                else
                {
                    return;
                }
            }
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                DialogResult dialogresult = folderBrowserDialog1.ShowDialog();
                folderBrowserDialog1.ShowDialog();
                string folderName = "";
                if (dialogresult == DialogResult.OK)
                {
                    folderName = folderBrowserDialog1.SelectedPath;
                }
            }
        }
    }
}
