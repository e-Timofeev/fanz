using lib;

using Microsoft.Win32;

using RSA;

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Management;
using System.Runtime.Serialization.Formatters.Binary;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;


namespace Лицензатор_ФАНЗ
{
    public partial class Form1 : Form
    {
        public static string put = Path.Combine(Application.StartupPath);
        public static string News = Path.Combine(Application.StartupPath + @"\content\News.txt");
        public string path = string.Empty;
        public string temp_pas = string.Empty;

        public class HardDrive
        {
            private string model = null;
            private string type = null;
            private string serialNo = null;

            public string Model
            {
                get { return model; }
                set { model = value; }
            }

            public string Type
            {
                get { return type; }
                set { type = value; }
            }

            public string SerialNo
            {
                get { return serialNo; }
                set { serialNo = value; }
            }
        }
        public void InfoPC()
        {
            string key = string.Empty;
            Dictionary<string, string> ids =
            new Dictionary<string, string>();
            ArrayList hdCollection = new ArrayList();
            string login = Environment.UserName.ToString();
            string hostname = Environment.UserDomainName.ToString();
            ids.Add("HostName", hostname + ":");
            ids.Add("UserName", login + ":");
            ManagementObjectSearcher searcher;

            try
            {
                searcher = new
                    ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive");

                foreach (ManagementObject wmi_HD in searcher.Get())
                {
                    HardDrive hd = new HardDrive();
                    hdCollection.Add(hd);
                }
                searcher = new
                    ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMedia");
                int i = 0;
                foreach (ManagementObject wmi_HD in searcher.Get())
                {
                    HardDrive hd = (HardDrive)hdCollection[i];
                    if (wmi_HD["SerialNumber"] == null)
                    {
                        hd.SerialNo = "None";
                    }
                    else
                    {
                        hd.SerialNo = wmi_HD["SerialNumber"].ToString();
                        ++i;
                    }
                }
                foreach (HardDrive hd in hdCollection)
                {
                    ids.Add("HDDid", hd.SerialNo + ":");
                }
            }
            catch { }
            try
            {
                searcher = new ManagementObjectSearcher("root\\CIMV2",
                       "SELECT * FROM Win32_Processor");
                foreach (ManagementObject queryObj in searcher.Get())
                    ids.Add("ProcessorId", queryObj["ProcessorId"].ToString() + ":");
            }
            catch { }
            try
            {
                searcher = new ManagementObjectSearcher("root\\CIMV2",
                       "SELECT * FROM CIM_Card");
                foreach (ManagementObject queryObj in searcher.Get())
                    ids.Add("CardID", queryObj["SerialNumber"].ToString() + ":");
            }
            catch { }
            try
            {
                searcher = new ManagementObjectSearcher("root\\CIMV2",
                       "SELECT * FROM CIM_OperatingSystem");
                foreach (ManagementObject queryObj in searcher.Get())
                    ids.Add("OSSerialNumber", queryObj["SerialNumber"].ToString() + ":");
            }
            catch { }
            try
            {
                searcher = new ManagementObjectSearcher("root\\CIMV2",
                       "SELECT UUID FROM Win32_ComputerSystemProduct");
                foreach (ManagementObject queryObj in searcher.Get())
                    ids.Add("UUID", queryObj["UUID"].ToString());
            }
            catch { }
            foreach (var x in ids)
            {
                key += x.Value;
            }
            string password = key;
            string hash = string.Empty;
            byte[] bytes = Encoding.Unicode.GetBytes(password);
            MD5CryptoServiceProvider CSP = new MD5CryptoServiceProvider();
            byte[] byteHash = CSP.ComputeHash(bytes);
            foreach (byte b in byteHash)
            {
                hash += string.Format("{0:x2}", b);
            }
            temp_pas = Convert.ToString(hash);
            textBox1.Text = temp_pas;
        }
        public Form1()
        {
            InitializeComponent();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            label14.Text = "";
            label15.Text = "";
            label16.Text = "";
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
            label14.Text = "";
            label15.Text = "";
            label16.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog opf = new OpenFileDialog();
            opf.Filter = "All files (*.*)|*.*";
            opf.InitialDirectory = @put;
            opf.ShowDialog();
            string filename = opf.FileName;
            if (filename != "")
            {
                #region загрузка первичного ключа из файла
                if (File.Exists(@filename))
                {
                    string file_hash = string.Empty;
                    using (FileStream fstream = File.OpenRead(@filename))
                    {
                        byte[] array = new byte[fstream.Length];
                        fstream.Read(array, 0, array.Length);
                        file_hash = System.Text.Encoding.Default.GetString(array);
                        fstream.Close();
                    }
                    textBox1.Text = file_hash;
                }
                else { }
                #endregion
            }
            else
            {
                MessageBox.Show("Не выбран файл для загрузки.", "ФАНЗ");
                return;
            }
            opf.RestoreDirectory = true;   
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox8.Clear();

            if (textBox1.Text != string.Empty)
            {
                int index = 0;
                string password = textBox1.Text;
                string version = string.Empty;
                string hash = string.Empty;
                string hashver = string.Empty;
                string tp = string.Empty;
                string oneKey = string.Empty;
                string twoKey = string.Empty;
                string RSApas = string.Empty;
                    #region получение хэш-сумм через MD5 первичного ключа и типа версии
                if (radioButton1.Checked || radioButton2.Checked || radioButton3.Checked || radioButton4.Checked)
                {
                    for (int i = 0; i < 4; i++)
                    {
                        if (((RadioButton)groupBox1.Controls[i]).Checked == true)
                        {
                            index = i;
                            if (index == 2)
                                textBox3.Text = index.ToString();
                            label10.Text = ((RadioButton)groupBox1.Controls[i]).Text;
                            if (index == 1)
                                textBox8.Text = index.ToString();
                            label10.Text = ((RadioButton)groupBox1.Controls[i]).Text;
                            if (index == 0)
                                textBox5.Text = index.ToString();
                            label10.Text = ((RadioButton)groupBox1.Controls[i]).Text;
                            if (index == 3)
                                textBox4.Text = index.ToString();
                            label10.Text = ((RadioButton)groupBox1.Controls[i]).Text;
                        }
                    }

                    byte[] bytes = Encoding.Unicode.GetBytes(password);
                    MD5CryptoServiceProvider CSP = new MD5CryptoServiceProvider();
                    byte[] byteHash = CSP.ComputeHash(bytes);
                    foreach (byte b in byteHash)
                    {
                        hash += string.Format("{0:x2}", b);
                    }
                    byte[] bytesver = Encoding.Unicode.GetBytes(index.ToString());
                    MD5CryptoServiceProvider CSPver = new MD5CryptoServiceProvider();
                    byte[] byteHashver = CSPver.ComputeHash(bytesver);
                    foreach (byte b in byteHashver)
                    {
                        hashver += string.Format("{0:x2}", b);
                    }
                    tp = Convert.ToString(hashver);
                    version = hashver;
                    textBox7.Text = tp;
                #endregion
                    RSAlib rb = new RSAlib();
                    RSApas = rb.encode(version);
                    oneKey = "" + rb.GetNKey();
                    twoKey = "" + rb.GetDKey();
                    textBox6.Text = oneKey + " : " + twoKey;
                    textBox2.Text = hash + ":" + RSApas + ":" + oneKey + "-" + twoKey;
                }
                else
                {
                    MessageBox.Show("Не выбрана версия для генерации ключа.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }
            else
            {
                MessageBox.Show("Загрузите первичный ключ, поле пустое.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        private void Error()
        {
            try
            {
                var keys = Registry.LocalMachine.OpenSubKey("Software\\Fanz", true);
                if (keys != null)
                {
                    Registry.LocalMachine.DeleteSubKey("Software\\Fanz");
                    label15.Text = "успешно";
                }
            }
            catch { label15.Text = "ошибка"; }
        }
        private void regCreateSubKey(string apps)
        {
            try
            {
                proba t = new proba();
                string encrypt = t.Encode(textBox2.Text);
                BinaryFormatter formatter = new BinaryFormatter();
                using (FileStream fs = new FileStream(@apps, FileMode.OpenOrCreate))
                {
                    formatter.Serialize(fs, encrypt);
                }
            }
            catch
            {
                MessageBox.Show("Ошибка при формировании ключа.\nОбратитесь к системному администратору.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            try
            {
                if (File.Exists(@apps))
                {
                    DateTime reg = File.GetCreationTime(apps);
                    string value = Convert.ToString(reg);
                    using (var keySoftware = Registry.LocalMachine.OpenSubKey("Software", true))
                    {
                        keySoftware.CreateSubKey("Fanz").Close();
                        RegistryKey rk = Registry.LocalMachine.OpenSubKey("Software\\Fanz", true);
                        rk.SetValue("RegDate", value, RegistryValueKind.String);
                        label16.Text = "успешно";
                    }
                }
                else
                {
                    label16.Text = "ошибка";
                    MessageBox.Show("Ошибка при формировании ключа - файл не создан.\nОбратитесь к системному администратору.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
            }
            catch
            {
                try
                {
                    File.Delete(@apps);
                }
                catch { }
                MessageBox.Show("У вас не хватает прав для установки ключа.\nОбратитесь к системному администратору.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            #region Удаление ключа и записей в реестре и установка нового
            try
            {
                if (radioButton6.Checked && textBox2.Text!=string.Empty)
                {
                    if (Directory.Exists(@"C:\Program Files (x86)\ФАНЗ"))
                    {
                        path = @"C:\Program Files (x86)\ФАНЗ\keys.licence";
                        if (File.Exists(path))
                        {
                            File.Delete(path);
                            label14.Text = "успешно";
                            Error();
                            regCreateSubKey(path);
                        }
                        else
                        {
                            label14.Text = "ошибка";
                            return;
                        }
                    }
                    else
                    {
                        if (Directory.Exists(@"C:\Program Files\ФАНЗ"))
                        {
                            path = @"C:\Program Files (x86)\ФАНЗ\keys.licence";
                            if (File.Exists(path))
                            {
                                File.Delete(path);
                                label14.Text = "успешно";
                                Error();
                                regCreateSubKey(path);
                            }
                            else
                            {
                                label14.Text = "ошибка";
                                return;
                            }
                        }
                    }
                }
                else if (radioButton5.Checked && textBox2.Text != string.Empty)
                {
                    path = folderBrowserDialog1.SelectedPath + @"\keys.licence";
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                        label14.Text = "успешно";
                        Error();
                        regCreateSubKey(path);
                    }
                    else
                    {
                        label14.Text = "ошибка";
                        return;
                    }
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
            #endregion

        }
        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked)
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

        private void button4_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            InfoPC();
        }
    }
}
