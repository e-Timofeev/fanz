using System;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Security.Cryptography;


namespace Simplex_2.Формы_проекта
{
    public partial class Form2 : Form
    {
        public static string pas = Path.Combine(Application.StartupPath, "pas.txt");
        public static string host = Path.Combine(Application.StartupPath, "host.txt");
        
        public Form2()
        {
            InitializeComponent();
        }
     
        private void Form2_InputLanguageChanged(object sender, InputLanguageChangedEventArgs e)
        {
           toolStripStatusLabel1.Text = InputLanguage.CurrentInputLanguage.LayoutName;
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            string login = Environment.UserName.ToString();
            string password = maskedTextBox1.Text;
            string temp_pas = string.Empty;
            string file_hash = string.Empty;

            string hostname = Environment.UserDomainName.ToString();
            string host_tmp = string.Empty;
            string hash_host = string.Empty;
            string hash_real = string.Empty;
            DialogResult result;
            if (!File.Exists(@host))
            {
                if (File.Exists(@pas))
                {
                    #region Проверка пароля
                    using (FileStream fstream = File.OpenRead(@pas))
                    {
                        byte[] array = new byte[fstream.Length];
                        fstream.Read(array, 0, array.Length);
                        file_hash = System.Text.Encoding.Default.GetString(array);
                    }

                    if (radioButton2.Checked == true)
                    {
                        string s = maskedTextBox1.Text;
                        byte[] bytes = Encoding.Unicode.GetBytes(s);
                        MD5CryptoServiceProvider CSP = new MD5CryptoServiceProvider();
                        byte[] byteHash = CSP.ComputeHash(bytes);
                        string hash = string.Empty;
                        foreach (byte b in byteHash)
                            hash += string.Format("{0:x2}", b);
                        temp_pas = Convert.ToString(new Guid(hash));

                        if (file_hash == temp_pas)
                        {
                            this.Close();
                        }
                        else
                        {
                            result = MessageBox.Show("Неверно указан пароль администратора.", "ФАНЗ", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                            if (result == System.Windows.Forms.DialogResult.Cancel)
                            {
                                this.Close();
                                try
                                {
                                    Environment.Exit(0);
                                }
                                catch { }
                            }
                            else
                            {
                                maskedTextBox1.Text = "";
                            }
                        }
                    }
                    else
                    {
                        this.Close();
                    }
                    #endregion
                }
                else
                {
                    MessageBox.Show("Доступ запрещен. Обратитесь к администратору.\nПриложение будет закрыто.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    this.Close();
                    try
                    {
                        Environment.Exit(0);
                    }
                    catch { }
                }
            }
            else if (File.Exists(@host))
            {
                using (FileStream fstream = File.OpenRead(@host))
                {
                    byte[] array1 = new byte[fstream.Length];
                    fstream.Read(array1, 0, array1.Length);
                    hash_host = System.Text.Encoding.Default.GetString(array1);
                }

                byte[] bytes3 = Encoding.Unicode.GetBytes(hostname);
                MD5CryptoServiceProvider CSP3 = new MD5CryptoServiceProvider();
                byte[] byteHash3 = CSP3.ComputeHash(bytes3);
                foreach (byte b3 in byteHash3)
                {
                    hash_real += string.Format("{0:x2}", b3);
                }
                host_tmp = Convert.ToString(new Guid(hash_real));

                if (hash_host == host_tmp)
                {
                    if (File.Exists(@pas))
                    {
                        #region Просто проверка пароля
                        using (FileStream fstream = File.OpenRead(@pas))
                        {
                            byte[] array = new byte[fstream.Length];
                            fstream.Read(array, 0, array.Length);
                            file_hash = System.Text.Encoding.Default.GetString(array);
                        }

                        if (radioButton2.Checked == true)
                        {
                            string s = maskedTextBox1.Text;
                            byte[] bytes2 = Encoding.Unicode.GetBytes(s);
                            MD5CryptoServiceProvider CSP2 = new MD5CryptoServiceProvider();
                            byte[] byteHash2 = CSP2.ComputeHash(bytes2);
                            string hash = string.Empty;
                            foreach (byte b in byteHash2)
                                hash += string.Format("{0:x2}", b);
                            temp_pas = Convert.ToString(new Guid(hash));

                            if (file_hash == temp_pas)
                            {
                                this.Close();
                            }
                            else
                            {
                                result = MessageBox.Show("Неверно указан пароль администратора.", "ФАНЗ", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                                if (result == System.Windows.Forms.DialogResult.Cancel)
                                {
                                    this.Close();
                                    try
                                    {
                                        Environment.Exit(0);
                                    }
                                    catch { }
                                }
                                else
                                {
                                    maskedTextBox1.Text = "";
                                }
                            }
                        }
                        else
                        {
                            this.Close();
                        }
                        #endregion
                    }
                    else
                    {
                        MessageBox.Show("Доступ запрещен. Обратитесь к администратору.\nПриложение будет закрыто.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        this.Close();
                        try
                        {
                            Environment.Exit(0);
                        }
                        catch { }
                    }
                }
                else
                {
                    MessageBox.Show("Несанционированная попытка копирования!", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    this.Close();
                    try
                    {
                        Environment.Exit(0);
                    }
                    catch { }
                }
            }
        }     
        private void Form2_Load(object sender, EventArgs e)
        {
          radioButton1.Checked= true;
          maskedTextBox1.Enabled = false;
          textEdit1.Enabled = false;
          toolStripStatusLabel1.Text = InputLanguage.CurrentInputLanguage.LayoutName;
          maskedTextBox1.Text = string.Empty;
          textEdit1.Text = Environment.UserName.ToString();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (maskedTextBox1.Enabled != true)
            {
                maskedTextBox1.Enabled = true;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (maskedTextBox1.Enabled == true)
            {
                maskedTextBox1.Enabled = false;
            }
        }
    }
}
