using System;
using System.Text;
using System.Windows.Forms;
using System.Security.Cryptography;
using System.IO;

namespace Simplex_2.Формы_проекта
{
    public partial class Password : Form
    {
        public static string pas = Path.Combine(Application.StartupPath, "pas.txt");

        public Password()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != string.Empty)
            {
                string password = textBox1.Text;
                string temp_pas = string.Empty;
                string hash = string.Empty;

                byte[] bytes = Encoding.Unicode.GetBytes(password);
                MD5CryptoServiceProvider CSP = new MD5CryptoServiceProvider();
                byte[] byteHash = CSP.ComputeHash(bytes);               
                foreach (byte b in byteHash)
                {
                    hash += string.Format("{0:x2}", b);
                }
                temp_pas = Convert.ToString(new Guid(hash));

                try
                {
                    StreamWriter writer1 = new StreamWriter(@pas, false, Encoding.GetEncoding(1251));   //запись в файл пароля
                    writer1.Write(temp_pas);
                    writer1.Close();
                }
                catch (Exception ez)
                {
                    MessageBox.Show(ez.Message, "ФАНЗ");
                    return;
                }
                MessageBox.Show("Пароль успешно сохранен.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Close();
            }
            else
            {
                MessageBox.Show("Введите пароль, поле пустое.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
