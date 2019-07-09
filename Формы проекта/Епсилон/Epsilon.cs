using System;
using System.Windows.Forms;

namespace Simplex_2.Формы_проекта
{
    public partial class Epsilon : Form
    {
        public Epsilon()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void t_KeyPress(object sender, KeyPressEventArgs e)
        {
                if (e.KeyChar == '.')
                    e.KeyChar = ',';
                if (e.KeyChar != 22)
                    e.Handled = !Char.IsDigit(e.KeyChar) && (e.KeyChar != ',' || (((System.Windows.Forms.TextBox)sender).Text.Contains(",") && !((System.Windows.Forms.TextBox)sender).SelectedText.Contains(","))) && e.KeyChar != (char)Keys.Back && (e.KeyChar != '-' || ((System.Windows.Forms.TextBox)sender).SelectionStart != 0 || (((System.Windows.Forms.TextBox)sender).Text.Contains("-") && !((System.Windows.Forms.TextBox)sender).SelectedText.Contains("-")));
                else
                {
                    double d;
                    e.Handled = !double.TryParse(Clipboard.GetText(), out d) || (d < 0 && (((System.Windows.Forms.TextBox)sender).SelectionStart != 0 || ((System.Windows.Forms.TextBox)sender).Text.Contains("-") && !((System.Windows.Forms.TextBox)sender).SelectedText.Contains("-"))) || ((d - (int)d) != 0 && ((System.Windows.Forms.TextBox)sender).Text.Contains(",") && !((System.Windows.Forms.TextBox)sender).SelectedText.Contains(","));
                    MessageBox.Show("Не удалось вставить содержимое буфера обмена.","ФАНЗ",MessageBoxButtons.OK,MessageBoxIcon.Stop);
                }
        }

        private void textBox1_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox1, e);
        }
    }
}
