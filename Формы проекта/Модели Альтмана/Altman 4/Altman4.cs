using System;
using System.Drawing;
using System.Windows.Forms;

namespace Simplex_2.Формы_проекта
{
    

    public partial class Altman4 : DevExpress.XtraBars.Ribbon.RibbonForm
    {              
        public Altman4()
        {
            InitializeComponent();
        }
             
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            Form1 main = this.Owner as Form1;
            double P = 0;
            double X1 = 0;
            double X2 = 0;
            double X3 = 0;
            double X4 = 0;
            double X5 = 0;
            if (!textEdit1.Text.Equals(string.Empty))
            {
                if (main.gridView2.RowCount != 0)
                {
                    double var1 = Convert.ToDouble(main.gridView2.GetRowCellValue(10, main.gridView2.Columns[1])); // Итого по разделу II
                    double var2 = Convert.ToDouble(main.gridView2.GetRowCellValue(23, main.gridView2.Columns[1])); // Краткосрочные обязательства
                    double var3 = Convert.ToDouble(main.gridView2.GetRowCellValue(12, main.gridView2.Columns[1])); // Итого актив
                    double var4 = Convert.ToDouble(main.gridView2.GetRowCellValue(38, main.gridView2.Columns[1])); // Чистая прибыль
                    double var5 = Convert.ToDouble(main.gridView2.GetRowCellValue(37, main.gridView2.Columns[1])); // Прибыль до налогообложения
                    double var6 = Convert.ToDouble(textEdit1.Text);                                                // Проценты к уплате
                    double var7 = Convert.ToDouble(main.gridView2.GetRowCellValue(30, main.gridView2.Columns[1])); // Выручка 
                    double var8 = Convert.ToDouble(main.gridView2.GetRowCellValue(8, main.gridView2.Columns[1]));  // Денежные средства
                    double var9 = Convert.ToDouble(main.gridView2.GetRowCellValue(19, main.gridView2.Columns[1])); // Итого по разделу III

                    if (!var3.Equals(0))
                    {
                        X1 = Math.Round(var5 / var3);
                        X3 = Math.Round(var4 / var3);
                        X4 = Math.Round(var8 / var3);
                    }
                    else
                    {
                        MessageBox.Show("Значение показателя «Итого актив» равно 0.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    if (!var9.Equals(0))
                    {
                        X2 = Math.Round(var2 / var9, 2);
                    }
                    else
                    {
                        MessageBox.Show("Значение показателя «Итого по разделу III» равно 0.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    if (!var6.Equals(0))
                    {
                        X5 = Math.Round(var5 / var6);
                    }
                    else
                    {
                        X5 = 0;
                    }
                    double Y = Math.Round(Convert.ToDouble(4.28 + 0.18 * X1 - 0.01 * X2 + 0.08 * X3 + 0.02 * X4 + 0.19 * X5), 2);
                    P = 1.0 / (1 + Math.Exp(Y));
                    double pp = P * 100;
                    textEdit2.Text = P.ToString("0.##############");
                    richTextBox1.SelectionAlignment = HorizontalAlignment.Left;
                    richTextBox1.Text = "Вероятность банкротства составляет " + pp.ToString("0.##############") + " %.";
                }
                else
                {
                    MessageBox.Show("Таблица результатов оптимизации пустая, нельзя обновить данные.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }
            else
            {
                MessageBox.Show("Для продолжения расчетов введите значение показателя «Проценты к уплате» в соответствии с отчетом о финансовых результатах.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void textEdit2_EnabledChanged(object sender, EventArgs e)
        {
            ((TextBox)sender).ForeColor = Color.Black;
        }
    }
 
}
