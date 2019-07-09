using System;
using System.Drawing;
using System.Windows.Forms;

namespace Simplex_2.Формы_проекта
{
    

    public partial class Altman1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {              
        public Altman1()
        {
            InitializeComponent();
        }
             
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            Form1 main = this.Owner as Form1;
            double Z = 0;
            if (!textEdit1.Text.Equals(string.Empty))
            {
                if (main.gridView2.RowCount != 0)
                {
                    double var1 = Convert.ToDouble(main.gridView2.GetRowCellValue(10, main.gridView2.Columns[1])); // Итого по разделу II
                    double var2 = Convert.ToDouble(main.gridView2.GetRowCellValue(23, main.gridView2.Columns[1])); // Краткосрочные обязательства
                    double var3 = Convert.ToDouble(main.gridView2.GetRowCellValue(12, main.gridView2.Columns[1])); // Итого актив
                    double var4 = Convert.ToDouble(main.gridView2.GetRowCellValue(38, main.gridView2.Columns[1])); // Чистая прибыль
                    double var5 = Convert.ToDouble(main.gridView2.GetRowCellValue(37, main.gridView2.Columns[1])); // Прибыль до налогообложения
                    double var6 = Convert.ToDouble(textEdit1.Text);                                                // Рыночная стоимость акций
                    double var7 = Convert.ToDouble(main.gridView2.GetRowCellValue(24, main.gridView2.Columns[1])); // Итого по разделу IV 
                    double var8 = Convert.ToDouble(main.gridView2.GetRowCellValue(30, main.gridView2.Columns[1])); // Выручка 
                    if (var3 != 0 && var7 != 0 && var2 != 0)
                    {
                        Z = Math.Round(1.2 * (var1 - var2) / var3 + 1.4 * (var4 / var3) + 3.3 * (var5 / var3) + 0.6 * (var6 / var7) + 1.0 * (var8 / var3), 2);
                        textEdit2.Text = Z.ToString();
                        richTextBox1.SelectionAlignment = HorizontalAlignment.Left;
                        if (Z < 1.81)
                            richTextBox1.Text = "Значение индекса Альтмана " + Z.ToString() + " соответствует высокой вероятности банкротства \n80-100% в течение ближайших двух лет.";
                        if (Z >= 1.81 && Z <= 2.77)
                            richTextBox1.Text = "Значение индекса Альтмана " + Z.ToString() + " соответствует cредней вероятности краха компании \n35-50% в течение ближайших двух лет.";
                        if (Z > 2.77 && Z <= 2.99)
                            richTextBox1.Text = "Значение индекса Альтмана " + Z.ToString() + " соответствует вероятности банкротства \n15-20% в течение ближайших двух лет.";
                        if (Z > 2.99)
                            richTextBox1.Text = "Ситуация на предприятии стабильна, риск неплатежеспособности в течение ближайших двух лет крайне мал.";
                    }
                    else
                    {
                        MessageBox.Show("Среди указанных показателей есть равные 0:\n«Итого актив», «Краткосрочные обязательства», «Итого по разделу IV».", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }

                else
                {
                    MessageBox.Show("Таблица результатов оптимизации пустая, нельзя обновить данные.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }
            else
            {
                MessageBox.Show("Для продолжения расчетов введите значение показателя «Рыночная стоимость акций» (Рыночная стоимость \nакционерного капитала MVE).", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void textEdit2_EnabledChanged(object sender, EventArgs e)
        {
            ((TextBox)sender).ForeColor = Color.Black;
        }
    }
 
}
