using System;
using System.Drawing;
using System.Windows.Forms;

namespace Simplex_2.Формы_проекта
{


    public partial class Altman2 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public Altman2()
        {
            InitializeComponent();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            Form1 main = this.Owner as Form1;
            double Z = 0;
            if (main.gridView2.RowCount != 0)
            {
                double var1 = Convert.ToDouble(main.gridView2.GetRowCellValue(10, main.gridView2.Columns[1])); // Итого по разделу II
                double var2 = Convert.ToDouble(main.gridView2.GetRowCellValue(23, main.gridView2.Columns[1])); // Краткосрочные
                double var3 = Convert.ToDouble(main.gridView2.GetRowCellValue(12, main.gridView2.Columns[1])); // Итого актив
                double var4 = Convert.ToDouble(main.gridView2.GetRowCellValue(38, main.gridView2.Columns[1])); // Чистая прибыль
                double var5 = Convert.ToDouble(main.gridView2.GetRowCellValue(37, main.gridView2.Columns[1])); // Прибыль до налогообложения
                double var6 = Convert.ToDouble(main.gridView2.GetRowCellValue(19, main.gridView2.Columns[1])); // Итого по разделу III
                double var7 = Convert.ToDouble(main.gridView2.GetRowCellValue(24, main.gridView2.Columns[1])); // Итого по разделу IV 
                double var8 = Convert.ToDouble(main.gridView2.GetRowCellValue(30, main.gridView2.Columns[1])); // Выручка 
                if (var3 != 0 && var7 != 0)
                {
                    Z = Math.Round(0.717 * (var1 - var2) / var3 + 0.847 * (var4 / var3) + 3.107 * (var5 / var3) + 0.42 * (var6 / var7) + 0.995 * (var8 / var3), 2);
                    textEdit2.Text = Z.ToString();
                    richTextBox1.SelectionAlignment = HorizontalAlignment.Left;
                    if (Z < 1.23)
                        richTextBox1.Text = "Значение индекса Альтмана " + Z.ToString() + " соответствует высокой вероятности банкротства \n80-100% в течение ближайших двух лет.";
                    if (Z >= 1.23 && Z <= 2.9)
                        richTextBox1.Text = "Ситуация неопределенна.";
                    if (Z > 2.9)
                        richTextBox1.Text = "Ситуация на предприятии стабильна, риск неплатежеспособности в течение ближайших двух лет крайне мал.";
                }
                else
                {
                    MessageBox.Show("Среди указанных показателей есть равные 0:\n«Итого актив», «Итого по разделу IV».", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }
            else
            {
                MessageBox.Show("Таблица результатов оптимизации пустая, нельзя обновить данные.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void textEdit2_EnabledChanged(object sender, EventArgs e)
        {
            ((TextBox)sender).ForeColor = Color.Black;
        }
    }

}
