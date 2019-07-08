using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlServerCe;

namespace Simplex_2.Формы_проекта
{
    public partial class Dupon : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public bool result;

        public Dupon()
        {
            InitializeComponent();
        }
       
        private void Dupon_Load(object sender, EventArgs e)
        {
            Factor main = this.Owner as Factor;
            string year = string.Empty;
            string filter = string.Empty;
            int countRow = 0;
            #region Получаем год
            if (main.barEditItem1.EditValue.ToString()!=string.Empty)
            {
                try
                {
                    year = PeriodYear(main.barEditItem1.EditValue.ToString());
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    gridControl1.DataSource = null;
                    return;
                }
            }
            #endregion
            #region Формируем строку для фильтра по вьюхе
            try
            {
                SqlCeConnection conn;
                conn = new SqlCeConnection(@"Data Source = |DataDirectory|\FANZ.sdf");
                conn.Open();
                SqlCeDataAdapter adapter;
                adapter = new SqlCeDataAdapter("SELECT * FROM Altman", conn);
                DataTable dt = new DataTable();               
                adapter.Fill(dt);
                DataView MyDataView = new DataView(dt);

                filter = year;

                MyDataView.RowFilter = filter;

                countRow = MyDataView.Count;
                if (countRow == 0)
                {
                    gridControl1.DataSource = null;
                    MessageBox.Show("Некорректная загрузка данных из базы. Закройте текущее окно.\nНа форме  «Факторный анализ» выполните очистку и повторите алгоритм расчетов.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
                else
                {
                    gridControl1.DataSource = MyDataView;
                }
                conn.Close();
            }
            catch (Exception q)
            {
                gridControl1.DataSource = null;
                MessageBox.Show(q.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            #endregion
            gridView1.Columns[0].Visible = false;
            gridView1.Columns[4].Visible = false;
            for (int t = 0; t < gridView1.Columns.Count; t++)
            {
                gridView1.Columns[t].AppearanceHeader.Options.UseTextOptions = true;
                gridView1.Columns[t].AppearanceHeader.BackColor = Color.LightSteelBlue;
                gridView1.Columns[t].AppearanceHeader.BackColor2 = Color.White;
                gridView1.Columns[t].AppearanceHeader.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.Vertical;
                gridView1.Columns[t].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                gridView1.Columns[t].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                gridView1.Columns[t].BestFit();
            }
            gridView1.Columns[1].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Default;
            gridView1.Columns[1].Caption = "Показатель";
            gridView1.Columns[3].Caption = "Дата";
            gridView1.Columns[3].DisplayFormat.FormatString = "g";
        }
        public string PeriodYear(string id)                                               //год
        {
            string year = string.Empty;
            switch (id)
            {
                case "2015":
                    year = String.Format("([Дата_1] >= '{0}') AND ([Дата_1] < '{1}')", ("2015-01-01"), ("2016-01-01"));
                    break;
                case "2016":
                    year = String.Format("([Дата_1] >= '{0}') AND ([Дата_1] < '{1}')", ("2016-01-01"), ("2017-01-01"));
                    break;
                case "2017":
                    year = String.Format("([Дата_1] >= '{0}') AND ([Дата_1] < '{1}')", ("2017-01-01"), ("2018-01-01"));
                    break;
                case "2018":
                    year = String.Format("([Дата_1] >= '{0}') AND ([Дата_1] < '{1}')", ("2018-01-01"), ("2019-01-01"));
                    break;
                case "2019":
                    year = String.Format("([Дата_1] >= '{0}') AND ([Дата_1] < '{1}')", ("2019-01-01"), ("2020-01-01"));
                    break;
                case "2020":
                    year = String.Format("([Дата_1] >= '{0}') AND ([Дата_1] < '{1}')", ("2020-01-01"), ("2021-01-01"));
                    break;
            }
            return year;
        }

        private void TextBoxs()
        {
            if (gridView1.RowCount != 0)
            {
                textEdit1.Text = Convert.ToString(gridView1.GetRowCellValue(3, gridView1.Columns[2]));
                textEdit2.Text = Convert.ToString(gridView1.GetRowCellValue(2, gridView1.Columns[2]));
                textEdit3.Text = Convert.ToString(gridView1.GetRowCellValue(0, gridView1.Columns[2]));
                textEdit4.Text = Convert.ToString(gridView1.GetRowCellValue(1, gridView1.Columns[2]));
            }
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            TextBoxs();
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            if (textEdit1.Text != "" && textEdit2.Text != "" && textEdit3.Text != "" && textEdit4.Text != "")
            {
                this.Close();
            }
            else
            {
                MessageBox.Show("Заполнены не все показатели.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void Dupon_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult res;
            if (textEdit1.Text != "" && textEdit2.Text != "" && textEdit3.Text != "" && textEdit4.Text != "")
            {
                e.Cancel = false;
            }
            else
            {
               res = MessageBox.Show("Заполнены не все показатели. Продолжить закрытие?\nРасчетные данные не будут записаны в модель. ", "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
               if (res == System.Windows.Forms.DialogResult.Yes)
               {
                   e.Cancel = false;
               }
               else
               {
                   e.Cancel = true;
               }
            }

        }

        private void textEdit1_EnabledChanged(object sender, EventArgs e)
        {
            ((TextBox)sender).ForeColor = Color.Black;
        }

        private void textEdit2_EnabledChanged(object sender, EventArgs e)
        {
            ((TextBox)sender).ForeColor = Color.Black;
        }

        private void textEdit3_EnabledChanged(object sender, EventArgs e)
        {
            ((TextBox)sender).ForeColor = Color.Black;
        }

        private void textEdit4_EnabledChanged(object sender, EventArgs e)
        {
            ((TextBox)sender).ForeColor = Color.Black;
        }
    }
}
