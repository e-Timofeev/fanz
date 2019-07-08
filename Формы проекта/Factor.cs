using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraBars;
using System.Data.SqlServerCe;

namespace Simplex_2.Формы_проекта
{
    public partial class Factor : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public int KOL = 0;
        public string s = System.Windows.Forms.Application.StartupPath;
        public Factor()
        {
            InitializeComponent();
        }
        private void Combo()
        {
            try
            {
                repositoryItemComboBox3.Items.Clear();
                barEditItem3.EditValue = string.Empty;
                SqlCeConnection conn;
                conn = new SqlCeConnection(@"Data Source = |DataDirectory|\FANZ.sdf");
                conn.Open();
                SqlCeDataAdapter adapter;
                adapter = new SqlCeDataAdapter("SELECT DISTINCT Sysdate FROM Altman", conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                int row = dt.Rows.Count;

                for (int j = 0; j < row; j++)
                {
                    repositoryItemComboBox3.Items.AddRange(dt.Rows[j].ItemArray.ToList());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void Factor_Load(object sender, EventArgs e)
        {
            this.Width = Screen.PrimaryScreen.Bounds.Width;
            this.Height = 600;
            this.Location = new Point((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2,
                (Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
            Combo();
        }
        #region Перерисовка бордюра у меток
        private void label3_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label3.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label2_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label2.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label1_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label1.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label4_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label4.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label5_Paint_1(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label5.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label6_Paint_1(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label6.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label7_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label7.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label8_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label8.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label9_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label9.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label10_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label10.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label11_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label11.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label12_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label12.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label13_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label13.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label14_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label14.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label15_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label15.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label16_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label16.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label17_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label17.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label18_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label18.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label19_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label19.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label20_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label20.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label21_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label21.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label22_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label22.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label23_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label23.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label24_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label24.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label25_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label25.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }


        private void label26_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label26.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label28_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label28.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label29_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label29.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label31_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label31.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label32_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label32.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label34_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label34.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label35_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label35.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label37_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label37.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label38_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label38.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label27_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label27.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label33_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label33.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label36_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label36.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label39_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label39.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label40_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label40.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label41_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label41.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label42_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label42.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label30_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label30.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label44_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label44.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label45_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label45.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label43_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label43.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label48_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label48.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label49_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label49.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label52_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label52.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label53_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label53.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label51_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label51.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label54_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label54.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label47_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label47.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label50_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label50.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label56_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label56.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label55_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label55.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label58_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label58.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label57_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label57.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label59_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label59.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label60_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label60.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label61_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label61.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label62_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label62.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }

        private void label46_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, label46.DisplayRectangle, Color.Black, ButtonBorderStyle.Solid);
        }
        #endregion
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
        private void barButtonItem2_ItemClick(object sender, ItemClickEventArgs e)
        {
            string date1 = DateTime.Today.ToShortDateString();
            string date2 = "Cоздан в " + Convert.ToString(DateTime.Now.Hour + "." + DateTime.Now.Minute);
            string filename = date2 + ".png";
            DialogResult results;
            //------------------------------------------------------------------------------------------------------------------------              
            try
            {
                if (!Directory.Exists(s + @"\Отчеты\Факторный анализ\" + "" + date1))
                {
                    Directory.CreateDirectory(s + @"\Отчеты\Факторный анализ\" + "" + date1);
                    string m = s + "\\Отчеты\\Факторный анализ\\" + date1;
                    using (Bitmap printImage = new Bitmap(tableLayoutPanel1.Width, tableLayoutPanel1.Height))
                    {
                        tableLayoutPanel1.DrawToBitmap(printImage, new Rectangle(0, 0, printImage.Width, printImage.Height));
                        printImage.Save(m + @"\" + filename, System.Drawing.Imaging.ImageFormat.Png);
                    }
                    MessageBox.Show("Данные успешно сохранены.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    string m = s + "\\Отчеты\\Факторный анализ\\" + date1;
                    using (Bitmap printImage = new Bitmap(tableLayoutPanel1.Width, tableLayoutPanel1.Height))
                    {
                        tableLayoutPanel1.DrawToBitmap(printImage, new Rectangle(0, 0, printImage.Width, printImage.Height));
                        printImage.Save(m + @"\" + filename, System.Drawing.Imaging.ImageFormat.Png);
                    }
                    MessageBox.Show("Данные успешно сохранены.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //------------------------------------------------------------------------------------------------------------------------            
            results = MessageBox.Show("Открыть корневую папку с сохраненными результатами?", "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (results == System.Windows.Forms.DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();

                openFileDialog1.InitialDirectory = s + @"\Отчеты\Факторный анализ\" + date1;
                openFileDialog1.Filter = "Png files (*.png)|*";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        System.IO.StreamReader sr = new
                        System.IO.StreamReader(openFileDialog1.FileName);
                        System.Diagnostics.Process myProc = new System.Diagnostics.Process();
                        myProc = System.Diagnostics.Process.Start(openFileDialog1.FileName);
                        sr.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
        }
        #region Очищалки
        private void Cleanlabel()
        {
            //if (barEditItem4.EditValue.ToString() == Boolean.TrueString)
            //{
            #region clear
            label8.Text = string.Empty;
            label9.Text = string.Empty;
            label28.Text = string.Empty;
            label29.Text = string.Empty;
            label44.Text = string.Empty;
            label45.Text = string.Empty;
            label59.Text = string.Empty;
            label60.Text = string.Empty;

            label14.Text = string.Empty;
            label17.Text = string.Empty;
            label20.Text = string.Empty;
            label23.Text = string.Empty;
            label15.Text = string.Empty;
            label18.Text = string.Empty;
            label21.Text = string.Empty;
            label24.Text = string.Empty;

            label16.Text = string.Empty;
            label19.Text = string.Empty;
            label22.Text = string.Empty;
            label25.Text = string.Empty;
            label16.BackColor = Color.WhiteSmoke;
            label19.BackColor = Color.WhiteSmoke;
            label22.BackColor = Color.WhiteSmoke;
            label25.BackColor = Color.WhiteSmoke;

            label31.Text = string.Empty;
            label34.Text = string.Empty;
            label37.Text = string.Empty;
            label32.Text = string.Empty;
            label35.Text = string.Empty;
            label38.Text = string.Empty;
            label33.Text = string.Empty;
            label36.Text = string.Empty;
            label39.Text = string.Empty;

            label33.BackColor = Color.WhiteSmoke;
            label36.BackColor = Color.WhiteSmoke;
            label39.BackColor = Color.WhiteSmoke;

            label48.Text = string.Empty;
            label49.Text = string.Empty;
            label52.Text = string.Empty;
            label53.Text = string.Empty;
            label51.Text = string.Empty;
            label54.Text = string.Empty;

            label51.BackColor = Color.WhiteSmoke;
            label54.BackColor = Color.WhiteSmoke;

            label50.Text = string.Empty;
            label55.Text = string.Empty;
            label56.Text = string.Empty;

            label55.BackColor = Color.WhiteSmoke;
            #endregion
            //}
        }
        #endregion
        private void Period1()
        {
            if (barEditItem1.EditValue.ToString() != "")
            {
                if (label14.Text != "" && label17.Text != "" && label20.Text != "" && label23.Text != "" && label15.Text != "" && label18.Text != "" && label21.Text != "" && label24.Text != "")
                {
                    #region Расчеты процентов, проверка и закрашивание
                    double c14 = Convert.ToDouble(label14.Text);
                    double c17 = Convert.ToDouble(label17.Text);
                    double c20 = Convert.ToDouble(label20.Text);
                    double c23 = Convert.ToDouble(label23.Text);

                    double d15 = Convert.ToDouble(label15.Text);
                    double d18 = Convert.ToDouble(label18.Text);
                    double d21 = Convert.ToDouble(label21.Text);
                    double d24 = Convert.ToDouble(label24.Text);

                    if (c14 != 0)
                    {
                        double c16 = ((d15 - c14) / c14) * 100;
                        label16.Text = c16.ToString();
                        if (c16 < 0)
                        {
                            label16.BackColor = Color.Salmon;
                        }
                        else
                        {
                            label16.BackColor = Color.LightGreen;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Значение показателя «Чистая прибыль» равно 0.\nПродолжение расчета отменено.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    if (c17 != 0)
                    {
                        double c19 = ((d18 - c17) / c17) * 100;
                        label19.Text = c19.ToString();
                        if (c19 < 0)
                        {
                            label19.BackColor = Color.Salmon;
                        }
                        else
                        {
                            label19.BackColor = Color.LightGreen;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Значение показателя «Выручка» равно 0.\nПродолжение расчета отменено.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    if (c20 != 0)
                    {
                        double c22 = ((d21 - c20) / c20) * 100;
                        label22.Text = c22.ToString();
                        if (c22 < 0)
                        {
                            label22.BackColor = Color.Salmon;
                        }
                        else
                        {
                            label22.BackColor = Color.LightGreen;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Значение показателя «Активы» равно 0.\nПродолжение расчета отменено.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    if (c23 != 0)
                    {
                        double c25 = ((d24 - c23) / c23) * 100;
                        label25.Text = c25.ToString();
                        if (c25 < 0)
                        {
                            label25.BackColor = Color.Salmon;
                        }
                        else
                        {
                            label25.BackColor = Color.LightGreen;
                        }
                    }
                    #endregion
                    else
                    {
                        MessageBox.Show("Значение показателя «Собственный капитал» равно 0.\nПродолжение расчета отменено.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }
                else
                {
                    if ((barEditItem4.EditValue.ToString() == Boolean.TrueString))
                    {
                        Cleanlabel();
                    }
                    MessageBox.Show("Среди показателей есть пустые значения.\nПродолжение расчета отменено.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }
        }
        private void Period2()
        {
            #region Расчеты процентов, проверка и закрашивание
            double c14 = Convert.ToDouble(label14.Text);
            double c17 = Convert.ToDouble(label17.Text);
            double c20 = Convert.ToDouble(label20.Text);
            double c23 = Convert.ToDouble(label23.Text);

            double d15 = Convert.ToDouble(label15.Text);
            double d18 = Convert.ToDouble(label18.Text);
            double d21 = Convert.ToDouble(label21.Text);
            double d24 = Convert.ToDouble(label24.Text);

            //-----------------------------------------                    
            if (c17 != 0)
            {
                double c31 = (c14 / c17) * 100;
                label31.Text = c31.ToString();
            }
            else
            {
                MessageBox.Show("Значение показателя «Выручка» равно 0.\nПродолжение расчета отменено.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (c20 != 0)
            {
                double c34 = c17 / c20;
                label34.Text = c34.ToString();
            }
            else
            {
                MessageBox.Show("Значение показателя «Активы» равно 0.\nПродолжение расчета отменено.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (c23 != 0)
            {
                double c37 = c20 / c23;
                label37.Text = c37.ToString();
            }
            else
            {
                MessageBox.Show("Значение показателя «Собственный капитал» равно 0.\nПродолжение расчета отменено.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (d18 != 0)
            {
                double c32 = (d15 / d18) * 100;
                label32.Text = c32.ToString();
            }
            else
            {
                MessageBox.Show("Значение показателя «Выручка» равно 0.\nПродолжение расчета отменено.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (d21 != 0)
            {
                double c35 = d18 / d21;
                label35.Text = c35.ToString();
            }
            else
            {
                MessageBox.Show("Значение показателя «Активы» равно 0.\nПродолжение расчета отменено.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (d24 != 0)
            {
                double c38 = d21 / d24;
                label38.Text = c38.ToString();
            }
            else
            {
                MessageBox.Show("Значение показателя «Собственный капитал» равно 0.\nПродолжение расчета отменено.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            double c33 = Math.Round(Convert.ToDouble(label32.Text) - Convert.ToDouble(label31.Text), 2);
            double c36 = Math.Round(Convert.ToDouble(label35.Text) - Convert.ToDouble(label34.Text), 2);
            double c39 = Math.Round(Convert.ToDouble(label38.Text) - Convert.ToDouble(label37.Text), 2);
            label33.Text = c33.ToString();
            label36.Text = c36.ToString();
            label39.Text = c39.ToString();
            if (c33 < 0)
            {
                label33.BackColor = Color.Salmon;
            }
            else
            {
                label33.BackColor = Color.LightGreen;
            }
            if (c36 < 0)
            {
                label36.BackColor = Color.Salmon;
            }
            else
            {
                label36.BackColor = Color.LightGreen;
            }
            if (c39 < 0)
            {
                label39.BackColor = Color.Salmon;
            }
            else
            {
                label39.BackColor = Color.LightGreen;
            }
            #endregion
        }
        private void Period3()
        {
            #region Расчеты процентов, проверка и закрашивание
            double c31 = Convert.ToDouble(label31.Text);
            double c34 = Convert.ToDouble(label34.Text);
            double c37 = Convert.ToDouble(label37.Text);

            double c32 = Convert.ToDouble(label32.Text);
            double c35 = Convert.ToDouble(label35.Text);
            double c38 = Convert.ToDouble(label38.Text);
            //-----------------------------------------                    
            double c48 = c31 * c34;
            double c52 = c34 * c37;

            double c49 = c32 * c35;
            double c53 = c35 * c38;

            double c51 = c49 - c48;
            double c54 = c53 - c52;

            label48.Text = c48.ToString();
            label49.Text = c49.ToString();
            label51.Text = c51.ToString();
            label52.Text = c52.ToString();
            label53.Text = c53.ToString();
            label54.Text = c54.ToString();

            if (c51 < 0)
            {
                label51.BackColor = Color.Salmon;
            }
            else
            {
                label51.BackColor = Color.LightGreen;
            }
            if (c54 < 0)
            {
                label54.BackColor = Color.Salmon;
            }
            else
            {
                label54.BackColor = Color.LightGreen;
            }

            #endregion
        }
        private void Period4()
        {
            #region Расчеты процентов, проверка и закрашивание
            double c31 = Convert.ToDouble(label31.Text);
            double c34 = Convert.ToDouble(label34.Text);
            double c37 = Convert.ToDouble(label37.Text);

            double c32 = Convert.ToDouble(label32.Text);
            double c35 = Convert.ToDouble(label35.Text);
            double c38 = Convert.ToDouble(label38.Text);
            //-----------------------------------------                    
            double c50 = Math.Round((c31 * c34 * c37), 1);
            double c56 = Math.Round((c32 * c35 * c38), 1);
            double c55 = Math.Round((c56 - c50), 2);

            label50.Text = c50.ToString();
            label55.Text = c55.ToString();
            label56.Text = c56.ToString();

            if (c55 < 0)
            {
                label55.BackColor = Color.Salmon;
            }
            else
            {
                label55.BackColor = Color.LightGreen;
            }
            #endregion
        }
        private void Round()
        {
            try
            {
                label16.Text = Math.Round(Convert.ToDouble(label16.Text), 0).ToString();
                label19.Text = Math.Round(Convert.ToDouble(label19.Text), 0).ToString();
                label22.Text = Math.Round(Convert.ToDouble(label22.Text), 0).ToString();
                label25.Text = Math.Round(Convert.ToDouble(label25.Text), 0).ToString();

                label31.Text = Math.Round(Convert.ToDouble(label31.Text), 1).ToString();
                label34.Text = Math.Round(Convert.ToDouble(label34.Text), 2).ToString();
                label37.Text = Math.Round(Convert.ToDouble(label37.Text), 2).ToString();
                label32.Text = Math.Round(Convert.ToDouble(label32.Text), 1).ToString();
                label35.Text = Math.Round(Convert.ToDouble(label35.Text), 2).ToString();
                label38.Text = Math.Round(Convert.ToDouble(label38.Text), 2).ToString();
                label33.Text = Math.Round(Convert.ToDouble(label33.Text), 1).ToString();
                label36.Text = Math.Round(Convert.ToDouble(label36.Text), 2).ToString();
                label39.Text = Math.Round(Convert.ToDouble(label39.Text), 2).ToString();

                label48.Text = Math.Round(Convert.ToDouble(label48.Text), 1).ToString();
                label49.Text = Math.Round(Convert.ToDouble(label49.Text), 1).ToString();
                label51.Text = Math.Round(Convert.ToDouble(label51.Text), 1).ToString();
                label52.Text = Math.Round(Convert.ToDouble(label52.Text), 2).ToString();
                label53.Text = Math.Round(Convert.ToDouble(label53.Text), 2).ToString();
                label54.Text = Math.Round(Convert.ToDouble(label54.Text), 2).ToString();
            }
            catch
            {
                MessageBox.Show("Заданы не все данные для расчета. Повторите расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        private void barEditItem1_EditValueChanged(object sender, EventArgs e)
        {
            string year1 = string.Empty;
            string filter1 = string.Empty;
            int count = 0;

            if (barEditItem1.EditValue.ToString() != string.Empty)
            {
                #region Получаем год
                try
                {
                    year1 = PeriodYear(barEditItem1.EditValue.ToString());
                    #region Проверка наличия данных по дате
                    SqlCeConnection conn;
                    conn = new SqlCeConnection(@"Data Source = |DataDirectory|\FANZ.sdf");
                    conn.Open();
                    SqlCeDataAdapter adapter;
                    adapter = new SqlCeDataAdapter("SELECT * FROM Altman", conn);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    DataView MyDataView = new DataView(dt);
                    conn.Close();
                    filter1 = year1;
                    MyDataView.RowFilter = filter1;
                    count = MyDataView.Count;
                    #endregion
                    #region Обработка и заполнение
                    if (count == 0)
                    {
                        if ((barEditItem4.EditValue.ToString() == Boolean.TrueString))
                        {
                            Cleanlabel();
                        }
                        MessageBox.Show("По заданным параметрам данные отсутствуют.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    if (count == 4)
                    {
                        label8.Text = barEditItem1.EditValue.ToString();               //Период с
                        label28.Text = barEditItem1.EditValue.ToString();
                        label44.Text = barEditItem1.EditValue.ToString();
                        label59.Text = barEditItem1.EditValue.ToString();

                        label14.Text = AltmanTableAdapter.Select4().ToString();        //Чистая прибыль
                        label17.Text = AltmanTableAdapter.Select3().ToString();        //Выручка
                        label20.Text = AltmanTableAdapter.Select1().ToString();        //ИТОГО АКТИВ
                        label23.Text = AltmanTableAdapter.Select2().ToString();        //Собственный капитал
                    }
                    if (count > 4)
                    {
                        Dupon dup = new Dupon();
                        dup.Owner = this;
                        dup.ShowDialog();

                        if (dup.textEdit1.Text != "" && dup.textEdit2.Text != "" && dup.textEdit3.Text != "" && dup.textEdit4.Text != "")
                        {
                            label8.Text = barEditItem1.EditValue.ToString();               //Период с
                            label28.Text = barEditItem1.EditValue.ToString();
                            label44.Text = barEditItem1.EditValue.ToString();
                            label59.Text = barEditItem1.EditValue.ToString();

                            this.label14.Text = dup.textEdit1.Text;       //чистая прибыль
                            this.label17.Text = dup.textEdit2.Text;       //выручка
                            this.label20.Text = dup.textEdit3.Text;       //активы
                            this.label23.Text = dup.textEdit4.Text;       //капитал
                        }
                        else
                        {
                            label8.Text = "";                             //Период с
                            label28.Text = "";
                            label44.Text = "";
                            label59.Text = "";
                        }
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    if ((barEditItem4.EditValue.ToString() == Boolean.TrueString))
                    {
                        Cleanlabel();
                    }
                    MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                #endregion
            }
        }
        private void barEditItem3_EditValueChanged(object sender, EventArgs e)
        {
            contextMenuStrip1.Show(MousePosition.X - barEditItem3.Width / 2, MousePosition.Y - barEditItem3.Height);
        }
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                string p = barEditItem3.EditValue.ToString();
                this.AltmanTableAdapter.Delete(p);
                if ((barEditItem4.EditValue.ToString() == Boolean.TrueString))
                {
                    Cleanlabel();
                }
                Combo();
                MessageBox.Show("Все записи по этой дате удалены.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception t)
            {
                if ((barEditItem4.EditValue.ToString() == Boolean.TrueString))
                {
                    Cleanlabel();
                }
                MessageBox.Show(t.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void barEditItem2_EditValueChanged(object sender, EventArgs e)
        {
            if (barEditItem2.EditValue.ToString() != string.Empty)
            {
                string year1 = string.Empty;
                string filter1 = string.Empty;
                int count = 0;

                #region Получаем год
                try
                {
                    year1 = PeriodYear(barEditItem2.EditValue.ToString());
                    #region Проверка наличия данных по дате
                    SqlCeConnection conn;
                    conn = new SqlCeConnection(@"Data Source = |DataDirectory|\FANZ.sdf");
                    conn.Open();
                    SqlCeDataAdapter adapter;
                    adapter = new SqlCeDataAdapter("SELECT * FROM Altman", conn);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    DataView MyDataView = new DataView(dt);
                    conn.Close();
                    filter1 = year1;
                    MyDataView.RowFilter = filter1;
                    count = MyDataView.Count;
                    #endregion
                    #region Обработка
                    if (count == 0)
                    {
                        if ((barEditItem4.EditValue.ToString() == Boolean.TrueString))
                        {
                            Cleanlabel();
                        }
                        MessageBox.Show("По заданным параметрам данные отсутствуют.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    if (count == 4)
                    {
                        label9.Text = barEditItem2.EditValue.ToString();               //Период с
                        label29.Text = barEditItem2.EditValue.ToString();
                        label45.Text = barEditItem2.EditValue.ToString();
                        label60.Text = barEditItem2.EditValue.ToString();

                        label15.Text = AltmanTableAdapter.Select4().ToString();        //Чистая прибыль
                        label18.Text = AltmanTableAdapter.Select3().ToString();        //Выручка
                        label21.Text = AltmanTableAdapter.Select1().ToString();        //ИТОГО АКТИВ
                        label24.Text = AltmanTableAdapter.Select2().ToString();        //Собственный капитал
                    }
                    if (count > 4)
                    {
                        Dupon dup = new Dupon();
                        dup.Owner = this;
                        dup.ShowDialog();
                        if (dup.textEdit1.Text != "" && dup.textEdit2.Text != "" && dup.textEdit3.Text != "" && dup.textEdit4.Text != "")
                        {
                            label9.Text = barEditItem2.EditValue.ToString();               //Период с
                            label29.Text = barEditItem2.EditValue.ToString();
                            label45.Text = barEditItem2.EditValue.ToString();
                            label60.Text = barEditItem2.EditValue.ToString();

                            this.label15.Text = dup.textEdit1.Text;       //чистая прибыль
                            this.label18.Text = dup.textEdit2.Text;       //выручка
                            this.label21.Text = dup.textEdit3.Text;       //активы
                            this.label24.Text = dup.textEdit4.Text;       //капитал
                        }
                        else
                        {
                            label9.Text = "";               //Период с
                            label29.Text = "";
                            label45.Text = "";
                            label60.Text = "";
                        }
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    if ((barEditItem4.EditValue.ToString() == Boolean.TrueString))
                    {
                        Cleanlabel();
                    }
                    MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                #endregion
            }
        }
        private void barButtonItem1_ItemClick(object sender, ItemClickEventArgs e)
        {
            try
            {
                Period1();
                Period2();
                Period3();
                Period4();
                Round();
            }
            catch
            {
                Cleanlabel();
                return;
            }
        }

        private void barButtonItem3_ItemClick(object sender, ItemClickEventArgs e)
        {

            Cleanlabel();
            barEditItem1.EditValue = string.Empty;
            barEditItem2.EditValue = string.Empty;
        }

    }
}