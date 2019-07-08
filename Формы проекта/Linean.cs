using System;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.XtraCharts;


namespace Simplex_2.Формы_проекта
{

    public partial class Linean : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public Int32 i = 0;
        public int nul = 0;

        public Linean()
        {
            InitializeComponent();

        }
        private void chartControl1_Zoom(object sender, ChartZoomEventArgs e)
        {
            XYDiagram diagram = (XYDiagram)chartControl1.Diagram;
            diagram.SecondaryAxesY[0].VisualRange.SetMinMaxValues(Convert.ToDouble(e.NewYRange.MinValue) / 2, Convert.ToDouble(e.NewYRange.MaxValue));
        }
        private void dataGridView1_MouseHover(object sender, EventArgs e)
        {
            dataGridView1.Focus();
        }
        private void Grid()
        {
            #region Формирование всех свойств и методов для таблицы
            dataGridView1.Columns[0].HeaderCell.Style.BackColor = Color.LightBlue;
            dataGridView1.Columns[1].HeaderCell.Style.BackColor = Color.LightBlue;
            dataGridView1.Columns[2].HeaderCell.Style.BackColor = Color.LightBlue;
            dataGridView1.Columns[3].HeaderCell.Style.BackColor = Color.LightBlue;

            dataGridView1.Columns[0].HeaderText = "ID";
            dataGridView1.Columns[1].HeaderText = "Дата расчета";
            dataGridView1.Columns[2].HeaderText = "Значение";
            dataGridView1.Columns[3].HeaderText = "Номер";

            dataGridView1.Columns[0].Width = 50;
            dataGridView1.Columns[1].Width = 100;

            dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;

            dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            #endregion
        }
        private void LineanTrade()
        {
            #region Линейный тип
            Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.WorksheetFunction wsf = xl.WorksheetFunction;
            Series series3 = new Series("Линейный тренд", ViewType.Line);
            chartControl1.Series.Add(series3);
            series3.ArgumentScaleType = ScaleType.Numerical;
            series3.ValueScaleType = ScaleType.Numerical;
            
            int ch = 0;
            int ch1 = 0;
            int ch2 = 0;
            int ch3 = 1;
            int count = dataGridView1.RowCount;
            double[] x = new double[count];
            double x1 = 0;
            double[] y = new double[count];
            double[] yY = new double[count];
            for (int k = 0; k < count; k++)
            {
                x[k] = Convert.ToDouble(dataGridView1.Rows[k].Cells[3].Value);
                y[k] = Convert.ToDouble(dataGridView1.Rows[k].Cells[2].Value);
                ch++;
            }
            for (int k = 0; k < count; k++)
            {
                object[,] lin1 = wsf.LinEst(y, x, 1, 1);
                var b1 = lin1[1, 1];
                var a1 = lin1[1, 2];
                x1 = Convert.ToDouble(dataGridView1.Rows[k].Cells[3].Value);
                yY[k] = Math.Round(Convert.ToDouble(a1) + Convert.ToDouble(x1 * Convert.ToDouble(b1)), 2);
                ch3++;
            }
            for (int k = 0; k < yY.Length; k++)
            {
                ch2++;
                series3.Points.Add(new SeriesPoint(ch2, yY[k]));
            }
            TrendLine trendline1 = new TrendLine("Линейный тренд");
            LineSeriesView myView = ((LineSeriesView)series3.View);
            myView.AxisY.WholeRange.AlwaysShowZeroLevel = false;
            trendline1.ExtrapolateToInfinity = false;
            trendline1.Color = Color.Red;
            int t = trendline1.Weight;
            t = 3;           
            trendline1.ShowInLegend = false;
            trendline1.Visible = false;
            trendline1.LineStyle.DashStyle = DashStyle.Dash;
            myView.Indicators.Add(trendline1);

            try
            {
                object[,] lin1 = wsf.LinEst(y, x, 1, 1);
                var b1 = lin1[1, 1];
                var a1 = lin1[1, 2];
                double yY1 = Math.Round(Convert.ToDouble(a1) + Convert.ToDouble((ch + 1) * Convert.ToDouble(b1)), 2);
                textBox2.Text = yY1.ToString();
                chartControl1.Series[1].Points.Add(new SeriesPoint((ch + 1), yY1));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ФАНЗ");
                return;
            }

            for (int jk = 0; jk < x.Length; jk++)
            {
                ch1++;
                chartControl1.Series[0].Points.Add(new SeriesPoint(ch1, Convert.ToDouble(dataGridView1[2, jk].Value)));
            }
            #endregion
        }
        private void LogTrade()
        {
            #region Логарифмический тип
            Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.WorksheetFunction wsf = xl.WorksheetFunction;

            Series series3 = new Series("Логарифмический тренд", ViewType.Spline);
            chartControl1.Series.Add(series3);
            series3.ArgumentScaleType = ScaleType.Numerical;
            series3.ValueScaleType = ScaleType.Numerical;
            int ch = 0;
            int ch1 = 0;
            int ch2 = 0;
            int ch3 = 1;
            int count = dataGridView1.RowCount;
            double[] x = new double[count];
            double x1 = 0;
            double[] y = new double[count];
            double[] yY = new double[count];

            for (int k = 0; k < count; k++)
            {
                x[k] = Math.Log(Convert.ToDouble(dataGridView1.Rows[k].Cells[3].Value));
                y[k] = Convert.ToDouble(dataGridView1.Rows[k].Cells[2].Value);
                ch++;
            }

            for (int k = 0; k < count; k++)
            {
                object[,] lin1 = wsf.LinEst(y, x, 1, 1);
                var b1 = lin1[1, 1];
                var a1 = lin1[1, 2];
                x1 = Convert.ToDouble(dataGridView1.Rows[k].Cells[3].Value);
                yY[k] = Math.Round(Convert.ToDouble(a1) + Convert.ToDouble(Math.Log(x1) * Convert.ToDouble(b1)), 2);
                ch3++;
            }
            for (int k = 0; k < yY.Length; k++)
            {
                ch2++;
                series3.Points.Add(new SeriesPoint(ch2, yY[k]));
            }
            TrendLine trendline1 = new TrendLine("Логарифмический тренд");
            SplineSeriesView myView = ((SplineSeriesView)series3.View);
            trendline1.ExtrapolateToInfinity = false;
            trendline1.Color = Color.Red;
            trendline1.ShowInLegend = false;
            trendline1.Visible = false;
            trendline1.LineStyle.DashStyle = DashStyle.Dash;
            myView.Indicators.Add(trendline1);
            try
            {
                object[,] lin = wsf.LinEst(y, x, 1, 1);
                var b = lin[1, 1];
                var a = lin[1, 2];
                double yY1 = Math.Round(Convert.ToDouble(a) + Convert.ToDouble(Math.Log(ch + 1) * Convert.ToDouble(b)), 2);
                textBox2.Text = yY1.ToString();
                chartControl1.Series[1].Points.Add(new SeriesPoint((ch + 1), yY1));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ФАНЗ");
                return;
            }

            for (int jk = 0; jk < x.Length; jk++)
            {
                ch1++;
                chartControl1.Series[0].Points.Add(new SeriesPoint(ch1, Convert.ToDouble(dataGridView1[2, jk].Value)));
            }
            #endregion
        }
        private void PowTrade()
        {
            #region Степенная регрессия
            Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.WorksheetFunction wsf = xl.WorksheetFunction;
            Microsoft.Office.Interop.Excel.WorksheetFunction index = xl.WorksheetFunction;
            Microsoft.Office.Interop.Excel.WorksheetFunction exp = xl.WorksheetFunction;

            Series series3 = new Series("Степенной тренд", ViewType.Spline);
            chartControl1.Series.Add(series3);
            series3.ArgumentScaleType = ScaleType.Numerical;
            series3.ValueScaleType = ScaleType.Numerical;
            int ch = 0;
            int ch1 = 0;
            int ch2 = 0;
            int ch3 = 1;
            int count = dataGridView1.RowCount;
            double[] x = new double[count];
            double x1 = 0;
            double[] y = new double[count];
            double[] yY = new double[count];

            for (int k = 0; k < count; k++)
            {
                x[k] = Math.Log(Convert.ToDouble(dataGridView1.Rows[k].Cells[3].Value));
                y[k] = Math.Log(Convert.ToDouble(dataGridView1.Rows[k].Cells[2].Value));
                ch++;
            }

            for (int k = 0; k < count; k++)
            {
                object[,] lin1 = wsf.LinEst(y, x, 1, 1);           
                double c = Math.Exp(wsf.Index((lin1), 1, 2));     
                var b = lin1[1, 1];
                x1 = Convert.ToDouble(dataGridView1.Rows[k].Cells[3].Value);
                yY[k] = Math.Round(c * Math.Pow(x1, Convert.ToDouble(b)), 2);
                ch3++;
            }
            for (int k = 0; k < yY.Length; k++)
            {
                ch2++;
                series3.Points.Add(new SeriesPoint(ch2, yY[k]));
            }
            TrendLine trendline1 = new TrendLine("Степенной тренд");
            SplineSeriesView myView = ((SplineSeriesView)series3.View);
            trendline1.ExtrapolateToInfinity = false;
            trendline1.Color = Color.Red;
            trendline1.ShowInLegend = false;
            trendline1.Visible = false;
            trendline1.LineStyle.DashStyle = DashStyle.Dash;
            myView.Indicators.Add(trendline1);
            try
            {
                object[,] lin1 = wsf.LinEst(y, x, 1, 1);           
                double c = Math.Exp(wsf.Index((lin1), 1, 2));  
                var b = lin1[1, 1];
                double yY1 = Math.Round(c * Math.Pow(ch + 1, Convert.ToDouble(b)), 2);
                textBox2.Text = yY1.ToString();
                chartControl1.Series[1].Points.Add(new SeriesPoint((ch + 1), yY1));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ФАНЗ");
                return;
            }

            for (int jk = 0; jk < x.Length; jk++)
            {
                ch1++;
                chartControl1.Series[0].Points.Add(new SeriesPoint(ch1, Convert.ToDouble(dataGridView1[2, jk].Value)));
            }
            #endregion
        }
        private void ExTrade()
        {
            #region Экспоненциальная регрессия
            Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.WorksheetFunction wsf = xl.WorksheetFunction;

            Series series3 = new Series("Экспоненциальный тренд", ViewType.Spline);
            chartControl1.Series.Add(series3);
            series3.ArgumentScaleType = ScaleType.Numerical;
            series3.ValueScaleType = ScaleType.Numerical;
            int ch = 0;
            int ch1 = 0;
            int ch2 = 0;
            int ch3 = 1;
            double x1 = 0;
            int count = dataGridView1.RowCount;
            double[] x = new double[count];
            double[] y = new double[count];
            double[] yY = new double[count];

            for (int k = 0; k < count; k++)
            {
                x[k] = Convert.ToDouble(dataGridView1.Rows[k].Cells[3].Value);
                y[k] = Convert.ToDouble(dataGridView1.Rows[k].Cells[2].Value);
                ch++;
            }

            for (int k = 0; k < count; k++)
            {
                object[,] lin1 = wsf.LogEst(y, x, 1, 1);
                double b = Convert.ToDouble(lin1[1, 2]);
                double m = Convert.ToDouble(lin1[1, 1]);
                double bb = Math.Round(b, 3);                                 
                double mm = Convert.ToDouble(Math.Round(Math.Log(m), 4));     
                x1 = Convert.ToDouble(dataGridView1.Rows[k].Cells[3].Value);
                yY[k] = Math.Round(bb * Math.Exp(x1 * mm), 2);
                ch3++;
            }
            for (int k = 0; k < yY.Length; k++)
            {
                ch2++;
                series3.Points.Add(new SeriesPoint(ch2, yY[k]));
            }
            TrendLine trendline1 = new TrendLine("Экспоненциальный тренд");
            SplineSeriesView myView = ((SplineSeriesView)series3.View);
            trendline1.ExtrapolateToInfinity = false;
            trendline1.ShowInLegend = false;
            trendline1.Visible = false;
            trendline1.Color = Color.Red;
            trendline1.LineStyle.DashStyle = DashStyle.Dash;
            myView.Indicators.Add(trendline1);
            try
            {
                object[,] lin1 = wsf.LogEst(y, x, 1, 1);
                double b = Convert.ToDouble(lin1[1, 2]);
                double m = Convert.ToDouble(lin1[1, 1]);
                double bb = Math.Round(b, 3);                                  
                double mm = Convert.ToDouble(Math.Round(Math.Log(m), 4));      

                double yY1 = Math.Round(bb * Math.Exp(mm * (ch + 1)), 2);
                textBox2.Text = yY1.ToString();
                chartControl1.Series[1].Points.Add(new SeriesPoint((ch + 1), yY1));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ФАНЗ");
                return;
            }

            for (int jk = 0; jk < x.Length; jk++)
            {
                ch1++;
                chartControl1.Series[0].Points.Add(new SeriesPoint(ch1, Convert.ToDouble(dataGridView1[2, jk].Value)));
            }
            #endregion
        }
        private void Linean_Load(object sender, EventArgs e)
        {
            try
            {
                this.lineanTableAdapter.Fill(this.fANZDataSet.Linean);
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Columns[3].Visible = false;
                double count = dataGridView1.RowCount;
                int ch = 0;
                for (int u = 0; u < count; u++)
                {
                    ch++;
                    dataGridView1[3, u].Value = ch;
                }
                Grid();
             }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "ФАНЗ");
            }      
        }
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            int kol = chartControl1.Series.Count;
            for (int j = 0; j < kol; j++)
            {
                chartControl1.Series[j].Points.Clear();                
                if (j > 1)
                {
                    chartControl1.Series[j].ShowInLegend = false;
                }
            }
            textBox2.Clear();

            Int32 sthet = i;

            if (dataGridView1.RowCount != 0)
            {
                switch (sthet)
                {
                    case 1:
                        LineanTrade();
                        break;
                    case 2:
                        LogTrade();
                        break;
                    case 3:
                        PowTrade();
                        break;
                    case 4:
                        ExTrade();
                        break;
                }
            }
            else
            {
                MessageBox.Show("Не заполнена таблица координат.", "Cистема", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        private void simpleButton2_Click(object sender, EventArgs e)
        {
            int kol = chartControl1.Series.Count;
            for (int j = 0; j < kol; j++)
            {
                chartControl1.Series[j].Points.Clear();
                if (j > 1)
                {
                    chartControl1.Series[j].ShowInLegend = false;
                }
            }
           
            textBox2.Clear();
        }

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            i = 1;
        }

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            i = 2;
        }

        private void chartControl1_Zoom_1(object sender, ChartZoomEventArgs e)
        {
            XYDiagram diagram = (XYDiagram)chartControl1.Diagram;
            diagram.SecondaryAxesY[0].VisualRange.SetMinMaxValues(Convert.ToDouble(e.NewYRange.MinValue) / 2, Convert.ToDouble(e.NewYRange.MaxValue));
        }

        private void barButtonItem3_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            i = 3;
        }

        private void barButtonItem4_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            i = 4;
        }

        private void barButtonItem5_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            HelpToLinean help = new HelpToLinean();
            help.Show();
        }
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DeleteID o = new DeleteID();
            o.ShowDialog();
            if (o.result == true)
            {
                try
                {
                    lineanTableAdapter.DeleteQuery((o.comboBoxEdit1.Text));
                    this.Validate();
                    this.lineanBindingSource.EndEdit();
                    this.fANZDataSet.AcceptChanges();
                    lineanTableAdapter.Update(fANZDataSet);
                    lineanTableAdapter.Fill(fANZDataSet.Linean);
                    dataGridView1.Refresh();
                }
                catch (Exception f1) 
                {
                    MessageBox.Show(f1.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
                }
            }
            double count = dataGridView1.RowCount;
            int ch = 0;
            for (int u = 0; u < count; u++)
            {
                ch++;
                dataGridView1[3, u].Value = ch;
            }
        }
    }
}
