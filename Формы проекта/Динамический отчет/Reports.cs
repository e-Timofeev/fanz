using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using System.Data.SqlServerCe;
using DevExpress.XtraCharts;
using DevExpress.XtraPrinting;
using System.IO;

namespace Simplex_2.Формы_проекта
{
    public partial class Reports : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public double step;
        public double step1;
        public int step2;
        public string s = System.Windows.Forms.Application.StartupPath;
        public Reports()
        {
            InitializeComponent();
        }
        public static string Find(CheckedComboBoxEdit gp)
        {
            string query = string.Empty;
            string str = string.Empty;
            foreach (CheckedListBoxItem cnt in gp.Properties.Items)
            {
                if (cnt.CheckState == CheckState.Checked)
                {
                    query += "'" + cnt.Value.ToString() + "',";
                }
            }
            str = "(" + query.Substring(0, query.Length - 1) + ")";
            return str;
        }
        public static string FindDate(DevExpress.XtraLayout.LayoutControlGroup lg)
        {
            int c = 0;
            List<string> list = new List<string>();
            string output = string.Empty;
            foreach (DevExpress.XtraLayout.LayoutControlItem cnt in lg.Items)
            {
                foreach (Control dat in cnt.Control.Controls)
                {
                    if (dat.Text != string.Empty)
                    {
                        list.Add(dat.Parent.Name);
                        c++;
                    }
                }
            }
            output = String.Join(",", list);
            return output;
        }
        private void barButtonItem4_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            panelControl1.Visible = true;
        }
        private void barButtonItem5_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            panelControl1.Hide();
        }
        private string PeriodQuarter(string year, int id)                                //квартал
        {
            string result = string.Empty;
            switch (id)
            {
                case 0:
                    result = year + "-" + "01-01" + "," + year + "-" + "03-31";
                    break;
                case 1:
                    result = year + "-" + "04-01" + "," + year + "-" + "06-30";
                    break;
                case 2:
                    result = year + "-" + "07-01" + "," + year + "-" + "09-30";
                    break;
                case 3:
                    result = year + "-" + "10-01" + "," + year + "-" + "12-31";
                    break;
            }
            return result;
        }
        private string PeriodYear(int id)                                               //год
        {
            string year = string.Empty;
            switch (id)
            {
                case 0:
                    year = String.Format("([Дата] >= '{0}') AND ([Дата] < '{1}')", ("2015-01-01"), ("2016-01-01"));
                    break;
                case 1:
                    year = String.Format("([Дата] >= '{0}') AND ([Дата] < '{1}')", ("2016-01-01"), ("2017-01-01"));
                    break;
                case 2:
                    year = String.Format("([Дата] >= '{0}') AND ([Дата] < '{1}')", ("2017-01-01"), ("2018-01-01"));
                    break;
                case 3:
                    year = String.Format("([Дата] >= '{0}') AND ([Дата] < '{1}')", ("2018-01-01"), ("2019-01-01"));
                    break;
                case 4:
                    year = String.Format("([Дата] >= '{0}') AND ([Дата] < '{1}')", ("2019-01-01"), ("2020-01-01"));
                    break;
                case 5:
                    year = String.Format("([Дата] >= '{0}') AND ([Дата] < '{1}')", ("2020-01-01"), ("2021-01-01"));
                    break;
            }
            return year;
        }
        private static string DateBegin(DateTime value1, DateTime value2)
        {
            string row = string.Empty;
            int result = DateTime.Compare(value1, value2);
            if (result >= 0)
            {
                return row;
            }
            else
            {
                row = String.Format("([Дата] >= '{0}') AND ([Дата] < '{1}')", value1, value2);
                return row;
            }
        }           //произвольная дата
        private void Report()
        {
            //#region Сбор данных, конвертация
            string year = string.Empty;
            string quarter = string.Empty;
            string begin = string.Empty;
            string end = string.Empty;
            string time = string.Empty;
            string filter = string.Empty;
            List<string> col = new List<string>();
            List<string> union = new List<string>();
            List<string> value = new List<string>();
            List<string> name = new List<string>();
            List<string> spisok_1 = new List<string>();
            List<string> spisok_2 = new List<string>();
            //делаем очистку-----------------
            union.Clear();
            value.Clear();
            col.Clear();
            name.ToArray();
            spisok_1.Clear();
            spisok_2.Clear();
            //попытка подчистить метки перед любым запуском 
            string type = chartControl1.Diagram.GetType().ToString();
            string type2d = "DevExpress.XtraCharts.XYDiagram";
            if (type == type2d)
            {
                XYDiagram diagram = (XYDiagram)chartControl1.Diagram;
                diagram.AxisX.CustomLabels.Clear();
                if (chartControl1.Series.Count > 1)
                {
                    for (int e = 1; e < chartControl1.Series.Count; e++)
                    {
                        chartControl1.Series.Clear();
                    }
                    chartControl1.Series.Add(new Series("Группировка по дате", ViewType.Bar));
                    ((SideBySideBarSeriesView)chartControl1.Series[0].View).BarDistance = 0;
                    ((SideBySideBarSeriesView)chartControl1.Series[0].View).BarDistanceFixed = 1;
                    ((SideBySideBarSeriesView)chartControl1.Series[0].View).BarWidth = 0.6;
                }
                //------------------------------
                #region Получаем год
                if (comboBoxEdit1.Text != string.Empty)
                {
                    try
                    {
                        int index1 = comboBoxEdit1.SelectedIndex;
                        year = PeriodYear(index1);
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        return;
                    }
                }
                #endregion
                #region Получаем квартал, если задан год
                if (comboBoxEdit1.Text != string.Empty && comboBoxEdit2.Text != string.Empty)
                {
                    int index2 = comboBoxEdit2.SelectedIndex;
                    string period = PeriodQuarter(comboBoxEdit1.Text, index2);
                    string zp = ",";
                    int index = period.IndexOf(zp);
                    begin = period.Substring(0, index);
                    end = period.Substring(index + 1);
                    quarter = String.Format("([Дата] >= '{0}') AND ([Дата] < '{1}')", Convert.ToDateTime(begin), Convert.ToDateTime(end));
                }
                if ((comboBoxEdit1.Text == string.Empty && comboBoxEdit2.Text != string.Empty))
                {
                    gridControl1.DataSource = null;
                    gridControl2.DataSource = null;
                    ClearDiagram();
                    chartControl1.Series[0].LegendText = "Группировка по дате";
                    MessageBox.Show("Не выбран год. Укажите год для фильтрации по кварталу.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if ((comboBoxEdit1.Text != string.Empty && checkEdit3.CheckState == CheckState.Checked) && comboBoxEdit2.Text == string.Empty)
                {
                    gridControl1.DataSource = null;
                    gridControl2.DataSource = null;
                    ClearDiagram();
                    chartControl1.Series[0].LegendText = "Группировка по дате";
                    MessageBox.Show("Для указанного года включен фильтр по кварталу с пустым значением. Заполните поле или отключите фильтр.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                #endregion
                #region Выбираем периоды
                if (dateEdit1.Text != string.Empty && dateEdit2.Text != string.Empty)
                {
                    string rows = DateBegin(dateEdit1.DateTime, dateEdit2.DateTime);
                    if (rows != string.Empty)
                    {
                        time = rows;
                    }
                    else
                    {
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        MessageBox.Show("Дата начала периода больше или совпадает с датой окончания периода. Установите корректные значения.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }
                else if (dateEdit1.Text == string.Empty || dateEdit2.Text == string.Empty)
                {
                    if (dateEdit1.Text == string.Empty && dateEdit2.Text != string.Empty)
                    {
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        MessageBox.Show("Не задана дата начала периода.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    else
                        if (dateEdit2.Text == string.Empty && dateEdit1.Text != string.Empty)
                        {
                            {
                                gridControl1.DataSource = null;
                                gridControl2.DataSource = null;
                                ClearDiagram();
                                chartControl1.Series[0].LegendText = "Группировка по дате";
                                MessageBox.Show("Не задана дата окончания периода.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                return;
                            }
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
                    adapter = new SqlCeDataAdapter("SELECT * FROM Reports", conn);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    DataView MyDataView = new DataView(dt);

                    string res1 = Find(checkedComboBoxEdit1);
                    string res2 = FindDate(layoutControlGroup2);
                    if (res2.Contains("comboBoxEdit1"))
                    {
                        filter = "AND" + year;
                    }
                    if (res2.Contains("comboBoxEdit2"))
                    {
                        filter += "AND" + quarter;
                    }
                    if (res2.Contains("dateEdit1") && res2.Contains("dateEdit2"))
                    {
                        filter += "AND" + time;
                    }
                    MyDataView.RowFilter = "[Показатель] IN " + res1 + filter;
                    int countRow = MyDataView.Count;
                    if (countRow == 0)
                    {
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        MessageBox.Show("По заданным параметрам данные отсутствуют.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    else
                    {
                        gridControl1.DataSource = MyDataView;
                        gridControl2.DataSource = MyDataView;
                        MessageBox.Show("Отчет построен.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    conn.Close();
                }
                catch (Exception q)
                {
                    gridControl1.DataSource = null;
                    gridControl2.DataSource = null;
                    ClearDiagram();
                    chartControl1.Series[0].LegendText = "Группировка по дате";
                    MessageBox.Show(q.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                #endregion
                #region Формируем данные для графика
                for (int i = 0; i < gridView2.DataRowCount; i++)
                {
                    if (gridView2.IsDataRow(i))
                    {
                        string v = gridView2.GetRowCellValue(i, "Дата").ToString();
                        col.Add(v);
                    }
                }
                union = col.Distinct().ToList();
                int count = union.Count;
                foreach (string item in union)
                {
                    for (int i = 0; i < gridView2.DataRowCount; i++)
                    {
                        if (gridView2.IsDataRow(i))
                        {
                            string v = gridView2.GetRowCellValue(i, "Дата").ToString();
                            if (item == v)
                            {
                                value.Add(gridView2.GetRowCellValue(i, "Значение").ToString());
                                name.Add(gridView2.GetRowCellValue(i, "Показатель").ToString());
                            }
                        }
                    }
                }
                gridView1.CollapseAllGroups();
                if (!count.Equals(0))
                {
                    int q = value.Count / count;
                    int sum = 0;
                    int[] ss = new int[count];
                    for (int s = 0; s < count; s++)
                    {
                        ss[s] = sum;
                        sum += q;
                    }
                    if (ss.Length > 1)
                    {
                        for (int point = 1; point < ss.Length; point++)
                        {
                            chartControl1.Series.Add(new Series(DateTime.Now.ToLongDateString(), ViewType.Bar));
                            ((SideBySideBarSeriesView)chartControl1.Series[point].View).BarDistanceFixed = 1;
                            ((SideBySideBarSeriesView)chartControl1.Series[point].View).BarDistance = 0;
                            ((SideBySideBarSeriesView)chartControl1.Series[point].View).BarWidth = 0.6;
                        }
                    }
                    for (int t = 0; t < chartControl1.Series.Count; t++)
                    {
                        diagram.AxisY.VisualRange.Auto = true;
                        diagram.AxisX.VisualRange.Auto = true;
                        chartControl1.Series[t].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                        chartControl1.Series[t].Label.TextAlignment = System.Drawing.StringAlignment.Far;
                        chartControl1.Series[t].Label.ResolveOverlappingMode = ResolveOverlappingMode.Default;
                        chartControl1.Series[t].Label.FillStyle.FillMode = FillMode.Gradient;
                        chartControl1.Series[t].Label.TextColor = System.Drawing.Color.FromArgb(64, 64, 64);
                        chartControl1.Series[t].Label.Border.Color = System.Drawing.Color.FromArgb(255, 128, 128);
                        ((BarSeriesLabel)chartControl1.Series[t].Label).ShowForZeroValues = true;
                        ((SideBySideBarSeriesView)chartControl1.Series[t].View).Shadow.Visible = true;
                    }
                #endregion
                #region Cтроим график
                    try
                    {
                        spisok_2 = name.GetRange(ss[0], q);
                        for (int n = 1; n <= spisok_2.Count; n++)
                        {
                            diagram.AxisX.CustomLabels.Add(new CustomAxisLabel(spisok_2[n - 1].ToString()));
                            diagram.AxisX.CustomLabels[n - 1].AxisValue = n;
                        }
                    }
                    catch (Exception g)
                    {
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        MessageBox.Show(g.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    for (int suka = 0; suka < ss.Length; suka++)
                    {
                        spisok_1.Clear();
                        spisok_1 = value.GetRange(ss[suka], q);

                        for (int n = 1; n <= spisok_1.Count; n++)
                        {
                            chartControl1.Series[suka].Points.Add(new SeriesPoint(n, Convert.ToDouble(spisok_1[n - 1])));
                        }
                    }
                    for (int t = 0; t < chartControl1.Series.Count; t++)
                    {
                        try
                        {
                            chartControl1.Series[t].Name = union[t].ToString();
                        }
                        catch (Exception w)
                        {
                            gridControl1.DataSource = null;
                            gridControl2.DataSource = null;
                            ClearDiagram();
                            chartControl1.Series[0].LegendText = "Группировка по дате";
                            MessageBox.Show(w.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    if (chartControl1.Series[0].LegendText.Equals("Группировка по дате"))
                    {
                        chartControl1.Series[0].LegendText = union[0].ToString();
                    }
                }
                else
                {
                    gridControl1.DataSource = null;
                    gridControl2.DataSource = null;
                    ClearDiagram();
                    chartControl1.Series[0].LegendText = "Группировка по дате";
                    MessageBox.Show("Отчет не содержит ни одной даты, диаграмма не может быть построена.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                    #endregion
            }
            else
            {
                XYDiagram3D diagram = (XYDiagram3D)chartControl1.Diagram;
                if (chartControl1.Series.Count > 1)
                {
                    for (int e = 1; e < chartControl1.Series.Count; e++)
                    {
                        chartControl1.Series.Clear();
                    }
                    chartControl1.Series.Add(new Series("Группировка по дате", ViewType.Bar3D));
                    ((SideBySideBar3DSeriesView)chartControl1.Series[0].View).BarDistance = 0;
                    ((SideBySideBar3DSeriesView)chartControl1.Series[0].View).BarDistanceFixed = 1;
                    ((SideBySideBar3DSeriesView)chartControl1.Series[0].View).BarWidth = 0.6;
                }
                //------------------------------
                #region Получаем год
                if (comboBoxEdit1.Text != string.Empty)
                {
                    try
                    {
                        int index1 = comboBoxEdit1.SelectedIndex;
                        year = PeriodYear(index1);
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        return;
                    }
                }
                #endregion
                #region Получаем квартал, если задан год
                if (comboBoxEdit1.Text != string.Empty && comboBoxEdit2.Text != string.Empty)
                {
                    int index2 = comboBoxEdit2.SelectedIndex;
                    string period = PeriodQuarter(comboBoxEdit1.Text, index2);
                    string zp = ",";
                    int index = period.IndexOf(zp);
                    begin = period.Substring(0, index);
                    end = period.Substring(index + 1);
                    quarter = String.Format("([Дата] >= '{0}') AND ([Дата] < '{1}')", Convert.ToDateTime(begin), Convert.ToDateTime(end));
                }
                if ((comboBoxEdit1.Text == string.Empty && comboBoxEdit2.Text != string.Empty))
                {
                    gridControl1.DataSource = null;
                    gridControl2.DataSource = null;
                    ClearDiagram();
                    chartControl1.Series[0].LegendText = "Группировка по дате";
                    MessageBox.Show("Не выбран год. Укажите год для фильтрации по кварталу.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if ((comboBoxEdit1.Text != string.Empty && checkEdit3.CheckState == CheckState.Checked) && comboBoxEdit2.Text == string.Empty)
                {
                    gridControl1.DataSource = null;
                    gridControl2.DataSource = null;
                    ClearDiagram();
                    chartControl1.Series[0].LegendText = "Группировка по дате";
                    MessageBox.Show("Для указанного года включен фильтр по кварталу с пустым значением. Заполните поле или отключите фильтр.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                #endregion
                #region Выбираем периоды
                if (dateEdit1.Text != string.Empty && dateEdit2.Text != string.Empty)
                {
                    string rows = DateBegin(dateEdit1.DateTime, dateEdit2.DateTime);
                    if (rows != string.Empty)
                    {
                        time = rows;
                    }
                    else
                    {
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        MessageBox.Show("Дата начала периода больше или совпадает с датой окончания периода. Установите корректные значения.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }
                else if (dateEdit1.Text == string.Empty || dateEdit2.Text == string.Empty)
                {
                    if (dateEdit1.Text == string.Empty && dateEdit2.Text != string.Empty)
                    {
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        MessageBox.Show("Не задана дата начала периода.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    else
                        if (dateEdit2.Text == string.Empty && dateEdit1.Text != string.Empty)
                        {
                            {
                                gridControl1.DataSource = null;
                                gridControl2.DataSource = null;
                                ClearDiagram();
                                chartControl1.Series[0].LegendText = "Группировка по дате";
                                MessageBox.Show("Не задана дата окончания периода.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                return;
                            }
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
                    adapter = new SqlCeDataAdapter("SELECT * FROM Reports", conn);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    DataView MyDataView = new DataView(dt);

                    string res1 = Find(checkedComboBoxEdit1);
                    string res2 = FindDate(layoutControlGroup2);
                    if (res2.Contains("comboBoxEdit1"))
                    {
                        filter = "AND" + year;
                    }
                    if (res2.Contains("comboBoxEdit2"))
                    {
                        filter += "AND" + quarter;
                    }
                    if (res2.Contains("dateEdit1") && res2.Contains("dateEdit2"))
                    {
                        filter += "AND" + time;
                    }
                    MyDataView.RowFilter = "[Показатель] IN " + res1 + filter;
                    int countRow = MyDataView.Count;
                    if (countRow == 0)
                    {
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        MessageBox.Show("По заданным параметрам данные отсутствуют.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    else
                    {
                        gridControl1.DataSource = MyDataView;
                        gridControl2.DataSource = MyDataView;
                        MessageBox.Show("Отчет построен.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    conn.Close();
                }
                catch (Exception q)
                {
                    gridControl1.DataSource = null;
                    gridControl2.DataSource = null;
                    ClearDiagram();
                    chartControl1.Series[0].LegendText = "Группировка по дате";
                    MessageBox.Show(q.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                #endregion
                #region Формируем данные для графика
                for (int i = 0; i < gridView2.DataRowCount; i++)
                {
                    if (gridView2.IsDataRow(i))
                    {
                        string v = gridView2.GetRowCellValue(i, "Дата").ToString();
                        col.Add(v);
                    }
                }
                union = col.Distinct().ToList();
                int count = union.Count;
                foreach (string item in union)
                {
                    for (int i = 0; i < gridView2.DataRowCount; i++)
                    {
                        if (gridView2.IsDataRow(i))
                        {
                            string v = gridView2.GetRowCellValue(i, "Дата").ToString();
                            if (item == v)
                            {
                                value.Add(gridView2.GetRowCellValue(i, "Значение").ToString());
                                name.Add(gridView2.GetRowCellValue(i, "Показатель").ToString());
                            }
                        }
                    }
                }
                gridView1.CollapseAllGroups();
                if (!count.Equals(0))
                {
                    int q = value.Count / count;
                    int sum = 0;
                    int[] ss = new int[count];
                    for (int s = 0; s < count; s++)
                    {
                        ss[s] = sum;
                        sum += q;
                    }
                    if (ss.Length > 1)
                    {
                        for (int point = 1; point < ss.Length; point++)
                        {
                            chartControl1.Series.Add(new Series(DateTime.Now.ToLongDateString(), ViewType.Bar3D));
                            ((SideBySideBar3DSeriesView)chartControl1.Series[point].View).BarDistanceFixed = 1;
                            ((SideBySideBar3DSeriesView)chartControl1.Series[point].View).BarDistance = 0;
                            ((SideBySideBar3DSeriesView)chartControl1.Series[point].View).BarWidth = 0.6;
                        }
                    }
                    for (int t = 0; t < chartControl1.Series.Count; t++)
                    {
                        diagram.AxisY.VisualRange.Auto = true;
                        diagram.AxisX.VisualRange.Auto = true;
                        chartControl1.Series[t].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                        chartControl1.Series[t].Label.TextAlignment = System.Drawing.StringAlignment.Far;
                        chartControl1.Series[t].Label.ResolveOverlappingMode = ResolveOverlappingMode.Default;
                        chartControl1.Series[t].Label.FillStyle.FillMode = FillMode.Gradient;
                        chartControl1.Series[t].Label.TextColor = System.Drawing.Color.FromArgb(64, 64, 64);
                        chartControl1.Series[t].Label.Border.Color = System.Drawing.Color.FromArgb(255, 128, 128);
                        ((Bar3DSeriesLabel)chartControl1.Series[t].Label).ShowForZeroValues = true;
                    }
                #endregion
                #region Cтроим график
                    for (int suka = 0; suka < ss.Length; suka++)
                    {
                        spisok_1.Clear();
                        spisok_1 = value.GetRange(ss[suka], q);

                        for (int n = 1; n <= spisok_1.Count; n++)
                        {
                            chartControl1.Series[suka].Points.Add(new SeriesPoint(n, Convert.ToDouble(spisok_1[n - 1])));
                        }
                    }
                    for (int t = 0; t < chartControl1.Series.Count; t++)
                    {
                        try
                        {
                            chartControl1.Series[t].Name = union[t].ToString();
                        }
                        catch (Exception w)
                        {
                            gridControl1.DataSource = null;
                            gridControl2.DataSource = null;
                            ClearDiagram();
                            chartControl1.Series[0].LegendText = "Группировка по дате";
                            MessageBox.Show(w.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    if (chartControl1.Series[0].LegendText.Equals("Группировка по дате"))
                    {
                        chartControl1.Series[0].LegendText = union[0].ToString();
                    }
                }
                else
                {
                    gridControl1.DataSource = null;
                    gridControl2.DataSource = null;
                    ClearDiagram();
                    chartControl1.Series[0].LegendText = "Группировка по дате";
                    MessageBox.Show("Отчет не содержит ни одной даты, диаграмма не может быть построена.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                    #endregion
            }
        }
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            int c = 0;
            int c1 = 0;
            foreach (CheckedListBoxItem cnt in checkedComboBoxEdit1.Properties.Items)
            {
                if (cnt.CheckState == CheckState.Checked)
                {
                    c++;
                }
            }
            if (c == 0)
            {
                gridControl1.DataSource = null;
                ClearDiagram();
                chartControl1.Series[0].LegendText = "Группировка по дате";
                MessageBox.Show("Не выбраны показатели для формирования отчета.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            foreach (DevExpress.XtraLayout.LayoutControlItem cnt in layoutControlGroup2.Items)
            {
                foreach (Control dat in cnt.Control.Controls)
                {
                    if (dat.Text != string.Empty)
                    {
                        c1++;
                    }
                }
            }
            if (c1 == 0)
            {
                gridControl1.DataSource = null;
                gridControl2.DataSource = null;
                ClearDiagram();
                chartControl1.Series[0].LegendText = "Группировка по дате";
                MessageBox.Show("Не заданы даты для определения периода.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            try
            {
                Report();
                gridView1.Columns[3].DisplayFormat.FormatString = "g";
            }
            catch (Exception w)
            {
                gridControl1.DataSource = null;
                gridControl2.DataSource = null;
                ClearDiagram();
                chartControl1.Series[0].LegendText = "Группировка по дате";
                MessageBox.Show(w.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void ClearDiagram()
        {
            int kol = chartControl1.Series.Count;
            string type = chartControl1.Diagram.GetType().ToString();
            string type2d = "DevExpress.XtraCharts.XYDiagram";
            if (type == type2d)
            {
                XYDiagram diagram = (XYDiagram)chartControl1.Diagram;
                diagram.AxisX.CustomLabels.Clear();
            }
            else
            {
                string type1 = chartControl1.Diagram.GetType().ToString();
                string type3d = "DevExpress.XtraCharts.XYDiagram3D";
                if (type == type3d)
                {
                    for (int t = 0; t < chartControl1.Series.Count; t++)
                    {
                        chartControl1.Series[t].ChangeView(ViewType.Bar);
                    }
                    XYDiagram diagram = (XYDiagram)chartControl1.Diagram;
                    diagram.AxisX.CustomLabels.Clear();
                    barButtonItem23.Enabled = false;
                }
            }
                for (int j = 0; j < kol; j++)
                {
                    chartControl1.Series[j].Points.Clear();
                }
                for (int j = 1; j < kol; j++)
                {
                    chartControl1.Series[j].ShowInLegend = false;
                }
        }
        private void Clining()
        {
            try
            {
                gridControl1.DataSource = null;
                ClearDiagram();
                foreach (CheckedListBoxItem cnt in checkedComboBoxEdit1.Properties.Items)
                {
                    if (cnt.CheckState == CheckState.Checked)
                    {
                        cnt.CheckState = CheckState.Unchecked;
                    }
                }
                comboBoxEdit1.Text = string.Empty;
                comboBoxEdit2.Text = string.Empty;
                comboBoxEdit3.Text = string.Empty;
                dateEdit1.Text = string.Empty;
                dateEdit2.Text = string.Empty;
            }
            catch (Exception ex)
            {
                gridControl1.DataSource = null;
                ClearDiagram();
                chartControl1.Series[0].LegendText = "Группировка по дате";
                MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void barButtonItem15_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Clining();
            chartControl1.Series[0].LegendText = "Группировка по дате";
        }
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                int p = comboBoxEdit3.SelectedIndex;
                string dates = comboBoxEdit3.Properties.Items[p].ToString();
                this.reports1TableAdapter.Delete(dates);
                Combo();
                Clining();
                chartControl1.Series[0].LegendText = "Группировка по дате";
                MessageBox.Show("Все записи по этой дате удалены.","ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception t)
            {
                gridControl1.DataSource = null;
                ClearDiagram();
                chartControl1.Series[0].LegendText = "Группировка по дате";
                MessageBox.Show(t.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void comboBoxEdit3_SelectedIndexChanged(object sender, EventArgs e)
        {            
           contextMenuStrip1.Show(MousePosition.X - comboBoxEdit3.Width / 2, MousePosition.Y - comboBoxEdit3.Height);
        }
        private void Combo()
        {
            try
            {
                comboBoxEdit3.Properties.Items.Clear();
                comboBoxEdit3.Text = string.Empty;
                SqlCeConnection conn;
                conn = new SqlCeConnection(@"Data Source = |DataDirectory|\FANZ.sdf");
                conn.Open();
                SqlCeDataAdapter adapter;
                adapter = new SqlCeDataAdapter("SELECT DISTINCT sysdate FROM Reports", conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                int row = dt.Rows.Count;

                for (int j = 0; j < row; j++)
                {
                    comboBoxEdit3.Properties.Items.AddRange(dt.Rows[j].ItemArray.ToList());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void Reports_Load(object sender, EventArgs e)
        {
            gridView1.Columns[0].Visible = false;
            xtraTabPage3.PageVisible = false;
            comboBoxEdit1.Properties.ReadOnly = true;
            comboBoxEdit2.Properties.ReadOnly = true;
            dateEdit1.Properties.ReadOnly = true;
            dateEdit2.Properties.ReadOnly = true;
            Combo();
        }
        #region Обработка чекбоксов
        private void checkEdit1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit1.CheckState == CheckState.Checked)
            {
                comboBoxEdit1.Properties.ReadOnly = false;
            }
            else
            {
                comboBoxEdit1.Properties.ReadOnly = true;
                comboBoxEdit1.Text = string.Empty;
            }
        }
        private void checkEdit2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit2.CheckState == CheckState.Checked)
            {
                dateEdit1.Properties.ReadOnly = false;
            }
            else
            {
                dateEdit1.Properties.ReadOnly = true;
                dateEdit1.Text = string.Empty;
            }
        }
        private void checkEdit3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit3.CheckState == CheckState.Checked)
            {
                comboBoxEdit2.Properties.ReadOnly = false;
            }
            else
            {
                comboBoxEdit2.Properties.ReadOnly = true;
                comboBoxEdit2.Text = string.Empty;
            }
        }
        private void checkEdit4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit4.CheckState == CheckState.Checked)
            {
                dateEdit2.Properties.ReadOnly = false;
            }
            else
            {
                dateEdit2.Properties.ReadOnly = true;
                dateEdit2.Text = string.Empty;
            }
        }
        #endregion
        #region Ползунок ширины
        private void barEditItem3_EditValueChanged(object sender, EventArgs e)
        {
            step = Convert.ToDouble(barEditItem3.EditValue);
            string type = chartControl1.Diagram.GetType().ToString();
            string type2d = "DevExpress.XtraCharts.XYDiagram";
            if (type == type2d)
            {
                TextAnnotation Annotation = (TextAnnotation)chartControl1.AnnotationRepository.GetElementByName("Вид");
                if (Annotation.Text == "Классический вид")
                {
                    if (step > 0)
                    {
                        if (chartControl1.Series.Count > 0)
                        {
                            for (int e1 = 0; e1 < chartControl1.Series.Count; e1++)
                            {
                                ((SideBySideBarSeriesView)chartControl1.Series[e1].View).BarWidth = step;
                                ((SideBySideBarSeriesView)chartControl1.Series[e1].View).EqualBarWidth = true;
                            }
                        }
                    }
                    else
                    {
                        if (chartControl1.Series.Count > 0)
                        {
                            for (int e1 = 0; e1 < chartControl1.Series.Count; e1++)
                            {
                                ((SideBySideBarSeriesView)chartControl1.Series[e1].View).BarWidth = 0.6;
                                ((SideBySideBarSeriesView)chartControl1.Series[e1].View).BarDistanceFixed = 1;
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Ширину рядов можно менять только в классическом виде.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    return;
                }
            }
            else
            {
                TextAnnotation Annotation = (TextAnnotation)chartControl1.AnnotationRepository.GetElementByName("Вид");
                if (Annotation.Text == "Классический вид")
                {
                    if (step > 0)
                    {
                        if (chartControl1.Series.Count > 0)
                        {
                            for (int e1 = 0; e1 < chartControl1.Series.Count; e1++)
                            {
                                ((SideBySideBar3DSeriesView)chartControl1.Series[e1].View).BarWidth = step;
                                ((SideBySideBar3DSeriesView)chartControl1.Series[e1].View).EqualBarWidth = true;
                            }
                        }
                    }
                    else
                    {
                        if (chartControl1.Series.Count > 0)
                        {
                            for (int e1 = 0; e1 < chartControl1.Series.Count; e1++)
                            {
                                ((SideBySideBar3DSeriesView)chartControl1.Series[e1].View).BarWidth = 0.6;
                                ((SideBySideBar3DSeriesView)chartControl1.Series[e1].View).BarDistanceFixed = 1;
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Ширину рядов можно менять только в классическом виде.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    return;
                }
            }
        }
        #endregion

        #region Переключаем на некрасивый вид
        private void barButtonItem18_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            TextAnnotation Annotation = (TextAnnotation)chartControl1.AnnotationRepository.GetElementByName("Вид");
            Annotation.Text = "Классический вид";
            // #region Сбор данных, конвертация
            string year = string.Empty;
            string quarter = string.Empty;
            string begin = string.Empty;
            string end = string.Empty;
            string time = string.Empty;
            string filter = string.Empty;
            List<string> col = new List<string>();
            List<string> union = new List<string>();
            List<string> value = new List<string>();
            List<string> name = new List<string>();
            List<string> spisok_1 = new List<string>();
            List<string> spisok_2 = new List<string>();
            //делаем очистку-----------------
            union.Clear();
            value.Clear();
            col.Clear();
            name.ToArray();
            spisok_1.Clear();
            spisok_2.Clear();
            //попытка подчистить метки перед любым запуском
            string type = chartControl1.Diagram.GetType().ToString();
            string type2d = "DevExpress.XtraCharts.XYDiagram";
            if (type == type2d)
            {
                XYDiagram diagram = (XYDiagram)chartControl1.Diagram;
                diagram.AxisX.CustomLabels.Clear();

                if (chartControl1.Series.Count > 1)
                {
                    for (int eq = 1; eq < chartControl1.Series.Count; eq++)
                    {
                        chartControl1.Series.Clear();
                    }
                    chartControl1.Series.Add(new Series("Группировка по дате", ViewType.Bar));
                    ((SideBySideBarSeriesView)chartControl1.Series[0].View).BarDistance = 0;
                    ((SideBySideBarSeriesView)chartControl1.Series[0].View).BarDistanceFixed = 4;
                    ((SideBySideBarSeriesView)chartControl1.Series[0].View).BarWidth = 3;
                }
                //------------------------------
                #region Получаем год
                if (comboBoxEdit1.Text != string.Empty)
                {
                    try
                    {
                        int index1 = comboBoxEdit1.SelectedIndex;
                        year = PeriodYear(index1);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        return;
                    }
                }
                #endregion
                #region Получаем квартал, если задан год
                if (comboBoxEdit1.Text != string.Empty && comboBoxEdit2.Text != string.Empty)
                {
                    int index2 = comboBoxEdit2.SelectedIndex;
                    string period = PeriodQuarter(comboBoxEdit1.Text, index2);
                    string zp = ",";
                    int index = period.IndexOf(zp);
                    begin = period.Substring(0, index);
                    end = period.Substring(index + 1);
                    quarter = String.Format("([Дата] >= '{0}') AND ([Дата] < '{1}')", Convert.ToDateTime(begin), Convert.ToDateTime(end));
                }
                if ((comboBoxEdit1.Text == string.Empty && comboBoxEdit2.Text != string.Empty))
                {
                    gridControl1.DataSource = null;
                    gridControl2.DataSource = null;
                    ClearDiagram();
                    chartControl1.Series[0].LegendText = "Группировка по дате";
                    MessageBox.Show("Не выбран год. Укажите год для фильтрации по кварталу.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if ((comboBoxEdit1.Text != string.Empty && checkEdit3.CheckState == CheckState.Checked) && comboBoxEdit2.Text == string.Empty)
                {
                    gridControl1.DataSource = null;
                    gridControl2.DataSource = null;
                    ClearDiagram();
                    chartControl1.Series[0].LegendText = "Группировка по дате";
                    MessageBox.Show("Для указанного года включен фильтр по кварталу с пустым значением. Заполните поле или отключите фильтр.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                #endregion
                #region Выбираем периоды
                if (dateEdit1.Text != string.Empty && dateEdit2.Text != string.Empty)
                {
                    string rows = DateBegin(dateEdit1.DateTime, dateEdit2.DateTime);
                    if (rows != string.Empty)
                    {
                        time = rows;
                    }
                    else
                    {
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        MessageBox.Show("Дата начала периода больше или совпадает с датой окончания периода. Установите корректные значения.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }
                else if (dateEdit1.Text == string.Empty || dateEdit2.Text == string.Empty)
                {
                    if (dateEdit1.Text == string.Empty && dateEdit2.Text != string.Empty)
                    {
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        MessageBox.Show("Не задана дата начала периода.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    else
                        if (dateEdit2.Text == string.Empty && dateEdit1.Text != string.Empty)
                        {
                            {
                                gridControl1.DataSource = null;
                                gridControl2.DataSource = null;
                                ClearDiagram();
                                chartControl1.Series[0].LegendText = "Группировка по дате";
                                MessageBox.Show("Не задана дата окончания периода.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                return;
                            }
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
                    adapter = new SqlCeDataAdapter("SELECT * FROM Reports", conn);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    DataView MyDataView = new DataView(dt);

                    string res1 = Find(checkedComboBoxEdit1);
                    string res2 = FindDate(layoutControlGroup2);
                    if (res2.Contains("comboBoxEdit1"))
                    {
                        filter = "AND" + year;
                    }
                    if (res2.Contains("comboBoxEdit2"))
                    {
                        filter += "AND" + quarter;
                    }
                    if (res2.Contains("dateEdit1") && res2.Contains("dateEdit2"))
                    {
                        filter += "AND" + time;
                    }
                    MyDataView.RowFilter = "[Показатель] IN " + res1 + filter;
                    int countRow = MyDataView.Count;
                    if (countRow == 0)
                    {
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        MessageBox.Show("По заданным параметрам данные отсутствуют.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    else
                    {
                        gridControl1.DataSource = MyDataView;
                        gridControl2.DataSource = MyDataView;
                    }
                    conn.Close();
                }
                catch (Exception q)
                {
                    gridControl1.DataSource = null;
                    gridControl2.DataSource = null;
                    ClearDiagram();
                    chartControl1.Series[0].LegendText = "Группировка по дате";
                    MessageBox.Show(q.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                #endregion
                #region Первичные данные
                for (int i = 0; i < gridView2.DataRowCount; i++)
                {
                    if (gridView2.IsDataRow(i))
                    {
                        string v = gridView2.GetRowCellValue(i, "Дата").ToString();
                        col.Add(v);
                    }
                }
                union = col.Distinct().ToList();
                int count = union.Count;
                foreach (string item in union)
                {
                    for (int i = 0; i < gridView2.DataRowCount; i++)
                    {
                        if (gridView2.IsDataRow(i))
                        {
                            string v = gridView2.GetRowCellValue(i, "Дата").ToString();
                            if (item == v)
                            {
                                value.Add(gridView2.GetRowCellValue(i, "Значение").ToString());
                                name.Add(gridView2.GetRowCellValue(i, "Показатель").ToString());
                            }
                        }
                    }
                }
                gridView1.CollapseAllGroups();
                if (!count.Equals(0))
                {
                    int q = value.Count / count;
                    int sum = 0;
                    int[] ss = new int[count];
                    for (int s = 0; s < count; s++)
                    {
                        ss[s] = sum;
                        sum += q;
                    }
                    if (ss.Length > 1)
                    {
                        for (int point = 1; point < ss.Length; point++)
                        {
                            chartControl1.Series.Add(new Series(DateTime.Now.ToLongDateString(), ViewType.Bar));
                            ((SideBySideBarSeriesView)chartControl1.Series[point].View).BarDistance = 0;
                            ((SideBySideBarSeriesView)chartControl1.Series[point].View).BarDistanceFixed = 4;
                            ((SideBySideBarSeriesView)chartControl1.Series[point].View).BarWidth = 3;
                        }
                    }
                    for (int t = 0; t < chartControl1.Series.Count; t++)
                    {
                        diagram.AxisY.VisualRange.Auto = true;
                        diagram.AxisX.VisualRange.Auto = true;
                        chartControl1.Series[t].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                        chartControl1.Series[t].Label.TextAlignment = System.Drawing.StringAlignment.Far;
                        chartControl1.Series[t].Label.ResolveOverlappingMode = ResolveOverlappingMode.Default;
                        chartControl1.Series[t].Label.FillStyle.FillMode = FillMode.Gradient;
                        chartControl1.Series[t].Label.TextColor = System.Drawing.Color.FromArgb(64, 64, 64);
                        chartControl1.Series[t].Label.Border.Color = System.Drawing.Color.FromArgb(255, 128, 128);
                        ((BarSeriesLabel)chartControl1.Series[t].Label).ShowForZeroValues = true;
                        ((SideBySideBarSeriesView)chartControl1.Series[t].View).Shadow.Visible = true;
                    }
                #endregion
                #region Cтроим второй график
                    for (int suka = 0; suka < ss.Length; suka++)
                    {
                        spisok_1.Clear();
                        spisok_1 = value.GetRange(ss[suka], q);
                        for (int n = ss[suka] + 1; n <= (suka + 1) * q; n++)
                        {
                            chartControl1.Series[suka].Points.Add(new SeriesPoint(n, Convert.ToDouble(spisok_1[n - (ss[suka] + 1)])));
                        }
                    }
                }
                for (int t = 0; t < chartControl1.Series.Count; t++)
                {
                    try
                    {
                        chartControl1.Series[t].Name = union[t].ToString();
                    }
                    catch (Exception w)
                    {
                        gridControl1.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        MessageBox.Show(w.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                    #endregion
            }
            else
            {
                XYDiagram3D diagram = (XYDiagram3D)chartControl1.Diagram;
                if (chartControl1.Series.Count > 1)
                {
                    for (int eq = 1; eq < chartControl1.Series.Count; eq++)
                    {
                        chartControl1.Series.Clear();
                    }
                    chartControl1.Series.Add(new Series("Группировка по дате", ViewType.Bar3D));
                    ((SideBySideBar3DSeriesView)chartControl1.Series[0].View).BarDistance = 0;
                    ((SideBySideBar3DSeriesView)chartControl1.Series[0].View).BarDistanceFixed = 4;
                    ((SideBySideBar3DSeriesView)chartControl1.Series[0].View).BarWidth = 3;
                }
                //------------------------------
                #region Получаем год
                if (comboBoxEdit1.Text != string.Empty)
                {
                    try
                    {
                        int index1 = comboBoxEdit1.SelectedIndex;
                        year = PeriodYear(index1);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        return;
                    }
                }
                #endregion
                #region Получаем квартал, если задан год
                if (comboBoxEdit1.Text != string.Empty && comboBoxEdit2.Text != string.Empty)
                {
                    int index2 = comboBoxEdit2.SelectedIndex;
                    string period = PeriodQuarter(comboBoxEdit1.Text, index2);
                    string zp = ",";
                    int index = period.IndexOf(zp);
                    begin = period.Substring(0, index);
                    end = period.Substring(index + 1);
                    quarter = String.Format("([Дата] >= '{0}') AND ([Дата] < '{1}')", Convert.ToDateTime(begin), Convert.ToDateTime(end));
                }
                if ((comboBoxEdit1.Text == string.Empty && comboBoxEdit2.Text != string.Empty))
                {
                    gridControl1.DataSource = null;
                    gridControl2.DataSource = null;
                    ClearDiagram();
                    chartControl1.Series[0].LegendText = "Группировка по дате";
                    MessageBox.Show("Не выбран год. Укажите год для фильтрации по кварталу.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if ((comboBoxEdit1.Text != string.Empty && checkEdit3.CheckState == CheckState.Checked) && comboBoxEdit2.Text == string.Empty)
                {
                    gridControl1.DataSource = null;
                    gridControl2.DataSource = null;
                    ClearDiagram();
                    chartControl1.Series[0].LegendText = "Группировка по дате";
                    MessageBox.Show("Для указанного года включен фильтр по кварталу с пустым значением. Заполните поле или отключите фильтр.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                #endregion
                #region Выбираем периоды
                if (dateEdit1.Text != string.Empty && dateEdit2.Text != string.Empty)
                {
                    string rows = DateBegin(dateEdit1.DateTime, dateEdit2.DateTime);
                    if (rows != string.Empty)
                    {
                        time = rows;
                    }
                    else
                    {
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        MessageBox.Show("Дата начала периода больше или совпадает с датой окончания периода. Установите корректные значения.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }
                else if (dateEdit1.Text == string.Empty || dateEdit2.Text == string.Empty)
                {
                    if (dateEdit1.Text == string.Empty && dateEdit2.Text != string.Empty)
                    {
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        MessageBox.Show("Не задана дата начала периода.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    else
                        if (dateEdit2.Text == string.Empty && dateEdit1.Text != string.Empty)
                        {
                            {
                                gridControl1.DataSource = null;
                                gridControl2.DataSource = null;
                                ClearDiagram();
                                chartControl1.Series[0].LegendText = "Группировка по дате";
                                MessageBox.Show("Не задана дата окончания периода.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                return;
                            }
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
                    adapter = new SqlCeDataAdapter("SELECT * FROM Reports", conn);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    DataView MyDataView = new DataView(dt);

                    string res1 = Find(checkedComboBoxEdit1);
                    string res2 = FindDate(layoutControlGroup2);
                    if (res2.Contains("comboBoxEdit1"))
                    {
                        filter = "AND" + year;
                    }
                    if (res2.Contains("comboBoxEdit2"))
                    {
                        filter += "AND" + quarter;
                    }
                    if (res2.Contains("dateEdit1") && res2.Contains("dateEdit2"))
                    {
                        filter += "AND" + time;
                    }
                    MyDataView.RowFilter = "[Показатель] IN " + res1 + filter;
                    int countRow = MyDataView.Count;
                    if (countRow == 0)
                    {
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        MessageBox.Show("По заданным параметрам данные отсутствуют.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    else
                    {
                        gridControl1.DataSource = MyDataView;
                        gridControl2.DataSource = MyDataView;
                    }
                    conn.Close();
                }
                catch (Exception q)
                {
                    gridControl1.DataSource = null;
                    gridControl2.DataSource = null;
                    ClearDiagram();
                    chartControl1.Series[0].LegendText = "Группировка по дате";
                    MessageBox.Show(q.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                #endregion
                #region Первичные данные
                for (int i = 0; i < gridView2.DataRowCount; i++)
                {
                    if (gridView2.IsDataRow(i))
                    {
                        string v = gridView2.GetRowCellValue(i, "Дата").ToString();
                        col.Add(v);
                    }
                }
                union = col.Distinct().ToList();
                int count = union.Count;
                foreach (string item in union)
                {
                    for (int i = 0; i < gridView2.DataRowCount; i++)
                    {
                        if (gridView2.IsDataRow(i))
                        {
                            string v = gridView2.GetRowCellValue(i, "Дата").ToString();
                            if (item == v)
                            {
                                value.Add(gridView2.GetRowCellValue(i, "Значение").ToString());
                                name.Add(gridView2.GetRowCellValue(i, "Показатель").ToString());
                            }
                        }
                    }
                }
                gridView1.CollapseAllGroups();
                if (!count.Equals(0))
                {
                    int q = value.Count / count;
                    int sum = 0;
                    int[] ss = new int[count];
                    for (int s = 0; s < count; s++)
                    {
                        ss[s] = sum;
                        sum += q;
                    }
                    if (ss.Length > 1)
                    {
                        for (int point = 1; point < ss.Length; point++)
                        {
                            chartControl1.Series.Add(new Series(DateTime.Now.ToLongDateString(), ViewType.Bar3D));
                            ((SideBySideBar3DSeriesView)chartControl1.Series[point].View).BarDistance = 0;
                            ((SideBySideBar3DSeriesView)chartControl1.Series[point].View).BarDistanceFixed = 4;
                            ((SideBySideBar3DSeriesView)chartControl1.Series[point].View).BarWidth = 3;
                        }
                    }
                    for (int t = 0; t < chartControl1.Series.Count; t++)
                    {
                        diagram.AxisY.VisualRange.Auto = true;
                        diagram.AxisX.VisualRange.Auto = true;
                        chartControl1.Series[t].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                        chartControl1.Series[t].Label.TextAlignment = System.Drawing.StringAlignment.Far;
                        chartControl1.Series[t].Label.ResolveOverlappingMode = ResolveOverlappingMode.Default;
                        chartControl1.Series[t].Label.FillStyle.FillMode = FillMode.Gradient;
                        chartControl1.Series[t].Label.TextColor = System.Drawing.Color.FromArgb(64, 64, 64);
                        chartControl1.Series[t].Label.Border.Color = System.Drawing.Color.FromArgb(255, 128, 128);
                        ((Bar3DSeriesLabel)chartControl1.Series[t].Label).ShowForZeroValues = true;
                    }
                #endregion
                #region Cтроим второй график
                    for (int suka = 0; suka < ss.Length; suka++)
                    {
                        spisok_1.Clear();
                        spisok_1 = value.GetRange(ss[suka], q);
                        for (int n = ss[suka] + 1; n <= (suka + 1) * q; n++)
                        {
                            chartControl1.Series[suka].Points.Add(new SeriesPoint(n, Convert.ToDouble(spisok_1[n - (ss[suka] + 1)])));
                        }
                    }
                }
                for (int t = 0; t < chartControl1.Series.Count; t++)
                {
                    try
                    {
                        chartControl1.Series[t].Name = union[t].ToString();
                    }
                    catch (Exception w)
                    {
                        gridControl1.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        MessageBox.Show(w.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                    #endregion
            }
        }
        #endregion`
        #region Возвращаем обратно
        private void barButtonItem19_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Report();
            TextAnnotation Annotation = (TextAnnotation)chartControl1.AnnotationRepository.GetElementByName("Вид");
            Annotation.Text = "Основной вид";
        }
        #endregion

        #region Отчет в 3d
        private void barButtonItem20_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string type = chartControl1.Diagram.GetType().ToString();
            string type2d = "DevExpress.XtraCharts.XYDiagram";
            if (type == type2d)
            {
                for (int t = 0; t < chartControl1.Series.Count; t++)
                {
                    chartControl1.Series[t].ChangeView(ViewType.Bar3D);
                }
                barButtonItem23.Enabled = true;
                ((XYDiagram3D)chartControl1.Diagram).RuntimeZooming = true;
                ((XYDiagram3D)chartControl1.Diagram).RuntimeScrolling = true;
                ((XYDiagram3D)chartControl1.Diagram).AxisY.Label.NumericOptions.Format = NumericFormat.General;
            }
        }
        #endregion
        #region Вкл вращение 3D
        private void barButtonItem23_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            ((XYDiagram3D)chartControl1.Diagram).RuntimeRotation = true;
        }
        #endregion
        #region Отчет в 2d
        private void barButtonItem21_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            List<string> spisok_2 = new List<string>();
            List<string> col = new List<string>();
            List<string> union = new List<string>();
            List<string> value = new List<string>();
            List<string> name = new List<string>();
            TextAnnotation Annotation = (TextAnnotation)chartControl1.AnnotationRepository.GetElementByName("Вид");

            string type = chartControl1.Diagram.GetType().ToString();
            string type3d = "DevExpress.XtraCharts.XYDiagram3D";
            if (type == type3d)
            {
                barButtonItem23.Enabled = false;                
                #region Формируем данные для графика
                for (int i = 0; i < gridView2.DataRowCount; i++)
                {
                    if (gridView2.IsDataRow(i))
                    {
                        string v = gridView2.GetRowCellValue(i, "Дата").ToString();
                        col.Add(v);
                    }
                }
                union = col.Distinct().ToList();               
                foreach (string item in union)
                {
                    for (int i = 0; i < gridView2.DataRowCount; i++)
                    {
                        if (gridView2.IsDataRow(i))
                        {
                            string v = gridView2.GetRowCellValue(i, "Дата").ToString();
                            if (item == v)
                            {
                                value.Add(gridView2.GetRowCellValue(i, "Значение").ToString());
                                name.Add(gridView2.GetRowCellValue(i, "Показатель").ToString());
                            }
                        }
                    }
                }
                int count = union.Count;
                int q = value.Count / count;
                #endregion
                for (int t = 0; t < chartControl1.Series.Count; t++)
                {
                    chartControl1.Series[t].ChangeView(ViewType.Bar);
                }
                XYDiagram diagram = (XYDiagram)chartControl1.Diagram;
                diagram.AxisX.CustomLabels.Clear();
                if (Annotation.Text == "Основной вид")
                {
                    try
                    {
                        spisok_2 = name.GetRange(0, q);
                        for (int n = 1; n <= spisok_2.Count; n++)
                        {
                            diagram.AxisX.CustomLabels.Add(new CustomAxisLabel(spisok_2[n - 1].ToString()));
                            diagram.AxisX.CustomLabels[n - 1].AxisValue = n;
                        }                      
                    }
                    catch (Exception g)
                    {
                        gridControl1.DataSource = null;
                        gridControl2.DataSource = null;
                        ClearDiagram();
                        chartControl1.Series[0].LegendText = "Группировка по дате";
                        MessageBox.Show(g.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }             
                for (int t = 0; t < chartControl1.Series.Count; t++)
                {
                    diagram.AxisY.VisualRange.Auto = true;
                    diagram.AxisX.VisualRange.Auto = true;
                    chartControl1.Series[t].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    chartControl1.Series[t].Label.TextAlignment = System.Drawing.StringAlignment.Far;
                    chartControl1.Series[t].Label.ResolveOverlappingMode = ResolveOverlappingMode.Default;
                    chartControl1.Series[t].Label.FillStyle.FillMode = FillMode.Gradient;
                    chartControl1.Series[t].Label.TextColor = System.Drawing.Color.FromArgb(64, 64, 64);
                    chartControl1.Series[t].Label.Border.Color = System.Drawing.Color.FromArgb(255, 128, 128);
                    ((BarSeriesLabel)chartControl1.Series[t].Label).ShowForZeroValues = true;
                    ((SideBySideBarSeriesView)chartControl1.Series[t].View).Shadow.Visible = true;
                }                
            }
        }
        #endregion

        private void gridView1_PrintInitialize(object sender, DevExpress.XtraGrid.Views.Base.PrintInitializeEventArgs e)
        {
            PrintingSystemBase pb = e.PrintingSystem as PrintingSystemBase;
            pb.PageSettings.PaperKind = System.Drawing.Printing.PaperKind.A4;
            pb.PageSettings.Landscape = false;
            pb.PageSettings.BottomMargin = 10;
            pb.PageSettings.TopMargin = 10;
            pb.PageSettings.RightMargin = 10;
            pb.PageSettings.LeftMargin = 10; 
        }

        private void barButtonItem24_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0)
            {
                if (!gridControl1.IsPrintingAvailable)
                {
                    MessageBox.Show("The 'DevExpress.XtraPrinting' library is not found", "Error");
                    return;
                }
                gridControl1.ShowPrintPreview();
            }
            else
            {
                MessageBox.Show("Таблица пустая, предпросмотр отмененен.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }          
        }

        private void barButtonItem13_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0)
            {
                if (!gridControl1.IsPrintingAvailable)
                {
                    MessageBox.Show("The 'DevExpress.XtraPrinting' library is not found", "Error");
                    return;
                }
                gridControl1.Print();
            }
            else
            {
                MessageBox.Show("Таблица пустая, печать отменена.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #region Экспорт в ексель (3 вида)
        private void barButtonItem6_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0)
            {
                string date1 = DateTime.Today.ToShortDateString();
                string date2 = "Cоздан в " + Convert.ToString(DateTime.Now.Hour + "." + DateTime.Now.Minute);
                string filename = date2 + ".xlsx";
                DialogResult results;
                //------------------------------------------------------------------------------------------------------------------------              
                try
                {
                    if (!Directory.Exists(s + @"\Отчеты\Динамический отчет\" + "" + date1))
                    {
                        Directory.CreateDirectory(s + @"\Отчеты\Динамический отчет\" + "" + date1);
                        string m = s + "\\Отчеты\\Динамический отчет\\" + date1;
                        gridView1.ExportToXlsx(m + @"\" + filename);
                        MessageBox.Show("Данные в Ексель успешно экспортированы.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string m = s + "\\Отчеты\\Динамический отчет\\" + date1;
                        gridView1.ExportToXlsx(m + @"\" + filename);
                        MessageBox.Show("Данные в Ексель успешно экспортированы.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                    openFileDialog1.InitialDirectory = s + @"\Отчеты\Динамический отчет\" + date1;
                    openFileDialog1.Filter = "txt files (*.xlsх)|*.txt|All files (*.*)|*.*";
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
            else
            {
                MessageBox.Show("Таблица пустая. Экспорт запрещен.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }            
        }
        private void barButtonItem10_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0)
            {
                string date1 = DateTime.Today.ToShortDateString();
                string date2 = "Cоздан в " + Convert.ToString(DateTime.Now.Hour + "." + DateTime.Now.Minute);
                string filename = date2 + ".xls";
                DialogResult results;
                //------------------------------------------------------------------------------------------------------------------------              
                try
                {
                    if (!Directory.Exists(s + @"\Отчеты\Динамический отчет\" + "" + date1))
                    {
                        Directory.CreateDirectory(s + @"\Отчеты\Динамический отчет\" + "" + date1);
                        string m = s + "\\Отчеты\\Динамический отчет\\" + date1;
                        gridView1.ExportToXls(m + @"\" + filename);
                        MessageBox.Show("Данные в Ексель успешно экспортированы.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string m = s + "\\Отчеты\\Динамический отчет\\" + date1;
                        gridView1.ExportToXls(m + @"\" + filename);
                        MessageBox.Show("Данные в Ексель успешно экспортированы.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                    openFileDialog1.InitialDirectory = s + @"\Отчеты\Динамический отчет\" + date1;
                    openFileDialog1.Filter = "txt files (*.xlsх)|*.txt|All files (*.*)|*.*";
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
            else
            {
                MessageBox.Show("Таблица пустая. Экспорт запрещен.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }            
        }
        private void barButtonItem12_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0)
            {
                string date1 = DateTime.Today.ToShortDateString();
                string date2 = "Cоздан в " + Convert.ToString(DateTime.Now.Hour + "." + DateTime.Now.Minute);
                string filename = date2 + ".csv";
                DialogResult results;
                //------------------------------------------------------------------------------------------------------------------------              
                try
                {
                    if (!Directory.Exists(s + @"\Отчеты\Динамический отчет\" + "" + date1))
                    {
                        Directory.CreateDirectory(s + @"\Отчеты\Динамический отчет\" + "" + date1);
                        string m = s + "\\Отчеты\\Динамический отчет\\" + date1;
                        gridView1.ExportToCsv(m + @"\" + filename);
                        MessageBox.Show("Данные в Ексель успешно экспортированы.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string m = s + "\\Отчеты\\Динамический отчет\\" + date1;
                        gridView1.ExportToCsv(m + @"\" + filename);
                        MessageBox.Show("Данные в Ексель успешно экспортированы.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                    openFileDialog1.InitialDirectory = s + @"\Отчеты\Динамический отчет\" + date1;
                    openFileDialog1.Filter = "txt files (*.xlsх)|*.txt|All files (*.*)|*.*";
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
            else
            {
                MessageBox.Show("Таблица пустая. Экспорт запрещен.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }           
        }
        #endregion
        #region Экспорт в rtf
        private void barButtonItem8_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0)
            {
                string date1 = DateTime.Today.ToShortDateString();
                string date2 = "Cоздан в " + Convert.ToString(DateTime.Now.Hour + "." + DateTime.Now.Minute);
                string filename = date2 + ".rtf";
                DialogResult results;
                //------------------------------------------------------------------------------------------------------------------------              
                try
                {
                    if (!Directory.Exists(s + @"\Отчеты\Динамический отчет\" + "" + date1))
                    {
                        Directory.CreateDirectory(s + @"\Отчеты\Динамический отчет\" + "" + date1);
                        string m = s + "\\Отчеты\\Динамический отчет\\" + date1;
                        gridView1.ExportToRtf(m + @"\" + filename);
                        MessageBox.Show("Данные успешно экспортированы.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string m = s + "\\Отчеты\\Динамический отчет\\" + date1;
                        gridView1.ExportToRtf(m + @"\" + filename);
                        MessageBox.Show("Данные успешно экспортированы.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                    openFileDialog1.InitialDirectory = s + @"\Отчеты\Динамический отчет\" + date1;
                    openFileDialog1.Filter = "RTF Files(*.rtf)|*.rtf";
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
            else
            {
                MessageBox.Show("Таблица пустая. Экспорт запрещен.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion
        #region Экспорт в pdf
        private void barButtonItem7_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0)
            {
                string date1 = DateTime.Today.ToShortDateString();
                string date2 = "Cоздан в " + Convert.ToString(DateTime.Now.Hour + "." + DateTime.Now.Minute);
                string filename = date2 + ".pdf";
                DialogResult results;
                //------------------------------------------------------------------------------------------------------------------------              
                try
                {
                    if (!Directory.Exists(s + @"\Отчеты\Динамический отчет\" + "" + date1))
                    {
                        Directory.CreateDirectory(s + @"\Отчеты\Динамический отчет\" + "" + date1);
                        string m = s + "\\Отчеты\\Динамический отчет\\" + date1;
                        gridView1.ExportToPdf(m + @"\" + filename);
                        MessageBox.Show("Данные в Ексель успешно экспортированы.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string m = s + "\\Отчеты\\Динамический отчет\\" + date1;
                        gridView1.ExportToPdf(m + @"\" + filename);
                        MessageBox.Show("Данные успешно экспортированы.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                    openFileDialog1.InitialDirectory = s + @"\Отчеты\Динамический отчет\" + date1;
                    openFileDialog1.Filter = "PDF Files(*.pdf)|*.pdf";
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
            else
            {
                MessageBox.Show("Таблица пустая. Экспорт запрещен.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion
        #region Экспорт в html
        private void barButtonItem9_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0)
            {
                string date1 = DateTime.Today.ToShortDateString();
                string date2 = "Cоздан в " + Convert.ToString(DateTime.Now.Hour + "." + DateTime.Now.Minute);
                string filename = date2 + ".htm";
                DialogResult results;
                //------------------------------------------------------------------------------------------------------------------------              
                try
                {
                    if (!Directory.Exists(s + @"\Отчеты\Динамический отчет\" + "" + date1))
                    {
                        Directory.CreateDirectory(s + @"\Отчеты\Динамический отчет\" + "" + date1);
                        string m = s + "\\Отчеты\\Динамический отчет\\" + date1;
                        gridView1.ExportToHtml(m + @"\" + filename);
                        MessageBox.Show("Данные в Ексель успешно экспортированы.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string m = s + "\\Отчеты\\Динамический отчет\\" + date1;
                        gridView1.ExportToHtml(m + @"\" + filename);
                        MessageBox.Show("Данные успешно экспортированы.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                    openFileDialog1.InitialDirectory = s + @"\Отчеты\Динамический отчет\" + date1;
                    openFileDialog1.Filter = "HTML Files(*.htm)|*.htm";
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
            else
            {
                MessageBox.Show("Таблица пустая. Экспорт запрещен.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion
        #region Экспорт в текстовый формат
        private void barButtonItem11_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0)
            {
                string date1 = DateTime.Today.ToShortDateString();
                string date2 = "Cоздан в " + Convert.ToString(DateTime.Now.Hour + "." + DateTime.Now.Minute);
                string filename = date2 + ".txt";
                DialogResult results;
                //------------------------------------------------------------------------------------------------------------------------              
                try
                {
                    if (!Directory.Exists(s + @"\Отчеты\Динамический отчет\" + "" + date1))
                    {
                        Directory.CreateDirectory(s + @"\Отчеты\Динамический отчет\" + "" + date1);
                        string m = s + "\\Отчеты\\Динамический отчет\\" + date1;
                        gridView1.ExportToText(m + @"\" + filename);
                        MessageBox.Show("Данные в Ексель успешно экспортированы.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string m = s + "\\Отчеты\\Динамический отчет\\" + date1;
                        gridView1.ExportToText(m + @"\" + filename);
                        MessageBox.Show("Данные успешно экспортированы.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                    openFileDialog1.InitialDirectory = s + @"\Отчеты\Динамический отчет\" + date1;
                    openFileDialog1.Filter = "TXT Files(*.txt)|*.txt";
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
            else
            {
                MessageBox.Show("Таблица пустая. Экспорт запрещен.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion
    }
}
