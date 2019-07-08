using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Simplex_2.Формы_проекта;
using System.Xml;
using System.Security.Cryptography;
using System.Data.SqlServerCe;
using DevExpress.XtraCharts;

using Invention;
using DevExpress.Utils;
using DevExpress.XtraBars;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraPrinting;

namespace Simplex_2
{
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public static string host = Path.Combine(System.Windows.Forms.Application.StartupPath, "host.txt");
        public string s = System.Windows.Forms.Application.StartupPath;
        public static string style_name = Path.Combine(System.Windows.Forms.Application.StartupPath, "style_name.txt");

        public static string shablon = Path.Combine(Application.StartupPath, "result_new.xml");
        public static string shablon_2 = Path.Combine(Application.StartupPath, "result2_new.xml");
        public static string shablon_3 = Path.Combine(Application.StartupPath, "result3.xml");
       
        public double ms = 0;
        public string key = string.Empty;
        DevExpress.XtraCharts.Printing.ChartPrinter cp;
        public Form1()
        {
            InitializeComponent();
        }
        #region Все остальные методы
        private void button1_Click(object sender, EventArgs e)
        {
            BindingSource bindingSource1 = new BindingSource();
            try
            {
                DataTable dataT;
                BindingSource bindS;
                using (SqlCeConnection conn = new SqlCeConnection("Data Source=|DataDirectory|\\Fanz.sdf"))
                {
                    dataT = new DataTable();
                    bindS = new BindingSource();
                    string query = "SELECT x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11, x12, x13, x14, x15, x16, x17, x18, x19, x20, x21, x22, x23, x24, x25, X26, X27, X28, X29, X30, X31, X32, X33, X34, X35, B_i_ FROM Gauss";
                    SqlCeDataAdapter dA = new SqlCeDataAdapter(query, conn);
                    SqlCeCommandBuilder cBuilder = new SqlCeCommandBuilder(dA);
                    dA.Fill(dataT);
                    bindS.DataSource = dataT;
                    dataGridView1.DataSource = bindS;
                }
            }
            catch (Exception quest)
            {
                MessageBox.Show(quest.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        #region Пока что заполнение векторов ограничений
        public void Vector_S(DataGridView datagriedview1)
        {
            int last = dataGridView1.ColumnCount - 1;
            try
            {
                if (textBox12.Text != "" && textBox13.Text != "" && textBox14.Text != "")
                {
                    dataGridView1.Rows[7].Cells[last].Value = Convert.ToDouble(textBox12.Text);
                    dataGridView1.Rows[9].Cells[last].Value = Convert.ToDouble(textBox13.Text);
                    dataGridView1.Rows[10].Cells[last].Value = Convert.ToDouble(textBox14.Text);
                }
                else
                {
                    MessageBox.Show("Проверьте значения БА, УК, БП!", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (textBox16.Text != "")
                {
                    dataGridView1.Rows[16].Cells[last].Value = Convert.ToDouble(textBox12.Text) * Convert.ToDouble(textBox16.Text);
                }
                else
                {
                    MessageBox.Show("Заполните к-т имущества производств. назнач!", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (textBox15.Text != "" && textBox25.Text != "")
                {
                    dataGridView1.Rows[33].Cells[last].Value = Convert.ToDouble(textBox25.Text) * Convert.ToDouble(textBox15.Text);
                }
                else
                {
                    MessageBox.Show("Не заполнены ФОТ или СебФОТ!", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (textBox6.Text != "" && textBox14.Text != "")
                {
                    dataGridView1.Rows[34].Cells[last].Value = Convert.ToDouble(textBox6.Text) * Convert.ToDouble(textBox14.Text);
                }
                else
                {
                    MessageBox.Show("Не заполнены БП или ОСовК!", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ФАНЗ");
                return;
            }
        }
        public void Vector_M(DataGridView datagriedview)
        {
            try
            {
                if (textBox12.Text != "" && textBox13.Text != "" && textBox14.Text != "")
                {
                    dataGridView2.Rows[7].Cells[83].Value = Convert.ToDouble(textBox12.Text);
                    dataGridView2.Rows[9].Cells[83].Value = Convert.ToDouble(textBox13.Text);
                    dataGridView2.Rows[10].Cells[83].Value = Convert.ToDouble(textBox14.Text);
                }
                else
                {
                    MessageBox.Show("Проверьте значения БА, УК, БП!", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                if (textBox16.Text != "")
                {
                    dataGridView2.Rows[16].Cells[83].Value = Convert.ToDouble(textBox12.Text) * Convert.ToDouble(textBox16.Text);
                }
                else
                {
                    MessageBox.Show("Заполните к-т имущества производств. назнач!", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (textBox15.Text != "" && textBox25.Text != "")
                {
                    dataGridView2.Rows[33].Cells[83].Value = Convert.ToDouble(textBox25.Text) * Convert.ToDouble(textBox15.Text);
                }
                else
                {
                    MessageBox.Show("Не заполнены ФОТ или СебФОТ!", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (textBox6.Text != "" && textBox14.Text != "")
                {
                    dataGridView2.Rows[34].Cells[83].Value = Convert.ToDouble(textBox6.Text) * Convert.ToDouble(textBox14.Text);
                }
                else
                {
                    MessageBox.Show("Не заполнены БП или ОСовК!", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ФАНЗ");
                return;
            }
        }
        #endregion
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                Metods gauss2 = new Metods();
                gauss2.SolveGauss(dataGridView1, dataGridView5);
            }
            catch (Exception quest)
            {
                MessageBox.Show(quest.Message, "ФАНЗ");
                return;
            }
        }
        private void dataGridView1_MouseHover(object sender, EventArgs e)
        {
            dataGridView1.Focus();
        }
        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            object head = this.dataGridView1.Rows[e.RowIndex].HeaderCell.Value;

            if (head == null || !head.Equals((e.RowIndex + 1).ToString()))
                this.dataGridView1.Rows[e.RowIndex].HeaderCell.Value = (e.RowIndex + 1).ToString();

        }
        #region Задаем погрешность епсилон для простого симплекс метода
        private void button6_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount != 0)
            {
                Epsilon eps = new Epsilon();
                eps.ShowDialog();

                if (eps.textBox1.Text != "")
                {
                    int M = dataGridView1.RowCount;
                    double epsil = Convert.ToDouble(eps.textBox1.Text);
                    for (int i = 0; i < M; i++)
                    {
                        double tmp = Convert.ToDouble(dataGridView1.Rows[i].Cells[i].Value);

                        if (tmp == 0)
                        {
                            dataGridView1.Rows[i].Cells[i].Value = epsil;
                            dataGridView1.Rows[i].Cells[i].Style.BackColor = System.Drawing.Color.LightGreen;
                        }
                        else dataGridView1.Rows[i].Cells[i].Style.BackColor = System.Drawing.Color.OrangeRed;
                    }
                }
            }
            else
            {
                MessageBox.Show("Таблица пустая, нельзя задать погрешность", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        #endregion
        private void dataGridView2_MouseHover(object sender, EventArgs e)
        {
            dataGridView2.Focus();
        }
        private void dataGridView2_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            object head = this.dataGridView2.Rows[e.RowIndex].HeaderCell.Value;

            if (head == null || !head.Equals((e.RowIndex + 1).ToString()))
                this.dataGridView2.Rows[e.RowIndex].HeaderCell.Value = (e.RowIndex + 1).ToString();
        }     
        private void button7_Click(object sender, EventArgs e)
        {
            BindingSource bindingSource1 = new BindingSource();
            try
            {
                DataTable dataT;
                BindingSource bindS;
                using (SqlCeConnection conn = new SqlCeConnection("Data Source=|DataDirectory|\\Fanz.sdf"))
                {
                    dataT = new DataTable();
                    bindS = new BindingSource();
                    string query = "SELECT Базис, x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11, x12, x13, x14, x15, x16, x17, x18, x19, x20, x21, x22, x23, x24, x25, X26, X27, X28, X29, X30, X31, X32, X33, X34, X35, X36, X37, X38, X39, X40, X41, X42, X43, X44, X45, X46, X47, X48, R1, R2, R3, R4, R5, R6, R7, R8, R9, R10, R11, R12, R13, R14, R15, R16, R17, R18, R19, R20, R21, R22, R23, R24, R25, R26, R27, R28, R29, R30, R31, R32, R33, R34, B_i_, [Min Bi/Ai] FROM Simplex"; 
                    SqlCeDataAdapter dA = new SqlCeDataAdapter(query, conn);
                    SqlCeCommandBuilder cBuilder = new SqlCeCommandBuilder(dA);
                    dA.Fill(dataT);
                    bindS.DataSource = dataT;
                    dataGridView2.DataSource = bindS;
                }
            }
            catch (Exception quest)
            {
                MessageBox.Show(quest.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        #endregion
        #region При загрузке формы и при функции изменения контролов показателей
        #region Отключение кнопок при автозагрузке
        private void LoadPro()
        {
            barButtonItem8.Enabled = false;
            barButtonItem10.Enabled = false;
            barButtonItem11.Enabled = false;
            barButtonItem13.Enabled = false;

            barButtonItem19.Enabled = false;
            barSubItem10.Enabled = false;
            barSubItem11.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button17.Enabled = false;
            button21.Enabled = false;
        }
        #endregion
        private void Form1_Load_1(object sender, EventArgs e)
        {
            this.Visible = true;
            string stile = string.Empty;
            this.Hide();
            Form2 f = new Form2();
            this.Hide();
            f.ShowDialog();

            if (f.radioButton1.Checked == true)
            {
                xtraTabPage3.PageVisible = false;
                xtraTabPage4.PageVisible = false;
                ribbonPageGroup3.Visible = false;
            }
            this.Show();
            LoadPro();
            if (File.Exists(@style_name))
            {
                StreamReader reader = new StreamReader(@style_name, Encoding.GetEncoding(1251));
                {
                    while (!reader.EndOfStream)
                        stile = reader.ReadLine();
                }
                reader.Close();
            }
            defaultLookAndFeel1.LookAndFeel.SetSkinStyle(stile);

            #region После авторизации проверяем, что нам требуется показывать
            
            XYDiagram diagram = (XYDiagram)chartControl2.Diagram;
            diagram.AxisX.Visible = false;
            toolStripStatusLabel2.Text = "Проект (загрузка данных, заполнение показателей)";
            groupControl1.Height = 20;
            groupControl2.Height = 20;
            groupControl3.Height = 20;
            groupControl4.Height = 20;
            groupControl5.Height = 20;
            groupControl1.Width = 215;
            groupControl2.Width = 215;
            groupControl3.Width = 215;
            groupControl4.Width = 215;
            groupControl5.Width = 215;
            if (dataGridView4.Rows.Count == 0)
            {
                dataGridView4.Rows.Add();
                for (int i = 0; i < dataGridView4.ColumnCount; i++)
                {
                    dataGridView4[i, 0].Value = 0;
                }
            }
            else
            {
                dataGridView4.Rows.Clear();
                dataGridView4.Rows.Add();

                for (int i = 0; i < dataGridView4.ColumnCount; i++)
                {
                    dataGridView4[i, 0].Value = 0;
                }
            }
            #endregion
        }
        #region Навигация по аккардиону
        private void groupControl1_DoubleClick(object sender, EventArgs e)
        {
            if (groupControl1.Height == 271)
            {
                groupControl1.Height = 20;
                groupControl1.Width = 215;
                if (groupControl2.Height != 128)
                { groupControl2.Width = 215; }
                if (groupControl3.Height != 176)
                { groupControl3.Width = 215; }
                if (groupControl4.Height != 89)
                { groupControl4.Width = 215; }
                if (groupControl5.Height != 250)
                { groupControl5.Width = 215; }
            }
            else if (groupControl1.Height == 20)
            {
                groupControl1.Height = 271;
                groupControl1.Width = 408;
                groupControl2.Width = 408;
                groupControl3.Width = 408;
                groupControl4.Width = 408;
                groupControl5.Width = 408;
            }
        }
        private void groupControl2_DoubleClick_1(object sender, EventArgs e)
        {
            if (groupControl2.Height == 128)
            {
                groupControl2.Height = 20;
                groupControl2.Width = 215;
                if (groupControl1.Height != 271)
                { groupControl1.Width = 215; }
                if (groupControl3.Height != 176)
                { groupControl3.Width = 215; }
                if (groupControl4.Height != 89)
                { groupControl4.Width = 215; }
                if (groupControl5.Height != 250)
                { groupControl5.Width = 215; }
            }
            else if (groupControl2.Height == 20)
            {
                groupControl2.Height = 128;
                groupControl1.Width = 408;
                groupControl2.Width = 408;
                groupControl3.Width = 408;
                groupControl4.Width = 408;
                groupControl5.Width = 408;
            }
        }
        private void groupControl3_DoubleClick(object sender, EventArgs e)
        {
            if (groupControl3.Height == 176)
            {
                groupControl3.Height = 20;
                groupControl3.Width = 215;
                if (groupControl1.Height != 271)
                { groupControl1.Width = 215; }
                if (groupControl2.Height != 128)
                { groupControl2.Width = 215; }
                if (groupControl4.Height != 89)
                { groupControl4.Width = 215; }
                if (groupControl5.Height != 250)
                { groupControl5.Width = 215; }
            }
            else if (groupControl3.Height == 20)
            {
                groupControl3.Height = 176;
                groupControl1.Width = 408;
                groupControl2.Width = 408;
                groupControl3.Width = 408;
                groupControl4.Width = 408;
                groupControl5.Width = 408;
            }
        }
        private void groupControl4_DoubleClick(object sender, EventArgs e)
        {
            if (groupControl4.Height == 89)
            {
                groupControl4.Height = 20;
                groupControl4.Width = 215;
                if (groupControl1.Height != 271)
                { groupControl1.Width = 215; }
                if (groupControl2.Height != 128)
                { groupControl2.Width = 215; }
                if (groupControl3.Height != 176)
                { groupControl3.Width = 215; }
                if (groupControl5.Height != 250)
                { groupControl5.Width = 215; }
            }
            else if (groupControl4.Height == 20)
            {
                groupControl4.Height = 89;
                groupControl1.Width = 408;
                groupControl2.Width = 408;
                groupControl3.Width = 408;
                groupControl4.Width = 408;
                groupControl5.Width = 408;
            }
        }
        private void groupControl5_DoubleClick(object sender, EventArgs e)
        {
            if (groupControl5.Height == 250)
            {
                groupControl5.Height = 20;
                groupControl5.Width = 215;
                if (groupControl1.Height != 271)
                { groupControl1.Width = 215; }
                if (groupControl2.Height != 128)
                { groupControl2.Width = 215; }
                if (groupControl3.Height != 176)
                { groupControl3.Width = 215; }
                if (groupControl4.Height != 89)
                { groupControl4.Width = 215; }
            }
            else if (groupControl5.Height == 20)
            {
                groupControl5.Height = 250;
                groupControl1.Width = 408;
                groupControl2.Width = 408;
                groupControl3.Width = 408;
                groupControl4.Width = 408;
                groupControl5.Width = 408;
            }
        #endregion
        }
            
        #endregion
        #region Проверка, что заполнены все показатели и проверка диапазона вводимых значений
        private void proverka_ogr()
        {
            List<string> list = new List<string>();
            List<string> message = new List<string>();
            DialogResult res;

            button2.Enabled = true;
            button3.Enabled = true;
            barButtonItem11.Enabled = true;
           

            if (textBox2.Text == "")
            {
                list.AddRange(new String[] { "МСОС" });
            }

            if (textBox4.Text == "")
            {
                list.AddRange(new String[] { "КОСОС" });
            }
            else if (Convert.ToDouble(textBox4.Text) < 0.1)
            {
                message.AddRange(new String[] { "Рекомендуемое значение для КОСОС должно быть больше 0,1." });
            }

            if (textBox3.Text == "")
            {
                list.AddRange(new String[] { "КАЛ" });
            }
            else
            {
                if (Convert.ToDouble(textBox3.Text) > 0.5 || Convert.ToDouble(textBox3.Text) < 0.2)
                {
                    message.AddRange(new String[] { "Рекомендуемый диапазон для КАЛ от 0,2 до 0,5." });
                }
            }

            if (textBox5.Text == "")
            {
                list.AddRange(new String[] { "КБЛ" });
            }
            else if (Convert.ToDouble(textBox5.Text) > 1.2 || Convert.ToDouble(textBox5.Text) < 0.7)
            {
                message.AddRange(new String[] { "Рекомендуемый диапазон для КБЛ от 0,7 до 1,2." });
            }

            if (textBox1.Text == "")
            {
                list.AddRange(new String[] { "КТЛ" });
            }
            else if (Convert.ToDouble(textBox1.Text) > 2.5 || Convert.ToDouble(textBox1.Text) < 1.5)
            {
                message.AddRange(new String[] { "Рекомендуемый диапазон для КТЛ от 1,5 до 2,5." });
            }

            if (textBox16.Text == "")
            {
                list.AddRange(new String[] { "КИПН" });
            }
            else if (Convert.ToDouble(textBox16.Text) < 0.5)
            {
                message.AddRange(new String[] { "Рекомендуемое значение для КИПН должно быть больше 0,5." });
            }

            if (textBox8.Text == "")
            {
                list.AddRange(new String[] { "ДСОСЗ" });
            }
            else if (Convert.ToDouble(textBox8.Text) < 0.9)
            {
                message.AddRange(new String[] { "Рекомендуемое значение для ДСОЗС должно быть больше 0,9." });
            }

            if (textBox18.Text == "")
            {
                list.AddRange(new String[] { "КФЛ" });
            }
            else if (Convert.ToDouble(textBox18.Text) > 2.0 || Convert.ToDouble(textBox18.Text) < 1.0)
            {
                message.AddRange(new String[] { "Рекомендуемый диапазон для КФЛ от 1,0 до 2,0." });
            }

            if (textBox17.Text == "")
            {
                list.AddRange(new String[] { "СДВ" });
            }

            if (textBox7.Text == "")
            {
                list.AddRange(new String[] { "КМСК" });
            }
            else if (Convert.ToDouble(textBox7.Text) < 0.05)
            {
                message.AddRange(new String[] { "Рекомендуемое значение для КМСК должно быть больше 0,05." });
            }

            if (textBox19.Text == "")
                list.AddRange(new String[] { "РПВП" });
            if (textBox20.Text == "")
                list.AddRange(new String[] { "ROE" });
            if (textBox9.Text == "")
                list.AddRange(new String[] { "ОДЗ" });
            if (textBox22.Text == "")
                list.AddRange(new String[] { "ОЗ" });
            if (textBox11.Text == "")
                list.AddRange(new String[] { "ФО" });
            if (textBox21.Text == "")
                list.AddRange(new String[] { "ОСК" });
            if (textBox6.Text == "")
                list.AddRange(new String[] { "ОСовК" });
            if (textBox12.Text == "")
                list.AddRange(new String[] { "БА" });
            if (textBox13.Text == "")
                list.AddRange(new String[] { "УК" });
            if (textBox14.Text == "")
                list.AddRange(new String[] { "БП" });
            if (textBox15.Text == "")
                list.AddRange(new String[] { "ФОТ" });
            if (textBox25.Text == "")
                list.AddRange(new String[] { "СебФОТ" });
            if (textBox26.Text == "")
                list.AddRange(new String[] { "Н.пр" });
            if (textBox27.Text == "")
                list.AddRange(new String[] { "ОЧП" });
            if (textBox24.Text == "")
                list.AddRange(new String[] { "КрУр" });

            if (message.Count != 0)
            {
                string output1 = string.Join("\n", message);
                MessageBox.Show(output1, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            if (list.Count != 0)
            {
                button2.Enabled = false;
                button3.Enabled = false;
                barButtonItem11.Enabled = false;
                

                string output = string.Join(", ", list);
                res = MessageBox.Show("Не заполнены обязательные поля.\nВывести список пустых полей?", "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (res == System.Windows.Forms.DialogResult.Yes)
                {
                    MessageBox.Show(output, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                else return;
            }
        }
        private void Logical_Control()
        {
            DialogResult result;
            DialogResult res;
            ClearColor();
            #region Актив = Пассив
            if (textBox12.Text != textBox14.Text)
            {
                result = MessageBox.Show("Валюта актива не равна валюте пассива.\nСкорректируйте указанные показатели.", "ФАНЗ", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    textBox12.Text = "";
                    textBox14.Text = "";
                    button2.Enabled = false;
                    button3.Enabled = false;
                    barButtonItem11.Enabled = false;
                   
                    textBox12.BackColor = Color.LightBlue;
                    textBox14.BackColor = Color.LightBlue;
                    return;
                }
                else
                {
                    button2.Enabled = false;
                    button3.Enabled = false;
                    barButtonItem11.Enabled = false;
                   
                }
            }
            #endregion
            #region  ОСовК <= ОСК
            if (textBox6.Text != "" && textBox21.Text != "" && (Convert.ToDouble(textBox6.Text) > Convert.ToDouble(textBox21.Text)))
            {
                res = MessageBox.Show("Не выполнено следующее неравенство: ОСовК <= ОСК.\nВозможно будут получены некорректные результаты.\nПродолжить расчеты?",
                   "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (res == System.Windows.Forms.DialogResult.No)
                {
                    button2.Enabled = false;
                    button3.Enabled = false;
                    barButtonItem11.Enabled = false;
                    
                    textBox6.Text = "";
                    textBox21.Text = "";
                    return;
                }
                else
                {
                    textBox6.BackColor = Color.LightCoral;
                    textBox21.BackColor = Color.LightCoral;
                }
            }
            #endregion
            #region РСК (ROE) < ОСК
            if (textBox20.Text != "" && textBox21.Text != "" && (Convert.ToDouble(textBox20.Text) >= Convert.ToDouble(textBox21.Text)))
            {
                res = MessageBox.Show("Не выполнено следующее неравенство: РСК (ROE) < ОСК.\nВозможно будут получены некорректные результаты.\nПродолжить расчеты?",
                   "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (res == System.Windows.Forms.DialogResult.No)
                {
                    button2.Enabled = false;
                    button3.Enabled = false;
                    barButtonItem11.Enabled = false;
                   

                    textBox20.Text = "";
                    textBox21.Text = "";
                    return;
                }
                else
                {
                    textBox20.BackColor = Color.LightCoral;
                    textBox21.BackColor = Color.LightCoral;
                }
            }
            #endregion
            #region ОДЗ > OСовК
            if (textBox9.Text != "" && textBox6.Text != "" && (Convert.ToDouble(textBox9.Text) < Convert.ToDouble(textBox6.Text)))
            {
                res = MessageBox.Show("Не выполнено следующее неравенство: ОДЗ > OСовК.\nВозможно будут получены некорректные результаты.\nПродолжить расчеты?",
                   "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (res == System.Windows.Forms.DialogResult.No)
                {
                    button2.Enabled = false;
                    button3.Enabled = false;
                    barButtonItem11.Enabled = false;
                   
                    textBox6.Text = "";
                    textBox9.Text = "";
                    return;
                }
                else
                {
                    textBox6.BackColor = Color.LightCoral;
                    textBox9.BackColor = Color.LightCoral;
                }
            }
            #endregion
            #region ФО > ОСовК
            if (textBox11.Text != "" && textBox6.Text != "" && (Convert.ToDouble(textBox11.Text) < Convert.ToDouble(textBox6.Text)))
            {
                res = MessageBox.Show("Не выполнено следующее неравенство: ФО > ОСовК.\nВозможно будут получены некорректные результаты.\nПродолжить расчеты?",
                   "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (res == System.Windows.Forms.DialogResult.No)
                {
                    button2.Enabled = false;
                    button3.Enabled = false;
                    barButtonItem11.Enabled = false;
                    
                    textBox6.Text = "";
                    textBox11.Text = "";
                    return;
                }
                else
                {
                    textBox6.BackColor = Color.LightCoral;
                    textBox11.BackColor = Color.LightCoral;
                }
            }
            #endregion
            #region ОЗ > ОСовК
            if (textBox22.Text != "" && textBox6.Text != "" && (Convert.ToDouble(textBox22.Text) < Convert.ToDouble(textBox6.Text)))
            {
                res = MessageBox.Show("Не выполнено следующее неравенство: ОЗ > ОСовК.\nВозможно будут получены некорректные результаты.\nПродолжить расчеты?",
                   "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (res == System.Windows.Forms.DialogResult.No)
                {
                    button2.Enabled = false;
                    button3.Enabled = false;
                    barButtonItem11.Enabled = false;
                    
                    textBox6.Text = "";
                    textBox22.Text = "";
                    return;
                }
                else
                {
                    textBox6.BackColor = Color.LightCoral;
                    textBox22.BackColor = Color.LightCoral;
                }
            }
            #endregion
            #region РПВП < 100%
            if (textBox19.Text != "")
            {
                double value = Convert.ToDouble(textBox19.Text);

                if ((value < 0) || (value > 1))
                {
                    res = MessageBox.Show("Показатель РПВП указывается в долях. За 1 принимается 100%.\nДопустимый диапазон показателя от 0 до 1.\nПродолжить расчеты?.", "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                    if (res == System.Windows.Forms.DialogResult.No)
                    {
                        button2.Enabled = false;
                        button3.Enabled = false;
                        barButtonItem11.Enabled = false;
                        
                        textBox19.Clear();
                        return;
                    }
                    else
                    {
                        textBox19.BackColor = Color.LightCoral;
                    }
                }
            }
            #endregion
            #region КТЛ > КБЛ > КАЛ
            if (textBox1.Text != "" && textBox3.Text != "" && textBox5.Text != "")
            {
                bool pokaz1 = (Convert.ToDouble(textBox1.Text) > Convert.ToDouble(textBox5.Text));
                bool pokaz2 = (Convert.ToDouble(textBox5.Text) > Convert.ToDouble(textBox3.Text));

                if (pokaz1 == false || pokaz2 == false || (pokaz1 == false && pokaz2 == false))
                {
                    textBox1.BackColor = Color.White;
                    textBox3.BackColor = Color.White;
                    textBox5.BackColor = Color.White;
                    res = MessageBox.Show("Не выполнено следующее неравенство: КТЛ > КБЛ > КАЛ.\nВозможно будут получены некорректные результаты.\nПродолжить расчеты?",
                    "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                    if (res == System.Windows.Forms.DialogResult.No)
                    {
                        button2.Enabled = false;
                        button3.Enabled = false;
                        barButtonItem11.Enabled = false;
                       
                        textBox1.Text = "";
                        textBox3.Text = "";
                        textBox5.Text = "";
                        return;
                    }
                    else
                    {
                        textBox1.BackColor = Color.LightCoral;
                        textBox3.BackColor = Color.LightCoral;
                        textBox5.BackColor = Color.LightCoral;
                    }
                }
            }
            #endregion
        }
        #endregion
        #region Запуск алгоритма
        public bool Vivod_M()
        {
            while (button3.Enabled == true)
            {
                Bazis res = new Bazis();

                if (res.dopustim(dataGridView2) == true)
                {
                    return true;
                }
                else
                {
                    res.dopustim(dataGridView2);

                    if (res.optimalnost(dataGridView2) == true)
                    {
                        res.Remove_M(dataGridView2);
                        button3.Enabled = false;
                        return true;
                    }
                    if (res.Simpex_res(dataGridView2) == true)
                    {
                        return true;
                    }
                    else res.Simpex_res(dataGridView2);
                }
            }
            return false;
        }
        public bool Vivod_Z()
        {
            while (dataGridView2.Columns.Count > 2)
            {
                Bazis otvet = new Bazis();

                if (otvet.dopustim_Z(dataGridView2) == true)
                {
                    return true;
                }
                else
                {
                    otvet.dopustim_Z(dataGridView2);

                    if (otvet.optimalnost_Z(dataGridView2) == true)
                    {
                        dataGridView2.Columns.RemoveAt(84);
                        for (int v = 1; v < 83; v++)
                        {
                            dataGridView2.Columns[v].Visible = false;
                        }
                        Zamena(dataGridView2);
                        Stroka_otvet(dataGridView2, dataGridView4);
                        InsertResult(gridView2, dataGridView4);     
                        return true;
                    }

                    if (otvet.Simpex_res_Z(dataGridView2) == true)
                    {
                        return true;
                    }
                    else otvet.Simpex_res_Z(dataGridView2);
                }
            }

            return false;
        }
        #region Сохранение итогового отклонения в базу данных

        private void ImportToDataTable()
        {
            using (SqlCeConnection connect = new SqlCeConnection(@"Data Source = |DataDirectory|\FANZ.sdf"))
            {
                using (SqlCeCommand command = new SqlCeCommand())
                {
                    command.Connection = connect;
                    command.CommandText = "INSERT INTO Reports (Показатель, Значение, Дата, sysdate) VALUES (@Показатель,@Значение,@Дата,@sys)";

                    command.Parameters.Add(new SqlCeParameter("@Показатель", SqlDbType.NVarChar));
                    command.Parameters.Add(new SqlCeParameter("@Значение", SqlDbType.Float));
                    command.Parameters.Add(new SqlCeParameter("@Дата", SqlDbType.DateTime));
                    command.Parameters.Add(new SqlCeParameter("@sys", SqlDbType.NVarChar));
                    connect.Open();
                    for (int i = 0; i < gridView1.DataRowCount; i++)
                    {
                        if (gridView1.IsDataRow(i))
                        {
                            command.Parameters["@Показатель"].Value = gridView1.GetRowCellValue(i, gridView1.Columns[1]);
                            command.Parameters["@Значение"].Value = gridView1.GetRowCellValue(i, gridView1.Columns[5]);
                            command.Parameters["@Дата"].Value = DateTime.Now;
                            command.Parameters["@sys"].Value = DateTime.Now.ToString();
                            command.ExecuteNonQuery();
                        }
                    }
                }
            }
            MessageBox.Show("Данные успешно сохранены", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void InsertToBZ()
        {
            DialogResult ExportSQL;
            ExportSQL = MessageBox.Show("Сохранить данные в базу данных для отчета?", "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (ExportSQL == System.Windows.Forms.DialogResult.Yes)
            {
                try
                {
                    ImportToDataTable();
                }
                catch (Exception r)
                {
                    MessageBox.Show(r.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (label32.Text != string.Empty)
                {
                    string itog = label32.Text;
                    string date = DateTime.Now.ToString();

                    try
                    {
                        lineanTableAdapter1.InsertQuery(date, itog, null);
                        this.Validate();
                        lineanTableAdapter1.Update(fanzDataSet1);
                    }
                    catch (Exception r1)
                    {
                        MessageBox.Show(r1.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            else return;
        }

        #endregion
        public void ExecuteSolve()
        {
            #region barButtonItem
            barButtonItem11.Enabled = false;
            barButtonItem10.Enabled = false;
            barSubItem1.Enabled = false;
            barButtonItem8.Enabled = false;
            barButtonItem1.Enabled = false;
            barButtonItem19.Enabled = true;
            barSubItem10.Enabled = true;
            barSubItem11.Enabled = true;
            button17.Enabled = true;
            button21.Enabled = true;
            button2.Enabled = false;
            button3.Enabled = false;            
            barButtonItem13.Enabled = true;
            #endregion
        }
        private void button3_Click_1(object sender, EventArgs e)
        {
            Vivod_M();

            if (button3.Enabled == false)
            {
                Vivod_Z();
            }
            else if (button3.Enabled == true)
            {
                return;
            }
            ExecuteSolve();
            toolStripStatusLabel2.Text = "Зарегистрирован (найдено решение)";
        }
        #endregion
        //---------------------------------------------------------------------------
        #region САМ РАСЧЕТ МЕТОДА ЗАКОНЧЕН. ОСТАЛЬНОЕ ЛОГИКА ПРИЛОЖЕНИЯ
        #region Замена переменных по оптимизации на соответствующие им текстовые обозначения
        private void Zamena(DataGridView d)
        {
            int N = d.RowCount - 1;
            int pos = 0;
            string stroka = string.Empty;
            for (int y = 0; y < N; y++)
            {
                stroka = d[0, y].Value.ToString();
                if (stroka.Equals("x1"))
                {
                    stroka = "ТА";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x2"))
                {
                    stroka = "КП";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x3"))
                {
                    stroka = "СОС";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x4"))
                {
                    stroka = "ДС";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x5"))
                {
                    stroka = "ДЗ";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x6"))
                {
                    stroka = "ВА";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x7"))
                {
                    stroka = "ЗЗ";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x8"))
                {
                    stroka = "ПК";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x9"))
                {
                    stroka = "СК";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x10"))
                {
                    stroka = "ФР";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x11"))
                {
                    stroka = "ДП";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x12"))
                {
                    stroka = "Оср";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x13"))
                {
                    stroka = "ПР.ТА";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x14"))
                {
                    stroka = "ПР.ВА";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x15"))
                {
                    stroka = "ВР";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x16"))
                {
                    stroka = "Вал.пр.";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x17"))
                {
                    stroka = "КрУр";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x18"))
                {
                    stroka = "Пр.прод";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x19"))
                {
                    stroka = "Проч.д";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x20"))
                {
                    stroka = "Проч.р";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x21"))
                {
                    stroka = "ЧП";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x22"))
                {
                    stroka = "RСовК";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x23"))
                {
                    stroka = "Пр.до_нал.";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x24"))
                {
                    stroka = "Себ.прод.";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x25"))
                {
                    stroka = "НерПр";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x26"))
                {
                    stroka = "d1";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x27"))
                {
                    stroka = "d2";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x28"))
                {
                    stroka = "d3";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x29"))
                {
                    stroka = "d4";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x30"))
                {
                    stroka = "d5";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x31"))
                {
                    stroka = "d6";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x32"))
                {
                    stroka = "d7";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x33"))
                {
                    stroka = "d8";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x34"))
                {
                    stroka = "d9";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x35"))
                {
                    stroka = "d10";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x36"))
                {
                    stroka = "d11";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x37"))
                {
                    stroka = "d12";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x38"))
                {
                    stroka = "d13";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x39"))
                {
                    stroka = "d14";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x40"))
                {
                    stroka = "d15";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x41"))
                {
                    stroka = "d16";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x42"))
                {
                    stroka = "d17";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x43"))
                {
                    stroka = "d18";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x44"))
                {
                    stroka = "d19";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x45"))
                {
                    stroka = "d20";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x46"))
                {
                    stroka = "d21";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x47"))
                {
                    stroka = "d22";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
                if (stroka.Equals("x48"))
                {
                    stroka = "d23";
                    pos = y;
                    DataGridViewCell x_i = d[0, pos];
                    x_i.Value = stroka;
                }
            }
        }
        #endregion
        #region Сравниваем с рассчетом и заполняем строку для ответа
        private void Stroka_otvet(DataGridView grid1, DataGridView grid2)
        {
            int N = grid1.RowCount;
            int M = grid2.ColumnCount;
            int pos = 0;
            int um = 0;
            string stroka = string.Empty;
            string header = string.Empty;

            try
            {
                for (int y = 0; y < M; y++)
                {
                    header = grid2.Columns[y].HeaderCell.Value.ToString();

                    for (int x = 0; x < N; x++)
                    {
                        stroka = grid1[0, x].Value.ToString();

                        if (header.Equals(stroka))
                        {
                            pos = x;
                            um = y;
                            if (header != "RСовК")
                            {
                                header = Math.Round(Convert.ToDouble(grid1[83, pos].Value)).ToString();
                            }
                            else header = Math.Round(Convert.ToDouble(grid1[83, pos].Value), 2).ToString();
                            grid2[um, 0].Value = header;
                        }
                    }
                }
            }
            catch (Exception r)
            {
                MessageBox.Show(r.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        #endregion
        #region Из локализованного решения пишем значения в результирующие таблицы
        private  void InsertResult(GridView grid3, DataGridView grid4)
        {
            
            if (grid3.RowCount != 0 && grid4.RowCount != 0 && textBox13.Text != "")
            {
                grid3.SetRowCellValue(1, gridView2.Columns[2], grid4[11, 0].Value);
                grid3.SetRowCellValue(2, gridView2.Columns[2], grid4[13, 0].Value);
                grid3.SetRowCellValue(3, gridView2.Columns[2], grid4[5, 0].Value);
                grid3.SetRowCellValue(6, gridView2.Columns[2], grid4[6, 0].Value);
                grid3.SetRowCellValue(7, gridView2.Columns[2], grid4[4, 0].Value);
                grid3.SetRowCellValue(8, gridView2.Columns[2], grid4[3, 0].Value);
                grid3.SetRowCellValue(9, gridView2.Columns[2], grid4[12, 0].Value);
                grid3.SetRowCellValue(10, gridView2.Columns[2], grid4[0, 0].Value);
                grid3.SetRowCellValue(12, gridView2.Columns[2], (Convert.ToDouble(grid4[5, 0].Value) + Convert.ToDouble(grid4[0, 0].Value)));
                grid3.SetRowCellValue(16, gridView2.Columns[2], Convert.ToDouble(textBox13.Text));
                grid3.SetRowCellValue(17, gridView2.Columns[2], grid4[9, 0].Value);
                grid3.SetRowCellValue(18, gridView2.Columns[2], grid4[24, 0].Value);
                grid3.SetRowCellValue(19, gridView2.Columns[2], grid4[8, 0].Value);
                grid3.SetRowCellValue(22, gridView2.Columns[2], grid4[10, 0].Value);
                grid3.SetRowCellValue(23, gridView2.Columns[2], grid4[1, 0].Value);
                grid3.SetRowCellValue(24, gridView2.Columns[2], grid4[7, 0].Value);
                grid3.SetRowCellValue(26, gridView2.Columns[2], (Convert.ToDouble(grid4[8, 0].Value) + Convert.ToDouble(grid4[7, 0].Value)));
                grid3.SetRowCellValue(30, gridView2.Columns[2], grid4[14, 0].Value);
                grid3.SetRowCellValue(31, gridView2.Columns[2], grid4[23, 0].Value);
                grid3.SetRowCellValue(32, gridView2.Columns[2], grid4[15, 0].Value);
                grid3.SetRowCellValue(33, gridView2.Columns[2], grid4[16, 0].Value);
                grid3.SetRowCellValue(34, gridView2.Columns[2], grid4[17, 0].Value);
                grid3.SetRowCellValue(35, gridView2.Columns[2], grid4[18, 0].Value);
                grid3.SetRowCellValue(36, gridView2.Columns[2], grid4[19, 0].Value);
                grid3.SetRowCellValue(37, gridView2.Columns[2], grid4[22, 0].Value);
                grid3.SetRowCellValue(38, gridView2.Columns[2], grid4[20, 0].Value);
                grid3.SetRowCellValue(39, gridView2.Columns[2], grid4[24, 0].Value);
            }
        }
        #endregion
        
        #region Загрузка результирующей таблицы Результаты оптимизации из XML в dataGridView
        private void button16_Click_1(object sender, EventArgs e)
        {
            if (gridView2.RowCount == 0)
            {
                try
                {
                    #region загрузка в в gridcontrol2
                    try
                    {
                        XmlReader xmlFile;
                        xmlFile = XmlReader.Create(@shablon, new XmlReaderSettings());
                        DataSet ds = new DataSet();
                        ds.ReadXml(xmlFile);
                        gridControl2.DataSource = ds.Tables[0];
                        #region Формирование всех свойств и методов для таблицы
                        gridView2.Columns[0].Caption = "Бухгалтерский баланс. Основные показатели.\nАКТИВ";
                        gridView2.Columns[1].Caption = "Фактические значения";
                        gridView2.Columns[2].Caption = "Метод оптимизации на основе\n заданных показателей";

                        for (int t = 0; t < gridView2.Columns.Count; t++)
                        {
                            gridView2.Columns[t].AppearanceHeader.Options.UseTextOptions = true;
                            gridView2.Columns[t].AppearanceHeader.BackColor = Color.LightSteelBlue;
                            gridView2.Columns[t].AppearanceHeader.BackColor2 = Color.White;
                            gridView2.Columns[t].AppearanceHeader.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.Vertical;
                            gridView2.Columns[t].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                            gridView2.Columns[t].BestFit();
                        }
                        foreach (DevExpress.XtraGrid.Columns.GridColumn column in gridView2.Columns)
                        { column.OptionsColumn.AllowSort = DefaultBoolean.False; }

                        gridView2.Columns[1].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                        gridView2.Columns[2].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                        gridView2.Columns[1].AppearanceCell.BackColor = Color.LightGreen;
                        gridView2.Columns[1].AppearanceCell.BackColor2 = Color.White;
                        gridView2.Columns[1].AppearanceCell.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.Horizontal;
                        gridView2.Appearance.VertLine.BackColor = Color.LightSteelBlue;
                        gridView2.Appearance.HorzLine.BackColor = Color.LightSteelBlue;

                        gridView2.OptionsPrint.UsePrintStyles = true;
                        gridView2.Appearance.HeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
                        gridView2.ColumnPanelRowHeight = 50;
                        DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit Grid_MemoEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit(); // инициализируем MemoEdit с именем “MyGrid_MemoEdit”
                        Grid_MemoEdit.WordWrap = true;

                        gridView2.OptionsView.RowAutoHeight = true;
                        gridView2.Columns[0].ColumnEdit = Grid_MemoEdit;
                        gridView2.Columns[2].ColumnEdit = Grid_MemoEdit;

                        for (int i = 0; i < gridView2.DataRowCount; i++)
                        {
                            if (i == 0 || i == 4 || i == 5 || i == 11 || i == 13 || i == 14 || i == 15 || i == 20 || i == 21 || i == 25 || i == 27 || i == 28 || i == 29)
                            {
                                gridView2.SetRowCellValue(i, gridView2.Columns[2], null);
                            }
                        }
                        #endregion
                        #region Формулы рассчета по ячейкам
                        gridView2.SetRowCellValue(3, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(1, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(2, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(10, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(6, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(7, gridView2.Columns[1])) +
                                           Convert.ToDouble(gridView2.GetRowCellValue(8, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(9, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(12, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(3, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(10, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(19, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(16, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(17, gridView2.Columns[1])) +
                                           Convert.ToDouble(gridView2.GetRowCellValue(18, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(24, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(22, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(23, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(26, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(19, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(24, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(32, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(30, gridView2.Columns[1])) - Convert.ToDouble(gridView2.GetRowCellValue(31, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(34, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(32, gridView2.Columns[1])) - Convert.ToDouble(gridView2.GetRowCellValue(33, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(37, gridView2.Columns[1], ((Convert.ToDouble(gridView2.GetRowCellValue(34, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(35, gridView2.Columns[1]))) -
                            Convert.ToDouble(gridView2.GetRowCellValue(36, gridView2.Columns[1]))));
                        //новое---------------------
                       
                        gridView2.SetRowCellValue(39, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(18, gridView2.Columns[1]))));
                        ////--------------------------

                        gridView2.Columns[0].OptionsColumn.ReadOnly = true;
                        gridView2.Columns[2].OptionsColumn.ReadOnly = true;

                        #endregion
                        xmlFile.Dispose();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                MessageBox.Show("Таблица содержит данные, требуетя очистка.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion
        #region Загрузка из XML таблицы Результаты расчетов по результатам оптимизации
        public void button15_Click(object sender, EventArgs e)
        {
            if (gridView3.RowCount == 0)
            {
                try
                {
                    #region загрузка в в gridcontrol3
                    try
                    {
                        XmlReader xmlFile;
                        xmlFile = XmlReader.Create(@shablon_2, new XmlReaderSettings());
                        DataSet ds = new DataSet();
                        ds.ReadXml(xmlFile);
                        gridControl3.DataSource = ds.Tables[0];
                        #region Формирование всех свойств и методов для таблицы

                        gridView3.Columns[0].Caption = "№ по\n справочнику";
                        gridView3.Columns[1].Caption = "Наименование\n показателя";
                        gridView3.Columns[2].Caption = "Сокращенное\n наименование\n в модели";
                        gridView3.Columns[3].Caption = "Заданные значения\n показателей";
                        gridView3.Columns[4].Caption = "Расчетные значения по\n оптимизированным статьям\n баланса и отчета\n о прибылях и убытках";
                        gridView3.Columns[5].Caption = "Абсолютное\n отклонение";
                        gridView3.Columns[6].Caption = "Относительное\n отклонение, %";

                        for (int t = 0; t < gridView3.Columns.Count; t++)
                        {
                            gridView3.Columns[t].AppearanceHeader.Options.UseTextOptions = true;
                            gridView3.Columns[t].AppearanceHeader.BackColor = Color.LightSteelBlue;
                            gridView3.Columns[t].AppearanceHeader.BackColor2 = Color.White;
                            gridView3.Columns[t].AppearanceHeader.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.Vertical;
                            gridView3.Columns[t].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                            gridView3.Columns[t].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                            gridView3.Columns[t].BestFit();
                        }
                        gridView3.Columns[1].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Default;

                        foreach (DevExpress.XtraGrid.Columns.GridColumn column in gridView3.Columns)
                        { column.OptionsColumn.AllowSort = DefaultBoolean.False; }

                        gridView3.Appearance.VertLine.BackColor = Color.LightSteelBlue;
                        gridView3.Appearance.HorzLine.BackColor = Color.LightSteelBlue;

                        gridView3.OptionsPrint.UsePrintStyles = true;
                        gridView3.ColumnPanelRowHeight = 60;
                        DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit Grid_MemoEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit(); // инициализируем MemoEdit с именем “MyGrid_MemoEdit”
                        Grid_MemoEdit.WordWrap = true;

                        gridView3.OptionsView.RowAutoHeight = true;
                        gridView3.OptionsView.ColumnAutoWidth = true;
                        gridView3.Columns[1].ColumnEdit = Grid_MemoEdit;
                        #endregion
                        xmlFile.Dispose();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                MessageBox.Show("Таблица содержит данные, требуетя очистка.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion

        #region Обновление итоговых сумм
        public void UpdateValue()
        {
            try
            {
                if (gridView2.RowCount != 0)
                {

                    gridView2.SetRowCellValue(3, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(1, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(2, gridView2.Columns[1]))));
                    gridView2.SetRowCellValue(10, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(6, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(7, gridView2.Columns[1])) +
                                                 Convert.ToDouble(gridView2.GetRowCellValue(8, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(9, gridView2.Columns[1]))));
                    gridView2.SetRowCellValue(12, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(3, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(10, gridView2.Columns[1]))));
                    gridView2.SetRowCellValue(19, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(16, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(17, gridView2.Columns[1])) +
                                                 Convert.ToDouble(gridView2.GetRowCellValue(18, gridView2.Columns[1]))));
                    gridView2.SetRowCellValue(24, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(22, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(23, gridView2.Columns[1]))));
                    gridView2.SetRowCellValue(26, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(19, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(24, gridView2.Columns[1]))));
                    gridView2.SetRowCellValue(32, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(30, gridView2.Columns[1])) - Convert.ToDouble(gridView2.GetRowCellValue(31, gridView2.Columns[1]))));
                    gridView2.SetRowCellValue(34, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(32, gridView2.Columns[1])) - Convert.ToDouble(gridView2.GetRowCellValue(33, gridView2.Columns[1]))));
                    gridView2.SetRowCellValue(37, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(34, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(35, gridView2.Columns[1]))) -
                                                 Convert.ToDouble(gridView2.GetRowCellValue(36, gridView2.Columns[1])));
                }
            }
            catch (Exception q)
            {
                MessageBox.Show(q.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            try
            {
                if (gridView2.RowCount != 0)
                {
                    UpdateValue();
                    InsertResult(gridView2, dataGridView4);
                }
                else
                {
                    MessageBox.Show("Таблица пустая, нельзя обновить данные!", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }
            catch (Exception q)
            {
                MessageBox.Show(q.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        #endregion
        #region Фокус и подсказки по итоговым суммам
        private void dataGridView5_MouseHover(object sender, EventArgs e)
        {
            dataGridView5.Focus();
        }
        private void dataGridView5_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            object head = this.dataGridView5.Rows[e.RowIndex].HeaderCell.Value;

            if (head == null || !head.Equals((e.RowIndex + 1).ToString()))
                this.dataGridView5.Rows[e.RowIndex].HeaderCell.Value = (e.RowIndex + 1).ToString();
        }
        private void dataGridView5_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DialogResult result;

            if (e.Button == MouseButtons.Right)
            {
                result = MessageBox.Show("Эспортировать данные таблицы в Excel?", "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
                    Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
                    object _missingObj = System.Reflection.Missing.Value;
                    ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
                    ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

                    for (int i = 0; i < dataGridView5.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView5.ColumnCount; j++)
                        {
                            ExcelApp.Cells[i + 1, j + 1] = dataGridView5.Rows[i].Cells[j].Value;
                        }
                    }
                    ExcelWorkSheet.Name = "Симплекс-метод";
                    ExcelWorkSheet.Cells[1, 1] = "Показатели";
                    ExcelWorkSheet.Cells[1, 2] = "Результаты";
                    ExcelWorkSheet.Cells[1, 3] = "A*x-b=0?";
                    ExcelWorkSheet.get_Range("A1").ColumnWidth = 40;
                    ExcelWorkSheet.get_Range("B1").ColumnWidth = 20;
                    ExcelWorkSheet.get_Range("C1").ColumnWidth = 40;

                    ExcelApp.Visible = true;
                    ExcelApp.UserControl = true;
                    ExcelApp.DisplayAlerts = true;
                }
            }
        }
        private void flowLayoutPanel1_MouseHover(object sender, EventArgs e)
        {
            flowLayoutPanel1.Focus();
        }
        #endregion
        #region Контроль ввода значений в textBox, обработка и органичения на ввод показателей
        private void t_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (dataGridView2.RowCount != 0)
            {
                if (e.KeyChar == '.')
                    e.KeyChar = ',';
                if (e.KeyChar != 22)
                    e.Handled = !Char.IsDigit(e.KeyChar) && (e.KeyChar != ',' || (((System.Windows.Forms.TextBox)sender).Text.Contains(",") && !((System.Windows.Forms.TextBox)sender).SelectedText.Contains(","))) && e.KeyChar != (char)Keys.Back && (e.KeyChar != '-' || ((System.Windows.Forms.TextBox)sender).SelectionStart != 0 || (((System.Windows.Forms.TextBox)sender).Text.Contains("-") && !((System.Windows.Forms.TextBox)sender).SelectedText.Contains("-")));
                else
                {
                    double d;
                    e.Handled = !double.TryParse(Clipboard.GetText(), out d) || (d < 0 && (((System.Windows.Forms.TextBox)sender).SelectionStart != 0 || ((System.Windows.Forms.TextBox)sender).Text.Contains("-") && !((System.Windows.Forms.TextBox)sender).SelectedText.Contains("-"))) || ((d - (int)d) != 0 && ((System.Windows.Forms.TextBox)sender).Text.Contains(",") && !((System.Windows.Forms.TextBox)sender).SelectedText.Contains(","));
                    MessageBox.Show("Не удалось вставить содержимое буфера обмена");
                }
            }
            else
            {
                e.Handled = true;
                MessageBox.Show("Невозможно ввести показатели, т.к. симплекс-таблица не загружена. Загрузите данные в систему.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        private void textBox2_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox2, e);
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (textBox2.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[1].Cells[3].Value = -Convert.ToDouble(textBox2.Text);
                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[1].Cells[2].Value = -Convert.ToDouble(textBox2.Text);
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[1].Cells[3].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[1].Cells[2].Value = 0;
                    }
                }
            }
        }

        private void textBox4_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox4, e);
        }

        private void textBox4_TextChanged_1(object sender, EventArgs e)
        {
            if (textBox4.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[12].Cells[1].Value = -Convert.ToDouble(textBox4.Text);
                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[12].Cells[0].Value = -Convert.ToDouble(textBox4.Text);
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[12].Cells[1].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[12].Cells[0].Value = 0;
                    }
                }
            }
        }

        private void textBox3_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox3, e);
        }

        private void textBox3_TextChanged_1(object sender, EventArgs e)
        {
            if (textBox3.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[4].Cells[2].Value = textBox3.Text;
                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[4].Cells[1].Value = textBox3.Text;
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[4].Cells[2].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[4].Cells[1].Value = 0;
                    }
                }
            }
        }

        private void textBox5_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox5, e);
        }

        private void textBox5_TextChanged_1(object sender, EventArgs e)
        {
            if (textBox5.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[3].Cells[2].Value = textBox5.Text;
                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[3].Cells[1].Value = textBox5.Text;
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[3].Cells[2].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[3].Cells[1].Value = 0;
                    }
                }
            }
        }

        private void textBox1_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox1, e);
        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[2].Cells[2].Value = textBox1.Text;
                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[2].Cells[1].Value = textBox1.Text;
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[2].Cells[2].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[2].Cells[1].Value = 0;
                    }
                }
            }
        }

        private void textBox16_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox16, e);
        }

        private void textBox16_TextChanged_1(object sender, EventArgs e)
        {
            if (textBox16.Text != "" && dataGridView2.RowCount != 0)
            {
                double kipn = Convert.ToDouble(textBox16.Text);
            }
        }

        private void textBox8_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox8, e);
        }

        private void textBox8_TextChanged_1(object sender, EventArgs e)
        {
            if (textBox8.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[6].Cells[7].Value = -Convert.ToDouble(textBox8.Text);
                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[6].Cells[6].Value = -Convert.ToDouble(textBox8.Text);
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[6].Cells[7].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[6].Cells[6].Value = 0;
                    }
                }
            }
        }

        #region Печать, фокусы контролов, работа с екселем
        private void flowLayoutPanel1_MouseMove(object sender, MouseEventArgs e)
        {
            flowLayoutPanel1.Focus();
        }

        #endregion

        private void textBox18_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox18, e);
        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {
            if (textBox18.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[11].Cells[9].Value = -Convert.ToDouble(textBox18.Text);
                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[11].Cells[8].Value = -Convert.ToDouble(textBox18.Text);
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[11].Cells[9].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[11].Cells[8].Value = 0;
                    }
                }
            }
        }

        private void textBox17_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox17, e);
        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {
            if (textBox17.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[14].Cells[6].Value = -Convert.ToDouble(textBox17.Text);
                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[14].Cells[5].Value = -Convert.ToDouble(textBox17.Text);
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[14].Cells[6].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[14].Cells[5].Value = 0;
                    }
                }
            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox7, e);
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            if (textBox7.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[13].Cells[9].Value = -Convert.ToDouble(textBox7.Text);
                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[13].Cells[8].Value = -Convert.ToDouble(textBox7.Text);
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[13].Cells[9].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[13].Cells[8].Value = 0;
                    }
                }
            }
        }

        private void textBox19_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox19, e);
        }

        private void textBox19_TextChanged(object sender, EventArgs e)
        {
            if (textBox19.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[17].Cells[15].Value = Convert.ToDouble(textBox19.Text);
                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[17].Cells[14].Value = Convert.ToDouble(textBox19.Text);
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[17].Cells[15].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[17].Cells[14].Value = 0;
                    }
                }
            }
        }

        private void textBox20_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox20, e);
        }

        private void textBox20_TextChanged(object sender, EventArgs e)
        {
            if (textBox20.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[19].Cells[9].Value = Convert.ToDouble(textBox20.Text);
                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[19].Cells[8].Value = Convert.ToDouble(textBox20.Text);
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[19].Cells[9].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[19].Cells[8].Value = 0;
                    }
                }
            }
        }

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox9, e);
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            if (textBox9.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[26].Cells[5].Value = -Convert.ToDouble(textBox9.Text);
                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[26].Cells[4].Value = -Convert.ToDouble(textBox9.Text);
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[26].Cells[5].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[26].Cells[4].Value = 0;
                    }
                }
            }
        }

        private void textBox22_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox22, e);
        }

        private void textBox22_TextChanged(object sender, EventArgs e)
        {
            if (textBox22.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[27].Cells[7].Value = -Convert.ToDouble(textBox22.Text);
                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[27].Cells[6].Value = -Convert.ToDouble(textBox22.Text);
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[27].Cells[7].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[27].Cells[6].Value = 0;
                    }
                }
            }
        }

        private void textBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox11, e);
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            if (textBox11.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[25].Cells[12].Value = -Convert.ToDouble(textBox11.Text);
                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[25].Cells[11].Value = -Convert.ToDouble(textBox11.Text);
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[25].Cells[12].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[25].Cells[11].Value = 0;
                    }
                }
            }
        }

        private void textBox21_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox21, e);
        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            if (textBox21.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[28].Cells[9].Value = -Convert.ToDouble(textBox21.Text);
                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[28].Cells[8].Value = -Convert.ToDouble(textBox21.Text);
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[28].Cells[9].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[28].Cells[8].Value = 0;
                    }
                }
            }
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox6, e);
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            if (textBox6.Text != "" && dataGridView2.RowCount != 0)
            {
                double ocovk = Convert.ToDouble(textBox6.Text);
            }
        }

        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox12, e);
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            if (textBox12.Text != "" && dataGridView2.RowCount != 0)
            {
                double ba = Convert.ToDouble(textBox12.Text);
            }
        }

        private void textBox13_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox13, e);
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            if (textBox13.Text != "" && dataGridView2.RowCount != 0)
            {
                double uk = Convert.ToDouble(textBox13.Text);
            }
        }

        private void textBox14_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox14, e);
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            if (textBox14.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[18].Cells[22].Value = -Convert.ToDouble(textBox14.Text);
                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[18].Cells[21].Value = -Convert.ToDouble(textBox14.Text);
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[18].Cells[22].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[18].Cells[21].Value = 0;
                    }
                }
            }
        }

        private void textBox15_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox15, e);
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            if (textBox15.Text != "" && dataGridView2.RowCount != 0)
            {
                double fot = Convert.ToDouble(textBox15.Text);
            }
        }

        private void textBox25_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox25, e);
        }

        private void textBox25_TextChanged(object sender, EventArgs e)
        {
            if (textBox25.Text != "" && dataGridView2.RowCount != 0)
            {
                double oz = Convert.ToDouble(textBox25.Text);
            }
        }

        private void textBox26_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox26, e);
        }

        private void textBox26_TextChanged(object sender, EventArgs e)
        {
            if (textBox26.Text != "" && dataGridView2.RowCount != 0)
            {
                if (Convert.ToDouble(textBox26.Text) < 1)
                {
                    dataGridView2.Rows[23].Cells[23].Value = -1 + Convert.ToDouble(textBox26.Text);
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[23].Cells[22].Value = -1 + Convert.ToDouble(textBox26.Text);
                    }
                }
                else
                {
                    if (Convert.ToDouble(textBox26.Text) > 1 && Convert.ToDouble(textBox26.Text) < 100)
                    {
                        double temp = Convert.ToDouble(textBox26.Text) / 100;
                        dataGridView2.Rows[23].Cells[23].Value = -1 + temp;
                        if (dataGridView1.RowCount != 0)
                        {
                            dataGridView1.Rows[23].Cells[22].Value = -1 + temp;
                        }
                    }
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[23].Cells[23].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[23].Cells[22].Value = 0;
                    }
                }
            }
        }

        private void textBox27_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox27, e);
        }

        private void textBox27_TextChanged(object sender, EventArgs e)
        {
            if (textBox27.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[24].Cells[21].Value = -Convert.ToDouble(textBox27.Text);
                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[24].Cells[20].Value = -Convert.ToDouble(textBox27.Text);
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[24].Cells[21].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[24].Cells[20].Value = 0;
                    }
                }
            }
        }

        private void textBox24_KeyPress(object sender, KeyPressEventArgs e)
        {
            t_KeyPress(textBox24, e);
        }

        private void textBox24_TextChanged(object sender, EventArgs e)
        {
            if (textBox24.Text != "" && dataGridView2.RowCount != 0)
            {
                dataGridView2.Rows[30].Cells[24].Value = -Convert.ToDouble(textBox24.Text);
                Vector_M(dataGridView2);

                try
                {
                    int[] index_r = new int[] { 0, 5, 7, 8, 9, 10, 15, 20, 21, 22, 23, 29 };

                    for (int j = 1; j < 49; j++)
                    {
                        double sum = 0;
                        int index = 0;

                        for (int i = 0; i < index_r.GetLength(0); i++)
                        {
                            index = index_r[i];
                            sum += Convert.ToDouble(dataGridView2.Rows[index].Cells[j].Value);
                        }
                        DataGridViewCell sum1 = dataGridView2.Rows[35].Cells[j];
                        sum1.Value = -sum;
                    }
                }
                catch (Exception q)
                {
                    MessageBox.Show(q.Message, "ФАНЗ");
                }

                if (dataGridView1.RowCount != 0)
                {
                    dataGridView1.Rows[30].Cells[23].Value = -Convert.ToDouble(textBox24.Text);
                    Vector_S(dataGridView1);
                }
            }
            else
            {
                if (dataGridView2.RowCount != 0)
                {
                    dataGridView2.Rows[30].Cells[24].Value = 0;
                    if (dataGridView1.RowCount != 0)
                    {
                        dataGridView1.Rows[30].Cells[23].Value = 0;
                    }
                }
            }
        }

        #endregion
        #region Кнопка "Проверка показателей"
        private void barButtonItem10_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            proverka_ogr();
            Logical_Control();
        }
        #endregion
        #region Кнопка "Удалить из таблиц исследования м.Гауссом.
        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.DataSource != null && dataGridView5.RowCount != 0)
                {
                    DialogResult result;
                    result = MessageBox.Show("Таблицы содержат данные, \n Продолжить удаление?", "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == System.Windows.Forms.DialogResult.Yes)
                    {
                        dataGridView1.DataSource = null;
                        dataGridView1.Rows.Clear();
                        dataGridView5.Rows.Clear();
                    }

                }
                else
                {
                    MessageBox.Show("Таблицы пустые, нельзя удалить данные.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        #endregion
        #region Кнопка "Удалить из таблицы М-задачи.
        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView2.DataSource != null)
                {
                    DialogResult result;
                    result = MessageBox.Show("Таблица содержит данные, \n Продолжить удаление?", "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == System.Windows.Forms.DialogResult.Yes)
                    {
                        dataGridView2.DataSource = null;
                        dataGridView2.Rows.Clear();
                    }
                }
                else
                {
                    MessageBox.Show("Таблица пустая, нельзя удалить данные.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        #endregion
        
        #region Инициализация значений показателей в текстбоксах
        private void LoadFinValue()
        {
            textBox2.Text = Convert.ToString(0.2);
            textBox4.Text = Convert.ToString(0.3);
            textBox3.Text = Convert.ToString(0.2);
            textBox5.Text = Convert.ToString(1.2);
            textBox1.Text = Convert.ToString(2);
            textBox16.Text = Convert.ToString(0.63);
            textBox8.Text = Convert.ToString(1);

            textBox18.Text = Convert.ToString(2);
            textBox17.Text = Convert.ToString(0.5);
            textBox7.Text = Convert.ToString(1);
            
            textBox19.Text = Convert.ToString(0.5);
            textBox20.Text = Convert.ToString(0.5);

            textBox9.Text = Convert.ToString(3);
            textBox22.Text = Convert.ToString(2);
            textBox11.Text = Convert.ToString(2);
            textBox21.Text = Convert.ToString(2);
            textBox6.Text = Convert.ToString(1);

            textBox12.Text = Convert.ToString(9508);
            textBox13.Text = Convert.ToString(120);
            textBox14.Text = Convert.ToString(9508);
            textBox15.Text = Convert.ToString(2500);
            textBox25.Text = Convert.ToString(2.5);
            textBox26.Text = Convert.ToString(0.2);
            textBox27.Text = Convert.ToString(0.5);
            textBox24.Text = Convert.ToString(0.25);
        }
        #endregion
       
        #region Кнопка "Загрузить \ все данные"
        private void Dpro()
        {

            barButtonItem11.Enabled = false;
            barButtonItem13.Enabled = false;

            barButtonItem19.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;

            barButtonItem8.Enabled = true;
            barButtonItem10.Enabled = true;
            barSubItem7.Enabled = true;
        }
        private void CleanTextBox()
        {
            textBox2.Clear();
            textBox4.Clear();
            textBox3.Clear();
            textBox5.Clear();
            textBox1.Clear();
            textBox16.Clear();
            textBox8.Clear();
            textBox18.Clear();
            textBox17.Clear();
            textBox7.Clear();
            textBox19.Clear();
            textBox20.Clear();
            textBox9.Clear();
            textBox22.Clear();
            textBox11.Clear();
            textBox21.Clear();
            textBox6.Clear();
            textBox12.Clear();
            textBox13.Clear();
            textBox14.Clear();
            textBox15.Clear();
            textBox25.Clear();
            textBox26.Clear();
            textBox27.Clear();
            textBox24.Clear();
        }

        private void Loading()
        {
            try
            {
                CleanTextBox();
                Dpro();
                DataTable dataT;
                BindingSource bindS;
                using (SqlCeConnection conn = new SqlCeConnection("Data Source=|DataDirectory|\\Fanz.sdf"))
                {
                    dataT = new DataTable();
                    bindS = new BindingSource();
                    string query = "SELECT Базис, x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11, x12, x13, x14, x15, x16, x17, x18, x19, x20, x21, x22, x23, x24, x25, X26, X27, X28, X29, X30, X31, X32, X33, X34, X35, X36, X37, X38, X39, X40, X41, X42, X43, X44, X45, X46, X47, X48, R1, R2, R3, R4, R5, R6, R7, R8, R9, R10, R11, R12, R13, R14, R15, R16, R17, R18, R19, R20, R21, R22, R23, R24, R25, R26, R27, R28, R29, R30, R31, R32, R33, R34, B_i_, [Min Bi/Ai] FROM Simplex"; 
                    SqlCeDataAdapter dA = new SqlCeDataAdapter(query, conn);
                    SqlCeCommandBuilder cBuilder = new SqlCeCommandBuilder(dA);
                    dA.Fill(dataT);
                    bindS.DataSource = dataT;
                    dataGridView2.DataSource = bindS;
                }

                BindingSource bindingSource1 = new BindingSource();
                try
                {
                    DataTable dataTs;
                    BindingSource bindSs;
                    using (SqlCeConnection conn = new SqlCeConnection("Data Source=|DataDirectory|\\Fanz.sdf"))
                    {
                        dataTs = new DataTable();
                        bindSs = new BindingSource();
                        string queryq = "SELECT x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11, x12, x13, x14, x15, x16, x17, x18, x19, x20, x21, x22, x23, x24, x25, X26, X27, X28, X29, X30, X31, X32, X33, X34, X35, B_i_ FROM Gauss";
                        SqlCeDataAdapter dAq = new SqlCeDataAdapter(queryq, conn);
                        SqlCeCommandBuilder cBuilderq = new SqlCeCommandBuilder(dAq);
                        dAq.Fill(dataTs);
                        bindSs.DataSource = dataTs;
                        dataGridView1.DataSource = bindSs;
                    }
                }
                catch (Exception quest)
                {
                    MessageBox.Show(quest.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (gridView2.RowCount != 0)
                {
                    gridControl2.DataSource = null;
                    gridView2.Columns.Clear();
                    gridView2.ColumnPanelRowHeight = 20;
                    #region загрузка в в gridcontrol2
                    try
                    {
                        XmlReader xmlFile;
                        xmlFile = XmlReader.Create(@shablon, new XmlReaderSettings());
                        DataSet ds = new DataSet();
                        ds.ReadXml(xmlFile);
                        gridControl2.DataSource = ds.Tables[0];
                        #region Формирование всех свойств и методов для таблицы
                        gridView2.Columns[0].Caption = "Бухгалтерский баланс. Основные показатели.\nАКТИВ";
                        gridView2.Columns[1].Caption = "Фактические значения";
                        gridView2.Columns[2].Caption = "Метод оптимизации на основе\n заданных показателей";

                        for (int t = 0; t < gridView2.Columns.Count; t++)
                        {
                            gridView2.Columns[t].AppearanceHeader.Options.UseTextOptions = true;
                            gridView2.Columns[t].AppearanceHeader.BackColor = Color.LightSteelBlue;
                            gridView2.Columns[t].AppearanceHeader.BackColor2 = Color.White;
                            gridView2.Columns[t].AppearanceHeader.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.Vertical;
                            gridView2.Columns[t].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                            gridView2.Columns[t].BestFit();
                        }
                        foreach (DevExpress.XtraGrid.Columns.GridColumn column in gridView2.Columns)
                        { column.OptionsColumn.AllowSort = DefaultBoolean.False; }

                        gridView2.Columns[1].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                        gridView2.Columns[2].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                        gridView2.Columns[1].AppearanceCell.BackColor = Color.LightGreen;
                        gridView2.Columns[1].AppearanceCell.BackColor2 = Color.White;
                        gridView2.Columns[1].AppearanceCell.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.Horizontal;
                        gridView2.Appearance.VertLine.BackColor = Color.LightSteelBlue;
                        gridView2.Appearance.HorzLine.BackColor = Color.LightSteelBlue;

                        gridView2.OptionsPrint.UsePrintStyles = true;
                        gridView2.Appearance.HeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
                        gridView2.ColumnPanelRowHeight = 50;
                        DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit Grid_MemoEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit(); // инициализируем MemoEdit с именем “MyGrid_MemoEdit”
                        Grid_MemoEdit.WordWrap = true;

                        gridView2.OptionsView.RowAutoHeight = true;
                        gridView2.OptionsView.ColumnAutoWidth = true;
                        gridView2.Columns[0].ColumnEdit = Grid_MemoEdit;
                        gridView2.Columns[2].ColumnEdit = Grid_MemoEdit;

                        for (int i = 0; i < gridView2.DataRowCount; i++)
                        {
                            if (i == 0 || i == 4 || i == 5 || i == 11 || i == 13 || i == 14 || i == 15 || i == 20 || i == 21 || i == 25 || i == 27 || i == 28)
                            {
                                gridView2.SetRowCellValue(i, gridView2.Columns[2], null);
                            }
                        }
                        #endregion
                        #region Формулы рассчета по ячейкам
                        gridView2.SetRowCellValue(3, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(1, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(2, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(10, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(6, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(7, gridView2.Columns[1])) +
                                           Convert.ToDouble(gridView2.GetRowCellValue(8, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(9, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(12, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(3, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(10, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(19, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(16, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(17, gridView2.Columns[1])) +
                                           Convert.ToDouble(gridView2.GetRowCellValue(18, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(24, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(22, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(23, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(26, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(19, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(24, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(32, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(30, gridView2.Columns[1])) - Convert.ToDouble(gridView2.GetRowCellValue(31, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(34, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(32, gridView2.Columns[1])) - Convert.ToDouble(gridView2.GetRowCellValue(33, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(37, gridView2.Columns[1], ((Convert.ToDouble(gridView2.GetRowCellValue(34, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(35, gridView2.Columns[1]))) -
                            Convert.ToDouble(gridView2.GetRowCellValue(36, gridView2.Columns[1]))));
                        //новое---------------------
                        
                        gridView2.SetRowCellValue(39, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(18, gridView2.Columns[1]))));
                        ////--------------------------

                        gridView2.Columns[0].OptionsColumn.ReadOnly = true;
                        gridView2.Columns[2].OptionsColumn.ReadOnly = true;

                        #endregion
                        xmlFile.Dispose();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    #endregion                  
                }
                else
                {
                    #region загрузка в в gridcontrol2
                    try
                    {
                        XmlReader xmlFile;
                        xmlFile = XmlReader.Create(@shablon, new XmlReaderSettings());
                        DataSet ds = new DataSet();
                        ds.ReadXml(xmlFile);
                        gridControl2.DataSource = ds.Tables[0];
                        #region Формирование всех свойств и методов для таблицы
                        gridView2.Columns[0].Caption = "Бухгалтерский баланс. Основные показатели.\nАКТИВ";
                        gridView2.Columns[1].Caption = "Фактические значения";
                        gridView2.Columns[2].Caption = "Метод оптимизации на основе\n заданных показателей";

                        for (int t = 0; t < gridView2.Columns.Count; t++)
                        {
                            gridView2.Columns[t].AppearanceHeader.Options.UseTextOptions = true;
                            gridView2.Columns[t].AppearanceHeader.BackColor = Color.LightSteelBlue;
                            gridView2.Columns[t].AppearanceHeader.BackColor2 = Color.White;
                            gridView2.Columns[t].AppearanceHeader.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.Vertical;
                            gridView2.Columns[t].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                            gridView2.Columns[t].BestFit();
                        }
                        foreach (DevExpress.XtraGrid.Columns.GridColumn column in gridView2.Columns)
                        { column.OptionsColumn.AllowSort = DefaultBoolean.False; }

                        gridView2.Columns[1].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                        gridView2.Columns[2].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                        gridView2.Columns[1].AppearanceCell.BackColor = Color.LightGreen;
                        gridView2.Columns[1].AppearanceCell.BackColor2 = Color.White;
                        gridView2.Columns[1].AppearanceCell.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.Horizontal;
                        gridView2.Appearance.VertLine.BackColor = Color.LightSteelBlue;
                        gridView2.Appearance.HorzLine.BackColor = Color.LightSteelBlue;

                        gridView2.OptionsPrint.UsePrintStyles = true;
                        gridView2.Appearance.HeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
                        gridView2.ColumnPanelRowHeight = 50;
                        DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit Grid_MemoEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit(); // инициализируем MemoEdit с именем “MyGrid_MemoEdit”
                        Grid_MemoEdit.WordWrap = true;

                        gridView2.OptionsView.RowAutoHeight = true;
                        gridView2.OptionsView.ColumnAutoWidth = true;
                        gridView2.Columns[0].ColumnEdit = Grid_MemoEdit;
                        gridView2.Columns[2].ColumnEdit = Grid_MemoEdit;

                        for (int i = 0; i < gridView2.DataRowCount; i++)
                        {
                            if (i == 0 || i == 4 || i == 5 || i == 11 || i == 13 || i == 14 || i == 15 || i == 20 || i == 21 || i == 25 || i == 27 || i == 28)
                            {
                                gridView2.SetRowCellValue(i, gridView2.Columns[2], null);
                            }
                        }
                        #endregion
                        #region Формулы рассчета по ячейкам
                        gridView2.SetRowCellValue(3, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(1, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(2, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(10, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(6, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(7, gridView2.Columns[1])) +
                                           Convert.ToDouble(gridView2.GetRowCellValue(8, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(9, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(12, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(3, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(10, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(19, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(16, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(17, gridView2.Columns[1])) +
                                           Convert.ToDouble(gridView2.GetRowCellValue(18, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(24, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(22, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(23, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(26, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(19, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(24, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(32, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(30, gridView2.Columns[1])) - Convert.ToDouble(gridView2.GetRowCellValue(31, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(34, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(32, gridView2.Columns[1])) - Convert.ToDouble(gridView2.GetRowCellValue(33, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(37, gridView2.Columns[1], ((Convert.ToDouble(gridView2.GetRowCellValue(34, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(35, gridView2.Columns[1]))) -
                            Convert.ToDouble(gridView2.GetRowCellValue(36, gridView2.Columns[1]))));
                        //новое---------------------
                        
                        gridView2.SetRowCellValue(39, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(18, gridView2.Columns[1]))));
                        ////--------------------------

                        gridView2.Columns[0].OptionsColumn.ReadOnly = true;
                        gridView2.Columns[2].OptionsColumn.ReadOnly = true;

                        #endregion
                        xmlFile.Dispose();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    #endregion
                }

                
                LoadFinValue();

                if (dataGridView4.Rows.Count == 0)
                {
                    dataGridView4.Rows.Add();
                    for (int i = 0; i < dataGridView4.ColumnCount; i++)
                    {
                        dataGridView4[i, 0].Value = 0;
                    }
                }
                else
                {
                    dataGridView4.Rows.Clear();
                    dataGridView4.Rows.Add();

                    for (int i = 0; i < dataGridView4.ColumnCount; i++)
                    {
                        dataGridView4[i, 0].Value = 0;
                    }
                }
                MessageBox.Show("Все данные успешно загружены.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception load)
            {
                MessageBox.Show("Произошла ошибка загрузки данных. Код ошибки указан ниже.\n\n" + load, "ФАНЗ", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop);
                return;
            }
        }
        private void barButtonItem17_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            toolStripStatusLabel2.Text = "Проект (загрузка данных, заполнение показателей)";
            //Очистим все таблицы перед загрузкой
            gridControl2.DataSource = null;
            gridView2.Columns.Clear();
            gridView2.ColumnPanelRowHeight = 20;

            gridControl3.DataSource = null;
            gridView3.Columns.Clear();
            gridView3.ColumnPanelRowHeight = 20;

            gridControl1.DataSource = null;
            gridView1.Columns.Clear();
            gridView1.ColumnPanelRowHeight = 10;

            if (dataGridView2.RowCount != 0)
            {
                CleanTextBox();
                dataGridView2.DataSource = null;
                dataGridView2.Columns.Clear();
                dataGridView2.Rows.Clear();
                Loading();
            }

            else
            {
                CleanTextBox();
                Loading();
            }
            #region Загрузка из XML таблицы Результаты расчетов по результатам оптимизации
            if (gridView3.RowCount == 0)
            {
                try
                {
                    #region загрузка в в gridcontrol3
                    try
                    {
                        XmlReader xmlFile;
                        xmlFile = XmlReader.Create(@shablon_2, new XmlReaderSettings());
                        DataSet ds = new DataSet();
                        ds.ReadXml(xmlFile);
                        gridControl3.DataSource = ds.Tables[0];
                        #region Формирование всех свойств и методов для таблицы

                        gridView3.Columns[0].Caption = "№ по\n справочнику";
                        gridView3.Columns[1].Caption = "Наименование\n показателя";
                        gridView3.Columns[2].Caption = "Сокращенное\n наименование\n в модели";
                        gridView3.Columns[3].Caption = "Заданные значения\n показателей";
                        gridView3.Columns[4].Caption = "Расчетные значения по\n оптимизированным статьям\n баланса и отчета\n о прибылях и убытках";
                        gridView3.Columns[5].Caption = "Абсолютное\n отклонение";
                        gridView3.Columns[6].Caption = "Относительное\n отклонение, %";

                        for (int t = 0; t < gridView3.Columns.Count; t++)
                        {
                            gridView3.Columns[t].AppearanceHeader.Options.UseTextOptions = true;
                            gridView3.Columns[t].AppearanceHeader.BackColor = Color.LightSteelBlue;
                            gridView3.Columns[t].AppearanceHeader.BackColor2 = Color.White;
                            gridView3.Columns[t].AppearanceHeader.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.Vertical;
                            gridView3.Columns[t].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                            gridView3.Columns[t].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                            gridView3.Columns[t].BestFit();
                        }
                        gridView3.Columns[1].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Default;

                        foreach (DevExpress.XtraGrid.Columns.GridColumn column in gridView3.Columns)
                        { column.OptionsColumn.AllowSort = DefaultBoolean.False; }

                        gridView3.Appearance.VertLine.BackColor = Color.LightSteelBlue;
                        gridView3.Appearance.HorzLine.BackColor = Color.LightSteelBlue;

                        gridView3.OptionsPrint.UsePrintStyles = true;
                        gridView3.ColumnPanelRowHeight = 60;
                        DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit Grid_MemoEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit(); // инициализируем MemoEdit с именем “MyGrid_MemoEdit”
                        Grid_MemoEdit.WordWrap = true;

                        gridView3.OptionsView.RowAutoHeight = true;
                        gridView3.OptionsView.ColumnAutoWidth = true;
                        gridView3.Columns[1].ColumnEdit = Grid_MemoEdit;
                        #endregion
                        xmlFile.Dispose();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                MessageBox.Show("Таблица содержит данные, требуетя очистка.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        #endregion
            #region Загрузка из XML таблицы отклонений
            if (gridView1.RowCount == 0)
            {
                try
                {
                    #region Загрузка в новый грид
                    XmlReader xmlFile;
                    xmlFile = XmlReader.Create(@shablon_3, new XmlReaderSettings());
                    DataSet ds = new DataSet();
                    ds.ReadXml(xmlFile);
                    gridControl1.DataSource = ds.Tables[0];
                    #region Формирование всех свойств и методов для таблицы

                    gridView1.Columns[0].Caption = "№";
                    gridView1.Columns[1].Caption = "Показатель";
                    gridView1.Columns[2].Caption = "Фактические\n значения";
                    gridView1.Columns[3].Caption = "Расчетные\n значения";
                    gridView1.Columns[4].Caption = "Абсолютное\n отклонение";
                    gridView1.Columns[5].Caption = "Относительное\n отклонение, %";

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

                    gridView1.AppearancePrint.HeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
                    gridView1.OptionsPrint.UsePrintStyles = true;
                    gridView1.ColumnPanelRowHeight = 50;

                    DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit Grid_MemoEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit(); // инициализируем MemoEdit с именем “MyGrid_MemoEdit”
                    Grid_MemoEdit.WordWrap = true;
                    gridView1.OptionsView.RowAutoHeight = true;
                    gridView1.Columns[1].ColumnEdit = Grid_MemoEdit;
                    #endregion
                    xmlFile.Dispose();
                    #endregion
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                MessageBox.Show("Таблица содержит данные, требуетя очистка.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            #endregion
        }
        #endregion
        #region Кнопка "Решение задачи"
        private void barButtonItem11_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
                Vivod_M();

                if (button3.Enabled == false)
                {
                    Vivod_Z();
                }
                else if (button3.Enabled == true)
                {
                    return;
                }
                ExecuteSolve();
                toolStripStatusLabel2.Text = "Зарегистрирован (найдено решение)";
        }
        #endregion
        #region "Функция очистки цвета всех показателей"
        private void ClearColor()
        {
            textBox2.BackColor = Color.White;
            textBox4.BackColor = Color.White;
            textBox3.BackColor = Color.White;
            textBox5.BackColor = Color.White;
            textBox1.BackColor = Color.White;
            textBox16.BackColor = Color.White;
            textBox8.BackColor = Color.White;
            textBox18.BackColor = Color.White;
            textBox17.BackColor = Color.White;
            textBox7.BackColor = Color.White;
            textBox19.BackColor = Color.White;
            textBox20.BackColor = Color.White;
            textBox9.BackColor = Color.White;
            textBox22.BackColor = Color.White;
            textBox11.BackColor = Color.White;
            textBox21.BackColor = Color.White;
            textBox6.BackColor = Color.White;
            textBox12.BackColor = Color.White;
            textBox13.BackColor = Color.White;
            textBox14.BackColor = Color.White;
            textBox15.BackColor = Color.White;
            textBox25.BackColor = Color.White;
            textBox26.BackColor = Color.White;
            textBox27.BackColor = Color.White;
            textBox24.BackColor = Color.White;
        }
        #endregion
        #region Кнопка "Удалить все данные"
        private void barButtonItem8_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            ClearColor();
            try
            {
                DialogResult result;
                result = MessageBox.Show("Удалить все загруженные и введеные данные?", "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    textBox2.Text = "";
                    textBox4.Text = "";
                    textBox3.Text = "";
                    textBox5.Text = "";
                    textBox1.Text = "";
                    textBox16.Text = "";
                    textBox8.Text = "";
                    textBox18.Text = "";
                    textBox17.Text = "";
                    textBox7.Text = "";
                    textBox19.Text = "";
                    textBox20.Text = "";
                    textBox9.Text = "";
                    textBox22.Text = "";
                    textBox11.Text = "";
                    textBox21.Text = "";
                    textBox6.Text = "";
                    textBox12.Text = "";
                    textBox13.Text = "";
                    textBox14.Text = "";
                    textBox15.Text = "";
                    textBox25.Text = "";
                    textBox26.Text = "";
                    textBox27.Text = "";
                    textBox24.Text = "";

                    dataGridView2.DataSource = null;
                    dataGridView2.Columns.Clear();
                    dataGridView2.Rows.Clear();

                    dataGridView1.DataSource = null;
                    dataGridView1.Columns.Clear();
                    dataGridView1.Rows.Clear();

                    dataGridView4.Rows.Clear();
                    dataGridView5.Rows.Clear();

                    gridControl2.DataSource = null;
                    gridView2.Columns.Clear();
                    gridView2.ColumnPanelRowHeight = 20;

                    gridControl3.DataSource = null;
                    gridView3.Columns.Clear();
                    gridView3.ColumnPanelRowHeight = 20;

                    gridControl1.DataSource = null;
                    gridView1.Columns.Clear();
                    gridView1.ColumnPanelRowHeight = 10;

                    #region После авторизации проверяем, что нам требуется показывать

                    if (toolStripStatusLabel2.Text != "Проект (загрузка данных, заполнение показателей)")
                    {
                        toolStripStatusLabel2.Text = "Проект (загрузка данных, заполнение показателей)";
                    }
                    button2.Enabled = false;
                    button3.Enabled = false;
                    barButtonItem11.Enabled = false;
                    barSubItem7.Enabled = true;
                    barButtonItem8.Enabled = false;
                    barButtonItem10.Enabled = false;
                    barButtonItem13.Enabled = false;
                    
                    #endregion
                }
            }
            catch (Exception load)
            {
                MessageBox.Show("Произошла ошибка. Код ошибки указан ниже.\n" + load, "ФАНЗ", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop);
                return;
            }
        }
        #endregion
        #region Кнопка "Вернуть в работу"
        private void barButtonItem13_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            toolStripStatusLabel2.Text = "В работе (на корректировке)";
            barSubItem1.Enabled = true;
            barSubItem10.Enabled = false;
            barSubItem11.Enabled = false;
            barButtonItem19.Enabled = false;          
            barButtonItem13.Enabled = false;
            button17.Enabled = false;
            button21.Enabled = false;
        }
        #endregion
        #region Кнопка "Корректировка показателей"

        private void edit_control()
        {
            List<string> list = new List<string>();
            DialogResult res;

            if (textBox2.Text == "")
            {
                list.AddRange(new String[] { "МСОС" });
            }

            if (textBox4.Text == "")
            {
                list.AddRange(new String[] { "КОСОС" });
            }

            if (textBox3.Text == "")
            {
                list.AddRange(new String[] { "КАЛ" });
            }

            if (textBox5.Text == "")
            {
                list.AddRange(new String[] { "КБЛ" });
            }

            if (textBox1.Text == "")
            {
                list.AddRange(new String[] { "КТЛ" });
            }

            if (textBox16.Text == "")
            {
                list.AddRange(new String[] { "КИПН" });
            }

            if (textBox8.Text == "")
            {
                list.AddRange(new String[] { "ДСОСЗ" });
            }

            if (textBox18.Text == "")
            {
                list.AddRange(new String[] { "КФЛ" });
            }

            if (textBox17.Text == "")
            {
                list.AddRange(new String[] { "СДВ" });
            }

            if (textBox7.Text == "")
            {
                list.AddRange(new String[] { "КМСК" });
            }

            if (textBox19.Text == "")
                list.AddRange(new String[] { "РПВП" });
            if (textBox20.Text == "")
                list.AddRange(new String[] { "ROE" });
            if (textBox9.Text == "")
                list.AddRange(new String[] { "ОДЗ" });
            if (textBox22.Text == "")
                list.AddRange(new String[] { "ОЗ" });
            if (textBox11.Text == "")
                list.AddRange(new String[] { "ФО" });
            if (textBox21.Text == "")
                list.AddRange(new String[] { "ОСК" });
            if (textBox6.Text == "")
                list.AddRange(new String[] { "ОСовК" });
            if (textBox12.Text == "")
                list.AddRange(new String[] { "БА" });
            if (textBox13.Text == "")
                list.AddRange(new String[] { "УК" });
            if (textBox14.Text == "")
                list.AddRange(new String[] { "БП" });
            if (textBox15.Text == "")
                list.AddRange(new String[] { "ФОТ" });
            if (textBox25.Text == "")
                list.AddRange(new String[] { "СебФОТ" });
            if (textBox26.Text == "")
                list.AddRange(new String[] { "Н.пр" });
            if (textBox27.Text == "")
                list.AddRange(new String[] { "ОЧП" });
            if (textBox24.Text == "")
                list.AddRange(new String[] { "КрУр" });

            if (list.Count != 0)
            {
                string output = string.Join(", ", list);
                res = MessageBox.Show("Не заполнены обязательные поля.\nВывести список пустых полей?", "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (res == System.Windows.Forms.DialogResult.Yes)
                {
                    MessageBox.Show(output, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                else return;
            }
        }
        #endregion       
        #endregion
        //---------------------------------------------------------------------------
        #region Алгоритмы корректировки решения!

        #region По Активу
        private void Aktiv(DataGridView grid1, GridView grid2)
        {
            double ves = 0;
            double TA = 0;
            double BA = 0;
            double new_TA = 0;
            double new_BA = 0;

            if (grid1.RowCount != 0 && grid2.RowCount != 0)
            {
                double aktiv = Convert.ToDouble(textBox12.Text);
                TA = Convert.ToDouble(dataGridView4[0, 0].Value);
                BA = Convert.ToDouble(dataGridView4[5, 0].Value);

                if (aktiv != (TA + BA))
                {
                    if (aktiv > (TA + BA))
                    {
                        ves = aktiv - (TA + BA);
                        new_TA = Math.Round((ves * TA) / (TA + BA));
                        new_BA = Math.Round((ves * BA) / (TA + BA));
                        TA = TA + new_TA;
                        BA = BA + new_BA;
                        dataGridView4[0, 0].Value = TA;
                        dataGridView4[5, 0].Value = BA;
                    }
                    else
                    {
                        ves = (TA + BA) - aktiv;
                        new_TA = Math.Round((ves * TA) / (TA + BA));
                        new_BA = Math.Round((ves * BA) / (TA + BA));
                        TA = TA - new_TA;
                        BA = BA - new_BA;
                        dataGridView4[0, 0].Value = TA;
                        dataGridView4[5, 0].Value = BA;
                    }
                }
            }
        }
        #endregion
        #region Раздел I Внеоборотные активы
        private void Razdel_1(DataGridView grid3, GridView grid4)
        {
            double ves = 0;
            double BA = 0;
            double os_sr = 0;
            double PrBa = 0;
            double new_os_sr = 0;
            double new_PrBa = 0;

            if (grid3.RowCount != 0 && grid4.RowCount != 0)
            {
                BA = Convert.ToDouble(dataGridView4[5, 0].Value);
                os_sr = Convert.ToDouble(dataGridView4[11, 0].Value);
                PrBa = Convert.ToDouble(dataGridView4[13, 0].Value);

                if (BA != (os_sr + PrBa))
                {
                    if (BA > (os_sr + PrBa))
                    {
                        ves = BA - (os_sr + PrBa);
                        new_os_sr = Math.Round((ves * os_sr) / (os_sr + PrBa));
                        new_PrBa = Math.Round((ves * PrBa) / (os_sr + PrBa));
                        os_sr = os_sr + new_os_sr;
                        PrBa = PrBa + new_PrBa;
                        dataGridView4[11, 0].Value = os_sr;
                        dataGridView4[13, 0].Value = PrBa;
                    }
                    else
                    {
                        ves = (os_sr + PrBa) - BA;
                        new_os_sr = Math.Round((ves * os_sr) / (os_sr + PrBa));
                        new_PrBa = Math.Round((ves * PrBa) / (os_sr + PrBa));
                        os_sr = os_sr - new_os_sr;
                        PrBa = PrBa - new_PrBa;
                        dataGridView4[11, 0].Value = os_sr;
                        dataGridView4[13, 0].Value = PrBa;
                    }
                }
            }
        }
        #endregion
        #region Раздел II Оборотные активы
        private void Razdel_2(DataGridView grid3, GridView grid4)
        {
            double ves = 0;
            double TA = 0;
            double zz = 0;
            double dz = 0;
            double ds = 0;
            double PrTA = 0;
            double summ = 0;

            double new_zz = 0;
            double new_dz = 0;
            double new_ds = 0;
            double new_PrTA = 0;

            if (grid3.RowCount != 0 && grid4.RowCount != 0)
            {
                TA = Convert.ToDouble(dataGridView4[0, 0].Value);
                zz = Convert.ToDouble(dataGridView4[6, 0].Value);
                dz = Convert.ToDouble(dataGridView4[4, 0].Value);
                ds = Convert.ToDouble(dataGridView4[3, 0].Value);
                PrTA = Convert.ToDouble(dataGridView4[12, 0].Value);
                summ = zz + dz + ds + PrTA;

                if (TA != (summ))
                {
                    if (TA > (summ))
                    {
                        ves = TA - (summ);
                        new_zz = Math.Round((ves * zz) / (summ));
                        new_dz = Math.Round((ves * dz) / (summ));
                        new_ds = Math.Round((ves * ds) / (summ));
                        new_PrTA = Math.Round((ves * PrTA) / (summ));
                        zz = zz + new_zz;
                        dz = dz + new_dz;
                        ds = ds + new_ds;
                        PrTA = PrTA + new_PrTA;

                        dataGridView4[6, 0].Value = zz;
                        dataGridView4[4, 0].Value = dz;
                        dataGridView4[3, 0].Value = ds;
                        dataGridView4[12, 0].Value = PrTA;
                    }
                    else
                    {
                        ves = (summ) - TA;
                        new_zz = Math.Round((ves * zz) / (summ));
                        new_dz = Math.Round((ves * dz) / (summ));
                        new_ds = Math.Round((ves * ds) / (summ));
                        new_PrTA = Math.Round((ves * PrTA) / (summ));
                        zz = zz - new_zz;
                        dz = dz - new_dz;
                        ds = ds - new_ds;
                        PrTA = PrTA - new_PrTA;

                        dataGridView4[6, 0].Value = zz;
                        dataGridView4[4, 0].Value = dz;
                        dataGridView4[3, 0].Value = ds;
                        dataGridView4[12, 0].Value = PrTA;
                    }
                }
            }
        }
        #endregion
        #region Равество актива и пассива. ПАССИВ (БП)
        private void BP(DataGridView grid3, GridView grid4)
        {
            double ves = 0;
            double TA = 0;
            double BA = 0;
            double artiv = 0;
            double passiv = 0;
            double sk = 0;
            double pk = 0;
            double new_sk = 0;
            double new_pk = 0;

            if (grid3.RowCount != 0 && grid4.RowCount != 0)
            {
                TA = Convert.ToDouble(dataGridView4[0, 0].Value);
                BA = Convert.ToDouble(dataGridView4[5, 0].Value);
                sk = Convert.ToDouble(dataGridView4[8, 0].Value);
                pk = Convert.ToDouble(dataGridView4[7, 0].Value);
                artiv = TA + BA;
                passiv = sk + pk;

                if (artiv != passiv)
                {
                    if (artiv > passiv)
                    {
                        ves = artiv - passiv;
                        new_sk = Math.Round((ves * sk) / (passiv));
                        new_pk = Math.Round((ves * pk) / (passiv));
                        sk += new_sk;
                        pk += new_pk;
                        dataGridView4[8, 0].Value = sk;
                        dataGridView4[7, 0].Value = pk;
                    }
                    else
                    {
                        ves = passiv - artiv;
                        new_sk = Math.Round((ves * sk) / (passiv));
                        new_pk = Math.Round((ves * pk) / (passiv));
                        sk -= new_sk;
                        pk -= new_pk;
                        dataGridView4[8, 0].Value = sk;
                        dataGridView4[7, 0].Value = pk;
                    }
                }
            }

        }
        #endregion
        #region Проверка СК
        private void SK(DataGridView grid3, GridView grid4)
        {
            double ves = 0;
            double sk = 0;
            double uk = 0;
            double fr = 0;
            double NerPr = 0;
            double summ = 0;
            double new_uk = 0;
            double new_fr = 0;
            double new_NerPr = 0;

            if (grid3.RowCount != 0 && grid4.RowCount != 0 && textBox13.Text != "")
            {
                sk = Convert.ToDouble(dataGridView4[8, 0].Value);
                uk = Convert.ToDouble(textBox13.Text);
                fr = Convert.ToDouble(dataGridView4[9, 0].Value);
                NerPr = Convert.ToDouble(dataGridView4[24, 0].Value);
                summ = uk + fr + NerPr;

                if (sk != summ)
                {
                    if (sk > summ)
                    {
                        ves = sk - summ;
                        new_uk = Math.Round((ves * uk) / (summ));
                        new_fr = Math.Round((ves * fr) / (summ));
                        new_NerPr = Math.Round((ves * NerPr) / (summ));

                        uk += new_uk;
                        fr += new_fr;
                        NerPr += new_NerPr;

                        dataGridView4[9, 0].Value = fr;
                        textBox13.Text = uk.ToString();
                        dataGridView4[24, 0].Value = NerPr;
                    }
                    else
                    {
                        ves = summ - sk;
                        new_uk = Math.Round((ves * uk) / (summ));
                        new_fr = Math.Round((ves * fr) / (summ));
                        new_NerPr = Math.Round((ves * NerPr) / (summ));

                        uk -= new_uk;
                        fr -= new_fr;
                        NerPr -= new_NerPr;

                        dataGridView4[9, 0].Value = fr;
                        textBox13.Text = uk.ToString();
                        dataGridView4[24, 0].Value = NerPr;
                    }
                }
            }
        }
        #endregion
        #region Проверка ПК
        private void PK(DataGridView grid3, GridView grid4)
        {
            double ves = 0;
            double pk = 0;
            double dp = 0;
            double kp = 0;
            double new_dp = 0;
            double new_kp = 0;
            double summ = 0;

            if (grid3.RowCount != 0 && grid4.RowCount != 0)
            {
                pk = Convert.ToDouble(dataGridView4[7, 0].Value);
                dp = Convert.ToDouble(dataGridView4[10, 0].Value);
                kp = Convert.ToDouble(dataGridView4[1, 0].Value);
                summ = dp + kp;

                if (pk != summ)
                {
                    if (pk > summ)
                    {
                        ves = pk - summ;
                        new_dp = Math.Round((ves * dp) / (summ));
                        new_kp = Math.Round((ves * kp) / (summ));

                        dp += new_dp;
                        kp += new_kp;

                        dataGridView4[1, 0].Value = kp;
                        dataGridView4[10, 0].Value = dp;
                    }
                    else
                    {
                        ves = summ - pk;
                        new_dp = Math.Round((ves * dp) / (summ));
                        new_kp = Math.Round((ves * kp) / (summ));

                        dp -= new_dp;
                        kp -= new_kp;

                        dataGridView4[1, 0].Value = kp;
                        dataGridView4[10, 0].Value = dp;
                    }
                }
            }
        }
        #endregion
        #region Первый пункт по ЧП
        private void One(GridView grid3, DataGridView grid4)
        {
            double NerPr = 0;
            double ochp = 0;
            double chp = 0;
            double ba = 0;
            double Rsovk = 0;

            if (textBox27.Text != "" && grid3.RowCount != 0 && grid4.RowCount != 0)
            {
                NerPr = Convert.ToDouble(dataGridView4[24, 0].Value);
                if (NerPr != 0)
                {
                    ochp = Convert.ToDouble(textBox27.Text);
                    if (ochp != 0)
                    {
                        chp = Math.Round(NerPr / ochp);
                        dataGridView4[20, 0].Value = chp;
                    }
                    else
                    {
                        MessageBox.Show("Показатель «Соотношение нераспределенной и чистой прибыли» равен 0.", "ФАНЗ");
                        return;
                    }
                }
                else
                {
                    ba = Convert.ToDouble(dataGridView4[5, 0].Value);
                    Rsovk = Convert.ToDouble(dataGridView4[21, 0].Value);
                    chp = Math.Round(ba * Rsovk);
                    dataGridView4[20, 0].Value = chp;
                }
            }
        }
        #endregion
        #region Второй пункт по Пр.до.нал
        private void Two(GridView grid3, DataGridView grid4)
        {
            double NalPr = 0;
            double PrDoNal = 0;
            double chp = 0;

            if (textBox26.Text != "" && grid3.RowCount != 0 && grid4.RowCount != 0)
            {
                chp = Convert.ToDouble(dataGridView4[20, 0].Value);
                NalPr = Convert.ToDouble(textBox26.Text);
                if (NalPr != 0)
                {
                    PrDoNal = Math.Round(chp / (1 - NalPr));
                    dataGridView4[22, 0].Value = PrDoNal;
                }
                else
                {
                    MessageBox.Show("Показатель «Налог на прибыль» равен 0.", "ФАНЗ");
                    return;
                }
            }
        }
        #endregion
        #region Третий пункт Неравенства
        private void Three(GridView grid3, DataGridView grid4)
        {
            double PrD = 0;
            double PrR = 0;
            double BA = 0;
            double PrDoNal = 0;
            double PrProd = 0;

            if (grid3.RowCount != 0 && grid4.RowCount != 0)
            {
                PrD = Convert.ToDouble(dataGridView4[18, 0].Value);
                PrR = Convert.ToDouble(dataGridView4[19, 0].Value);
                BA = Convert.ToDouble(dataGridView4[5, 0].Value);
                PrDoNal = Convert.ToDouble(dataGridView4[22, 0].Value);

                if (PrD <= 0.25 * BA && PrR <= 0.25 * BA)
                {
                    PrProd = Math.Round((PrDoNal - PrD) + PrR);
                    dataGridView4[17, 0].Value = PrProd;
                }
                else
                {
                    PrD = 0;
                    dataGridView4[18, 0].Value = PrD;
                    PrR = 0;
                    dataGridView4[19, 0].Value = PrR;
                    PrProd = PrDoNal;
                    dataGridView4[17, 0].Value = PrProd;
                }
            }
        }
        #endregion
        #region Четвертый пункт по Себ.прод
        private void Four(GridView grid3, DataGridView grid4)
        {
            double SebFot = 0;
            double FO = 0;
            double SebProd = 0;

            if (grid3.RowCount != 0 && grid4.RowCount != 0 && textBox15.Text != "" && textBox25.Text != "")
            {
                SebFot = Convert.ToDouble(textBox25.Text);
                FO = Convert.ToDouble(textBox15.Text);
                SebProd = Math.Round(SebFot * FO);
                dataGridView4[23, 0].Value = SebProd;
            }
        }
        #endregion
        #region Пятый пункт по КрУр
        private void Five(GridView grid3, DataGridView grid4)
        {
            double Krur = 0;
            double dolya = 0;
            double SebProd = 0;

            if (grid3.RowCount != 0 && grid4.RowCount != 0 && textBox24.Text != "")
            {
                dolya = Convert.ToDouble(textBox24.Text);
                SebProd = Convert.ToDouble(dataGridView4[23, 0].Value);
                Krur = Math.Round(SebProd * dolya);
                dataGridView4[16, 0].Value = Krur;
            }

        }
        #endregion
        #region Шестой пункт по ВП, ВР
        private void Sex(GridView grid3, DataGridView grid4)
        {
            double Krur = 0;
            double VP = 0;
            double VR = 0;
            double PrProd = 0;
            double SebProd = 0;

            if (grid3.RowCount != 0 && grid4.RowCount != 0)
            {
                Krur = Convert.ToDouble(dataGridView4[16, 0].Value);
                PrProd = Convert.ToDouble(dataGridView4[17, 0].Value);
                SebProd = Convert.ToDouble(dataGridView4[23, 0].Value);
                VP = Math.Round(Krur + PrProd);
                dataGridView4[15, 0].Value = VP;

                VR = Math.Round(VP + SebProd);
                dataGridView4[14, 0].Value = VR;
            }
        }
        #endregion

        private void barButtonItem19_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                #region Алгоритм корректировки таблиц Баланса и Отчета о прибылях и убытках
                if (dataGridView4.RowCount != 0 && gridView2.RowCount != 0)
                {
                    Aktiv(dataGridView4, gridView2);         
                    Razdel_1(dataGridView4, gridView2);        
                    Razdel_2(dataGridView4, gridView2);
                    BP(dataGridView4, gridView2);              
                    SK(dataGridView4, gridView2);              
                    PK(dataGridView4, gridView2);               
                    InsertResult(gridView2, dataGridView4);
                    //--------------------------------------------
                    One(gridView2, dataGridView4);
                    Two(gridView2, dataGridView4);
                    Three(gridView2, dataGridView4);
                    Four(gridView2, dataGridView4);
                    Five(gridView2, dataGridView4);
                    Sex(gridView2, dataGridView4);
                    InsertResult(gridView2, dataGridView4);
                }
                else
                {
                    MessageBox.Show("Алгоритм корректировки вернул ошибку.\nОбратитесь к администратору.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                #endregion
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            barButtonItem19.Enabled = false;
        }
        #endregion
        //---------------------------------------------------------------------------
        #region Заполнение таблиц отклонений            
        #region Загрузка из XML таблицы отклонений
        private void button20_Click(object sender, EventArgs e)
        {
            if (gridView1.RowCount == 0)
            {
                try
                {
                    #region Загрузка в новый грид
                    XmlReader xmlFile;
                    xmlFile = XmlReader.Create(@shablon_3, new XmlReaderSettings());
                    DataSet ds = new DataSet();
                    ds.ReadXml(xmlFile);
                    gridControl1.DataSource = ds.Tables[0];
                    #region Формирование всех свойств и методов для таблицы

                    gridView1.Columns[0].Caption = "№";
                    gridView1.Columns[1].Caption = "Показатель";
                    gridView1.Columns[2].Caption = "Фактические\n значения";
                    gridView1.Columns[3].Caption = "Расчетные\n значения";
                    gridView1.Columns[4].Caption = "Абсолютное\n отклонение";
                    gridView1.Columns[5].Caption = "Относительное\n отклонение, %";

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

                    gridView1.AppearancePrint.HeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
                    gridView1.OptionsPrint.UsePrintStyles = true;
                    gridView1.ColumnPanelRowHeight = 50;

                    DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit Grid_MemoEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit(); // инициализируем MemoEdit с именем “MyGrid_MemoEdit”
                    Grid_MemoEdit.WordWrap = true;
                    gridView1.OptionsView.RowAutoHeight = true;
                    gridView1.Columns[1].ColumnEdit = Grid_MemoEdit;
                    #endregion
                    xmlFile.Dispose();
                    #endregion
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                MessageBox.Show("Таблица содержит данные, требуетя очистка.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            
        }
        #endregion

        #region Заполнение первой таблицы
        private void ExportPokaz()
        {
            try
            {
                #region новый грид
                gridView3.SetRowCellValue(3, gridView3.Columns[3], textBox2.Text);
                gridView3.SetRowCellValue(4, gridView3.Columns[3], textBox4.Text);
                gridView3.SetRowCellValue(5, gridView3.Columns[3], textBox3.Text);
                gridView3.SetRowCellValue(6, gridView3.Columns[3], textBox5.Text);
                gridView3.SetRowCellValue(7, gridView3.Columns[3], textBox1.Text);
                gridView3.SetRowCellValue(8, gridView3.Columns[3], textBox16.Text);
                gridView3.SetRowCellValue(9, gridView3.Columns[3], textBox8.Text);

                gridView3.SetRowCellValue(11, gridView3.Columns[3], textBox18.Text);
                gridView3.SetRowCellValue(12, gridView3.Columns[3], textBox17.Text);
                gridView3.SetRowCellValue(13, gridView3.Columns[3], textBox7.Text);

                gridView3.SetRowCellValue(16, gridView3.Columns[3], Convert.ToDouble(textBox19.Text) * 100);
                gridView3.SetRowCellValue(17, gridView3.Columns[3], Convert.ToDouble(textBox20.Text) * 100);

                gridView3.SetRowCellValue(19, gridView3.Columns[3], textBox9.Text);
                gridView3.SetRowCellValue(20, gridView3.Columns[3], textBox22.Text);
                gridView3.SetRowCellValue(21, gridView3.Columns[3], textBox11.Text);
                gridView3.SetRowCellValue(22, gridView3.Columns[3], textBox21.Text);
                gridView3.SetRowCellValue(23, gridView3.Columns[3], textBox6.Text);

                gridView3.SetRowCellValue(25, gridView3.Columns[3], textBox12.Text);
                gridView3.SetRowCellValue(26, gridView3.Columns[3], textBox13.Text);
                gridView3.SetRowCellValue(27, gridView3.Columns[3], textBox14.Text);
                gridView3.SetRowCellValue(28, gridView3.Columns[3], textBox15.Text);
                gridView3.SetRowCellValue(29, gridView3.Columns[3], textBox25.Text);

                if (Convert.ToDouble(textBox26.Text) > 1 && Convert.ToDouble(textBox26.Text) < 100)
                {
                    string nalog = Convert.ToString(Convert.ToDouble(textBox26.Text) / 100);
                    gridView3.SetRowCellValue(30, gridView3.Columns[3], nalog);
                }
                else if (Convert.ToDouble(textBox26.Text) > 0 && Convert.ToDouble(textBox26.Text) < 1)
                {
                    gridView3.SetRowCellValue(30, gridView3.Columns[3], textBox26.Text);
                }

                gridView3.SetRowCellValue(31, gridView3.Columns[3], textBox27.Text);
                gridView3.SetRowCellValue(32, gridView3.Columns[3], textBox24.Text);
                #endregion
            }
            catch
            {
                MessageBox.Show("Таблица не загружена", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        private void Raschet()
        {
            try
            {
                #region Показатели для расчета

                double E6 = Convert.ToDouble(dataGridView4[11, 0].Value);
                double E7 = Convert.ToDouble(dataGridView4[13, 0].Value);
                double E8 = Convert.ToDouble(dataGridView4[5, 0].Value);

                double E11 = Convert.ToDouble(dataGridView4[6, 0].Value);
                double E12 = Convert.ToDouble(dataGridView4[4, 0].Value);
                double E13 = Convert.ToDouble(dataGridView4[3, 0].Value);
                double E14 = Convert.ToDouble(dataGridView4[12, 0].Value);
                double E15 = Convert.ToDouble(dataGridView4[0, 0].Value);

                double E17 = E15 + E8;

                double E21 = Convert.ToDouble(gridView2.GetRowCellValue(16, gridView2.Columns[2]));
                double E22 = Convert.ToDouble(dataGridView4[9, 0].Value);
                double E23 = Convert.ToDouble(dataGridView4[24, 0].Value);
                double E24 = Convert.ToDouble(dataGridView4[8, 0].Value);

                double E27 = Convert.ToDouble(dataGridView4[10, 0].Value);
                double E28 = Convert.ToDouble(dataGridView4[1, 0].Value);
                double E29 = Convert.ToDouble(dataGridView4[7, 0].Value);

                double E31 = E24 + E29;

                double E36 = Convert.ToDouble(dataGridView4[14, 0].Value);
                double E37 = Convert.ToDouble(dataGridView4[23, 0].Value);
                double E38 = Convert.ToDouble(dataGridView4[15, 0].Value);
                double E39 = Convert.ToDouble(dataGridView4[16, 0].Value);
                double E40 = Convert.ToDouble(dataGridView4[17, 0].Value);
                double E41 = Convert.ToDouble(dataGridView4[18, 0].Value);
                double E42 = Convert.ToDouble(dataGridView4[19, 0].Value);
                double E43 = Convert.ToDouble(dataGridView4[22, 0].Value);
                double E44 = Convert.ToDouble(dataGridView4[20, 0].Value);
                double E45 = Convert.ToDouble(dataGridView4[24, 0].Value);

                #endregion                
                #region Сами расчеты в новом гриде
                double coc = Math.Round(E15 - E28, 2);
                double tfp = Math.Round(coc - E13);
                gridView3.SetRowCellValue(2, gridView3.Columns[4], coc);
                gridView3.SetRowCellValue(1, gridView3.Columns[4], tfp);

                if (coc != 0)
                {
                    double mcoc = Math.Round(E13 / coc, 2);
                    gridView3.SetRowCellValue(3, gridView3.Columns[4], mcoc);
                }
                if (E15 != 0)
                {
                    double kococ = Math.Round(coc / E15, 2);
                    gridView3.SetRowCellValue(4, gridView3.Columns[4], kococ);
                }
                if (E28 != 0)
                {
                    double kal = Math.Round(E13 / E28, 2);
                    gridView3.SetRowCellValue(5, gridView3.Columns[4], kal);
                    double kbl = Math.Round((E12 + E13 + E14) / E28, 2);
                    gridView3.SetRowCellValue(6, gridView3.Columns[4], kbl);
                    double ktl = Math.Round(E15 / E28, 2);
                    gridView3.SetRowCellValue(7, gridView3.Columns[4], ktl);
                }
                if (E17 != 0)
                {
                    double kipn = Math.Round((E6 + E11) / E17, 2);
                    gridView3.SetRowCellValue(8, gridView3.Columns[4], kipn);
                }
                if (E11 != 0)
                {
                    double dcosz = Math.Round(coc / E11, 2);
                    gridView3.SetRowCellValue(9, gridView3.Columns[4], dcosz);
                }
                if (E24 != 0)
                {
                    double kfl = Math.Round(E29 / E24, 2);
                    gridView3.SetRowCellValue(11, gridView3.Columns[4], kfl);
                    double kmsk = Math.Round(coc / E24, 2);
                    gridView3.SetRowCellValue(13, gridView3.Columns[4], kmsk);
                }
                if (E8 != 0)
                {
                    double sdv = Math.Round(E27 / E8, 2);
                    gridView3.SetRowCellValue(12, gridView3.Columns[4], sdv);
                }
                if (E8 != 31)
                {
                    double roa = Math.Round((E44 / E31) * 100, 1);
                    gridView3.SetRowCellValue(15, gridView3.Columns[4], roa);
                }
                if (E36 != 0)
                {
                    double rpvp = Math.Round((E38 / E36) * 100, 1);
                    gridView3.SetRowCellValue(16, gridView3.Columns[4], rpvp);
                }
                if (E24 != 0)
                {
                    double roe = Math.Round((E44 / E24) * 100, 1);
                    gridView3.SetRowCellValue(17, gridView3.Columns[4], roe);
                }
                if (E12 != 0)
                {
                    double odz = Math.Round(E36 / E12, 2);
                    gridView3.SetRowCellValue(19, gridView3.Columns[4], odz);
                }
                if (E11 != 0)
                {
                    double oz = Math.Round(E37 / E11, 2);
                    gridView3.SetRowCellValue(20, gridView3.Columns[4], oz);
                }
                if (E6 != 0)
                {
                    double fo = Math.Round(E36 / E6, 2);
                    gridView3.SetRowCellValue(21, gridView3.Columns[4], fo);
                }
                if (E24 != 0)
                {
                    double osk = Math.Round(E36 / E24, 2);
                    gridView3.SetRowCellValue(22, gridView3.Columns[4], osk);
                }
                if (E31 != 0)
                {
                    double osvk = Math.Round(E36 / E31, 2);
                    gridView3.SetRowCellValue(23, gridView3.Columns[4], osvk);
                }
                double ba = Math.Round(E17, 2);
                double yk = Math.Round(E21, 2);
                double bp = Math.Round(E31, 2);
                double fot = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(28, gridView3.Columns[3])), 2);
                gridView3.SetRowCellValue(25, gridView3.Columns[4], ba);
                gridView3.SetRowCellValue(26, gridView3.Columns[4], yk);
                gridView3.SetRowCellValue(27, gridView3.Columns[4], bp);
                gridView3.SetRowCellValue(28, gridView3.Columns[4], fot);
                if (fot != 0)
                {
                    double sebfot = Math.Round(E37 / fot, 2);
                    gridView3.SetRowCellValue(29, gridView3.Columns[4], sebfot);
                }
                if (E43 != 0)
                {
                    double Npr = Math.Round(1 - (E44 / E43), 2);
                    gridView3.SetRowCellValue(30, gridView3.Columns[4], Npr);
                }
                if (E44 != 0)
                {
                    double ochp = Math.Round((E45 / E44), 2);
                    gridView3.SetRowCellValue(31, gridView3.Columns[4], ochp);
                }
                if (E37 != 0)
                {
                    double Krur = Math.Round(E39 / E37, 2);
                    gridView3.SetRowCellValue(32, gridView3.Columns[4], Krur);
                }

                #endregion
            }
            catch (Exception exx)
            {
                MessageBox.Show(exx.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        private void Abs_Otkl()
        {
            try
            {
                #region Расчеты в новом гриде
                double mcoc = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(3, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(3, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(3, gridView3.Columns[5], mcoc);
                double kococ = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(4, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(4, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(4, gridView3.Columns[5], kococ);
                double kal = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(5, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(5, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(5, gridView3.Columns[5], kal);
                double kbl = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(6, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(6, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(6, gridView3.Columns[5], kbl);
                double ktl = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(7, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(7, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(7, gridView3.Columns[5], ktl);
                double kipn = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(8, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(8, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(8, gridView3.Columns[5], kipn);
                double dcosz = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(9, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(9, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(9, gridView3.Columns[5], dcosz);

                double kfl = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(11, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(11, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(11, gridView3.Columns[5], kfl);
                double sdv = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(12, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(12, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(12, gridView3.Columns[5], sdv);
                double kmsk = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(13, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(13, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(13, gridView3.Columns[5], kmsk);

                double rpvp = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(16, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(16, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(16, gridView3.Columns[5], rpvp);
                double roe = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(17, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(17, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(17, gridView3.Columns[5], roe);

                double odz = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(19, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(19, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(19, gridView3.Columns[5], odz);
                double oz = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(20, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(20, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(20, gridView3.Columns[5], oz);
                double fo = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(21, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(21, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(21, gridView3.Columns[5], fo);
                double osk = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(22, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(22, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(22, gridView3.Columns[5], osk);
                double osvk = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(23, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(23, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(23, gridView3.Columns[5], osvk);

                double ba = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(25, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(25, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(25, gridView3.Columns[5], ba);
                double yk = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(26, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(26, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(26, gridView3.Columns[5], yk);
                double bp = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(27, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(27, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(27, gridView3.Columns[5], bp);
                double fot = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(28, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(28, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(28, gridView3.Columns[5], fot);
                double cebfot = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(29, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(29, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(29, gridView3.Columns[5], cebfot);
                double npr = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(30, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(30, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(30, gridView3.Columns[5], npr);
                double ochp = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(31, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(31, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(31, gridView3.Columns[5], ochp);
                double krur = Math.Round(Convert.ToDouble(gridView3.GetRowCellValue(32, gridView3.Columns[3])) - Convert.ToDouble(gridView3.GetRowCellValue(32, gridView3.Columns[4])), 2);
                gridView3.SetRowCellValue(32, gridView3.Columns[5], krur);
                #endregion
            }
            catch (Exception exx)
            {
                MessageBox.Show(exx.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        private void Otnos_Otkl()
        {
            try
            {                
                #region Расчеты в новом гриде
                double mcoc = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(3, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(3, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(3, gridView3.Columns[6], mcoc);
                double kococ = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(4, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(4, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(4, gridView3.Columns[6], kococ);
                double kal = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(5, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(5, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(5, gridView3.Columns[6], kal);
                double kbl = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(6, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(6, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(6, gridView3.Columns[6], kbl);
                double ktl = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(7, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(7, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(7, gridView3.Columns[6], ktl);
                double kipn = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(8, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(8, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(8, gridView3.Columns[6], kipn);
                double dcosz = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(9, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(9, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(9, gridView3.Columns[6], dcosz);

                double kfl = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(11, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(11, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(11, gridView3.Columns[6], kfl);
                double sdv = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(12, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(12, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(12, gridView3.Columns[6], sdv);
                double kmsk = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(13, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(13, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(13, gridView3.Columns[6], kmsk);

                double rpvp = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(16, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(16, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(16, gridView3.Columns[6], rpvp);
                double roe = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(17, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(17, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(17, gridView3.Columns[6], roe);

                double odz = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(19, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(19, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(19, gridView3.Columns[6], odz);
                double oz = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(20, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(20, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(20, gridView3.Columns[6], oz);
                double fo = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(21, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(21, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(21, gridView3.Columns[6], fo);
                double osk = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(22, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(22, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(22, gridView3.Columns[6], osk);
                double osvk = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(23, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(23, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(23, gridView3.Columns[6], osvk);

                double ba = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(25, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(25, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(25, gridView3.Columns[6], ba);
                double yk = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(26, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(26, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(26, gridView3.Columns[6], yk);
                double bp = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(27, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(26, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(27, gridView3.Columns[6], bp);
                double fot = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(28, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(28, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(28, gridView3.Columns[6], fot);
                double cebfot = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(29, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(29, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(29, gridView3.Columns[6], cebfot);
                double npr = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(30, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(30, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(30, gridView3.Columns[6], npr);
                double ochp = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(31, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(31, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(31, gridView3.Columns[6], ochp);
                double krur = Math.Round((Convert.ToDouble(gridView3.GetRowCellValue(32, gridView3.Columns[5])) / Convert.ToDouble(gridView3.GetRowCellValue(32, gridView3.Columns[3]))) * 100, 1);
                gridView3.SetRowCellValue(32, gridView3.Columns[6], krur);
                #endregion
            }
            catch (Exception exx)
            {
                MessageBox.Show(exx.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        #region Кнопка расчетов в таблице расчетов
        private void button17_Click(object sender, EventArgs e)
        {
            if (gridView2.RowCount != 0 && gridView3.RowCount != 0)
            {
                ExportPokaz();
                Raschet();
                Abs_Otkl();
                Otnos_Otkl();
            }
            else
            {
                MessageBox.Show("Не содержит необходимые данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion

        private void Raschet_2()
        {
            #region Показатели для расчета  в новом гриде
            try
            { 
                #region Показатели для gridview
                double D36 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(30, gridView2.Columns[1])), 2);
                double D38 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(32, gridView2.Columns[1])), 2);
                double D44 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(38, gridView2.Columns[1])), 2);
                double D45 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(39, gridView2.Columns[1])), 2);
                double D6 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(1, gridView2.Columns[1])), 2);
                double D7 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(2, gridView2.Columns[1])), 2);
                double D11 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(6, gridView2.Columns[1])), 2);
                double D12 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(7, gridView2.Columns[1])), 2);
                double D13 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(8, gridView2.Columns[1])), 2);
                double D14 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(9, gridView2.Columns[1])), 2);
                double D21 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(16, gridView2.Columns[1])), 2);
                double D22 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(17, gridView2.Columns[1])), 2);
                double D27 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(22, gridView2.Columns[1])), 2);
                double D28 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(23, gridView2.Columns[1])), 2);

                double E36 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(30, gridView2.Columns[2])), 2);
                double E38 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(32, gridView2.Columns[2])), 2);
                double E44 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(38, gridView2.Columns[2])), 2);
                double E45 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(39, gridView2.Columns[2])), 2);
                double E6 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(1, gridView2.Columns[2])), 2);
                double E7 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(2, gridView2.Columns[2])), 2);
                double E11 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(6, gridView2.Columns[2])), 2);
                double E12 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(7, gridView2.Columns[2])), 2);

                double E13 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(8, gridView2.Columns[2])), 2);
                double E14 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(9, gridView2.Columns[2])), 2);
                double E21 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(16, gridView2.Columns[2])), 2);

                double E22 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(17, gridView2.Columns[2])), 2); //!!!!!!!!!!!!!
                double E27 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(22, gridView2.Columns[2])), 2);
                double E28 = Math.Round(Convert.ToDouble(gridView2.GetRowCellValue(23, gridView2.Columns[2])), 2);

                gridView1.SetRowCellValue(0, gridView1.Columns[2], D36);
                gridView1.SetRowCellValue(1, gridView1.Columns[2], D38);
                gridView1.SetRowCellValue(2, gridView1.Columns[2], D44);
                gridView1.SetRowCellValue(3, gridView1.Columns[2], D45);
                gridView1.SetRowCellValue(4, gridView1.Columns[2], D6);
                gridView1.SetRowCellValue(5, gridView1.Columns[2], D7);
                gridView1.SetRowCellValue(6, gridView1.Columns[2], D11);
                gridView1.SetRowCellValue(7, gridView1.Columns[2], D12);
                gridView1.SetRowCellValue(8, gridView1.Columns[2], D13);
                gridView1.SetRowCellValue(9, gridView1.Columns[2], D14);
                gridView1.SetRowCellValue(10, gridView1.Columns[2], D21);
                gridView1.SetRowCellValue(11, gridView1.Columns[2], D22);
                gridView1.SetRowCellValue(12, gridView1.Columns[2], D27);
                gridView1.SetRowCellValue(13, gridView1.Columns[2], D28);

                gridView1.SetRowCellValue(0, gridView1.Columns[3], E36);
                gridView1.SetRowCellValue(1, gridView1.Columns[3], E38);
                gridView1.SetRowCellValue(2, gridView1.Columns[3], E44);
                gridView1.SetRowCellValue(3, gridView1.Columns[3], E45);
                gridView1.SetRowCellValue(4, gridView1.Columns[3], E6);
                gridView1.SetRowCellValue(5, gridView1.Columns[3], E7);
                gridView1.SetRowCellValue(6, gridView1.Columns[3], E11);
                gridView1.SetRowCellValue(7, gridView1.Columns[3], E12);
                gridView1.SetRowCellValue(8, gridView1.Columns[3], E13);
                gridView1.SetRowCellValue(9, gridView1.Columns[3], E14);
                gridView1.SetRowCellValue(10, gridView1.Columns[3], E21);
                gridView1.SetRowCellValue(11, gridView1.Columns[3], E22);
                gridView1.SetRowCellValue(12, gridView1.Columns[3], E27);
                gridView1.SetRowCellValue(13, gridView1.Columns[3], E28);
                #endregion
            }
            catch (Exception exx)
            {
                MessageBox.Show(exx.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            #endregion
        }
        private void Absolute_2()
        {
            try
            {
                #region Показатели абсолюта для gridview
                double e1 = Math.Round(Convert.ToDouble(gridView1.GetRowCellValue(0, gridView1.Columns[2])) - Convert.ToDouble(gridView1.GetRowCellValue(0, gridView1.Columns[3])), 2);
                gridView1.SetRowCellValue(0, gridView1.Columns[4], e1);
                double e2 = Math.Round(Convert.ToDouble(gridView1.GetRowCellValue(1, gridView1.Columns[2])) - Convert.ToDouble(gridView1.GetRowCellValue(1, gridView1.Columns[3])), 2);
                gridView1.SetRowCellValue(1, gridView1.Columns[4], e2);
                double e3 = Math.Round(Convert.ToDouble(gridView1.GetRowCellValue(2, gridView1.Columns[2])) - Convert.ToDouble(gridView1.GetRowCellValue(2, gridView1.Columns[3])), 2);
                gridView1.SetRowCellValue(2, gridView1.Columns[4], e3);
                double e4 = Math.Round(Convert.ToDouble(gridView1.GetRowCellValue(3, gridView1.Columns[2])) - Convert.ToDouble(gridView1.GetRowCellValue(3, gridView1.Columns[3])), 2);
                gridView1.SetRowCellValue(3, gridView1.Columns[4], e4);
                double e5 = Math.Round(Convert.ToDouble(gridView1.GetRowCellValue(4, gridView1.Columns[2])) - Convert.ToDouble(gridView1.GetRowCellValue(4, gridView1.Columns[3])), 2);
                gridView1.SetRowCellValue(4, gridView1.Columns[4], e5);
                double e6 = Math.Round(Convert.ToDouble(gridView1.GetRowCellValue(5, gridView1.Columns[2])) - Convert.ToDouble(gridView1.GetRowCellValue(5, gridView1.Columns[3])), 2);
                gridView1.SetRowCellValue(5, gridView1.Columns[4], e6);
                double e7 = Math.Round(Convert.ToDouble(gridView1.GetRowCellValue(6, gridView1.Columns[2])) - Convert.ToDouble(gridView1.GetRowCellValue(6, gridView1.Columns[3])), 2);
                gridView1.SetRowCellValue(6, gridView1.Columns[4], e7);
                double e8 = Math.Round(Convert.ToDouble(gridView1.GetRowCellValue(7, gridView1.Columns[2])) - Convert.ToDouble(gridView1.GetRowCellValue(7, gridView1.Columns[3])), 2);
                gridView1.SetRowCellValue(7, gridView1.Columns[4], e8);
                
                double e9 = Math.Round(Convert.ToDouble(gridView1.GetRowCellValue(8, gridView1.Columns[2])) - Convert.ToDouble(gridView1.GetRowCellValue(8, gridView1.Columns[3])), 2);
                gridView1.SetRowCellValue(8, gridView1.Columns[4], e9);
                
                double e10 = Math.Round(Convert.ToDouble(gridView1.GetRowCellValue(9, gridView1.Columns[2])) - Convert.ToDouble(gridView1.GetRowCellValue(9, gridView1.Columns[3])), 2);
                gridView1.SetRowCellValue(9, gridView1.Columns[4], e10);
                double e11 = Math.Round(Convert.ToDouble(gridView1.GetRowCellValue(10, gridView1.Columns[2])) - Convert.ToDouble(gridView1.GetRowCellValue(10, gridView1.Columns[3])), 2);
                gridView1.SetRowCellValue(10, gridView1.Columns[4], e11);
                double e12 = Math.Round(Convert.ToDouble(gridView1.GetRowCellValue(11, gridView1.Columns[2])) - Convert.ToDouble(gridView1.GetRowCellValue(11, gridView1.Columns[3])), 2);
                gridView1.SetRowCellValue(11, gridView1.Columns[4], e12);
                double e14 = Math.Round(Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[2])) - Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[3])), 2);    //12 вместо 13
                gridView1.SetRowCellValue(12, gridView1.Columns[4], e14);
                double e15 = Math.Round(Convert.ToDouble(gridView1.GetRowCellValue(13, gridView1.Columns[2])) - Convert.ToDouble(gridView1.GetRowCellValue(13, gridView1.Columns[3])), 2);    //13 вместо 14
                gridView1.SetRowCellValue(13, gridView1.Columns[4], e15);
                #endregion
            }
            catch (Exception exx)
            {
                MessageBox.Show(exx.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        private void Otnos_2()
        {
            try
            {
                #region Показатели относительные для gridview
                double e1 = Convert.ToDouble(gridView1.GetRowCellValue(0, gridView1.Columns[4]));
                double e2 = Convert.ToDouble(gridView1.GetRowCellValue(1, gridView1.Columns[4]));
                double e3 = Convert.ToDouble(gridView1.GetRowCellValue(2, gridView1.Columns[4]));
                double e4 = Convert.ToDouble(gridView1.GetRowCellValue(3, gridView1.Columns[4]));
                double e5 = Convert.ToDouble(gridView1.GetRowCellValue(4, gridView1.Columns[4]));
                double e6 = Convert.ToDouble(gridView1.GetRowCellValue(5, gridView1.Columns[4]));
                double e7 = Convert.ToDouble(gridView1.GetRowCellValue(6, gridView1.Columns[4]));
                double e8 = Convert.ToDouble(gridView1.GetRowCellValue(7, gridView1.Columns[4]));
                double e9 = Convert.ToDouble(gridView1.GetRowCellValue(8, gridView1.Columns[4]));
                double e10 = Convert.ToDouble(gridView1.GetRowCellValue(9, gridView1.Columns[4]));
                double e11 = Convert.ToDouble(gridView1.GetRowCellValue(10, gridView1.Columns[4]));
                double e12 = Convert.ToDouble(gridView1.GetRowCellValue(11, gridView1.Columns[4]));
                double e14 = Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[4]));
                double e15 = Convert.ToDouble(gridView1.GetRowCellValue(13, gridView1.Columns[4]));

                double c1 = Convert.ToDouble(gridView1.GetRowCellValue(0, gridView1.Columns[2]));
                double c2 = Convert.ToDouble(gridView1.GetRowCellValue(1, gridView1.Columns[2]));
                double c3 = Convert.ToDouble(gridView1.GetRowCellValue(2, gridView1.Columns[2]));
                double c4 = Convert.ToDouble(gridView1.GetRowCellValue(3, gridView1.Columns[2]));
                double c5 = Convert.ToDouble(gridView1.GetRowCellValue(4, gridView1.Columns[2]));
                double c6 = Convert.ToDouble(gridView1.GetRowCellValue(5, gridView1.Columns[2]));
                double c7 = Convert.ToDouble(gridView1.GetRowCellValue(6, gridView1.Columns[2]));
                double c8 = Convert.ToDouble(gridView1.GetRowCellValue(7, gridView1.Columns[2]));
                double c9 = Convert.ToDouble(gridView1.GetRowCellValue(8, gridView1.Columns[2]));
                double c10 = Convert.ToDouble(gridView1.GetRowCellValue(9, gridView1.Columns[2]));
                double c11 = Convert.ToDouble(gridView1.GetRowCellValue(10, gridView1.Columns[2]));
                double c12 = Convert.ToDouble(gridView1.GetRowCellValue(11, gridView1.Columns[2]));
                double c14 = Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[2]));
                double c15 = Convert.ToDouble(gridView1.GetRowCellValue(13, gridView1.Columns[2]));
                #region длинная хрень
                if (c1 == 0)
                {
                    gridView1.SetRowCellValue(0, gridView1.Columns[5], 0);
                }
                else
                {
                    gridView1.SetRowCellValue(0, gridView1.Columns[5], Math.Round((e1 / c1) * 100, 2));
                }
                if (c2 == 0)
                {
                    gridView1.SetRowCellValue(1, gridView1.Columns[5], 0);
                }
                else
                {
                    gridView1.SetRowCellValue(1, gridView1.Columns[5], Math.Round((e2 / c2) * 100, 2));
                }
                if (c3 == 0)
                {
                    gridView1.SetRowCellValue(2, gridView1.Columns[5], 0);
                }
                else
                {
                    gridView1.SetRowCellValue(2, gridView1.Columns[5], Math.Round((e3 / c3) * 100, 2));
                }
                if (c4 == 0)
                {
                    gridView1.SetRowCellValue(3, gridView1.Columns[5], 0);
                }
                else
                {
                    gridView1.SetRowCellValue(3, gridView1.Columns[5], Math.Round((e4 / c4) * 100, 2));
                }
                if (c5 == 0)
                {
                    gridView1.SetRowCellValue(4, gridView1.Columns[5], 0);
                }
                else
                {
                    gridView1.SetRowCellValue(4, gridView1.Columns[5], Math.Round((e5 / c5) * 100, 2));
                }
                if (c6 == 0)
                {
                    gridView1.SetRowCellValue(5, gridView1.Columns[5], 0);
                }
                else
                {
                    gridView1.SetRowCellValue(5, gridView1.Columns[5], Math.Round((e6 / c6) * 100, 2));
                }
                if (c7 == 0)
                {
                    gridView1.SetRowCellValue(6, gridView1.Columns[5], 0);
                }
                else
                {
                    gridView1.SetRowCellValue(6, gridView1.Columns[5], Math.Round((e7 / c7) * 100, 2));
                }
                if (c8 == 0)
                {
                    gridView1.SetRowCellValue(7, gridView1.Columns[5], 0);
                }
                else
                {
                    gridView1.SetRowCellValue(7, gridView1.Columns[5], Math.Round((e8 / c8) * 100, 2));
                }
                if (c9 == 0)
                {
                    gridView1.SetRowCellValue(8, gridView1.Columns[5], 0);
                }
                else
                {
                    gridView1.SetRowCellValue(8, gridView1.Columns[5], Math.Round((e9 / c9) * 100, 2));
                }
                if (c10 == 0)
                {
                    gridView1.SetRowCellValue(9, gridView1.Columns[5], 0);
                }
                else
                {
                    gridView1.SetRowCellValue(9, gridView1.Columns[5], Math.Round((e10 / c10) * 100, 2));
                }
                if (c11 == 0)
                {
                    gridView1.SetRowCellValue(10, gridView1.Columns[5], 0);
                }
                else
                {
                    gridView1.SetRowCellValue(10, gridView1.Columns[5], Math.Round((e11 / c11) * 100, 2));
                }
                if (c12 == 0)
                {
                    gridView1.SetRowCellValue(11, gridView1.Columns[5], 0);
                }
                else
                {
                    gridView1.SetRowCellValue(11, gridView1.Columns[5], Math.Round((e12 / c12) * 100, 2));
                }
                if (c14 == 0)
                {
                    gridView1.SetRowCellValue(12, gridView1.Columns[5], 0);
                }
                else
                {
                    gridView1.SetRowCellValue(12, gridView1.Columns[5], Math.Round((e14 / c14) * 100, 2));
                }
                if (c15 == 0)
                {
                    gridView1.SetRowCellValue(13, gridView1.Columns[5], 0);
                }
                else
                {
                    gridView1.SetRowCellValue(13, gridView1.Columns[5], Math.Round((e15 / c15) * 100, 2));
                }
                #endregion
                #endregion
            }
            catch (Exception exx)
            {
                MessageBox.Show(exx.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

        }
        private void squere()
        {
            try
            {
                #region Расчет расстояния 
                double e1 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(0, gridView1.Columns[4])), 2);
                double e2 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(1, gridView1.Columns[4])), 2);
                double e3 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(2, gridView1.Columns[4])), 2);
                double e4 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(3, gridView1.Columns[4])), 2);
                double e5 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(4, gridView1.Columns[4])), 2);
                double e6 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(5, gridView1.Columns[4])), 2);
                double e7 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(6, gridView1.Columns[4])), 2);
                double e8 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(7, gridView1.Columns[4])), 2);
                double e9 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(8, gridView1.Columns[4])), 2);
                double e10 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(9, gridView1.Columns[4])), 2);
                double e11 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(10, gridView1.Columns[4])), 2);
                double e12 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(11, gridView1.Columns[4])), 2);
                double e14 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[4])), 2);
                double e15 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(13, gridView1.Columns[4])), 2);

                double d1 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(0, gridView1.Columns[3])), 2);
                double d2 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(1, gridView1.Columns[3])), 2);
                double d3 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(2, gridView1.Columns[3])), 2);
                double d4 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(3, gridView1.Columns[3])), 2);
                double d5 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(4, gridView1.Columns[3])), 2);
                double d6 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(5, gridView1.Columns[3])), 2);
                double d7 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(6, gridView1.Columns[3])), 2);
                double d8 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(7, gridView1.Columns[3])), 2);
                double d9 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(8, gridView1.Columns[3])), 2);
                double d10 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(9, gridView1.Columns[3])), 2);
                double d11 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(10, gridView1.Columns[3])), 2);
                double d12 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(11, gridView1.Columns[3])), 2);
                double d14 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[3])), 2);
                double d15 = Math.Pow(Convert.ToDouble(gridView1.GetRowCellValue(13, gridView1.Columns[3])), 2);

                double Esumm = e1 + e2 + e3 + e4 + e5 + e6 + e7 + e8 + e9 + e10 + e11 + e12 + e14 + e15;
                double Dsumm = d1 + d2 + d3 + d4 + d5 + d6 + d7 + d8 + d9 + d10 + d11 + d12 + d14 + d15;
                double Equere = Math.Pow(Esumm, 0.5);
                double Dquere = Math.Pow(Dsumm, 0.5);
                double otklon = Math.Round(100 * (Math.Pow((Esumm), 0.5) / Math.Pow((Dsumm), 0.5)), 2);

                label32.Text = otklon.ToString();
                label36.Text = Convert.ToString(Math.Round((Convert.ToDouble(label34.Text) - otklon), 2));
                InsertToBZ();
                #endregion
            }

            catch (Exception exx)
            {
                MessageBox.Show(exx.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        private void diagramma2()
        {
            try
            {
                double f1 = Convert.ToDouble(gridView1.GetRowCellValue(0, gridView1.Columns[5]));
                double f2 = Convert.ToDouble(gridView1.GetRowCellValue(1, gridView1.Columns[5]));
                double f3 = Convert.ToDouble(gridView1.GetRowCellValue(2, gridView1.Columns[5]));
                double f4 = Convert.ToDouble(gridView1.GetRowCellValue(3, gridView1.Columns[5]));
                double f5 = Convert.ToDouble(gridView1.GetRowCellValue(4, gridView1.Columns[5]));
                double f6 = Convert.ToDouble(gridView1.GetRowCellValue(5, gridView1.Columns[5]));
                double f7 = Convert.ToDouble(gridView1.GetRowCellValue(6, gridView1.Columns[5]));
                double f8 = Convert.ToDouble(gridView1.GetRowCellValue(7, gridView1.Columns[5]));
                double f9 = Convert.ToDouble(gridView1.GetRowCellValue(8, gridView1.Columns[5]));
                double f10 = Convert.ToDouble(gridView1.GetRowCellValue(9, gridView1.Columns[5]));
                double f11 = Convert.ToDouble(gridView1.GetRowCellValue(10, gridView1.Columns[5]));
                double f12 = Convert.ToDouble(gridView1.GetRowCellValue(11, gridView1.Columns[5]));
                double f14 = Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[5]));
                double f15 = Convert.ToDouble(gridView1.GetRowCellValue(13, gridView1.Columns[5]));

                chartControl4.Series[0].Points.Add(new SeriesPoint(1, Convert.ToDouble(f1)));
                chartControl4.Series[0].Points.Add(new SeriesPoint(3, Convert.ToDouble(f2)));
                chartControl4.Series[0].Points.Add(new SeriesPoint(4, Convert.ToDouble(f3)));
                chartControl4.Series[0].Points.Add(new SeriesPoint(5, Convert.ToDouble(f4)));
                chartControl4.Series[0].Points.Add(new SeriesPoint(6, Convert.ToDouble(f5)));
                chartControl4.Series[0].Points.Add(new SeriesPoint(7, Convert.ToDouble(f6)));
                chartControl4.Series[0].Points.Add(new SeriesPoint(8, Convert.ToDouble(f7)));
                chartControl4.Series[0].Points.Add(new SeriesPoint(9, Convert.ToDouble(f8)));
                chartControl4.Series[0].Points.Add(new SeriesPoint(10, Convert.ToDouble(f9)));
                chartControl4.Series[0].Points.Add(new SeriesPoint(11, Convert.ToDouble(f10)));
                chartControl4.Series[0].Points.Add(new SeriesPoint(12, Convert.ToDouble(f11)));
                chartControl4.Series[0].Points.Add(new SeriesPoint(13, Convert.ToDouble(f12)));
                chartControl4.Series[0].Points.Add(new SeriesPoint(13, Convert.ToDouble(f14)));
                chartControl4.Series[0].Points.Add(new SeriesPoint(14, Convert.ToDouble(f15)));
            }
            catch (Exception exx)
            {
                MessageBox.Show(exx.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        #region Кнопка расчета в таблице "Относительных отклонений"
        private void button21_Click(object sender, EventArgs e)
        {
            if (gridView2.RowCount != 0 && gridView1.RowCount != 0)
            {
                Raschet_2();
                Absolute_2();
                Otnos_2();
                squere();
            }
            else
            {
                MessageBox.Show("Не содержит необходимые данные.\nНеобходимо загрузить данные или выполнить расчеты", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion

        #endregion
        #region Диаграмма в 3d
        private void barButtonItem22_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {            
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView2.Columns[2]).ToString() != "")
            {
                chartControl1.Series[0].Points.Clear();
                chartControl1.Series[1].Points.Clear();
                chartControl2.Series[0].Points.Clear();
                chartControl2.Series[1].Points.Clear();
                chartControl2.Annotations.Clear();
                chartControl4.Series[0].Points.Clear();

                string type = chartControl1.Diagram.GetType().ToString();
                string type2d = "DevExpress.XtraCharts.XYDiagram";
                if (type == type2d)
                {
                    chartControl1.Series[0].ChangeView(ViewType.FullStackedBar3D);
                    chartControl1.Series[1].ChangeView(ViewType.FullStackedBar3D);
                    XYDiagram3D diagram1 = (XYDiagram3D)chartControl1.Diagram;
                    diagram1.AxisY.VisualRange.Auto = true;
                    diagram1.AxisX.VisualRange.Auto = true;                   
                    ((XYDiagram3D)chartControl1.Diagram).RuntimeZooming = true;
                    ((XYDiagram3D)chartControl1.Diagram).RuntimeScrolling = true;
                    ((XYDiagram3D)chartControl1.Diagram).RuntimeRotation = false;
                    ((XYDiagram3D)chartControl1.Diagram).ZoomPercent = 100;
                    ((XYDiagram3D)chartControl1.Diagram).AxisY.Label.NumericOptions.Format = NumericFormat.Percent;
                }

                string type1 = chartControl2.Diagram.GetType().ToString();
                string type_2d = "DevExpress.XtraCharts.XYDiagram";
                if (type1 == type_2d)
                {
                    chartControl2.Series[0].ChangeView(ViewType.ManhattanBar);
                    chartControl2.Series[1].ChangeView(ViewType.ManhattanBar);
                    XYDiagram3D diagram2 = (XYDiagram3D)chartControl2.Diagram;
                    diagram2.AxisY.VisualRange.Auto = true;
                    diagram2.AxisX.VisualRange.Auto = true;
                    ((XYDiagram3D)chartControl2.Diagram).RuntimeZooming = true;
                    ((XYDiagram3D)chartControl2.Diagram).RuntimeScrolling = true;
                    ((XYDiagram3D)chartControl2.Diagram).RuntimeRotation = false;                    
                }                              
                #region Само построение
                #region Название меток по оси Х
                chartControl2.Annotations.AddTextAnnotation("Показатели", "1.   Выручка\n2.   Валовая прибыль (убыток)\n3.   Чистая прибыль (убыток)\n4.   Нераспределенная прибыль\n5.   Основные средства\n6.   Прочие внеоборотные активы\n7.   Запасы\n8.   Дебиторская задолженность\n9.   Денежные средства\n10.  Прочие текущие активы\n11.  Уставный капитал\n12.  Фонды и резервы\n13.  Долгосрочные обязательства\n14.  Краткосрочные обязательства").TextAlignment = StringAlignment.Near;
                #endregion            
                chartControl1.Series[0].Points.Add(new SeriesPoint(1, Convert.ToDouble(label36.Text)));
                chartControl1.Series[1].Points.Add(new SeriesPoint(1, Convert.ToDouble(label32.Text)));

                chartControl2.Series[0].Points.Add(new SeriesPoint(1, Convert.ToDouble(gridView1.GetRowCellValue(0, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(2, Convert.ToDouble(gridView1.GetRowCellValue(1, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(3, Convert.ToDouble(gridView1.GetRowCellValue(2, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(4, Convert.ToDouble(gridView1.GetRowCellValue(3, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(5, Convert.ToDouble(gridView1.GetRowCellValue(4, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(6, Convert.ToDouble(gridView1.GetRowCellValue(5, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(7, Convert.ToDouble(gridView1.GetRowCellValue(6, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(8, Convert.ToDouble(gridView1.GetRowCellValue(7, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(9, Convert.ToDouble(gridView1.GetRowCellValue(8, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(10, Convert.ToDouble(gridView1.GetRowCellValue(9, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(11, Convert.ToDouble(gridView1.GetRowCellValue(10, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(12, Convert.ToDouble(gridView1.GetRowCellValue(11, gridView1.Columns[2]))));

                chartControl2.Series[0].Points.Add(new SeriesPoint(13, Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(14, Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[2]))));

                chartControl2.Series[1].Points.Add(new SeriesPoint(1, Convert.ToDouble(gridView1.GetRowCellValue(0, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(2, Convert.ToDouble(gridView1.GetRowCellValue(1, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(3, Convert.ToDouble(gridView1.GetRowCellValue(2, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(4, Convert.ToDouble(gridView1.GetRowCellValue(3, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(5, Convert.ToDouble(gridView1.GetRowCellValue(4, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(6, Convert.ToDouble(gridView1.GetRowCellValue(5, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(7, Convert.ToDouble(gridView1.GetRowCellValue(6, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(8, Convert.ToDouble(gridView1.GetRowCellValue(7, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(9, Convert.ToDouble(gridView1.GetRowCellValue(8, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(10, Convert.ToDouble(gridView1.GetRowCellValue(9, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(11, Convert.ToDouble(gridView1.GetRowCellValue(10, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(12, Convert.ToDouble(gridView1.GetRowCellValue(11, gridView1.Columns[3]))));

                chartControl2.Series[1].Points.Add(new SeriesPoint(13, Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(14, Convert.ToDouble(gridView1.GetRowCellValue(13, gridView1.Columns[3]))));

                try
                {
                    diagramma2();
                }
                catch (Exception exx)
                {
                    MessageBox.Show(exx.Message, "ФАНЗ");
                    return;
                }
                #endregion
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion
        #region Диаграмма в 2d
        private void barButtonItem23_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView1.Columns[2]).ToString() != "")
            {
                chartControl1.Series[0].Points.Clear();
                chartControl1.Series[1].Points.Clear();
                chartControl2.Series[0].Points.Clear();
                chartControl2.Series[1].Points.Clear();
                chartControl2.Annotations.Clear();
                chartControl4.Series[0].Points.Clear();

                string type = chartControl1.Diagram.GetType().ToString();
                string type3d = "DevExpress.XtraCharts.XYDiagram3D";
                if (type == type3d)
                {
                    chartControl1.Series[0].ChangeView(ViewType.FullStackedBar);
                    chartControl1.Series[1].ChangeView(ViewType.FullStackedBar);
                    XYDiagram diagram1 = (XYDiagram)chartControl1.Diagram;
                    diagram1.AxisY.VisualRange.Auto = true;
                    diagram1.AxisX.VisualRange.Auto = true;
                }

                string type1 = chartControl2.Diagram.GetType().ToString();
                string type_3d = "DevExpress.XtraCharts.XYDiagram3D";
                if (type1 == type_3d)
                {
                    chartControl2.Series[0].ChangeView(ViewType.Bar);
                    chartControl2.Series[1].ChangeView(ViewType.Bar);
                    XYDiagram diagram2 = (XYDiagram)chartControl2.Diagram;
                    diagram2.AxisX.Visible = true;
                    diagram2.AxisY.VisualRange.Auto = true;
                    diagram2.AxisX.VisualRange.Auto = true;
                }
                else
                {
                    XYDiagram diagram2 = (XYDiagram)chartControl2.Diagram;
                    diagram2.AxisX.Visible = true;
                    diagram2.AxisY.VisualRange.Auto = true;
                    diagram2.AxisX.VisualRange.Auto = true;
                } 
                #region Само построение
                chartControl1.Series[0].Points.Add(new SeriesPoint(1, Convert.ToDouble(label36.Text)));
                chartControl1.Series[1].Points.Add(new SeriesPoint(1, Convert.ToDouble(label32.Text)));

                chartControl2.Series[0].Points.Add(new SeriesPoint(1, Convert.ToDouble(gridView1.GetRowCellValue(0, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(2, Convert.ToDouble(gridView1.GetRowCellValue(1, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(3, Convert.ToDouble(gridView1.GetRowCellValue(2, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(4, Convert.ToDouble(gridView1.GetRowCellValue(3, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(5, Convert.ToDouble(gridView1.GetRowCellValue(4, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(6, Convert.ToDouble(gridView1.GetRowCellValue(5, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(7, Convert.ToDouble(gridView1.GetRowCellValue(6, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(8, Convert.ToDouble(gridView1.GetRowCellValue(7, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(9, Convert.ToDouble(gridView1.GetRowCellValue(8, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(10, Convert.ToDouble(gridView1.GetRowCellValue(9, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(11, Convert.ToDouble(gridView1.GetRowCellValue(10, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(12, Convert.ToDouble(gridView1.GetRowCellValue(11, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(13, Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(14, Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[2]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(1, Convert.ToDouble(gridView1.GetRowCellValue(0, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(2, Convert.ToDouble(gridView1.GetRowCellValue(1, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(3, Convert.ToDouble(gridView1.GetRowCellValue(2, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(4, Convert.ToDouble(gridView1.GetRowCellValue(3, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(5, Convert.ToDouble(gridView1.GetRowCellValue(4, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(6, Convert.ToDouble(gridView1.GetRowCellValue(5, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(7, Convert.ToDouble(gridView1.GetRowCellValue(6, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(8, Convert.ToDouble(gridView1.GetRowCellValue(7, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(9, Convert.ToDouble(gridView1.GetRowCellValue(8, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(10, Convert.ToDouble(gridView1.GetRowCellValue(9, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(11, Convert.ToDouble(gridView1.GetRowCellValue(10, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(12, Convert.ToDouble(gridView1.GetRowCellValue(11, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(13, Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(14, Convert.ToDouble(gridView1.GetRowCellValue(13, gridView1.Columns[3]))));
                try
                {
                    diagramma2();
                }
                catch (Exception exx)
                {
                    MessageBox.Show(exx.Message, "ФАНЗ");
                    return;
                }
                #endregion
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion
        #endregion
        #region Экспорт в Ексель Таблицы "Результаты Оптимизации"
        private void barButtonItem21_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            #region NewGrid
            if (gridView2.RowCount != 0)
            {
                string date1 = DateTime.Today.ToShortDateString();
                string date2 = "Cоздан в " + Convert.ToString(DateTime.Now.Hour + "." + DateTime.Now.Minute);
                string filename = date2 + ".xlsx";
                DialogResult results;
                //------------------------------------------------------------------------------------------------------------------------               
                try
                {
                    if (!Directory.Exists(s + @"\Отчеты\Результаты оптимизации\" + "" + date1))
                    {
                        Directory.CreateDirectory(s + @"\Отчеты\Результаты оптимизации\" + "" + date1);
                        string m = s + "\\Отчеты\\Результаты оптимизации\\" + date1;
                        gridView2.OptionsView.ShowViewCaption = false;
                        gridView2.ExportToXlsx(m + @"\" + filename);
                        gridView2.OptionsView.ShowViewCaption = true;
                        MessageBox.Show("Данные в Ексель успешно сохранены", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string m = s + "\\Отчеты\\Результаты оптимизации\\" + date1;
                        gridView2.OptionsView.ShowViewCaption = false;
                        gridView2.ExportToXlsx(m + @"\" + filename);
                        gridView2.OptionsView.ShowViewCaption = true;
                        MessageBox.Show("Данные в Ексель успешно сохранены", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                    openFileDialog1.InitialDirectory = s + @"\Отчеты\Результаты оптимизации\" + date1;
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
            #endregion
        }
        #endregion
        #region Экспорт в Ексель Таблицы "Результаты расчетов"
        private void barButtonItem24_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView3.RowCount != 0)
            {
                string date1 = DateTime.Today.ToShortDateString();
                string date2 = "Cоздан в " + Convert.ToString(DateTime.Now.Hour + "." + DateTime.Now.Minute);
                string filename = date2 + ".xlsx";
                DialogResult results;
                //------------------------------------------------------------------------------------------------------------------------              
                try
                {
                    if (!Directory.Exists(s + @"\Отчеты\Результаты расчетов\" + "" + date1))
                    {
                        Directory.CreateDirectory(s + @"\Отчеты\Результаты расчетов\" + "" + date1); // создание папки по нужной дате
                        string m = s + "\\Отчеты\\Результаты расчетов\\" + date1;
                        gridView3.OptionsView.ShowViewCaption = false;
                        gridView3.ExportToXlsx(m + @"\" + filename);
                        gridView3.OptionsView.ShowViewCaption = true;
                        MessageBox.Show("Данные в Ексель успешно сохранены", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string m = s + "\\Отчеты\\Результаты расчетов\\" + date1;
                        gridView3.OptionsView.ShowViewCaption = false;
                        gridView3.ExportToXlsx(m + @"\" + filename);
                        gridView3.OptionsView.ShowViewCaption = true;
                        MessageBox.Show("Данные в Ексель успешно сохранены", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                    openFileDialog1.InitialDirectory = s + @"\Отчеты\Результаты расчетов\" + date1;
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
        #region Экспорт в Ексель Таблицы "Отклонение от идеального фин. состояния"
        private void barButtonItem25_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
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

                    if (!Directory.Exists(s + @"\Отчеты\Отклонение от идеального фин. состояния\" + "" + date1))
                    {
                        Directory.CreateDirectory(s + @"\Отчеты\Отклонение от идеального фин. состояния\" + "" + date1); // создание папки по нужной дате
                        string m = s + "\\Отчеты\\Отклонение от идеального фин. состояния\\" + date1;
                        gridView1.ExportToXlsx(m + @"\" + filename);
                        MessageBox.Show("Данные в Ексель успешно сохранены", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string m = s + "\\Отчеты\\Отклонение от идеального фин. состояния\\" + date1;
                        gridView1.ExportToXlsx(m + @"\" + filename);
                        MessageBox.Show("Данные в Ексель успешно сохранены", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                    openFileDialog1.InitialDirectory = s + @"\Отчеты\Отклонение от идеального фин. состояния\" + date1;
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

        #region Импорт из Екселя в "Результаты оптимизации"
        private void barButtonItem3_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView2.RowCount == 0)
            {
                string m = s + @"\Отчеты\Результаты оптимизации\";
                OpenFileDialog opf = new OpenFileDialog();
                opf.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                opf.InitialDirectory = (m);
                opf.ShowDialog();
                string filename = opf.FileName;
                if (filename != "")
                {                    
                    #region Загрузка из Екселя в новый грид
                    try
                    {
                        System.IO.FileStream stream = System.IO.File.Open(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                        Excel.IExcelDataReader IEDR;
                        int fileformat = opf.SafeFileName.IndexOf(".xlsx");
                        if (fileformat > -1)
                        {
                            IEDR = Excel.ExcelReaderFactory.CreateOpenXmlReader(stream);
                        }
                        else
                        {
                            IEDR = Excel.ExcelReaderFactory.CreateBinaryReader(stream);
                        }
                        IEDR.IsFirstRowAsColumnNames = true;
                        DataSet ds = IEDR.AsDataSet();                        
                        DataTable dt = ds.Tables[0];
                        gridControl2.DataSource = dt;
                        #region Формирование всех свойств и методов для таблицы

                        gridView2.Columns[0].Caption = "Бухгалтерский баланс. Основные показатели.\nАКТИВ";
                        gridView2.Columns[1].Caption = "Фактические значения";
                        gridView2.Columns[2].Caption = "Метод оптимизации на основе\n заданных показателей";

                        for (int t = 0; t < gridView2.Columns.Count; t++)
                        {
                            gridView2.Columns[t].AppearanceHeader.Options.UseTextOptions = true;
                            gridView2.Columns[t].AppearanceHeader.BackColor = Color.LightSteelBlue;
                            gridView2.Columns[t].AppearanceHeader.BackColor2 = Color.White;
                            gridView2.Columns[t].AppearanceHeader.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.Vertical;
                            gridView2.Columns[t].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                            gridView2.Columns[t].BestFit();
                        }
                        foreach (DevExpress.XtraGrid.Columns.GridColumn column in gridView2.Columns)
                        { column.OptionsColumn.AllowSort = DefaultBoolean.False; }
                        

                        gridView2.Columns[1].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                        gridView2.Columns[2].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                        gridView2.Columns[1].AppearanceCell.BackColor = Color.LightGreen;
                        gridView2.Columns[1].AppearanceCell.BackColor2 = Color.White;
                        gridView2.Columns[1].AppearanceCell.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.Horizontal;
                        gridView2.Appearance.VertLine.BackColor = Color.LightSteelBlue;
                        gridView2.Appearance.HorzLine.BackColor = Color.LightSteelBlue;

                        gridView2.OptionsPrint.UsePrintStyles = true;
                        gridView2.Appearance.HeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
                        gridView2.ColumnPanelRowHeight = 50;
                        DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit Grid_MemoEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit(); // инициализируем MemoEdit с именем “MyGrid_MemoEdit”
                        Grid_MemoEdit.WordWrap = true;

                        gridView2.OptionsView.RowAutoHeight = true;
                        gridView2.OptionsView.ColumnAutoWidth = true;
                        gridView2.Columns[0].ColumnEdit = Grid_MemoEdit;
                        gridView2.Columns[2].ColumnEdit = Grid_MemoEdit;

                        for (int i = 0; i < gridView2.DataRowCount; i++)
                        {
                            if (i == 0 || i == 4 || i == 5 || i == 11 || i == 13 || i == 14 || i == 15 || i == 20 || i == 21 || i == 25 || i == 27 || i == 28)
                            {
                                gridView2.SetRowCellValue(i, gridView2.Columns[2], null);
                            }
                        }
                        #endregion
                        #region Формулы рассчета по ячейкам
                        gridView2.SetRowCellValue(3, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(1, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(2, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(10, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(6, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(7, gridView2.Columns[1])) +
                                           Convert.ToDouble(gridView2.GetRowCellValue(8, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(9, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(12, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(3, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(10, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(19, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(16, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(17, gridView2.Columns[1])) +
                                           Convert.ToDouble(gridView2.GetRowCellValue(18, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(24, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(22, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(23, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(26, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(19, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(24, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(32, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(30, gridView2.Columns[1])) - Convert.ToDouble(gridView2.GetRowCellValue(31, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(34, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(32, gridView2.Columns[1])) - Convert.ToDouble(gridView2.GetRowCellValue(33, gridView2.Columns[1]))));
                        gridView2.SetRowCellValue(37, gridView2.Columns[1], ((Convert.ToDouble(gridView2.GetRowCellValue(34, gridView2.Columns[1])) + Convert.ToDouble(gridView2.GetRowCellValue(35, gridView2.Columns[1]))) -
                            Convert.ToDouble(gridView2.GetRowCellValue(36, gridView2.Columns[1]))));
                        //новое---------------------
                        
                        gridView2.SetRowCellValue(39, gridView2.Columns[1], (Convert.ToDouble(gridView2.GetRowCellValue(18, gridView2.Columns[1]))));
                        ////--------------------------

                        gridView2.Columns[0].OptionsColumn.ReadOnly = true;
                        gridView2.Columns[2].OptionsColumn.ReadOnly = true;

                        #endregion    
                        IEDR.Close();
                    }
                    catch (Exception load)
                    {
                        MessageBox.Show(load.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    #endregion
                }
                else
                {
                    MessageBox.Show("Не выбран файл для загрузки.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                opf.RestoreDirectory = true;
            }
            else
            {
                MessageBox.Show("Таблица содержит данные, требуетя очистка перед импортом.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion
        #region Импорт из Екселя в "Результаты расчетов"
        private void barButtonItem26_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView3.RowCount == 0)
            {
                string m = s + @"\Отчеты\Результаты расчетов\";
                OpenFileDialog opf = new OpenFileDialog();
                opf.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                opf.InitialDirectory = (m);
                opf.ShowDialog();
                string filename = opf.FileName;
                if (filename != "")
                {
                    #region Загрузка из Екселя в новый грид
                    try
                    {
                        FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read);
                        Excel.IExcelDataReader IEDR;
                        int fileformat = opf.SafeFileName.IndexOf(".xlsx");
                        if (fileformat > -1)
                        {
                            IEDR = Excel.ExcelReaderFactory.CreateOpenXmlReader(stream);
                        }
                        else
                        {
                            IEDR = Excel.ExcelReaderFactory.CreateBinaryReader(stream);
                        }
                        IEDR.IsFirstRowAsColumnNames = true;
                        DataSet ds = IEDR.AsDataSet();
                        DataTable dt = ds.Tables[0];
                        gridControl3.DataSource = dt;
                        #region Формирование всех свойств и методов для таблицы

                        gridView3.Columns[0].Caption = "№ по\n справочнику";
                        gridView3.Columns[1].Caption = "Наименование\n показателя";
                        gridView3.Columns[2].Caption = "Сокращенное\n наименование\n в модели";
                        gridView3.Columns[3].Caption = "Заданные\n значения\n показателей";
                        gridView3.Columns[4].Caption = "Расчетные значения по\n оптимизированным статьям\n баланса и отчета\n о прибылях и убытках";
                        gridView3.Columns[5].Caption = "Абсолютное\n отклонение";
                        gridView3.Columns[6].Caption = "Относительное_отклонение";

                        for (int t = 0; t < gridView3.Columns.Count; t++)
                        {
                            gridView3.Columns[t].AppearanceHeader.Options.UseTextOptions = true;
                            gridView3.Columns[t].AppearanceHeader.BackColor = Color.LightSteelBlue;
                            gridView3.Columns[t].AppearanceHeader.BackColor2 = Color.White;
                            gridView3.Columns[t].AppearanceHeader.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.Vertical;
                            gridView3.Columns[t].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                            gridView3.Columns[t].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                            gridView3.Columns[t].BestFit();
                        }
                        gridView3.Columns[1].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Default;

                        foreach (DevExpress.XtraGrid.Columns.GridColumn column in gridView3.Columns)
                        { column.OptionsColumn.AllowSort = DefaultBoolean.False; }

                        gridView3.Appearance.VertLine.BackColor = Color.LightSteelBlue;
                        gridView3.Appearance.HorzLine.BackColor = Color.LightSteelBlue;

                        gridView3.OptionsPrint.UsePrintStyles = true;
                        gridView3.ColumnPanelRowHeight = 60;
                        DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit Grid_MemoEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit(); // инициализируем MemoEdit с именем “MyGrid_MemoEdit”
                        Grid_MemoEdit.WordWrap = true;

                        gridView3.OptionsView.RowAutoHeight = true;
                        gridView3.OptionsView.ColumnAutoWidth = true;
                        gridView3.Columns[1].ColumnEdit = Grid_MemoEdit;
                        #endregion                        
                        IEDR.Close();
                    }
                    catch (Exception load)
                    {
                        MessageBox.Show(load.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    #endregion
                }
                else
                {
                    MessageBox.Show("Не выбран файл для загрузки.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                opf.RestoreDirectory = true;
            }
            else
            {
                MessageBox.Show("Таблица содержит данные, требуетя очистка перед импортом.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion
        #region Импорт из Екселя в "Отклонение от идеального фин. состояния"
        private void barButtonItem27_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount == 0)
            {
                string m = s + @"\Отчеты\Отклонение от идеального фин. состояния\";
                OpenFileDialog opf = new OpenFileDialog();
                opf.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                opf.InitialDirectory = (m);
                opf.ShowDialog();
                string filename = opf.FileName;
                if (filename != "")
                {
                    #region Загрузка из Екселя в новый грид
                    try
                    {
                        FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read);
                        Excel.IExcelDataReader IEDR;
                        int fileformat = opf.SafeFileName.IndexOf(".xlsx");
                        if (fileformat > -1)
                        {
                            IEDR = Excel.ExcelReaderFactory.CreateOpenXmlReader(stream);
                        }
                        else
                        {
                            IEDR = Excel.ExcelReaderFactory.CreateBinaryReader(stream);
                        }
                        IEDR.IsFirstRowAsColumnNames = true;
                        DataSet ds = IEDR.AsDataSet();
                        DataTable dt = ds.Tables[0];
                        gridControl1.DataSource = dt;
                        #region Формирование всех свойств и методов для таблицы

                        gridView1.Columns[0].Caption = "№";
                        gridView1.Columns[1].Caption = "Показатель";
                        gridView1.Columns[2].Caption = "Фактические\n значения";
                        gridView1.Columns[3].Caption = "Расчетные\n значения";
                        gridView1.Columns[4].Caption = "Абсолютное\n отклонение";
                        gridView1.Columns[5].Caption = "Относительное\n отклонение, %";

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

                        gridView1.AppearancePrint.HeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
                        gridView1.OptionsPrint.UsePrintStyles = true;
                        gridView1.ColumnPanelRowHeight = 50;

                        DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit Grid_MemoEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit(); // инициализируем MemoEdit с именем “MyGrid_MemoEdit”
                        Grid_MemoEdit.WordWrap = true;
                        gridView1.OptionsView.RowAutoHeight = true;
                        gridView1.Columns[1].ColumnEdit = Grid_MemoEdit;
                        #endregion
                        IEDR.Close();
                    }
                    catch (Exception load)
                    {
                        MessageBox.Show(load.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    #endregion
                }
                else
                {
                    MessageBox.Show("Не выбран файл для загрузки.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                opf.RestoreDirectory = true;
            }
            else
            {
                MessageBox.Show("Таблица содержит данные, требуетя очистка перед импортом.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion

        #region Очистка таблицы "Результаты оптимизации"
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (gridView2.DataSource != null || gridView2.RowCount != 0)
                {
                    DialogResult result;
                    result = MessageBox.Show("Таблица содержит данные, \n Продолжить удаление?", "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == System.Windows.Forms.DialogResult.Yes)
                    {
                        gridControl2.DataSource = null;
                        gridView2.Columns.Clear();
                        gridView2.ColumnPanelRowHeight = 20;
                    }
                }
                else
                {
                    MessageBox.Show("Таблица пустая, нельзя удалить данные", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        #endregion
        #region Очистка таблицы "Результаты расчетов"
        private void button18_Click(object sender, EventArgs e)
        {
            try
            {
                if (gridView3.DataSource != null || gridView3.RowCount != 0)
                {
                    DialogResult result;
                    result = MessageBox.Show("Таблица содержит данные, \n Продолжить удаление?", "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == System.Windows.Forms.DialogResult.Yes)
                    {
                        gridControl3.DataSource = null;
                        gridView3.Columns.Clear();
                        gridView3.ColumnPanelRowHeight = 20;
                    }
                }
                else
                    MessageBox.Show("Таблица пустая, нельзя удалить данные", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        #endregion
        #region Очистка таблицы "Отклонение от идеального фин. состояния"
        private void button22_Click(object sender, EventArgs e)
        {
            try
            {
                if (gridControl1.DataSource != null || gridView1.RowCount != 0)
                {
                    DialogResult result;
                    result = MessageBox.Show("Таблица содержит данные, \n Продолжить удаление?", "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == System.Windows.Forms.DialogResult.Yes)
                    {
                        gridControl1.DataSource = null;
                        gridView1.Columns.Clear();
                        gridView1.ColumnPanelRowHeight = 20;
                    }
                }
                else
                {
                    MessageBox.Show("Таблица пустая, нельзя удалить данные", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        #endregion

        #region Изменение пароля администратором
        private void barButtonItem20_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Password pas = new Password();
            pas.ShowDialog();
        }
        #endregion
        #region Вызов формы прогнозирования
        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Linean form = new Linean();
            form.ShowDialog();
        }
        #endregion
        #region Привязка хоста
        private void barButtonItem28_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            #region Кодируем текущий хост
            string hostname = Environment.UserDomainName.ToString();
            string host_tmp = string.Empty;
            string hash_host = string.Empty;

            byte[] bytes = Encoding.Unicode.GetBytes(hostname);
            MD5CryptoServiceProvider CSP = new MD5CryptoServiceProvider();
            byte[] byteHash = CSP.ComputeHash(bytes);
            foreach (byte b in byteHash)
            {
                hash_host += string.Format("{0:x2}", b);
            }
            host_tmp = Convert.ToString(new Guid(hash_host));
            #endregion

            if (!File.Exists(@host))
            {
                var file = System.IO.File.Create(@host);
                file.Close();
                try
                {
                    StreamWriter writer1 = new StreamWriter(@host, false, Encoding.GetEncoding(1251));   //запись в файл пароля
                    writer1.Write(host_tmp);
                    writer1.Close();
                }
                catch (Exception ez)
                {
                    MessageBox.Show(ez.Message, "ФАНЗ");
                    return;
                }
                MessageBox.Show("Привязка к хосту успешно произведена.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            else
            {
                try
                {
                    StreamWriter writer1 = new StreamWriter(@host, false, Encoding.GetEncoding(1251));   //запись в файл пароля
                    writer1.Write(host_tmp);
                    writer1.Close();
                }
                catch (Exception ez)
                {
                    MessageBox.Show(ez.Message, "ФАНЗ");
                    return;
                }
                MessageBox.Show("Привязка к хосту успешно произведена.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        #endregion
        #region Вызов справки
        private void Form1_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            Simplex_2.Формы_проекта.Help form = new Simplex_2.Формы_проекта.Help();
            form.Show();
        }
        #endregion
        #region Смена тем оформления
        private void barStaticItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            defaultLookAndFeel1.LookAndFeel.SetSkinStyle("DevExpress Style");
            string style = "DevExpress Style";
            try
            {
                StreamWriter writer1 = new StreamWriter(@style_name, false, Encoding.GetEncoding(1251));   //запись в файл темы
                writer1.Write(style);
                writer1.Close();
            }
            catch (Exception ez)
            {
                MessageBox.Show(ez.Message, "ФАНЗ");
                return;
            }
        }

        private void barStaticItem3_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            defaultLookAndFeel1.LookAndFeel.SetSkinStyle("DevExpress Dark Style");
            string style = "DevExpress Dark Style";
            try
            {
                StreamWriter writer1 = new StreamWriter(@style_name, false, Encoding.GetEncoding(1251));   //запись в файл темы
                writer1.Write(style);
                writer1.Close();
            }
            catch (Exception ez)
            {
                MessageBox.Show(ez.Message, "ФАНЗ");
                return;
            }
        }

        private void barStaticItem4_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            defaultLookAndFeel1.LookAndFeel.SetSkinStyle("VS2010");
            string style = "VS2010";
            try
            {
                StreamWriter writer1 = new StreamWriter(@style_name, false, Encoding.GetEncoding(1251));   //запись в файл темы
                writer1.Write(style);
                writer1.Close();
            }
            catch (Exception ez)
            {
                MessageBox.Show(ez.Message, "ФАНЗ");
                return;
            }
        }

        private void barStaticItem5_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            defaultLookAndFeel1.LookAndFeel.SetSkinStyle("Seven Classic");
            string style = "Seven Classic";
            try
            {
                StreamWriter writer1 = new StreamWriter(@style_name, false, Encoding.GetEncoding(1251));   //запись в файл темы
                writer1.Write(style);
                writer1.Close();
            }
            catch (Exception ez)
            {
                MessageBox.Show(ez.Message, "ФАНЗ");
                return;
            }
        }

        private void barStaticItem6_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            defaultLookAndFeel1.LookAndFeel.SetSkinStyle("Office 2010 Blue");
            string style = "Office 2010 Blue";
            try
            {
                StreamWriter writer1 = new StreamWriter(@style_name, false, Encoding.GetEncoding(1251));   //запись в файл темы
                writer1.Write(style);
                writer1.Close();
            }
            catch (Exception ez)
            {
                MessageBox.Show(ez.Message, "ФАНЗ");
                return;
            }
        }

        private void barStaticItem7_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            defaultLookAndFeel1.LookAndFeel.SetSkinStyle("Office 2010 Black");
            string style = "Office 2010 Black";
            try
            {
                StreamWriter writer1 = new StreamWriter(@style_name, false, Encoding.GetEncoding(1251));   //запись в файл темы
                writer1.Write(style);
                writer1.Close();
            }
            catch (Exception ez)
            {
                MessageBox.Show(ez.Message, "ФАНЗ");
                return;
            }
        }

        private void barStaticItem8_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            defaultLookAndFeel1.LookAndFeel.SetSkinStyle("Office 2010 Silver");
            string style = "Office 2010 Silver";
            try
            {
                StreamWriter writer1 = new StreamWriter(@style_name, false, Encoding.GetEncoding(1251));   //запись в файл темы
                writer1.Write(style);
                writer1.Close();
            }
            catch (Exception ez)
            {
                MessageBox.Show(ez.Message, "ФАНЗ");
                return;
            }
        }

        private void barStaticItem9_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            defaultLookAndFeel1.LookAndFeel.SetSkinStyle("Office 2013");
            string style = "Office 2013";
            try
            {
                StreamWriter writer1 = new StreamWriter(@style_name, false, Encoding.GetEncoding(1251));   //запись в файл темы
                writer1.Write(style);
                writer1.Close();
            }
            catch (Exception ez)
            {
                MessageBox.Show(ez.Message, "ФАНЗ");
                return;
            }
        }

        private void barStaticItem10_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            defaultLookAndFeel1.LookAndFeel.SetSkinStyle("Office 2013 Dark Gray");
            string style = "Office 2013 Dark Gray";
            try
            {
                StreamWriter writer1 = new StreamWriter(@style_name, false, Encoding.GetEncoding(1251));   //запись в файл темы
                writer1.Write(style);
                writer1.Close();
            }
            catch (Exception ez)
            {
                MessageBox.Show(ez.Message, "ФАНЗ");
                return;
            }
        }

        private void barStaticItem11_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            defaultLookAndFeel1.LookAndFeel.SetSkinStyle("Office 2013 Light Gray");
            string style = "Office 2013 Light Gray";
            try
            {
                StreamWriter writer1 = new StreamWriter(@style_name, false, Encoding.GetEncoding(1251));   //запись в файл темы
                writer1.Write(style);
                writer1.Close();
            }
            catch (Exception ez)
            {
                MessageBox.Show(ez.Message, "ФАНЗ");
                return;
            }
        }

        private void barStaticItem12_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            defaultLookAndFeel1.LookAndFeel.SetSkinStyle("Visual Studio 2013 Blue");
            string style = "Visual Studio 2013 Blue";
            try
            {
                StreamWriter writer1 = new StreamWriter(@style_name, false, Encoding.GetEncoding(1251));   //запись в файл темы
                writer1.Write(style);
                writer1.Close();
            }
            catch (Exception ez)
            {
                MessageBox.Show(ez.Message, "ФАНЗ");
                return;
            }
        }

        private void barStaticItem13_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            defaultLookAndFeel1.LookAndFeel.SetSkinStyle("Visual Studio 2013 Light");
            string style = "Visual Studio 2013 Light";
            try
            {
                StreamWriter writer1 = new StreamWriter(@style_name, false, Encoding.GetEncoding(1251));   //запись в файл темы
                writer1.Write(style);
                writer1.Close();
            }
            catch (Exception ez)
            {
                MessageBox.Show(ez.Message, "ФАНЗ");
                return;
            }
        }

        private void barStaticItem14_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            defaultLookAndFeel1.LookAndFeel.SetSkinStyle("Visual Studio 2013 Dark");
            string style = "Visual Studio 2013 Dark";
            try
            {
                StreamWriter writer1 = new StreamWriter(@style_name, false, Encoding.GetEncoding(1251));   //запись в файл темы
                writer1.Write(style);
                writer1.Close();
            }
            catch (Exception ez)
            {
                MessageBox.Show(ez.Message, "ФАНЗ");
                return;
            }
        }
        #endregion
        #region Печать 3х таблиц
        private void barButtonItem29_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView2.RowCount != 0)
            {
                if (!gridControl2.IsPrintingAvailable)
                {
                    MessageBox.Show("The 'DevExpress.XtraPrinting' library is not found", "Error");
                    return;
                }
                gridControl2.Print();
            }
            else
            {
                MessageBox.Show("Таблица пустая, печать отменена.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        private void barButtonItem32_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView2.RowCount != 0)
            {
                if (!gridControl2.IsPrintingAvailable)
                {
                    MessageBox.Show("The 'DevExpress.XtraPrinting' library is not found", "Error");
                    return;
                }
                gridControl2.ShowPrintPreview();
            }
            else
            {
                MessageBox.Show("Таблица пустая, предпросмотр недоступен.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        private void barButtonItem30_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
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
        private void barButtonItem33_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
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
                MessageBox.Show("Таблица пустая, печать отменена.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }                                
        }
        #region Риббон предпросмотр
        private void barButtonItem4_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Link l = new Link(new PrintingSystem());
            l.Landscape = true;
            l.PaperKind = System.Drawing.Printing.PaperKind.A4;
            l.Margins.Bottom = 10;
            l.Margins.Top = 10;
            l.Margins.Right = 10;
            l.Margins.Left = 10;
            cp = new DevExpress.XtraCharts.Printing.ChartPrinter(this.chartControl2);
            cp.Initialize(l.PrintingSystem, l);
            cp.SizeMode = DevExpress.XtraCharts.Printing.PrintSizeMode.Stretch;
            l.CreateDetailArea += new CreateAreaEventHandler(l_CreateDetailArea);
            l.ShowPreviewDialog();
            cp.Release();
        }
        void l_CreateDetailArea(object sender, CreateAreaEventArgs e)
        {
            cp.CreateDetail(e.Graph);
        }
        private void barButtonItem36_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            chartControl1.ShowRibbonPrintPreview();
        }
        private void barButtonItem35_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Link l = new Link(new PrintingSystem());
            l.Landscape = true;
            l.PaperKind = System.Drawing.Printing.PaperKind.A4;
            l.Margins.Bottom = 10;
            l.Margins.Top = 10;
            l.Margins.Right = 10;
            l.Margins.Left = 10;
            cp = new DevExpress.XtraCharts.Printing.ChartPrinter(this.chartControl4);
            cp.Initialize(l.PrintingSystem, l);
            cp.SizeMode = DevExpress.XtraCharts.Printing.PrintSizeMode.Stretch;
            l.CreateDetailArea += new CreateAreaEventHandler(l_CreateDetailArea);
            l.ShowPreviewDialog();
            cp.Release();
        }
        #endregion
        #region Печать
        private void barButtonItem37_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            chartControl1.Print();
        }
        private void barButtonItem38_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Link l = new Link(new PrintingSystem());
            l.Landscape = true;
            l.PaperKind = System.Drawing.Printing.PaperKind.A4;
            l.Margins.Bottom = 10;
            l.Margins.Top = 10;
            l.Margins.Right = 10;
            l.Margins.Left = 10;
            cp = new DevExpress.XtraCharts.Printing.ChartPrinter(this.chartControl2);
            cp.Initialize(l.PrintingSystem, l);
            cp.SizeMode = DevExpress.XtraCharts.Printing.PrintSizeMode.Stretch;
            l.CreateDetailArea += new CreateAreaEventHandler(l_CreateDetailArea);           
            l.Print(string.Empty);
            cp.Release();
        }
        private void barButtonItem39_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Link l = new Link(new PrintingSystem());
            l.Landscape = true;
            l.PaperKind = System.Drawing.Printing.PaperKind.A4;
            l.Margins.Bottom = 10;
            l.Margins.Top = 10;
            l.Margins.Right = 10;
            l.Margins.Left = 10;
            cp = new DevExpress.XtraCharts.Printing.ChartPrinter(this.chartControl4);
            cp.Initialize(l.PrintingSystem, l);
            cp.SizeMode = DevExpress.XtraCharts.Printing.PrintSizeMode.Stretch;
            l.CreateDetailArea += new CreateAreaEventHandler(l_CreateDetailArea);
            l.Print(string.Empty);
            cp.Release();
        }
        #endregion
        #endregion
        #region Всплывающие подсказки и обработка закрытия формы

        private void button9_MouseEnter(object sender, EventArgs e)
        {
            ToolTip toolTip1 = new System.Windows.Forms.ToolTip();
            toolTip1.SetToolTip(button9, "Сохраняет шаблон таблицы вместе с фактическими значениями,\nчтобы не вводить повторно при последующих расчетах.");
        }

        private void button16_MouseEnter(object sender, EventArgs e)
        {
            ToolTip toolTip2 = new System.Windows.Forms.ToolTip();
            toolTip2.SetToolTip(button16, "Загружает последний сохраненный шаблон.");
        }

        private void button14_MouseEnter(object sender, EventArgs e)
        {
            ToolTip toolTip3 = new System.Windows.Forms.ToolTip();
            toolTip3.SetToolTip(button14, "Обновление данных в таблице - суммирование итоговых \nи вычисляемых значений в обеих таблицах.");
        }

        private void button4_MouseEnter(object sender, EventArgs e)
        {
            ToolTip toolTip5 = new System.Windows.Forms.ToolTip();
            toolTip5.SetToolTip(button4, "Очистка таблицы.");
        }

        private void button17_MouseEnter(object sender, EventArgs e)
        {
            ToolTip toolTip6 = new System.Windows.Forms.ToolTip();
            toolTip6.SetToolTip(button17, "После получения решения и на основе\nрассчитанных данных - заполнение таблицы.");
        }

        private void button21_MouseEnter(object sender, EventArgs e)
        {
            ToolTip toolTip6 = new System.Windows.Forms.ToolTip();
            toolTip6.SetToolTip(button17, "После получения решения и на основе\nрассчитанных данных - заполнение таблицы.");
        }

        private void button15_MouseEnter(object sender, EventArgs e)
        {
            ToolTip toolTip7 = new System.Windows.Forms.ToolTip();
            toolTip7.SetToolTip(button15, "Загружает последний сохраненный шаблон.");
        }

        private void button13_MouseEnter(object sender, EventArgs e)
        {
            ToolTip toolTip8 = new System.Windows.Forms.ToolTip();
            toolTip8.SetToolTip(button13, "Сохраняет шаблон таблицы.");
        }

        private void button18_MouseEnter(object sender, EventArgs e)
        {
            ToolTip toolTip9 = new System.Windows.Forms.ToolTip();
            toolTip9.SetToolTip(button18, "Очистка таблицы.");
        }

        private void button20_MouseEnter(object sender, EventArgs e)
        {
            ToolTip toolTip10 = new System.Windows.Forms.ToolTip();
            toolTip10.SetToolTip(button20, "Загружает последний сохраненный шаблон.");
        }

        private void button19_MouseEnter(object sender, EventArgs e)
        {
            ToolTip toolTip11 = new System.Windows.Forms.ToolTip();
            toolTip11.SetToolTip(button19, "Сохраняет шаблон таблицы.");
        }

        private void button22_MouseEnter(object sender, EventArgs e)
        {
            ToolTip toolTip12 = new System.Windows.Forms.ToolTip();
            toolTip12.SetToolTip(button22, "Очистка таблицы.");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                DialogResult results;
                results = MessageBox.Show("Выйти из программы?", "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (results == System.Windows.Forms.DialogResult.Yes)
                {
                    e.Cancel = false;                            
                }
                else
                {
                    e.Cancel = true;
                }
            }
            catch (Exception er)
            {
                MessageBox.Show(er.Message,"ФАНЗ");
                return;
            }
        }
        #endregion
        private void barButtonItem14_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            AboutBox1 about = new AboutBox1();
            about.Show();
        }
        private void barButtonItem40_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Simplex_2.Формы_проекта.Help form = new Simplex_2.Формы_проекта.Help();
            form.Show();
        }
        private void barButtonItem41_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            InsertToBZ();
        }
        #region Диаграмма отклонений
        private void X_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView1.Columns[0]).ToString() != "")
            {
                string date = "Cоздан в " + Convert.ToString(DateTime.Now.Hour) + "." + Convert.ToString(DateTime.Now.Minute) + "." + Convert.ToString(DateTime.Now.Second);
                string filename = date + ".xls";

                try
                {
                    string save = s + @"\Экспорт\";
                    if (!Directory.Exists(save))
                    {
                        Directory.CreateDirectory(save);
                        chartControl1.ExportToXls(save + filename);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl1.ExportToXls(xlsStream);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        chartControl1.ExportToXls(save + filename);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl1.ExportToXls(xlsStream);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception www)
                {
                    MessageBox.Show(www.Message,"ФАНЗ");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        private void barButtonItem42_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView1.Columns[0]).ToString() != "")
            {
                string date = "Cоздан в " + Convert.ToString(DateTime.Now.Hour) + "." + Convert.ToString(DateTime.Now.Minute) + "." + Convert.ToString(DateTime.Now.Second);
                string filename = date + ".xlsx";
                try
                {
                    string save = s + @"\Экспорт\";
                    if (!Directory.Exists(save))
                    {
                        Directory.CreateDirectory(save);
                        chartControl1.ExportToXlsx(save + filename);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl1.ExportToXlsx(xlsStream);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        chartControl1.ExportToXlsx(save + filename);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl1.ExportToXlsx(xlsStream);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception www)
                {
                    MessageBox.Show(www.Message,"ФАНЗ");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

        }

        private void Э_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView1.Columns[0]).ToString() != "")
            {
                string date = "Cоздан в " + Convert.ToString(DateTime.Now.Hour) + "." + Convert.ToString(DateTime.Now.Minute) + "." + Convert.ToString(DateTime.Now.Second);
                string filename = date + ".pdf";
                try
                {
                    string save = s + @"\Экспорт\";
                    if (!Directory.Exists(save))
                    {
                        Directory.CreateDirectory(save);
                        chartControl1.ExportToPdf(save + filename);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl1.ExportToPdf(xlsStream);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        chartControl1.ExportToPdf(save + filename);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl1.ExportToPdf(xlsStream);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception www)
                {
                    MessageBox.Show(www.Message, "ФАНЗ");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void barButtonItem43_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView1.Columns[0]).ToString() != "")
            {
                string date = "Cоздан в " + Convert.ToString(DateTime.Now.Hour) + "." + Convert.ToString(DateTime.Now.Minute) + "." + Convert.ToString(DateTime.Now.Second);
                string filename = date + ".png";
                try
                {
                    string save = s + @"\Экспорт\";
                    if (!Directory.Exists(save))
                    {
                        Directory.CreateDirectory(save);
                        chartControl1.ExportToImage(save + filename, System.Drawing.Imaging.ImageFormat.Png);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl1.ExportToImage(xlsStream, System.Drawing.Imaging.ImageFormat.Png);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        chartControl1.ExportToImage(save + filename, System.Drawing.Imaging.ImageFormat.Png);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl1.ExportToImage(xlsStream, System.Drawing.Imaging.ImageFormat.Png);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception www)
                {
                    MessageBox.Show(www.Message, "ФАНЗ");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void barButtonItem44_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView1.Columns[0]).ToString() != "")
            {
                string date = "Cоздан в " + Convert.ToString(DateTime.Now.Hour) + "." + Convert.ToString(DateTime.Now.Minute) + "." + Convert.ToString(DateTime.Now.Second);
                string filename = date + ".jpg";
                try
                {
                    string save = s + @"\Экспорт\";
                    if (!Directory.Exists(save))
                    {
                        Directory.CreateDirectory(save);
                        chartControl1.ExportToImage(save + filename, System.Drawing.Imaging.ImageFormat.Jpeg);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl1.ExportToImage(xlsStream, System.Drawing.Imaging.ImageFormat.Jpeg);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        chartControl1.ExportToImage(save + filename, System.Drawing.Imaging.ImageFormat.Jpeg);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl1.ExportToImage(xlsStream, System.Drawing.Imaging.ImageFormat.Jpeg);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception www)
                {
                    MessageBox.Show(www.Message, "ФАНЗ");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion
        #region Диаграмма показателей
        private void barButtonItem45_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView1.Columns[0]).ToString() != "")
            {
                string date = "Cоздан в " + Convert.ToString(DateTime.Now.Hour) + "." + Convert.ToString(DateTime.Now.Minute) + "." + Convert.ToString(DateTime.Now.Second);
                string filename = date + ".xlsx";
                try
                {
                    string save = s + @"\Экспорт\";
                    if (!Directory.Exists(save))
                    {
                        Directory.CreateDirectory(save);
                        chartControl2.ExportToXlsx(save + filename);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl2.ExportToXlsx(xlsStream);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        chartControl2.ExportToXlsx(save + filename);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl2.ExportToXlsx(xlsStream);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception www)
                {
                    MessageBox.Show(www.Message, "ФАНЗ");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void barButtonItem46_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView1.Columns[0]).ToString() != "")
            {
                string date = "Cоздан в " + Convert.ToString(DateTime.Now.Hour) + "." + Convert.ToString(DateTime.Now.Minute) + "." + Convert.ToString(DateTime.Now.Second);
                string filename = date + ".xls";
                try
                {
                    string save = s + @"\Экспорт\";
                    if (!Directory.Exists(save))
                    {
                        Directory.CreateDirectory(save);
                        chartControl2.ExportToXls(save + filename);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl2.ExportToXls(xlsStream);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        chartControl2.ExportToXls(save + filename);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl2.ExportToXls(xlsStream);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }

                catch (Exception www)
                {
                    MessageBox.Show(www.Message, "ФАНЗ");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void barButtonItem47_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView1.Columns[0]).ToString() != "")
            {
                string date = "Cоздан в " + Convert.ToString(DateTime.Now.Hour) + "." + Convert.ToString(DateTime.Now.Minute) + "." + Convert.ToString(DateTime.Now.Second);
                string filename = date + ".pdf";
                try
                {
                    string save = s + @"\Экспорт\";
                    if (!Directory.Exists(save))
                    {
                        Directory.CreateDirectory(save);
                        chartControl2.ExportToPdf(save + filename);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl2.ExportToPdf(xlsStream);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        chartControl2.ExportToPdf(save + filename);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl2.ExportToPdf(xlsStream);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception www)
                {
                    MessageBox.Show(www.Message, "ФАНЗ");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void barButtonItem48_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView1.Columns[0]).ToString() != "")
            {
                string date = "Cоздан в " + Convert.ToString(DateTime.Now.Hour) + "." + Convert.ToString(DateTime.Now.Minute) + "." + Convert.ToString(DateTime.Now.Second);
                string filename = date + ".png";
                try
                {
                    string save = s + @"\Экспорт\";
                    if (!Directory.Exists(save))
                    {
                        Directory.CreateDirectory(save);
                        chartControl2.ExportToImage(save + filename, System.Drawing.Imaging.ImageFormat.Png);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl2.ExportToImage(xlsStream, System.Drawing.Imaging.ImageFormat.Png);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        chartControl2.ExportToImage(save + filename, System.Drawing.Imaging.ImageFormat.Png);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl2.ExportToImage(xlsStream, System.Drawing.Imaging.ImageFormat.Png);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception www)
                {
                    MessageBox.Show(www.Message, "ФАНЗ");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void barButtonItem49_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView1.Columns[0]).ToString() != "")
            {
                string date = "Cоздан в " + Convert.ToString(DateTime.Now.Hour) + "." + Convert.ToString(DateTime.Now.Minute) + "." + Convert.ToString(DateTime.Now.Second);
                string filename = date + ".jpg";
                try
                {
                    string save = s + @"\Экспорт\";
                    if (!Directory.Exists(save))
                    {
                        Directory.CreateDirectory(save);
                        chartControl2.ExportToImage(save + filename, System.Drawing.Imaging.ImageFormat.Jpeg);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl2.ExportToImage(xlsStream, System.Drawing.Imaging.ImageFormat.Jpeg);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        chartControl2.ExportToImage(save + filename, System.Drawing.Imaging.ImageFormat.Jpeg);
                        FileStream xlsStream = new FileStream(save + filename, FileMode.Create);
                        chartControl2.ExportToImage(xlsStream, System.Drawing.Imaging.ImageFormat.Jpeg);
                        xlsStream.Close();
                        MessageBox.Show("Сохранение выполнено успешно", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception www)
                {
                    MessageBox.Show(www.Message, "ФАНЗ");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion

        private void textBox26_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(textBox26, "Ввод данного показателя требуется осуществлять в %.");
        }

        private void textBox26_TextAlignChanged(object sender, EventArgs e)
        {
            if (Convert.ToDouble(textBox26.Text) > 1 && Convert.ToDouble(textBox26.Text) < 100)
            {
                double temp = Convert.ToDouble(textBox26.Text) / 100;
                textBox26.Text = Convert.ToString(temp);
            }
        }

        private void textBox26_Leave(object sender, EventArgs e)
        {
            if ((Convert.ToDouble(textBox26.Text) > 100))
            {
                MessageBox.Show("Некорректное значение. Было введено значение больше 100.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                textBox26.Clear();
                return;
            }
        }

        private void Дин_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                Reports report = new Reports();
                report.ShowDialog();
            }
            catch (Exception r)
            {
                MessageBox.Show(r.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void barButtonItem50_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            groupControl1.Height = 271;
            groupControl1.Width = 408;

            groupControl2.Height = 128;
            groupControl2.Width = 408;

            groupControl3.Height = 176;
            groupControl3.Width = 408;

            groupControl4.Height = 89;
            groupControl4.Width = 408;

            groupControl5.Height = 250;
            groupControl5.Width = 408;
        }
        private void barButtonItem51_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            groupControl1.Height = 20;
            groupControl2.Height = 20;
            groupControl3.Height = 20;
            groupControl4.Height = 20;
            groupControl5.Height = 20;

            groupControl1.Width = 215;
            groupControl2.Width = 215;
            groupControl3.Width = 215;
            groupControl4.Width = 215;
            groupControl5.Width = 215;
        }
        private void barButtonItem53_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            flowLayoutPanel1.Visible = false;
        }

        private void barButtonItem54_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            flowLayoutPanel1.Visible = true;
        }
        private void gridControl3_MouseHover(object sender, EventArgs e)
        {
            gridControl3.Focus();
        }
        #region Собственные настройки стиля строк и ячеек в таблице Результаты оптимизации
        private void gridView2_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle == 0 || e.RowHandle == 4 || e.RowHandle == 5 || e.RowHandle == 11 || e.RowHandle == 13 || e.RowHandle == 14 || e.RowHandle == 15 || e.RowHandle == 20 || e.RowHandle == 21 || e.RowHandle == 25 || e.RowHandle == 27 || e.RowHandle == 28)
            {
                e.HighPriority = true;
                e.Appearance.BackColor = Color.White;
            }
            if (e.RowHandle == 28 || e.RowHandle == 29)
            {
                e.HighPriority = true;
                e.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
                e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                e.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            }
            if (e.RowHandle == 28)
            {
                e.HighPriority = true;
                e.Appearance.BackColor = Color.LightSteelBlue;
                e.Appearance.BackColor2 = Color.White;
            }
            if (e.RowHandle == 12 || e.RowHandle == 26)
            {
                e.HighPriority = true;
                e.Appearance.BackColor = Color.LightBlue;
                e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;    
            }
            if (e.RowHandle == 3 || e.RowHandle == 10 || e.RowHandle == 19 || e.RowHandle == 24)
            {
                e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            }
        }
        private void gridView2_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            if (e.RowHandle == 0 || e.RowHandle == 3 || e.RowHandle == 5 || e.RowHandle == 10 || e.RowHandle == 12 || e.RowHandle == 14 || e.RowHandle == 15 || e.RowHandle == 19 || e.RowHandle == 21 || e.RowHandle == 24 || e.RowHandle == 26 || e.RowHandle == 28 || e.RowHandle == 29)
            {
                e.Appearance.Font = new Font("Microsoft Sans Serif", 9, FontStyle.Bold);
            }
        }
        #endregion
        #region Масштабирование шрифтов в таблице Результатах оптимизации
        private void zoomTrackBarControl1_EditValueChanged(object sender, EventArgs e)
        {
            const float defaultFontSize = 9;
            float fontSize = defaultFontSize;
            fontSize += Convert.ToInt32(zoomTrackBarControl1.EditValue);
            Font fnt = new Font(gridView2.Appearance.Row.Font.Name, fontSize, gridView2.Appearance.Row.Font.Style);
            gridView2.Appearance.HeaderPanel.Font = fnt;
            gridView2.Appearance.Row.Font = fnt;
        }
        #endregion
        #region Пользовательские настройки Readonly для строк в таблице Результаты оптимизации
        private void gridView2_CustomRowCellEdit(object sender, DevExpress.XtraGrid.Views.Grid.CustomRowCellEditEventArgs e)
        {
            var repositoryItemTextEditReadOnly = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
            repositoryItemTextEditReadOnly.Name = "repositoryItemTextEditReadOnly";
            repositoryItemTextEditReadOnly.ReadOnly = true;

            if (e.RowHandle == 0 || e.RowHandle == 3 || e.RowHandle == 4 || e.RowHandle == 5 || e.RowHandle == 10 || e.RowHandle == 11 || e.RowHandle == 12 || e.RowHandle == 13 || e.RowHandle == 14 || e.RowHandle == 15 || e.RowHandle == 19 || e.RowHandle == 20
             || e.RowHandle == 21 || e.RowHandle == 24 || e.RowHandle == 25 || e.RowHandle == 26 || e.RowHandle == 27 || e.RowHandle == 32 || e.RowHandle == 34 || e.RowHandle == 37)
            {
                e.RepositoryItem = repositoryItemTextEditReadOnly;
            }
        }
        #endregion
        #region Автонумерация -1-ой столбца в таблице Результаты оптимизации
        private void gridView2_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator && e.RowHandle >= 0)
            {
                e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }
        }
        private void gridView2_RowCountChanged(object sender, EventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView gridView = ((DevExpress.XtraGrid.Views.Grid.GridView)sender);
            if (!gridView.GridControl.IsHandleCreated) return;
            Graphics gr = Graphics.FromHwnd(gridView.GridControl.Handle);
            SizeF size = gr.MeasureString(gridView.RowCount.ToString(), gridView.PaintAppearance.Row.GetFont());
            gridView.IndicatorWidth = Convert.ToInt32(size.Width + 0.999f)
                + DevExpress.XtraGrid.Views.Grid.Drawing.GridPainter.Indicator.ImageSize.Width + 10;
        }
        #endregion
        #region Пользовательские настройки высоты строк в таблице Результаты оптимизации
        private void gridView2_CalcRowHeight(object sender, DevExpress.XtraGrid.Views.Grid.RowHeightEventArgs e)
        {
            if (e.RowHandle == 28)
            {
                e.RowHeight = 40;
            }
            if (e.RowHandle == 29)
            {
                e.RowHeight = 40;
            }
            if (e.RowHandle == 3 || e.RowHandle == 10 || e.RowHandle == 12 || e.RowHandle == 19 || e.RowHandle == 24 || e.RowHandle == 26
                || e.RowHandle == 32 || e.RowHandle == 34 || e.RowHandle == 37 || e.RowHandle == 38 || e.RowHandle == 39)
            {
                e.RowHeight = 20;
            }
        }
        private void gridControl2_MouseHover(object sender, EventArgs e)
        {
            gridControl2.Focus(); 
        }
        #endregion
        #region Всплывающие подсказки для таблицы Результаты оптимизации
        private void toolTipController1_GetActiveObjectInfo(object sender, DevExpress.Utils.ToolTipControllerGetActiveObjectInfoEventArgs e)
        {
            if (e.Info == null && e.SelectedControl == gridControl2)
            {
                GridView view = gridControl2.FocusedView as GridView;
                GridHitInfo info = view.CalcHitInfo(e.ControlMousePosition);
                try
                {
                    if (info.InRowCell)
                    {
                        string text = view.GetRowCellDisplayText(info.RowHandle, info.Column);
                        string cellKey = info.RowHandle.ToString() + " - " + info.Column.ToString();
                        if (info.RowHandle == 3 && info.Column == gridView2.Columns[1])
                        {
                            e.Info = new ToolTipControlInfo(cellKey, Convert.ToString(gridView2.GetRowCellValue(1, gridView2.Columns[1]) + " + " + Convert.ToString(gridView2.GetRowCellValue(2, gridView2.Columns[1]))), "ФАНЗ");
                        }
                        if (info.RowHandle == 10 && info.Column == gridView2.Columns[1])
                        {
                            e.Info = new ToolTipControlInfo(cellKey, Convert.ToString(gridView2.GetRowCellValue(6, gridView2.Columns[1]) + " + " + Convert.ToString(gridView2.GetRowCellValue(7, gridView2.Columns[1]))
                                + " + " + Convert.ToString(gridView2.GetRowCellValue(8, gridView2.Columns[1]) + " + " + Convert.ToString(gridView2.GetRowCellValue(9, gridView2.Columns[1])))), "ФАНЗ");
                        }
                        if (info.RowHandle == 12 && info.Column == gridView2.Columns[1])
                        {
                            e.Info = new ToolTipControlInfo(cellKey, Convert.ToString(gridView2.GetRowCellValue(3, gridView2.Columns[1]) + " + " + Convert.ToString(gridView2.GetRowCellValue(10, gridView2.Columns[1]))), "ФАНЗ");
                        }
                        if (info.RowHandle == 19 && info.Column == gridView2.Columns[1])
                        {
                            e.Info = new ToolTipControlInfo(cellKey, Convert.ToString(gridView2.GetRowCellValue(16, gridView2.Columns[1]) + " + " + Convert.ToString(gridView2.GetRowCellValue(17, gridView2.Columns[1]))
                                + " + " + Convert.ToString(gridView2.GetRowCellValue(18, gridView2.Columns[1]))), "ФАНЗ");
                        }
                        if (info.RowHandle == 24 && info.Column == gridView2.Columns[1])
                        {
                            e.Info = new ToolTipControlInfo(cellKey, Convert.ToString(gridView2.GetRowCellValue(22, gridView2.Columns[1]) + " + " + Convert.ToString(gridView2.GetRowCellValue(23, gridView2.Columns[1]))), "ФАНЗ");
                        }
                        if (info.RowHandle == 26 && info.Column == gridView2.Columns[1])
                        {
                            e.Info = new ToolTipControlInfo(cellKey, Convert.ToString(gridView2.GetRowCellValue(19, gridView2.Columns[1]) + " + " + Convert.ToString(gridView2.GetRowCellValue(24, gridView2.Columns[1]))),"ФАНЗ");
                        }
                        if (info.RowHandle == 32 && info.Column == gridView2.Columns[1])
                        {
                            e.Info = new ToolTipControlInfo(cellKey, Convert.ToString(gridView2.GetRowCellValue(30, gridView2.Columns[1]) + " - " + Convert.ToString(gridView2.GetRowCellValue(31, gridView2.Columns[1]))), "ФАНЗ");
                        }
                        if (info.RowHandle == 34 && info.Column == gridView2.Columns[1])
                        {
                            e.Info = new ToolTipControlInfo(cellKey, Convert.ToString(gridView2.GetRowCellValue(32, gridView2.Columns[1]) + " - " + Convert.ToString(gridView2.GetRowCellValue(33, gridView2.Columns[1]))), "ФАНЗ");
                        }
                        if (info.RowHandle == 37 && info.Column == gridView2.Columns[1])
                        {
                            e.Info = new ToolTipControlInfo(cellKey, Convert.ToString("(" + gridView2.GetRowCellValue(34, gridView2.Columns[1]) + " + " + Convert.ToString(gridView2.GetRowCellValue(35, gridView2.Columns[1])) + ")"
                                + " - " + Convert.ToString(gridView2.GetRowCellValue(36, gridView2.Columns[1]))), "ФАНЗ");
                        }
                        if (info.RowHandle == 39 && info.Column == gridView2.Columns[1])
                        {
                            e.Info = new ToolTipControlInfo(cellKey, "Показатель должен быть равен значению нераспределенной прибыли из бухгалтерского баланса!", "ФАНЗ");
                        }
                    }
                }
                catch (Exception ez)
                {
                    MessageBox.Show(ez.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        #endregion

        #region Собственные настройки стиля строк и ячеек в таблице Результаты расчетов
        private void gridView3_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if (e.RowHandle == 0 || e.RowHandle == 10 || e.RowHandle == 14 || e.RowHandle == 18 || e.RowHandle == 24)
            {
                e.Appearance.Font = new Font("Microsoft Sans Serif", 9, FontStyle.Bold);
            }
        }

        private void gridView3_RowStyle(object sender, RowStyleEventArgs e)
        {
            if (e.RowHandle == 0 || e.RowHandle == 10 || e.RowHandle == 14 || e.RowHandle == 18 || e.RowHandle == 24)
            {
                e.HighPriority = true;
                e.Appearance.BackColor = Color.LightBlue;
            }
        }
        #endregion
        #region Автонумерация -1-ой столбца в таблице Результаты расчетов
        private void gridView3_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator && e.RowHandle >= 0)
            {
                e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }
        }
        private void gridView3_RowCountChanged(object sender, EventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView gridView = ((DevExpress.XtraGrid.Views.Grid.GridView)sender);
            if (!gridView.GridControl.IsHandleCreated) return;
            Graphics gr = Graphics.FromHwnd(gridView.GridControl.Handle);
            SizeF size = gr.MeasureString(gridView.RowCount.ToString(), gridView.PaintAppearance.Row.GetFont());
            gridView.IndicatorWidth = Convert.ToInt32(size.Width + 0.999f)
                + DevExpress.XtraGrid.Views.Grid.Drawing.GridPainter.Indicator.ImageSize.Width + 10;
        }
        #endregion
        #region Отрисовка цветом ячеек в таблице Результаты расчетов
        private void gridView3_CustomDrawCell(object sender, RowCellCustomDrawEventArgs e)
        {
            GridView currentView = sender as GridView;
            int o = 0;
            if (e.RowHandle == currentView.FocusedRowHandle) return;

            if (e.Column == gridView3.Columns[6])
            {
                for (int y = 0; y < gridView3.RowCount; y++)
                {
                    if (gridView3.GetRowCellValue(y, gridView3.Columns[6]).ToString() == string.Empty)
                    {
                        o++;
                    }
                }
                if (o > 8)
                {
                    return;
                }
                else
                {
                    if ((e.RowHandle != 0 && e.RowHandle != 1 && e.RowHandle != 2 && e.RowHandle != 10 && e.RowHandle != 14 &&
                        e.RowHandle != 15 && e.RowHandle != 18 && e.RowHandle != 24))
                    {
                        if (Math.Abs((Convert.ToDouble(e.CellValue))) > 10)
                        {
                            e.Appearance.Options.UseBackColor = true;
                            e.Appearance.BackColor = Color.DarkSalmon;
                            e.Appearance.BackColor2 = Color.Empty;
                        }
                    }
                }
            }
        }
        #endregion
        #region Масштабирование шрифтов в таблице Результаты расчетов
        private void zoomTrackBarControl2_EditValueChanged(object sender, EventArgs e)
        {
            const float defaultFontSize = 9;
            float fontSize = defaultFontSize;
            fontSize += Convert.ToInt32(zoomTrackBarControl2.EditValue);
            Font fnt = new Font(gridView3.Appearance.Row.Font.Name, fontSize, gridView3.Appearance.Row.Font.Style);
            gridView3.Appearance.HeaderPanel.Font = fnt;
            gridView3.Appearance.Row.Font = fnt;
        }
        #endregion     

        #region Печать таблицы результов расчетов
        private void barButtonItem52_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView3.RowCount != 0)
            {
                if (!gridControl3.IsPrintingAvailable)
                {
                    MessageBox.Show("The 'DevExpress.XtraPrinting' library is not found", "Error");
                    return;
                }
                gridControl3.Print();
            }
            else
            {
                MessageBox.Show("Таблица пустая, печать отменена.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion
        #region Предпросмотр таблицы расчетов
        private void barButtonItem55_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView3.RowCount != 0)
            {
                if (!gridControl3.IsPrintingAvailable)
                {
                    MessageBox.Show("The 'DevExpress.XtraPrinting' library is not found", "Error");
                    return;
                }
                gridControl3.ShowPrintPreview();
            }
            else
            {
                MessageBox.Show("Таблица пустая, печать отменена.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }          
        }
        #endregion

        #region Кастомные настройки для параметров печати по 3 таблицам
        private void gridView3_PrintInitialize(object sender, PrintInitializeEventArgs e)
        {
            PrintingSystemBase pb = e.PrintingSystem as PrintingSystemBase;
            pb.PageSettings.PaperKind = System.Drawing.Printing.PaperKind.A4;
            pb.PageSettings.Landscape = true;
            pb.PageSettings.BottomMargin = 10;
            pb.PageSettings.TopMargin = 10;
            pb.PageSettings.RightMargin = 10;
            pb.PageSettings.LeftMargin = 10;            
        }
        private void gridView2_PrintInitialize(object sender, PrintInitializeEventArgs e)
        {
            PrintingSystemBase pb = e.PrintingSystem as PrintingSystemBase;
            pb.PageSettings.PaperKind = System.Drawing.Printing.PaperKind.A4;
            pb.PageSettings.BottomMargin = 10;
            pb.PageSettings.TopMargin = 10;
            pb.PageSettings.RightMargin = 10;
            pb.PageSettings.LeftMargin = 10;    
        }
        private void gridView1_PrintInitialize(object sender, PrintInitializeEventArgs e)
        {
            PrintingSystemBase pb = e.PrintingSystem as PrintingSystemBase;
            pb.PageSettings.PaperKind = System.Drawing.Printing.PaperKind.A4;
            pb.PageSettings.BottomMargin = 10;
            pb.PageSettings.TopMargin = 10;
            pb.PageSettings.RightMargin = 10;
            pb.PageSettings.LeftMargin = 10; 
        }
        //hide panel
        private void barButtonItem56_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            flowLayoutPanel1.Visible = false;
        }
        #endregion
        #region Обновить диаграмму показателей
        private void barButtonItem57_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView1.Columns[2]).ToString() != "")
            {
                chartControl2.Series[0].Points.Clear();
                chartControl2.Series[1].Points.Clear();
                chartControl2.Annotations.Clear();
                string type = chartControl2.Diagram.GetType().ToString();
                string type3d = "DevExpress.XtraCharts.XYDiagram3D";
                if (type == type3d)
                {
                    XYDiagram3D diagram2 = (XYDiagram3D)chartControl2.Diagram;
                    diagram2.AxisY.VisualRange.Auto = true;
                    diagram2.AxisX.VisualRange.Auto = true;
                    ((XYDiagram3D)chartControl2.Diagram).RuntimeZooming = true;
                    ((XYDiagram3D)chartControl2.Diagram).RuntimeScrolling = true;
                    ((XYDiagram3D)chartControl2.Diagram).RuntimeRotation = false;
                    chartControl2.Annotations.AddTextAnnotation("Показатели", "1.   Выручка\n2.   Валовая прибыль (убыток)\n3.   Чистая прибыль (убыток)\n4.   Нераспределенная прибыль\n5.   Основные средства\n6.   Прочие внеоборотные активы\n7.   Запасы\n8.   Дебиторская задолженность\n9.   Денежные средства\n10.  Прочие текущие активы\n11.  Уставный капитал\n12.  Фонды и резервы\n13.  Долгосрочные обязательства\n14.  Краткосрочные обязательства").TextAlignment = StringAlignment.Near;
                    for (int i = 0; i < chartControl2.Annotations.Count; i++)
                    {
                        chartControl2.Annotations[i].RuntimeMoving = true;
                        chartControl2.Annotations[i].RuntimeResizing = true;
                        chartControl2.Annotations[i].RuntimeRotation = true;
                    }
                }
                else
                {
                    XYDiagram diagram = (XYDiagram)chartControl2.Diagram;
                    diagram.AxisX.Visible = true;
                    diagram.AxisY.VisualRange.Auto = true;
                    diagram.AxisX.VisualRange.Auto = true;
                }
                #region Само построение
                chartControl2.Series[0].Points.Add(new SeriesPoint(1, Convert.ToDouble(gridView1.GetRowCellValue(0, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(2, Convert.ToDouble(gridView1.GetRowCellValue(1, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(3, Convert.ToDouble(gridView1.GetRowCellValue(2, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(4, Convert.ToDouble(gridView1.GetRowCellValue(3, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(5, Convert.ToDouble(gridView1.GetRowCellValue(4, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(6, Convert.ToDouble(gridView1.GetRowCellValue(5, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(7, Convert.ToDouble(gridView1.GetRowCellValue(6, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(8, Convert.ToDouble(gridView1.GetRowCellValue(7, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(9, Convert.ToDouble(gridView1.GetRowCellValue(8, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(10, Convert.ToDouble(gridView1.GetRowCellValue(9, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(11, Convert.ToDouble(gridView1.GetRowCellValue(10, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(12, Convert.ToDouble(gridView1.GetRowCellValue(11, gridView1.Columns[2]))));

                chartControl2.Series[0].Points.Add(new SeriesPoint(13, Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(14, Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[2]))));

                chartControl2.Series[1].Points.Add(new SeriesPoint(1, Convert.ToDouble(gridView1.GetRowCellValue(0, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(2, Convert.ToDouble(gridView1.GetRowCellValue(1, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(3, Convert.ToDouble(gridView1.GetRowCellValue(2, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(4, Convert.ToDouble(gridView1.GetRowCellValue(3, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(5, Convert.ToDouble(gridView1.GetRowCellValue(4, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(6, Convert.ToDouble(gridView1.GetRowCellValue(5, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(7, Convert.ToDouble(gridView1.GetRowCellValue(6, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(8, Convert.ToDouble(gridView1.GetRowCellValue(7, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(9, Convert.ToDouble(gridView1.GetRowCellValue(8, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(10, Convert.ToDouble(gridView1.GetRowCellValue(9, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(11, Convert.ToDouble(gridView1.GetRowCellValue(10, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(12, Convert.ToDouble(gridView1.GetRowCellValue(11, gridView1.Columns[3]))));

                chartControl2.Series[1].Points.Add(new SeriesPoint(13, Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(14, Convert.ToDouble(gridView1.GetRowCellValue(13, gridView1.Columns[3]))));
                #endregion
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion
        #region Диаграмма показателей в 3D
        private void barButtonItem58_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView2.Columns[2]).ToString() != "")
            {
                chartControl2.Series[0].Points.Clear();
                chartControl2.Series[1].Points.Clear();
                chartControl2.Annotations.Clear();
                string type = chartControl2.Diagram.GetType().ToString();
                string type2d = "DevExpress.XtraCharts.XYDiagram";
                if (type == type2d)
                {
                    chartControl2.Series[0].ChangeView(ViewType.ManhattanBar);
                    chartControl2.Series[1].ChangeView(ViewType.ManhattanBar);
                    XYDiagram3D diagram2 = (XYDiagram3D)chartControl2.Diagram;
                    diagram2.AxisY.VisualRange.Auto = true;
                    diagram2.AxisX.VisualRange.Auto = true;
                    ((XYDiagram3D)chartControl2.Diagram).RuntimeZooming = true;
                    ((XYDiagram3D)chartControl2.Diagram).RuntimeScrolling = true;
                    ((XYDiagram3D)chartControl2.Diagram).RuntimeRotation = false;
                    #region Название меток по оси Х
                    chartControl2.Annotations.AddTextAnnotation("Показатели", "1.   Выручка\n2.   Валовая прибыль (убыток)\n3.   Чистая прибыль (убыток)\n4.   Нераспределенная прибыль\n5.   Основные средства\n6.   Прочие внеоборотные активы\n7.   Запасы\n8.   Дебиторская задолженность\n9.   Денежные средства\n10.  Прочие текущие активы\n11.  Уставный капитал\n12.  Фонды и резервы\n13.  Долгосрочные обязательства\n14.  Краткосрочные обязательства").TextAlignment = StringAlignment.Near;
                    #endregion
                }
                else
                {
                    XYDiagram3D diagram2 = (XYDiagram3D)chartControl2.Diagram;
                    diagram2.AxisY.VisualRange.Auto = true;
                    diagram2.AxisX.VisualRange.Auto = true;
                    ((XYDiagram3D)chartControl2.Diagram).RuntimeZooming = true;
                    ((XYDiagram3D)chartControl2.Diagram).RuntimeScrolling = true;
                    ((XYDiagram3D)chartControl2.Diagram).RuntimeRotation = false;
                    #region Название меток по оси Х
                    chartControl2.Annotations.AddTextAnnotation("Показатели", "1.   Выручка\n2.   Валовая прибыль (убыток)\n3.   Чистая прибыль (убыток)\n4.   Нераспределенная прибыль\n5.   Основные средства\n6.   Прочие внеоборотные активы\n7.   Запасы\n8.   Дебиторская задолженность\n9.   Денежные средства\n10.  Прочие текущие активы\n11.  Уставный капитал\n12.  Фонды и резервы\n13.  Долгосрочные обязательства\n14.  Краткосрочные обязательства").TextAlignment = StringAlignment.Near;
                    for (int i = 0; i < chartControl2.Annotations.Count; i++)
                    {
                        chartControl2.Annotations[i].RuntimeMoving = true;
                        chartControl2.Annotations[i].RuntimeResizing = true;
                        chartControl2.Annotations[i].RuntimeRotation = true;
                    }
                    #endregion
                }
                #region Само построение
                chartControl2.Series[0].Points.Add(new SeriesPoint(1, Convert.ToDouble(gridView1.GetRowCellValue(0, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(2, Convert.ToDouble(gridView1.GetRowCellValue(1, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(3, Convert.ToDouble(gridView1.GetRowCellValue(2, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(4, Convert.ToDouble(gridView1.GetRowCellValue(3, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(5, Convert.ToDouble(gridView1.GetRowCellValue(4, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(6, Convert.ToDouble(gridView1.GetRowCellValue(5, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(7, Convert.ToDouble(gridView1.GetRowCellValue(6, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(8, Convert.ToDouble(gridView1.GetRowCellValue(7, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(9, Convert.ToDouble(gridView1.GetRowCellValue(8, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(10, Convert.ToDouble(gridView1.GetRowCellValue(9, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(11, Convert.ToDouble(gridView1.GetRowCellValue(10, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(12, Convert.ToDouble(gridView1.GetRowCellValue(11, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(13, Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(14, Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[2]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(1, Convert.ToDouble(gridView1.GetRowCellValue(0, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(2, Convert.ToDouble(gridView1.GetRowCellValue(1, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(3, Convert.ToDouble(gridView1.GetRowCellValue(2, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(4, Convert.ToDouble(gridView1.GetRowCellValue(3, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(5, Convert.ToDouble(gridView1.GetRowCellValue(4, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(6, Convert.ToDouble(gridView1.GetRowCellValue(5, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(7, Convert.ToDouble(gridView1.GetRowCellValue(6, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(8, Convert.ToDouble(gridView1.GetRowCellValue(7, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(9, Convert.ToDouble(gridView1.GetRowCellValue(8, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(10, Convert.ToDouble(gridView1.GetRowCellValue(9, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(11, Convert.ToDouble(gridView1.GetRowCellValue(10, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(12, Convert.ToDouble(gridView1.GetRowCellValue(11, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(13, Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(14, Convert.ToDouble(gridView1.GetRowCellValue(13, gridView1.Columns[3]))));
                #endregion
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion
        #region Диаграмма показателей в 2D
        private void barButtonItem59_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView1.Columns[2]).ToString() != "")
            {
                chartControl2.Series[0].Points.Clear();
                chartControl2.Series[1].Points.Clear();
                chartControl2.Annotations.Clear();
                string type = chartControl2.Diagram.GetType().ToString();
                string type3d = "DevExpress.XtraCharts.XYDiagram3D";
                if (type == type3d)
                {
                    chartControl2.Series[0].ChangeView(ViewType.Bar);
                    chartControl2.Series[1].ChangeView(ViewType.Bar);
                    XYDiagram diagram2 = (XYDiagram)chartControl2.Diagram;
                    diagram2.AxisX.Visible = true;
                    diagram2.AxisY.VisualRange.Auto = true;
                    diagram2.AxisX.VisualRange.Auto = true;
                }
                else
                {
                    XYDiagram diagram2 = (XYDiagram)chartControl2.Diagram;
                    diagram2.AxisX.Visible = true;
                    diagram2.AxisY.VisualRange.Auto = true;
                    diagram2.AxisX.VisualRange.Auto = true;
                }
                #region Само построение
                chartControl2.Series[0].Points.Add(new SeriesPoint(1, Convert.ToDouble(gridView1.GetRowCellValue(0, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(2, Convert.ToDouble(gridView1.GetRowCellValue(1, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(3, Convert.ToDouble(gridView1.GetRowCellValue(2, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(4, Convert.ToDouble(gridView1.GetRowCellValue(3, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(5, Convert.ToDouble(gridView1.GetRowCellValue(4, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(6, Convert.ToDouble(gridView1.GetRowCellValue(5, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(7, Convert.ToDouble(gridView1.GetRowCellValue(6, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(8, Convert.ToDouble(gridView1.GetRowCellValue(7, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(9, Convert.ToDouble(gridView1.GetRowCellValue(8, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(10, Convert.ToDouble(gridView1.GetRowCellValue(9, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(11, Convert.ToDouble(gridView1.GetRowCellValue(10, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(12, Convert.ToDouble(gridView1.GetRowCellValue(11, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(13, Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[2]))));
                chartControl2.Series[0].Points.Add(new SeriesPoint(14, Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[2]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(1, Convert.ToDouble(gridView1.GetRowCellValue(0, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(2, Convert.ToDouble(gridView1.GetRowCellValue(1, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(3, Convert.ToDouble(gridView1.GetRowCellValue(2, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(4, Convert.ToDouble(gridView1.GetRowCellValue(3, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(5, Convert.ToDouble(gridView1.GetRowCellValue(4, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(6, Convert.ToDouble(gridView1.GetRowCellValue(5, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(7, Convert.ToDouble(gridView1.GetRowCellValue(6, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(8, Convert.ToDouble(gridView1.GetRowCellValue(7, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(9, Convert.ToDouble(gridView1.GetRowCellValue(8, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(10, Convert.ToDouble(gridView1.GetRowCellValue(9, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(11, Convert.ToDouble(gridView1.GetRowCellValue(10, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(12, Convert.ToDouble(gridView1.GetRowCellValue(11, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(13, Convert.ToDouble(gridView1.GetRowCellValue(12, gridView1.Columns[3]))));
                chartControl2.Series[1].Points.Add(new SeriesPoint(14, Convert.ToDouble(gridView1.GetRowCellValue(13, gridView1.Columns[3]))));
                #endregion
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion

        #region Обновить диаграмму отклонений
        private void barButtonItem60_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView1.Columns[2]).ToString() != "")
            {
                chartControl1.Series[0].Points.Clear();
                chartControl1.Series[1].Points.Clear();

                string type = chartControl1.Diagram.GetType().ToString();
                string type3d = "DevExpress.XtraCharts.XYDiagram3D";
                if (type == type3d)
                {
                    XYDiagram3D diagram1 = (XYDiagram3D)chartControl1.Diagram;
                    diagram1.AxisY.VisualRange.Auto = true;
                    diagram1.AxisX.VisualRange.Auto = true;
                    ((XYDiagram3D)chartControl1.Diagram).RuntimeZooming = true;
                    ((XYDiagram3D)chartControl1.Diagram).RuntimeScrolling = true;
                    ((XYDiagram3D)chartControl1.Diagram).RuntimeRotation = false;
                    ((XYDiagram3D)chartControl1.Diagram).ZoomPercent = 100;
                    ((XYDiagram3D)chartControl1.Diagram).AxisY.Label.NumericOptions.Format = NumericFormat.Percent;
                }
                else
                {
                    XYDiagram diagram1 = (XYDiagram)chartControl1.Diagram;
                    diagram1.AxisY.VisualRange.Auto = true;
                    diagram1.AxisX.VisualRange.Auto = true;
                }
                #region Само построение
                chartControl1.Series[0].Points.Add(new SeriesPoint(1, Convert.ToDouble(label36.Text)));
                chartControl1.Series[1].Points.Add(new SeriesPoint(1, Convert.ToDouble(label32.Text)));
                #endregion
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion
        #region Диаграмма отклонений в 3D
        private void barButtonItem61_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView2.Columns[2]).ToString() != "")
            {
                chartControl1.Series[0].Points.Clear();
                chartControl1.Series[1].Points.Clear();

                string type = chartControl1.Diagram.GetType().ToString();
                string type2d = "DevExpress.XtraCharts.XYDiagram";
                if (type == type2d)
                {
                    chartControl1.Series[0].ChangeView(ViewType.FullStackedBar3D);
                    chartControl1.Series[1].ChangeView(ViewType.FullStackedBar3D);
                    XYDiagram3D diagram1 = (XYDiagram3D)chartControl1.Diagram;
                    diagram1.AxisY.VisualRange.Auto = true;
                    diagram1.AxisX.VisualRange.Auto = true;
                    ((XYDiagram3D)chartControl1.Diagram).RuntimeZooming = true;
                    ((XYDiagram3D)chartControl1.Diagram).RuntimeScrolling = true;
                    ((XYDiagram3D)chartControl1.Diagram).RuntimeRotation = false;
                    ((XYDiagram3D)chartControl1.Diagram).ZoomPercent = 100;
                    ((XYDiagram3D)chartControl1.Diagram).AxisY.Label.NumericOptions.Format = NumericFormat.Percent;
                }
                #region Само построение
                chartControl1.Series[0].Points.Add(new SeriesPoint(1, Convert.ToDouble(label36.Text)));
                chartControl1.Series[1].Points.Add(new SeriesPoint(1, Convert.ToDouble(label32.Text)));
                #endregion
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion
        #region Диаграмма отклонений в 2D
        private void barButtonItem62_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView1.Columns[2]).ToString() != "")
            {
                chartControl1.Series[0].Points.Clear();
                chartControl1.Series[1].Points.Clear();

                string type = chartControl1.Diagram.GetType().ToString();
                string type3d = "DevExpress.XtraCharts.XYDiagram3D";
                if (type == type3d)
                {
                    chartControl1.Series[0].ChangeView(ViewType.FullStackedBar);
                    chartControl1.Series[1].ChangeView(ViewType.FullStackedBar);
                    XYDiagram diagram1 = (XYDiagram)chartControl1.Diagram;
                    diagram1.AxisY.VisualRange.Auto = true;
                    diagram1.AxisX.VisualRange.Auto = true;
                }
                #region Само построение
                chartControl1.Series[0].Points.Add(new SeriesPoint(1, Convert.ToDouble(label36.Text)));
                chartControl1.Series[1].Points.Add(new SeriesPoint(1, Convert.ToDouble(label32.Text)));
                #endregion
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion
        #region Обновить диаграмму радар
        private void barButtonItem63_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridView1.RowCount != 0 && gridView1.GetRowCellValue(0, gridView1.Columns[2]).ToString() != "")
            {
                chartControl4.Series[0].Points.Clear();
                #region Само построение
                try
                {
                    diagramma2();
                }
                catch (Exception exx)
                {
                    MessageBox.Show(exx.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                #endregion
            }
            else
            {
                MessageBox.Show("Таблица отклонений не содержит данные.\nНеобходимо загрузить данные или выполнить расчеты.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        #endregion
        #region Масштабирование шрифтов в таблице отклонений
        private void zoomTrackBarControl3_EditValueChanged(object sender, EventArgs e)
        {
            const float defaultFontSize = 9;
            float fontSize = defaultFontSize;
            fontSize += Convert.ToInt32(zoomTrackBarControl3.EditValue);
            Font fnt = new Font(gridView1.Appearance.Row.Font.Name, fontSize, gridView1.Appearance.Row.Font.Style);
            gridView1.Appearance.HeaderPanel.Font = fnt;
            gridView1.Appearance.Row.Font = fnt;
        }
        #endregion

        private void barButtonItem64_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Altman1 al1 = new Altman1();
            al1.Owner = this;
            al1.ShowDialog();
        }

        private void barButtonItem65_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Altman2 al2 = new Altman2();
            al2.Owner = this;
            al2.ShowDialog();
        }

        private void barButtonItem66_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Altman3 al3 = new Altman3();
            al3.Owner = this;
            al3.ShowDialog();
        }

        private void barButtonItem67_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Altman4 al4 = new Altman4();
            al4.Owner = this;
            al4.ShowDialog();
        }

        private void сп_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            HelpAltman al = new HelpAltman();
            al.Show();
        }

        private void barButtonItem68_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Factor fac = new Factor();
            fac.Owner = this;
            fac.ShowDialog();
        }
        private void ImportFact()
        {
            int[] myArr = new int[] { 12, 19, 30, 38 };

            using (SqlCeConnection connect = new SqlCeConnection(@"Data Source = |DataDirectory|\FANZ.sdf"))
            {
                using (SqlCeCommand command = new SqlCeCommand())
                {
                    command.Connection = connect;
                    command.CommandText = "INSERT INTO Altman (Показатель_1, Значение, Дата_1, Sysdate) VALUES (@Показатель_1,@Значение_1,@Дата_1,@sys_1)";

                    command.Parameters.Add(new SqlCeParameter("@Показатель_1", SqlDbType.NVarChar));
                    command.Parameters.Add(new SqlCeParameter("@Значение_1", SqlDbType.Float));
                    command.Parameters.Add(new SqlCeParameter("@Дата_1", SqlDbType.DateTime));
                    command.Parameters.Add(new SqlCeParameter("@sys_1", SqlDbType.NVarChar));
                    connect.Open();
                    for (int i = 0; i < myArr.Length; i++)
                    {
                        if (gridView2.IsDataRow(i))
                        {
                            command.Parameters["@Показатель_1"].Value = gridView2.GetRowCellValue(myArr[i], gridView2.Columns[0]);
                            command.Parameters["@Значение_1"].Value = gridView2.GetRowCellValue(myArr[i], gridView2.Columns[1]);
                            command.Parameters["@Дата_1"].Value = DateTime.Now;
                            command.Parameters["@sys_1"].Value = DateTime.Now.ToString();
                            command.ExecuteNonQuery();
                        }
                    }
                }
            }
        }
        private void InsertFact()
        {
            DialogResult ExportSQL;
            ExportSQL = MessageBox.Show("Сохранить данные для использования в факторном анализе?", "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (ExportSQL == System.Windows.Forms.DialogResult.Yes)
            {
                try
                {
                    ImportFact();
                }
                catch (Exception r)
                {
                    MessageBox.Show(r.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else return;
        }
        /// <summary>
        /// Вызов формы администрирования симплекс-таблиц
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void barButtonItem70_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            SimplexTable main = new SimplexTable();
            main.ShowDialog();
        }
        /// <summary>
        /// Вызов формы руководства пользователя
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void barButtonItem69_ItemClick(object sender, ItemClickEventArgs e)
        {
            UserGuide manual = new UserGuide();
            manual.Show();
        }
    }
}
      