using Kernel;

using MathNet.Numerics.LinearAlgebra.Double;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace Invention
{
    public class Metods
    {
        public db db = new db();
        public db db1 = new db();
        public double[,] simplex;
        double[,] matrs;

        double[,] matrs1 = new double[36, 84];
        double[] vec = new double[35];

        public void GaussTable(DataGridView dt1)
        {
            dt1.DataSource = db.FetchAll("Gauss");
            matrs = new double[dt1.RowCount, dt1.ColumnCount];

            try
            {
                for (int i = 0; i < dt1.RowCount; i++)
                {
                    for (int j = 1; j < dt1.ColumnCount; j++)
                    {
                        matrs[i, j] = Convert.ToDouble(dt1.Rows[i].Cells[j].Value, CultureInfo.InvariantCulture);
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n(Использование букв и символов недопустимо!)");
                return;
            }
        }
        public void GaussTestValue(DataGridView dt1) 
        {
            if (dt1.RowCount != 0)
            {
                try
                {
                    DataGridViewCell Cell1 = dt1.Rows[1].Cells[2];
                    DataGridViewCell Cell2 = dt1.Rows[12].Cells[0];
                    DataGridViewCell Cell3 = dt1.Rows[4].Cells[1];
                    DataGridViewCell Cell4 = dt1.Rows[3].Cells[1];
                    DataGridViewCell Cell5 = dt1.Rows[2].Cells[1];
                    double kipn = 0.63;
                    DataGridViewCell Cell7 = dt1.Rows[6].Cells[6];

                    DataGridViewCell Cell8 = dt1.Rows[11].Cells[8];
                    DataGridViewCell Cell9 = dt1.Rows[14].Cells[5];
                    DataGridViewCell Cell10 = dt1.Rows[13].Cells[8];

                    DataGridViewCell Cell11 = dt1.Rows[17].Cells[14];
                    DataGridViewCell Cell12 = dt1.Rows[19].Cells[8];

                    DataGridViewCell Cell13 = dt1.Rows[26].Cells[4];
                    DataGridViewCell Cell14 = dt1.Rows[27].Cells[6];
                    DataGridViewCell Cell15 = dt1.Rows[25].Cells[11];
                    DataGridViewCell Cell16 = dt1.Rows[28].Cells[8];
                    double osvk = 1;

                    DataGridViewCell Cell17 = dt1.Rows[18].Cells[21];
                    DataGridViewCell Cell18 = dt1.Rows[23].Cells[22];
                    DataGridViewCell Cell19 = dt1.Rows[24].Cells[20];
                    DataGridViewCell Cell20 = dt1.Rows[30].Cells[23];
                    double ba = 9508;
                    double yk = 120;
                    double fot = 2500;
                    double cebfot = 2.5;


                    DataGridViewCell B8 = dt1.Rows[7].Cells[35];
                    DataGridViewCell B10 = dt1.Rows[9].Cells[35];
                    DataGridViewCell B11 = dt1.Rows[10].Cells[35];
                    DataGridViewCell B17 = dt1.Rows[16].Cells[35];
                    DataGridViewCell B34 = dt1.Rows[33].Cells[35];
                    DataGridViewCell B35 = dt1.Rows[34].Cells[35];

                    Cell1.Value = Convert.ToDouble(-0.2);
                    Cell2.Value = Convert.ToDouble(-0.3);
                    Cell3.Value = Convert.ToDouble(1);
                    Cell4.Value = Convert.ToDouble(1);
                    Cell5.Value = Convert.ToDouble(1.5);
                    Cell7.Value = Convert.ToDouble(-1);
                    Cell8.Value = Convert.ToDouble(-2);
                    Cell9.Value = Convert.ToDouble(-1);
                    Cell10.Value = Convert.ToDouble(-1);
                    Cell11.Value = Convert.ToDouble(0.5);
                    Cell12.Value = Convert.ToDouble(0.5);
                    Cell13.Value = Convert.ToDouble(-3);
                    Cell14.Value = Convert.ToDouble(-2);
                    Cell15.Value = Convert.ToDouble(-2);
                    Cell16.Value = Convert.ToDouble(-2);
                    Cell17.Value = Convert.ToDouble(-9508);
                    Cell18.Value = Convert.ToDouble(-0.8);
                    Cell19.Value = Convert.ToDouble(0.5);
                    Cell20.Value = Convert.ToDouble(-0.25);

                    B8.Value = Convert.ToDouble(ba);
                    B10.Value = Convert.ToDouble(yk);
                    B11.Value = Convert.ToDouble(9508);
                    B17.Value = Convert.ToDouble(ba * kipn);
                    B34.Value = Convert.ToDouble(fot * cebfot);
                    B35.Value = Convert.ToDouble(osvk * 9508);

                    //dt1.Rows[1].Cells[2].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[12].Cells[0].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[4].Cells[1].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[3].Cells[1].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[2].Cells[1].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[6].Cells[6].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[11].Cells[8].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[14].Cells[5].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[13].Cells[8].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[17].Cells[14].Style.BackColor = System.Drawing.Color.PaleTurquoise;

                    //dt1.Rows[19].Cells[8].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[26].Cells[4].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[27].Cells[6].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[25].Cells[11].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[28].Cells[8].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[18].Cells[21].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[23].Cells[22].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[24].Cells[20].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[30].Cells[23].Style.BackColor = System.Drawing.Color.PaleTurquoise;

                    //dt1.Rows[7].Cells[35].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[9].Cells[35].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[10].Cells[35].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[16].Cells[35].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[33].Cells[35].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                    //dt1.Rows[34].Cells[35].Style.BackColor = System.Drawing.Color.PaleTurquoise;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex + "Заданы не все показатели!");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Таблица пустая! Нельзя задать тестовые значения.", "Система", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
        }
        public void SolveGauss(DataGridView dt1, DataGridView dt2) 
        {
            if (dt1.RowCount != 0)
            {
                int M = dt1.RowCount;
                int N = dt1.ColumnCount - 1; // поправим из-за столбца с id
                List<string> list = new List<string>();
                ArrayList lis = new ArrayList();

                var matrs = new DenseMatrix(M, N);
                bool toOneUse = true;

                for (int i = 0; i < M; i++)
                {
                    for (int j = 0; j < N; j++)
                    {
                        matrs[i, j] = Convert.ToDouble(dt1.Rows[i].Cells[j].Value, CultureInfo.InvariantCulture);
                    }
                }

                int ch = 0;
                double[] vec = new double[35];
                for (int i = 0; i < M; i++)
                {
                    vec[i] = Convert.ToDouble(dt1[35, i].Value);
                    ++ch;
                }

                var v = new DenseVector(vec);
                int q = v.Count();
                var x = matrs.Solve(v);
                string output = string.Join(" ", x);
                string[] name = {"ТА","КП","СОС","ДС","ДЗ","ВА","ЗЗ","ПК","СК","ФР","ДП","Оср","ПР.ТА","ПР.ВА",
                "ВР","Вал.пр.","КрУр","Пр.прод","Проч.д","Проч.р","ЧП","RСовК","Пр.до_нал.","Себ.прод.","НерПр",
                "d1","d2","d3","d4","d5","d6","d7","d8","d9","d10"};

                String[] words = output.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                dt2.RowCount = N;
                int i1;

                for (i1 = 0; i1 < N; ++i1)
                {
                    double result = Convert.ToDouble(words[i1]);
                    Math.Round(result);
                    dt2.Rows[i1].Cells[0].Value = name[i1];
                    dt2.Rows[i1].Cells[1].Value = Math.Round(result);
                }

                int rang = matrs.Rank();
                var lu = matrs.LU();
                if (lu.Determinant == 0 && toOneUse)
                {
                    MessageBox.Show("Система несовместна и не имеет решений, т.к ее определитель равен " + lu.Determinant + ".",
                    "Исследование СЛАУ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    toOneUse = false;
                }
                if (rang > N && toOneUse)
                {
                    MessageBox.Show("Ранк матрицы= " + matrs.Rank() + " Система не имеет решений.", "Исследование системы", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    toOneUse = false;
                }
                else if (lu.Determinant != 0 && rang <= N)
                {
                    var r = matrs * x - v;
                    string output1 = string.Join(" ", r);
                    String[] word = output1.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    dt2.RowCount = N;
                    int j1;

                    for (j1 = 0; j1 < N; ++j1)
                    {
                        double result1 = Convert.ToDouble(word[j1]);
                        dt2.Rows[j1].Cells[2].Value = result1.ToString("0.##############");
                    }
                }
            }
            else
            {
                MessageBox.Show("Таблица пустая! Нельзя запустить расчет.", "Система", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
        }

        #region Загрузка данных с базы
        public void FillTable(DataGridView dt)
        {
            dt.DataSource = db1.FetchAll1("Simplex");
            simplex = new double[dt.RowCount, dt.ColumnCount];

            try
            {

                for (int i = 0; i < dt.RowCount; i++)
                {
                    for (int j = 1; j < dt.ColumnCount; j++)
                    {
                        simplex[i, j] = Convert.ToDouble(dt.Rows[i].Cells[j].Value, CultureInfo.InvariantCulture);
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n(Использование букв и символов недопустимо!)");
                return;
            }
        }   
        public void Value(DataGridView dt)
        {
            if (dt.RowCount != 0)
            {
                try
                {
                    DataGridViewCell Cell1 = dt.Rows[1].Cells[3];    
                    DataGridViewCell Cell2 = dt.Rows[12].Cells[1];   
                    DataGridViewCell Cell3 = dt.Rows[4].Cells[2];   
                    DataGridViewCell Cell4 = dt.Rows[3].Cells[2];   
                    DataGridViewCell Cell5 = dt.Rows[2].Cells[2];    
                    double kipn = 0.63;                                       
                    DataGridViewCell Cell7 = dt.Rows[6].Cells[7];   

                    DataGridViewCell Cell8 = dt.Rows[11].Cells[9];   
                    DataGridViewCell Cell9 = dt.Rows[14].Cells[6];  
                    DataGridViewCell Cell10 = dt.Rows[13].Cells[9];  

                    DataGridViewCell Cell11 = dt.Rows[17].Cells[15]; 
                    DataGridViewCell Cell12 = dt.Rows[19].Cells[9];  

                    DataGridViewCell Cell13 = dt.Rows[26].Cells[5];  
                    DataGridViewCell Cell14 = dt.Rows[27].Cells[7];  
                    DataGridViewCell Cell15 = dt.Rows[25].Cells[12]; 
                    DataGridViewCell Cell16 = dt.Rows[28].Cells[9];  
                    double osvk = 1;                                         

                    DataGridViewCell Cell17 = dt.Rows[18].Cells[22];  
                    DataGridViewCell Cell18 = dt.Rows[23].Cells[23];  
                    DataGridViewCell Cell19 = dt.Rows[24].Cells[21];  
                    DataGridViewCell Cell20 = dt.Rows[30].Cells[24];  
                    double ba = 9508;                                           
                    double yk = 120;                                           
                    double fot = 2500;                                         
                    double cebfot = 2.5;                                     

                    DataGridViewCell B8 = dt.Rows[7].Cells[83];
                    DataGridViewCell B10 = dt.Rows[9].Cells[83];
                    DataGridViewCell B11 = dt.Rows[10].Cells[83];
                    DataGridViewCell B17 = dt.Rows[16].Cells[83];
                    DataGridViewCell B34 = dt.Rows[33].Cells[83];
                    DataGridViewCell B35 = dt.Rows[34].Cells[83];

                    Cell1.Value = Convert.ToDouble(-0.2);
                    Cell2.Value = Convert.ToDouble(-0.3);
                    Cell3.Value = Convert.ToDouble(1);
                    Cell4.Value = Convert.ToDouble(1);
                    Cell5.Value = Convert.ToDouble(1.5);
                    Cell7.Value = Convert.ToDouble(-1);
                    Cell8.Value = Convert.ToDouble(-2);
                    Cell9.Value = Convert.ToDouble(-1);
                    Cell10.Value = Convert.ToDouble(-1);
                    Cell11.Value = Convert.ToDouble(0.5);
                    Cell12.Value = Convert.ToDouble(0.5);
                    Cell13.Value = Convert.ToDouble(-3);
                    Cell14.Value = Convert.ToDouble(-2);
                    Cell15.Value = Convert.ToDouble(-2);
                    Cell16.Value = Convert.ToDouble(-2);
                    Cell17.Value = Convert.ToDouble(-9508);
                    Cell18.Value = Convert.ToDouble(-0.8);
                    Cell19.Value = Convert.ToDouble(0.5);
                    Cell20.Value = Convert.ToDouble(-0.25);

                    B8.Value = Convert.ToDouble(ba);
                    B10.Value = Convert.ToDouble(yk);
                    B11.Value = Convert.ToDouble(9508);
                    B17.Value = Convert.ToDouble(ba * kipn);
                    B34.Value = Convert.ToDouble(fot * cebfot);
                    B35.Value = Convert.ToDouble(osvk * 9508);

                    //dt.Rows[1].Cells[3].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[12].Cells[1].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[4].Cells[2].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[3].Cells[2].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[2].Cells[2].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[6].Cells[7].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[11].Cells[9].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[14].Cells[6].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[13].Cells[9].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[17].Cells[15].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[19].Cells[9].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[26].Cells[5].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[27].Cells[7].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[25].Cells[12].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[28].Cells[9].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[18].Cells[22].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[23].Cells[23].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[24].Cells[21].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[30].Cells[24].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[7].Cells[83].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[9].Cells[83].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[10].Cells[83].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[16].Cells[83].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[33].Cells[83].Style.BackColor = System.Drawing.Color.Goldenrod;
                    //dt.Rows[34].Cells[83].Style.BackColor = System.Drawing.Color.Goldenrod;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex + "Заданы не все показатели!");
                    return;
                }    
                int[] index_r = new int[] { 0, 5, 7, 8, 9, 10, 15, 20, 21, 22, 23, 29 };

                for (int j = 1; j < 49; j++)            
                {
                    double sum = 0;
                    int index = 0;

                    for (int i = 0; i < index_r.GetLength(0); i++)
                    {
                        index = index_r[i];
                        sum += Convert.ToDouble(dt.Rows[index].Cells[j].Value);
                    }
                    DataGridViewCell sum1 = dt.Rows[35].Cells[j];
                    sum1.Value = -sum;
                    //dt.Rows[35].Cells[j].Style.BackColor = System.Drawing.Color.LightSteelBlue;
                }
            }

            else MessageBox.Show("Таблица пустая! Нельзя задать тестовые значения.", "Система", MessageBoxButtons.OK, MessageBoxIcon.Stop);


        }    
        #endregion
    }
   public class Bazis
    {    
        public int ved = 0;                         
        public int index = 0;                       
        public int free_last = 0;                  
        public int straf = 0;                       
        public int counter = 0;                  
        public int count = 0;
        public int count_straf = 0;

        #region Алгоритм решения М-задачи
        public bool dopustim(DataGridView data)
        {
            free_last = data.ColumnCount - 2;

            for (int i = 0; i < data.RowCount - 2; i++)
            {
                if (Convert.ToDouble(data[free_last, i].Value) < 0)
                {
                    counter++;
                }
            }
            if (counter > 0)
            {
                MessageBox.Show("Найдены отрицательные свободные члены", "Проверка на допустимость", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return true;
            }
            else
            {
                return false;
            }
        }
        public bool optimalnost(DataGridView data)
        {
            free_last = data.ColumnCount - 2;
            int counter = 0;

            for (int i = 1; i < 49; i++)         
            {
                if (Convert.ToDouble(data.Rows[35].Cells[i].Value) < 0)
                {
                    counter++;
                }
            }
            if (counter == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public bool Simpex_res(DataGridView data)
        {
            double temp = 0;              
            double tmp = 0;                
            double max = 0;                
            int mini = 0;                  
            double fix_plus = 0;           
            double fix = 0;                 
            double Rel = 0;                
            double min = 0;                 
            double minimal = 0;

            straf = data.RowCount - 2;      
            ArrayList arr = new ArrayList();
            double srav = double.MinValue;
            for (int f = 1; f < 49; f++)
            {
                if (Convert.ToDouble(data[f, straf].Value) < 0.0)
                {
                    count_straf++;
                }
            }
            if (count_straf > 0)
            {
                for (int k = 1; k < 49; k++)
                {
                    if (Convert.ToDouble(data[k, straf].Value) < 0.0)
                    {
                        temp = Math.Abs(Convert.ToDouble(data[k, straf].Value));
                        arr.Add(temp);
                    }
                }
                int ar = 0;
                for (int i = 0; i < arr.Count; i++)
                {
                    if (Convert.ToDouble(arr[i]) > srav)
                    {
                        srav = Convert.ToDouble(arr[i]);
                        ar = i;
                    }
                }

                for (int k = 1; k < 49; k++)
                {
                    if (Convert.ToDouble(data[k, straf].Value) < 0.0)
                    {
                        max = Math.Abs(Convert.ToDouble(data[k, straf].Value));
                        if (srav.Equals(max))
                        {
                            max = srav;
                            ved = k;
                            break;
                        }
                    }
                }
                arr.Clear();
            }

            for (int i = 0; i < straf; i++)
            {
                if (Convert.ToDouble(data[ved, i].Value) > 0.0)
                {
                    count++;
                }
            }
            if (count == 0)
            {
                MessageBox.Show("В ведущем столбце нет положительных элементов.\nТребуется улучшение столбца.", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                for (int k = 0; k < straf; k++)
                {
                    if (Convert.ToDouble(data[ved, k].Value) < 0.0)
                    {
                        tmp = Math.Abs(Convert.ToDouble(data[ved, k].Value));
                        arr.Add(tmp);
                    }
                }
                if (arr.Count == 0)
                {
                    MessageBox.Show("В ведущем столбце № " + ved.ToString() + " все элементы равны 0.\nДальнейшее решение невозможно", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return true;
                }
                else
                {
                    arr.Sort();                          
                    minimal = Convert.ToDouble(arr[0]);

                    for (int i = 0; i < straf; i++)
                    {
                        fix = Math.Abs(Convert.ToDouble(data[ved, i].Value));

                        if (minimal.Equals(fix))
                        {
                            minimal = fix;
                            mini = i;
                            break;
                        }
                    }
                    fix_plus = minimal;
                    DataGridViewCell step = data[ved, mini];
                    step.Value = fix_plus;

                    arr.Clear();  

                    for (int i = 0; i < straf; i++)
                    {
                        if (Convert.ToDouble(data[free_last, i].Value) == 0.0 && Convert.ToDouble(data[ved, i].Value) > 0.0)
                        {
                            Rel = 0.001 / Convert.ToDouble(data[ved, i].Value);
                            DataGridViewCell Real1 = data[free_last + 1, i];
                            Real1.Value = Rel;
                        }
                        else if (Convert.ToDouble(data[free_last, i].Value) > 0.0 && Convert.ToDouble(data[ved, i].Value) > 0.0)
                        {
                            Rel = Convert.ToDouble(data[free_last, i].Value) / Convert.ToDouble(data[ved, i].Value);
                            DataGridViewCell Real1 = data[free_last + 1, i];
                            Real1.Value = Rel;
                        }
                    }
                }
            }
            else
            {
                for (int i = 0; i < straf; i++)
                {
                    if (Convert.ToDouble(data[free_last, i].Value) == 0.0 && Convert.ToDouble(data[ved, i].Value) > 0.0)
                    {
                        Rel = 0.001 / Convert.ToDouble(data[ved, i].Value);
                        DataGridViewCell Real1 = data[free_last + 1, i];
                        Real1.Value = Rel;
                    }
                    else if (Convert.ToDouble(data[free_last, i].Value) > 0.0 && Convert.ToDouble(data[ved, i].Value) > 0.0)
                    {
                        Rel = Convert.ToDouble(data[free_last, i].Value) / Convert.ToDouble(data[ved, i].Value);
                        DataGridViewCell Real1 = data[free_last + 1, i];
                        Real1.Value = Rel;
                    }
                }
            }

            for (int k = 0; k < straf; k++)
            {
                if (Convert.ToDouble(data[free_last + 1, k].Value) != 1e+20)
                {
                    tmp = Convert.ToDouble(data[free_last + 1, k].Value);
                    arr.Add(tmp);
                }
            }
            arr.Sort();                           
            min = Convert.ToDouble(arr[0]);
            if (min != 1e+20)
            {
                for (int i = 0; i < straf; i++)
                {
                    double current = Convert.ToDouble(data[free_last + 1, i].Value);
                    if (min.Equals(current))
                    {
                        index = i + 1;
                        break;
                    }
                }
                double veduch = Convert.ToDouble(data[ved, index - 1].Value);      
                //data[ved, index - 1].Style.BackColor = Color.LightBlue;
                for (int p = 0; p < straf + 1; p++)
                {
                    DataGridViewCell minimum = data[free_last + 1, p];
                    minimum.Value = 1e+20;
                    //minimum.Style.BackColor = Color.Transparent;
                }
            }
            else if (min == 1e+20)
            {
                MessageBox.Show("Решение неограничено. Дальнейшее выполнение алгоритма невозможно.", "Поиск ведущей строки", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return true;
            }
            arr.Clear();      
            double item1 = 0;
            int M = data.RowCount;
            int N = data.ColumnCount;
            double[,] matrs1 = new double[M, N];

            for (int i = 0; i < M; i++)
            {
                for (int j = 1; j < N; j++)
                {
                    matrs1[i, j] = Convert.ToDouble(data.Rows[i].Cells[j].Value);
                }
            }

            double[,] clone = (double[,])matrs1.Clone();         

            DataGridViewCell x_i = data[0, index - 1];           
            x_i.Value = "x" + ved;                              
            //data[0, index - 1].Style.BackColor = System.Drawing.Color.Goldenrod;

            for (int i = 0; i < M; i++)
            {
                for (int j = 1; j < free_last; j++)
                {
                    if (i != index - 1)
                    {
                        matrs1[i, j] = matrs1[i, j] - (clone[i, ved] * matrs1[index - 1, j] / clone[index - 1, ved]);
                    }

                    double glav = Convert.ToDouble(clone[index - 1, j]); 
                    item1 = glav / clone[index - 1, ved];
                    DataGridViewCell str = data[j, index - 1];
                    str.Value = item1;

                    DataGridViewCell step = data[j, i];      
                    if (matrs1[i, j] != double.NaN)
                    {
                        step.Value = matrs1[i, j];
                    }
                    else
                    {
                        MessageBox.Show("Получены некорректные значения", "Пересчет таблицы");
                        return true;
                    }
                }
            }
            for (int i = 0; i < M; i++)
            {
                if (i != index - 1)
                {
                    matrs1[i, free_last] = matrs1[i, free_last] - (clone[i, ved] * matrs1[index - 1, free_last] / clone[index - 1, ved]);
                }
                double glav = Convert.ToDouble(clone[index - 1, free_last]); 
                item1 = glav / clone[index - 1, ved];
                DataGridViewCell str = data[free_last, index - 1];
                str.Value = item1;

                DataGridViewCell step = data[free_last, i];
                if (matrs1[i, free_last] > 0)
                {
                    step.Value = matrs1[i, free_last];
                }
            }
            return false;
        }
        public void Remove_M(DataGridView data)
        {
            data.Rows.RemoveAt(35);
        }
        public bool dopustim_Z(DataGridView data)
        {
            free_last = data.ColumnCount - 2;

            for (int i = 0; i < data.RowCount - 1; i++)
            {
                if (Convert.ToDouble(data[free_last, i].Value) < 0)
                {
                    counter++;
                }
            }
            if (counter > 0)
            {
                MessageBox.Show("Найдены отрицательные свободные члены", "Проверка на допустимость", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return true;
            }
            else
            {
                return false;
            }
        }
        public bool optimalnost_Z(DataGridView data)
        {
            free_last = data.ColumnCount - 2;
            int counter = 0;

            for (int i = 1; i < 49; i++)         
            {
                if (Convert.ToDouble(data.Rows[35].Cells[i].Value) < 0)
                {
                    counter++;
                }
            }
            if (counter == 0)
            {
                MessageBox.Show("Найдено оптимальное решение.", "Проверка на оптимальность", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return true;
            }
            else
            {
                return false;
            }
        }
 
        public bool Simpex_res_Z(DataGridView data)
        {
            double temp = 0;                
            double tmp = 0;                 
            double max = 0;                
            int mini = 0;                   
            double fix_plus = 0;            
            double fix = 0;                 
            double Rel = 0;                 
            double min = 0;                 
            double minimal = 0;

            straf = data.RowCount - 1;      
            ArrayList arr = new ArrayList();
            string list = string.Empty;
            double srav = double.MinValue;
            for (int f = 1; f < 49; f++)
            {
                if (Convert.ToDouble(data[f, straf].Value) < 0.0)
                {
                    count_straf++;
                }
            }
            if (count_straf > 0)
            {
                for (int k = 1; k < 49; k++)
                {
                    if (Convert.ToDouble(data[k, straf].Value) < 0.0)
                    {
                        temp = Math.Abs(Convert.ToDouble(data[k, straf].Value));
                        arr.Add(temp);
                    }
                }

                int ar = 0;
                for (int i = 0; i < arr.Count; i++)
                {
                    if (Convert.ToDouble(arr[i]) > srav)
                    {
                        srav = Convert.ToDouble(arr[i]);
                        ar = i;
                    }
                }

                for (int k = 1; k < 49; k++)
                {
                    if (Convert.ToDouble(data[k, straf].Value) < 0.0)
                    {
                        max = Math.Abs(Convert.ToDouble(data[k, straf].Value));
                        if (srav.Equals(max))
                        {
                            max = srav;
                            ved = k;
                            break;
                        }
                    }
                }
                arr.Clear();
            }

            for (int i = 0; i < straf; i++)
            {
                if (Convert.ToDouble(data[ved, i].Value) > 0.0)
                {
                    count++;
                }
            }
            if (count == 0)
            {
                MessageBox.Show("В ведущем столбце нет положительных элементов.\nТребуется улучшение столбца.", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                for (int k = 0; k < straf; k++)
                {
                    if (Convert.ToDouble(data[ved, k].Value) < 0.0)
                    {
                        tmp = Math.Abs(Convert.ToDouble(data[ved, k].Value));
                        arr.Add(tmp);
                    }
                }

                if (arr.Count != 0)
                {
                    arr.Sort();                          
                    minimal = Convert.ToDouble(arr[0]);
                }
                else
                {
                    MessageBox.Show("В ведущем столбце № " + ved.ToString() + " все элементы равны 0. \nДальнейшее решение невозможно.", "Система", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return true;
                }
                for (int i = 0; i < straf; i++)
                {
                    fix = Math.Abs(Convert.ToDouble(data[ved, i].Value));

                    if (minimal.Equals(fix))
                    {
                        minimal = fix;
                        mini = i;
                        break;
                    }
                }
                fix_plus = minimal;
                DataGridViewCell step = data[ved, mini];
                step.Value = fix_plus;

                arr.Clear();   

                for (int i = 0; i < straf; i++)
                {
                    if (Convert.ToDouble(data[free_last, i].Value) == 0.0 && Convert.ToDouble(data[ved, i].Value) > 0.0)
                    {
                        Rel = 0.001 / Convert.ToDouble(data[ved, i].Value);
                        DataGridViewCell Real1 = data[free_last + 1, i];
                        Real1.Value = Rel;
                    }
                    else if (Convert.ToDouble(data[free_last, i].Value) > 0.0 && Convert.ToDouble(data[ved, i].Value) > 0.0)
                    {
                        Rel = Convert.ToDouble(data[free_last, i].Value) / Convert.ToDouble(data[ved, i].Value);
                        DataGridViewCell Real1 = data[free_last + 1, i];
                        Real1.Value = Rel;
                    }
                }
            }
            else
            {
                for (int i = 0; i < straf; i++)
                {
                    if (Convert.ToDouble(data[free_last, i].Value) == 0.0 && Convert.ToDouble(data[ved, i].Value) > 0.0)
                    {
                        Rel = 0.001 / Convert.ToDouble(data[ved, i].Value);
                        DataGridViewCell Real1 = data[free_last + 1, i];
                        Real1.Value = Rel;
                    }
                    else if (Convert.ToDouble(data[free_last, i].Value) > 0.0 && Convert.ToDouble(data[ved, i].Value) > 0.0)
                    {
                        Rel = Convert.ToDouble(data[free_last, i].Value) / Convert.ToDouble(data[ved, i].Value);
                        DataGridViewCell Real1 = data[free_last + 1, i];
                        Real1.Value = Rel;
                    }
                }
            }

            for (int k = 0; k < straf; k++)
            {
                if (Convert.ToDouble(data[free_last + 1, k].Value) != 1e+20)
                {
                    tmp = Convert.ToDouble(data[free_last + 1, k].Value);
                    arr.Add(tmp);
                }
            }
            arr.Sort();                          
            min = Convert.ToDouble(arr[0]);
            if (min != 1e+20)
            {
                for (int i = 0; i < straf; i++)
                {
                    double current = Convert.ToDouble(data[free_last + 1, i].Value);
                    if (min.Equals(current))
                    {
                        index = i + 1;                                           
                        break;
                    }
                }
                double veduch = Convert.ToDouble(data[ved, index - 1].Value);                       
                //data[ved, index - 1].Style.BackColor = Color.LightBlue;
                for (int p = 0; p < straf + 1; p++)
                {
                    DataGridViewCell minimum = data[free_last + 1, p];
                    minimum.Value = 1e+20;
                    //minimum.Style.BackColor = Color.Transparent;
                }
            }
            else if (min == 1e+20)
            {
                MessageBox.Show("Решение неограничено. Дальнейшее выполнение алгоритма невозможно.", "Поиск ведущей строки", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return true;
            }
            arr.Clear();      

            double item1 = 0;
            int M = data.RowCount;
            int N = data.ColumnCount;
            double[,] matrs1 = new double[M, N];
            for (int i = 0; i < M; i++)
            {
                for (int j = 1; j < N; j++)
                {
                    matrs1[i, j] = Convert.ToDouble(data.Rows[i].Cells[j].Value);
                }
            }

            double[,] clone = (double[,])matrs1.Clone();   

            DataGridViewCell x_i = data[0, index - 1];           
            x_i.Value = "x" + ved;                              
            //data[0, index - 1].Style.BackColor = System.Drawing.Color.Goldenrod;

            for (int i = 0; i < M; i++)
            {
                for (int j = 1; j < free_last; j++)
                {
                    if (i != index - 1)
                    {
                        matrs1[i, j] = matrs1[i, j] - (clone[i, ved] * matrs1[index - 1, j] / clone[index - 1, ved]);
                    }

                    double glav = Convert.ToDouble(clone[index - 1, j]); 
                    item1 = glav / clone[index - 1, ved];
                    DataGridViewCell str = data[j, index - 1];
                    str.Value = item1;

                    DataGridViewCell step = data[j, i];        
                    if (matrs1[i, j] != double.NaN)
                    {
                        step.Value = matrs1[i, j];
                    }
                    else
                    {
                        MessageBox.Show("Получены некорректные значения", "Пересчет таблицы");
                        return true;
                    }
                }
            }
            for (int i = 0; i < M; i++)
            {
                if (i != index - 1)
                {
                    matrs1[i, free_last] = matrs1[i, free_last] - (clone[i, ved] * matrs1[index - 1, free_last] / clone[index - 1, ved]);
                }
                double glav = Convert.ToDouble(clone[index - 1, free_last]); 
                item1 = glav / clone[index - 1, ved];
                DataGridViewCell str = data[free_last, index - 1];
                str.Value = item1;

                DataGridViewCell step = data[free_last, i];
                if (matrs1[i, free_last] > 0)
                {
                    step.Value = matrs1[i, free_last];
                }
            }
            return false;
        }
        #endregion
    }
}
