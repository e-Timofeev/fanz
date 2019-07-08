using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraBars;
using System.Data.SqlServerCe;
using DevExpress.Utils;
using System.IO;
using DevExpress.XtraPrinting;

namespace Simplex_2.Формы_проекта
{
    public partial class SimplexTable : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public string s = System.Windows.Forms.Application.StartupPath;
        public SimplexTable()
        {
            InitializeComponent();
            gridView2.IndicatorWidth = 34; 
        }
        private void загру_ItemClick(object sender, ItemClickEventArgs e)
        {
            string[] name = {"ТА","КП","СОС","ДС","ДЗ","ВА","ЗЗ","ПК","СК","ФР","ДП","Оср","ПР.ТА","ПР.ВА",
                "ВР","Вал.пр.","КрУр","Пр.прод","Проч.д","Проч.р","ЧП","RСовК","Пр.до_нал.","Себ.прод.","НерПр"};
            
            string[] dop = {"d1","d2","d3","d4","d5","d6","d7","d8","d9","d10","d11","d13","d14","d15","d16","d17","d18","d19","d20","d22","d23","d24","d25"};

            try
            {
                SqlCeConnection conn;
                conn = new SqlCeConnection(@"Data Source = |DataDirectory|\FANZ.sdf");
                conn.Open();
                SqlCeDataAdapter adapter, adapter2;
                adapter = new SqlCeDataAdapter("SELECT x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11, x12, x13, x14, x15, x16, x17, x18, x19, x20, x21, x22, x23, x24, x25 FROM Simplex", conn);
                adapter2 = new SqlCeDataAdapter("SELECT X26, X27, X28, X29, X30, X31, X32, X33, X34, X35, X36, X37, X38, X39, X40, X41, X42, X43, X44, X45, X46, X47, X48 FROM Simplex", conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                DataTable dt2 = new DataTable();
                adapter2.Fill(dt2);
                gridControl1.DataSource = dt;
                gridControl2.DataSource = dt2;
                conn.Close();
                #region оформляем грид
                for (int r = 0; r < name.Count(); r++)
                {
                    gridView1.Columns[r].Caption = name[r];
                }
                for (int r = 0; r < dop.Count(); r++)
                {
                    gridView2.Columns[r].Caption = dop[r];
                }
                for (int t = 0; t < gridView1.Columns.Count; t++)
                {
                    gridView1.Columns[t].AppearanceHeader.Options.UseTextOptions = true;
                    gridView1.Columns[t].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gridView1.Columns[t].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gridView1.Columns[t].BestFit();
                }                
                for (int t = 0; t < gridView2.Columns.Count; t++)
                {
                    gridView2.Columns[t].AppearanceHeader.Options.UseTextOptions = true;
                    gridView2.Columns[t].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gridView2.Columns[t].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gridView2.Columns[t].BestFit();
                }
                foreach (DevExpress.XtraGrid.Columns.GridColumn column in gridView1.Columns)
                { column.OptionsColumn.AllowSort = DefaultBoolean.False; }
                foreach (DevExpress.XtraGrid.Columns.GridColumn column in gridView2.Columns)
                { column.OptionsColumn.AllowSort = DefaultBoolean.False; }
                #endregion
            }
            catch (Exception q)
            {
                gridControl1.DataSource = null;
                gridControl2.DataSource = null;
                MessageBox.Show(q.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        #region Удаления
        private void barButtonItem1_ItemClick(object sender, ItemClickEventArgs e)
        {
            try
            {
                if (gridView1.DataSource != null || gridView1.RowCount != 0)
                {
                    DialogResult result;
                    result = MessageBox.Show("Таблица содержит данные, \n Продолжить удаление?", "ФАНЗ", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == System.Windows.Forms.DialogResult.Yes)
                    {
                        gridControl1.DataSource = null;
                        gridView1.Columns.Clear();
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

        private void barButtonItem2_ItemClick(object sender, ItemClickEventArgs e)
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

        private void barButtonItem5_ItemClick(object sender, ItemClickEventArgs e)
        {
            gridControl1.DataSource = null;
            gridControl2.DataSource = null;
            gridControl3.DataSource = null;
            gridView1.Columns.Clear();            
            gridView2.Columns.Clear();
            gridView3.Columns.Clear();
        }
        #endregion
        #region Экспорт в ексель
        private void barButtonItem3_ItemClick(object sender, ItemClickEventArgs e)
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
                    if (!Directory.Exists(s + @"\Отчеты\Показатели Симплекс-таблицы\" + "" + date1))
                    {
                        Directory.CreateDirectory(s + @"\Отчеты\Показатели Симплекс-таблицы\" + "" + date1);
                        string m = s + "\\Отчеты\\Показатели Симплекс-таблицы\\" + date1;
                        gridView1.ExportToXlsx(m + @"\" + filename);
                        MessageBox.Show("Данные в Ексель успешно экспортированы.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string m = s + "\\Отчеты\\Показатели Симплекс-таблицы\\" + date1;
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

                    openFileDialog1.InitialDirectory = s + @"\Отчеты\Показатели Симплекс-таблицы\" + date1;
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

        private void barButtonItem4_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (gridView2.RowCount != 0)
            {
                string date1 = DateTime.Today.ToShortDateString();
                string date2 = "Cоздан в " + Convert.ToString(DateTime.Now.Hour + "." + DateTime.Now.Minute);
                string filename = date2 + ".xlsx";
                DialogResult results;
                //------------------------------------------------------------------------------------------------------------------------              
                try
                {
                    if (!Directory.Exists(s + @"\Отчеты\Доп. переменные Симплекс-таблицы\" + "" + date1))
                    {
                        Directory.CreateDirectory(s + @"\Отчеты\Доп. переменные Симплекс-таблицы\" + "" + date1);
                        string m = s + "\\Отчеты\\Доп. переменные Симплекс-таблицы\\" + date1;
                        gridView2.ExportToXlsx(m + @"\" + filename);
                        MessageBox.Show("Данные в Ексель успешно экспортированы.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string m = s + "\\Отчеты\\Доп. переменные Симплекс-таблицы\\" + date1;
                        gridView2.ExportToXlsx(m + @"\" + filename);
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

                    openFileDialog1.InitialDirectory = s + @"\Отчеты\Доп. переменные Симплекс-таблицы\" + date1;
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
        #region Предпросмотр, печать
        private void barButtonItem8_ItemClick(object sender, ItemClickEventArgs e)
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

        private void barButtonItem10_ItemClick(object sender, ItemClickEventArgs e)
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
                MessageBox.Show("Таблица пустая, предпросмотр отмененен.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }          
        }

        private void barButtonItem9_ItemClick(object sender, ItemClickEventArgs e)
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

        private void barButtonItem11_ItemClick(object sender, ItemClickEventArgs e)
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
        #endregion

        private void gridControl1_MouseHover(object sender, EventArgs e)
        {
            gridControl1.Focus();
        }

        private void gridControl2_MouseHover(object sender, EventArgs e)
        {
            gridControl2.Focus();
        }

        private void gridView2_PrintInitialize(object sender, DevExpress.XtraGrid.Views.Base.PrintInitializeEventArgs e)
        {
            PrintingSystemBase pb = e.PrintingSystem as PrintingSystemBase;
            pb.PageSettings.PaperKind = System.Drawing.Printing.PaperKind.A4;
            pb.PageSettings.Landscape = true;
            pb.PageSettings.BottomMargin = 10;
            pb.PageSettings.TopMargin = 10;
            pb.PageSettings.RightMargin = 10;
            pb.PageSettings.LeftMargin = 10; 
        }

        private void gridView1_PrintInitialize(object sender, DevExpress.XtraGrid.Views.Base.PrintInitializeEventArgs e)
        {
            PrintingSystemBase pb = e.PrintingSystem as PrintingSystemBase;
            pb.PageSettings.PaperKind = System.Drawing.Printing.PaperKind.A4;
            pb.PageSettings.Landscape = true;
            pb.PageSettings.BottomMargin = 10;
            pb.PageSettings.TopMargin = 10;
            pb.PageSettings.RightMargin = 10;
            pb.PageSettings.LeftMargin = 10; 
        }

        private void gridView1_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator && e.RowHandle >= 0)
            {
                e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
                if (e.RowHandle == 35)
                    e.Info.DisplayText = "M";
                if (e.RowHandle == 36)
                    e.Info.DisplayText = "Z";
            }
        }

        private void gridView1_RowCountChanged(object sender, EventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView gridView = ((DevExpress.XtraGrid.Views.Grid.GridView)sender);
            if (!gridView.GridControl.IsHandleCreated) return;
            Graphics gr = Graphics.FromHwnd(gridView.GridControl.Handle);
            SizeF size = gr.MeasureString(gridView.RowCount.ToString(), gridView.PaintAppearance.Row.GetFont());
            gridView.IndicatorWidth = Convert.ToInt32(size.Width + 0.999f)
                + DevExpress.XtraGrid.Views.Grid.Drawing.GridPainter.Indicator.ImageSize.Width + 10;
        }

        private void gridView2_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator && e.RowHandle >= 0)
            {
                e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
                if (e.RowHandle == 35)
                    e.Info.DisplayText = "M";
                if (e.RowHandle == 36)
                    e.Info.DisplayText = "Z";
            }
        }

        private void gridView2_RowCountChanged_1(object sender, EventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView gridView = ((DevExpress.XtraGrid.Views.Grid.GridView)sender);
            if (!gridView.GridControl.IsHandleCreated) return;
            Graphics gr = Graphics.FromHwnd(gridView.GridControl.Handle);
            SizeF size = gr.MeasureString(gridView.RowCount.ToString(), gridView.PaintAppearance.Row.GetFont());
            gridView.IndicatorWidth = Convert.ToInt32(size.Width + 0.999f)
                + DevExpress.XtraGrid.Views.Grid.Drawing.GridPainter.Indicator.ImageSize.Width + 10;
        }
        private void gridView1_CustomRowCellEdit(object sender, DevExpress.XtraGrid.Views.Grid.CustomRowCellEditEventArgs e)
        {
            var repositoryItemTextEditReadOnly = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
            repositoryItemTextEditReadOnly.Name = "repositoryItemTextEditReadOnly";
            repositoryItemTextEditReadOnly.ReadOnly = true;

            if (e.RowHandle == 35)
            {
                e.RepositoryItem = repositoryItemTextEditReadOnly;
            }
        }
        private void gridView2_CustomRowCellEdit(object sender, DevExpress.XtraGrid.Views.Grid.CustomRowCellEditEventArgs e)
        {
            var repositoryItemTextEditReadOnly = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
            repositoryItemTextEditReadOnly.Name = "repositoryItemTextEditReadOnly";
            repositoryItemTextEditReadOnly.ReadOnly = true;

            if (e.RowHandle == 35)
            {
                e.RepositoryItem = repositoryItemTextEditReadOnly;
            }
        }
        private void gridView1_RowStyle_1(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle == 35)
            {
                e.HighPriority = true;
                e.Appearance.BackColor = Color.Salmon;
            }
        }
        private void gridView2_RowStyle_1(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle == 35)
            {
                e.HighPriority = true;
                e.Appearance.BackColor = Color.Salmon;
            }
        }
        //Загрузка из таблицы гаусса
        private void barButtonItem13_ItemClick(object sender, ItemClickEventArgs e)
        {
            string[] gauss = {"ТА","КП","СОС","ДС","ДЗ","ВА","ЗЗ","ПК","СК","ФР","ДП","Оср","ПР.ТА","ПР.ВА",
                "ВР","Вал.пр.","КрУр","Пр.прод","Проч.д","Проч.р","ЧП","RСовК","Пр.до_нал.","Себ.прод.","НерПр",
                "d1","d2","d3","d4","d5","d6","d7","d8","d9","d10"};
            try
            {
                SqlCeConnection conn;
                conn = new SqlCeConnection(@"Data Source = |DataDirectory|\FANZ.sdf");
                conn.Open();
                SqlCeDataAdapter adapter3;            
                adapter3 = new SqlCeDataAdapter("SELECT x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11, x12, x13, x14, x15, x16, x17, x18, x19, x20, x21, x22, x23, x24, x25, X26, X27, X28, X29, X30, X31, X32, X33, X34, X35 FROM Gauss", conn);
                DataTable dt3 = new DataTable();
                adapter3.Fill(dt3);
                gridControl3.DataSource = dt3;
                conn.Close();
                #region оформляем грид
                for (int r = 0; r < gauss.Count(); r++)
                {
                    gridView3.Columns[r].Caption = gauss[r];
                }
                for (int t = 0; t < gridView3.Columns.Count; t++)
                {
                    gridView3.Columns[t].AppearanceHeader.Options.UseTextOptions = true;
                    gridView3.Columns[t].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gridView3.Columns[t].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gridView3.Columns[t].BestFit();
                }
                foreach (DevExpress.XtraGrid.Columns.GridColumn column in gridView3.Columns)
                { column.OptionsColumn.AllowSort = DefaultBoolean.False; }
                #endregion
            }
            catch (Exception q)
            {
                gridControl3.DataSource = null;
                MessageBox.Show(q.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        //Очистка таблицы
        private void barButtonItem16_ItemClick(object sender, ItemClickEventArgs e)
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
        //Переключение между вкладками
        private void ribbon_SelectedPageChanged(object sender, EventArgs e)
        {
            if (ribbon.SelectedPage == ribbonPage2)
            {
                xtraTabControl1.SelectedTabPage = xtraTabPage3;
            }
            else if (ribbon.SelectedPage == ribbonPage1)
            {
                xtraTabControl1.SelectedTabPage = xtraTabPage1;
            }
        }
        // Экспорт в Ексель
        private void barButtonItem17_ItemClick(object sender, ItemClickEventArgs e)
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
                    if (!Directory.Exists(s + @"\Отчеты\Симплекс-таблица Гаусса\" + "" + date1))
                    {
                        Directory.CreateDirectory(s + @"\Отчеты\Симплекс-таблица Гаусса\" + "" + date1);
                        string m = s + "\\Отчеты\\Симплекс-таблица Гаусса\\" + date1;
                        gridView3.ExportToXlsx(m + @"\" + filename);
                        MessageBox.Show("Данные в Ексель успешно экспортированы.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string m = s + "\\Отчеты\\Симплекс-таблица Гаусса\\" + date1;
                        gridView3.ExportToXlsx(m + @"\" + filename);
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

                    openFileDialog1.InitialDirectory = s + @"\Отчеты\Симплекс-таблица Гаусса\" + date1;
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
        // Предпросмотр
        private void barButtonItem18_ItemClick(object sender, ItemClickEventArgs e)
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
                MessageBox.Show("Таблица пустая, предпросмотр отмененен.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
        // Печать
        private void barButtonItem19_ItemClick(object sender, ItemClickEventArgs e)
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
        // Определение формы при печати
        private void gridView3_PrintInitialize(object sender, DevExpress.XtraGrid.Views.Base.PrintInitializeEventArgs e)
        {
            PrintingSystemBase pb = e.PrintingSystem as PrintingSystemBase;
            pb.PageSettings.PaperKind = System.Drawing.Printing.PaperKind.A4;
            pb.PageSettings.Landscape = true;
            pb.PageSettings.BottomMargin = 10;
            pb.PageSettings.TopMargin = 10;
            pb.PageSettings.RightMargin = 10;
            pb.PageSettings.LeftMargin = 10; 
        }
        // Счетчик
        private void gridView3_RowCountChanged(object sender, EventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView gridView = ((DevExpress.XtraGrid.Views.Grid.GridView)sender);
            if (!gridView.GridControl.IsHandleCreated) return;
            Graphics gr = Graphics.FromHwnd(gridView.GridControl.Handle);
            SizeF size = gr.MeasureString(gridView.RowCount.ToString(), gridView.PaintAppearance.Row.GetFont());
            gridView.IndicatorWidth = Convert.ToInt32(size.Width + 0.999f)
                + DevExpress.XtraGrid.Views.Grid.Drawing.GridPainter.Indicator.ImageSize.Width + 10;
        }
        // Индикатор
        private void gridView3_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator && e.RowHandle >= 0)
            {
                e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }
        }
    }
}