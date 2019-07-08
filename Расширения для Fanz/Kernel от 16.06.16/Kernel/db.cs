using System;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace Kernel
{
    public class db
    {
        #region Настройки, загрузка данных из базы
        public static string filename = Path.Combine(Application.StartupPath, "gauss.db");
        public static string filename1 = Path.Combine(Application.StartupPath, "data.db");
        public static string shablon_xml2 = Path.Combine(Application.StartupPath, "pokaz.xml");
        string ConnectionString = string.Format("data source={0};New=True;UseUTF8Encoding=True", filename);
        string ConnectionString1 = string.Format("data source={0};New=True;UseUTF8Encoding=True", filename1);

        public DataTable FetchAll(string databasename)
        {
            return FetchAll(databasename, "", "");
        }
        public DataTable FetchAll1(string databasename)
        {
            return FetchAll1(databasename, "", "");
        }

        public DataTable FetchAll(string databasename, string where, string etc)
        {
            DataTable dt = new DataTable();
            string sql = string.Format("SELECT * FROM {0} {1} {2}", databasename, where, etc);
            ConnectionState previousConnectionState = ConnectionState.Closed;

            using (SQLiteConnection connect = new SQLiteConnection(ConnectionString))
            {
                try
                {
                    previousConnectionState = connect.State;
                    if (connect.State == ConnectionState.Closed)
                    {
                        connect.Open();
                    }
                    SQLiteCommand command = new SQLiteCommand(sql, connect);
                    SQLiteDataAdapter adapter = new SQLiteDataAdapter(command);
                    adapter.Fill(dt);
                }
                catch (Exception error)
                {
                    System.Windows.Forms.MessageBox.Show(error.Message, "Ошибка при получении данных из базы", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
                finally
                {
                    if (previousConnectionState == ConnectionState.Closed)
                    {
                        connect.Close();
                    }
                }
            }
            return dt;
        }
        public DataTable FetchAll1(string databasename, string where, string etc)
        {
            DataTable dt = new DataTable();
            string sql = string.Format("SELECT * FROM {0} {1} {2}", databasename, where, etc);
            ConnectionState previousConnectionState = ConnectionState.Closed;

            using (SQLiteConnection connect = new SQLiteConnection(ConnectionString1))
            {
                try
                {
                    previousConnectionState = connect.State;
                    if (connect.State == ConnectionState.Closed)
                    {
                        connect.Open();
                    }
                    SQLiteCommand command = new SQLiteCommand(sql, connect);
                    SQLiteDataAdapter adapter = new SQLiteDataAdapter(command);
                    adapter.Fill(dt);
                }
                catch (Exception error)
                {
                    System.Windows.Forms.MessageBox.Show(error.Message, "Ошибка при получении данных из базы", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
                finally
                {
                    if (previousConnectionState == ConnectionState.Closed)
                    {
                        connect.Close();
                    }
                }
            }
            return dt;
        }
        #endregion

        #region Сохранение в Xml
        public void ToDataTable2(DataGridView dataGridView1, string tableName1)
        {
            DataGridView dgv1 = dataGridView1;
            DataTable table1 = new DataTable(tableName1);
            table1.Columns.Add(dgv1.Columns[0].Name);
            foreach (DataGridViewRow row in dgv1.Rows)
            {
                DataRow datarw = table1.NewRow();
                datarw[0] = row.Cells[0].Value;
                table1.Rows.Add(datarw);
            }
            table1.WriteXml(@shablon_xml2);
            dgv1.DataSource = null;
            table1.Rows.Clear();
            table1.Columns.Clear();
        }
        public void SaveData(DataGridView dgv1) // для показателей из грида 6
        {
            try
            {
                ToDataTable2(dgv1, "Pokaz");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }     
    
        #endregion

        #region Загрузка данных из Xml        
        public void Open2(DataGridView dgv)
        {
            try
            {
                XmlReader xmlFile;
                xmlFile = XmlReader.Create(@shablon_xml2, new XmlReaderSettings());
                DataSet ds = new DataSet();
                ds.ReadXml(xmlFile);
                dgv.DataSource = ds.Tables[0];
                xmlFile.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        #endregion       
    
    }
}
