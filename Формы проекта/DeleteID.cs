using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Data.SqlServerCe;

namespace Simplex_2.Формы_проекта
{
    public partial class DeleteID : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public bool result;

        public DeleteID()
        {
            InitializeComponent();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if (comboBoxEdit1.Text == "")
            {
                MessageBox.Show("Укажите дату для удаления.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            else
            {
                result = true;
                this.Close();
            }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            result = false;
            this.Close();
        }

        private void DeleteID_Load(object sender, EventArgs e)
        {
            try
            {
                comboBoxEdit1.Properties.Items.Clear();
                comboBoxEdit1.Text = string.Empty;
                SqlCeConnection conn;
                conn = new SqlCeConnection(@"Data Source = |DataDirectory|\FANZ.sdf");
                conn.Open();
                SqlCeDataAdapter adapter;
                adapter = new SqlCeDataAdapter("SELECT DISTINCT DateCalculation FROM Linean", conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                int row = dt.Rows.Count;

                for (int j = 0; j < row; j++)
                {
                    comboBoxEdit1.Properties.Items.AddRange(dt.Rows[j].ItemArray.ToList());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
    }
}
