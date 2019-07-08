using System;
using System.Windows.Forms;
using System.IO;

namespace Simplex_2.Формы_проекта
{
    public partial class HelpToLinean : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public static string file = Path.Combine(System.Windows.Forms.Application.StartupPath, "Справка для прогнозирования.pdf");

        public HelpToLinean()
        {
            InitializeComponent();
        }

        private void HelpToLinean_Load(object sender, EventArgs e)
        {
            try
            {
                pdfViewer1.DocumentFilePath = file;
            }
            catch
            {
                MessageBox.Show("Исходный файл справки удален.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
    }
}
