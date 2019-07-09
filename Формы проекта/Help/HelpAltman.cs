using System;
using System.Windows.Forms;
using System.IO;

namespace Simplex_2.Формы_проекта
{
    public partial class HelpAltman : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public static string file = Path.Combine(System.Windows.Forms.Application.StartupPath, "Модели банкротства Альтмана.pdf");

        public HelpAltman()
        {
            InitializeComponent();
        }

        private void HelpAltman_Load(object sender, EventArgs e)
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
