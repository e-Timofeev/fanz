using System;
using System.Windows.Forms;
using System.IO;

namespace Simplex_2.Формы_проекта
{
    public partial class Help : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public static string file1 = Path.Combine(System.Windows.Forms.Application.StartupPath, "Справка.pdf");

        public Help()
        {
            InitializeComponent();
        }

        protected void Help_Load(object sender, EventArgs e)
        {
            //try
            //{
            //    pdfViewer1.DocumentFilePath = file1;
            //}
            //catch
            //{
            //    MessageBox.Show("Исходный файл справки удален.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    return;
            //}
        }
    }
}
