using System;
using System.Windows.Forms;
using System.IO;

namespace Simplex_2.Формы_проекта
{
    public partial class UserGuide : Simplex_2.Формы_проекта.Help
    {
        public static string file = Path.Combine(System.Windows.Forms.Application.StartupPath, "Руководство пользователя_2.0.3.pdf");

        public UserGuide()
        {
            InitializeComponent();
        }

        private void UserGuide_Load(object sender, EventArgs e)
        {
            try
            {
                MessageBox.Show("jghfhfhk");
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
