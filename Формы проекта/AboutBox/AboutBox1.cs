using System;
using System.Reflection;

namespace Simplex_2.Формы_проекта
{
    partial class AboutBox1 : System.Windows.Forms.Form
    {      
        public AboutBox1()
        {
            InitializeComponent();
            this.Text = String.Format("О программе");
            this.labelProductName.Text = AssemblyProduct;
            this.labelVersion.Text = String.Format("Версия {0}", AssemblyVersion);
            this.textBoxDescription.Text = AssemblyDescription;
        }

        #region Методы доступа к атрибутам сборки

        public string AssemblyTitle
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                    {
                        return titleAttribute.Title;
                    }
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        public string AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }

        public string AssemblyDescription
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        public string AssemblyProduct
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public string AssemblyCopyright
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public string AssemblyCompany
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        #endregion

        private void AboutBox1_Load(object sender, EventArgs e)
        {
            labelProductName.Text = "Название: Финансовый анализ. Новые знания";
            labelVersion.Text = "Версия: 2.0.3" + " (демо-версия)";
            textBoxDescription.Text = "Программное приложение «ФАНЗ» предназначено для оценки финансового состояния предприятия (компании) на основе принципа поиска его идеального (оптимального) состояния и последовательного расчета общего относительного отклонения от этого состояния по специальному алгоритму. С математической точки зрения решается задача линейного программирования, в которой осуществляется максимизация (минимизация) целевой функции при заданных ограничениях. Целевая функция применительно к данной задаче представляет собой максимизацию рентабельности собственного капитала, а ограничения – это задаваемые пользователем значения финансовых показателей. Поиск оптимального решения осуществляется симплексным М-методом – одним из наиболее известных алгоритмов решения оптимизационных задач линейного программирования.";
        }
    }
}
