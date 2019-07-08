using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraCharts.UI;
using System.Data.SQLite;
using Kernel;
using Invention;

namespace Edition
{
    public class Versions
    {
        public List<BarButtonItem> barbutton = new List<BarButtonItem>();
        public List<BarSubItem> subbutton = new List<BarSubItem>();
        public List<Button> button = new List<Button>();

        public List<ChartRibbonPageGroup> page = new List<ChartRibbonPageGroup>();
        public List<ChartAppearanceRibbonPageGroup> style = new List<ChartAppearanceRibbonPageGroup>();
        public List<ChartWizardRibbonPageGroup> wiz = new List<ChartWizardRibbonPageGroup>();
        public List<ChartTemplatesRibbonPageGroup> template = new List<ChartTemplatesRibbonPageGroup>();
        public List<ChartPrintExportRibbonPageGroup> print = new List<ChartPrintExportRibbonPageGroup>();
       
        public static int num = 0;

        public List<string> solve = new List<string>
           {
            "barButtonItem19", "barButtonItem11", "barButtonItem8",  
            "barButtonItem10", "barButtonItem13", "barButtonItem1", "barButtonItem18"
           };
        public List<string> solve1 = new List<string>
           {
            "button2", "button17", "button21"
           };
        public List<string> solve2 = new List<string>
           {             
            "barSubItem7", "barSubItem10",
            "barSubItem11", "barSubItem1"
           };
        //-------------------------------------------
        public List<string> DSolse = new List<string>
           {
            "barButtonItem8",  "barButtonItem10", "barButtonItem11",  
            "barButtonItem13", "barButtonItem18", "barButtonItem19"
           };
        public List<string> DSolse1 = new List<string>
           {
            "button2", "button3"
           };
        public List<string> DSolse2 = new List<string>
           {             
            "barSubItem7"
           };
        //------------------------------------------
        public List<string> keys = new List<string> 
        { 
            "a46c3b54f2c9871cd81daf7a932499c0", "309fc7d3bc53bb63ac42e359260ac740",
            "6d5ababb65e9ff214b73e891b4afe6e8", "06d49632c9dc9bcb62aeaef99612ba6b" 
        };

        #region add
        public void AddBarButton(BarButtonItem bar)
        {
            barbutton.Add(bar);
        }
        public void AddBarSub(BarSubItem sub)
        {
            subbutton.Add(sub);
        }
        public void AddButton(Button bt)
        {
            button.Add(bt);
        }

        public void AddRibbonPage(ChartRibbonPageGroup pages)
        {
            page.Add(pages);
        }
        public void AddAppearance(ChartAppearanceRibbonPageGroup app)
        {
            style.Add(app);
        }        
        public void AddWizard(ChartWizardRibbonPageGroup wz)
        {
            wiz.Add(wz);
        }
        public void AddTemplate(ChartTemplatesRibbonPageGroup temp)
        {
            template.Add(temp);
        }
        public void AddPrint(ChartPrintExportRibbonPageGroup pr)
        {
            print.Add(pr);
        }
        #endregion
        #region show
        public void ShowBarButton()
        {
            foreach (var element in barbutton)
            {
                element.Enabled = true;
            }
        }
        public void ShowSubButton()
        {
            foreach (var element in subbutton)
            {
                element.Enabled = true;
            }
        }
        public void ShowButton()
        {
            foreach (var element in button)
            {
                element.Enabled = true;
            }
        }
        public void ShowRibbonPage()
        {
            foreach (var element in page)
            {
                element.Enabled = true;
            }
        }
        public void ShowAppearance()
        {
            foreach (var element in style)
            {
                element.Enabled = true;
            }
        }
        public void ShowWizard()
        {
            foreach (var element in wiz)
            {
                element.Enabled = true;
            }
        }
        public void ShowTemplate()
        {
            foreach (var element in template)
            {
                element.Enabled = true;
            }
        }
        public void ShowPrint()
        {
            foreach (var element in print)
            {
                element.Enabled = true;
            }
        }
        #endregion
        #region shadow
        public void ShadowBarButton()
        {
            foreach (var element in barbutton)
            {
                element.Enabled = false;
            }
        }
        public void ShadowSubButton()
        {
            foreach (var element in subbutton)
            {
                element.Enabled = false;
            }
        }
        public void ShadowButton()
        {
            foreach (var element in button)
            {
                element.Enabled = false;
            }
        }
        public void ShadowRibbonPage()
        {
            foreach (var element in page)
            {
                element.Enabled = false;
            }
        }
        public void ShadowAppearance()
        {
            foreach (var element in style)
            {
                element.Enabled = false;
            }
        }
        public void ShadowWizard()
        {
            foreach (var element in wiz)
            {
                element.Enabled = false;
            }
        }
        public void ShadowTemplate()
        {
            foreach (var element in template)
            {
                element.Enabled = false;
            }
        }
        public void ShadowPrint()
        {
            foreach (var element in print)
            {
                element.Enabled = false;
            }
        }
        #endregion
        #region clear
        public void ClearList()
        {
            barbutton.Clear();
            subbutton.Clear();
            button.Clear();
            page.Clear();
            style.Clear();
            wiz.Clear();
            template.Clear();
            print.Clear();
        }
        #endregion

        public void MNK(Form name, int key)
        {
            foreach (Control c in name.Controls)
            {                   
                if (key == 2 && c.Name == "panelControl1")
                {
                    c.Enabled = true;                    
                }
                else c.Enabled = false;
            }
            
        }

        public bool Validate(string key)
        {
            int j = 0;
            foreach (var i in keys)
            {
                if (i==key)
                {
                    return true;                    
                }
                j++;
            }
            return false;                
        }

        public int Zapros(string key)
        {            
            switch (key)
            {
                case "a46c3b54f2c9871cd81daf7a932499c0":  //демо
                    num = 0;
                    break;
                case "06d49632c9dc9bcb62aeaef99612ba6b":  //профессиональная
                    num = 1;
                    break;
                case "6d5ababb65e9ff214b73e891b4afe6e8":  //расширенная
                    num = 2;
                    break;
                case "309fc7d3bc53bb63ac42e359260ac740":  //базовая
                    num = 3;
                    break;
            }
            return num;
        }

        public void ExecuteSolve()
        {
            foreach (var find in barbutton)
            {
                for (int i = 0; i < solve.Count; i++)
                {
                    if (find.Name == solve[i].ToString())
                    {
                        #region barButtonItem
                        if (find.Name == "barButtonItem19")
                        {
                            find.Enabled = true;
                        }
                        if (find.Name == "barButtonItem11")
                        {
                            find.Enabled = false;
                        }
                        if (find.Name == "barButtonItem10")
                        {
                            find.Enabled = false;
                        }
                        if (find.Name == "barButtonItem8")
                        {
                            find.Enabled = false;
                        }
                        if (find.Name == "barButtonItem13")
                        {
                            find.Enabled = true;
                        }
                        if (find.Name == "barButtonItem1")
                        {
                            find.Enabled = false;
                        }
                        if (find.Name == "barButtonItem18")
                        {
                            find.Enabled = true;
                        }
                        #endregion
                    }
                }
            }
            foreach (var find in subbutton)
            {
                for (int i = 0; i < solve2.Count; i++)
                {
                    if (find.Name == solve2[i].ToString())
                    {
                        #region barSubItem
                        if (find.Name == "barSubItem1")
                        {
                            find.Enabled = false;
                        }
                        if (find.Name == "barSubItem7")
                        {
                            find.Enabled = true;
                        }
                        if (find.Name == "barSubItem10")
                        {
                            find.Enabled = true;
                        }
                        if (find.Name == "barSubItem11")
                        {
                            find.Enabled = true;
                        }
                        #endregion
                    }
                }
            }
            foreach (var find in button)
            {
                for (int i = 0; i < solve1.Count; i++)
                {
                    if (find.Name == solve1[i].ToString())
                    {
                        #region button
                        if (find.Name == "button2")
                        {
                            find.Enabled = false;
                        }
                        if (find.Name == "button17")
                        {
                            find.Enabled = true;
                        }
                        if (find.Name == "button21")
                        {
                            find.Enabled = true;
                        }
                        #endregion
                    }
                }
            }
        }
        public void DowloadSolve()
        {
            foreach (var find in barbutton)
            {
                for (int i = 0; i < DSolse.Count; i++)
                {
                    if (find.Name == DSolse[i].ToString())
                    {
                        #region barButtonItem
                        if (find.Name == "barButtonItem8")
                        {
                            find.Enabled = true;
                        }
                        if (find.Name == "barButtonItem10")
                        {
                            find.Enabled = true;
                        }
                        if (find.Name == "barButtonItem11")
                        {
                            find.Enabled = false;
                        }
                        if (find.Name == "barButtonItem13")
                        {
                            find.Enabled = false;
                        }
                        if (find.Name == "barButtonItem18")
                        {
                            find.Enabled = false;
                        }
                        if (find.Name == "barButtonItem19")
                        {
                            find.Enabled = false;
                        }
                        #endregion
                    }
                }
            }
            foreach (var find in subbutton)
            {
                for (int i = 0; i < DSolse2.Count; i++)
                {
                    if (find.Name == DSolse2[i].ToString())
                    {
                        #region barSubItem
                        if (find.Name == "barSubItem7")
                        {
                            find.Enabled = true;
                        } 
                        #endregion
                    }
                }
            }
            foreach (var find in button)
            {
                for (int i = 0; i < solve1.Count; i++)
                {
                    if (find.Name == solve1[i].ToString())
                    {
                        #region button
                        if (find.Name == "button2")
                        {
                            find.Enabled = false;
                        }
                        if (find.Name == "button3")
                        {
                            find.Enabled = false;
                        }
                        #endregion
                    }
                }
            }
        }
        public void Status(ToolStripStatusLabel stat)
        {
            stat.Text = "Зарегистрирован (найдено решение)";
        }
        public void Message()
        {
            MessageBox.Show("Используется демо-версия. Расчетный модуль не доступен.", "ФАНЗ", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
