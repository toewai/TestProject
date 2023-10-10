using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Excel= Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace TestProject
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, RoutedEventArgs e)
        {
            if (txtExcelPath.Text == "" || !File.Exists(txtExcelPath.Text))
            {
                MessageBox.Show("Please Check Your Excel file Paht");
                txtExcelPath.Focus();
                return;
            }
            if(txtMacro.Text == "")
            {
                MessageBox.Show("Please Enter Your Macro Name");
                txtMacro.Focus();
                return;
            }
            Excel.Application xlApp = new Excel.Application();
            Excel._Workbook xlWorkbook = xlApp.Workbooks.Open(@txtExcelPath.Text);
            if (xlWorkbook.HasVBProject)
            {
                bool macroExit = false;
                for (int i = 1; i <= xlWorkbook.VBProject.VBComponents.Count; i++)
                    if (xlWorkbook.VBProject.VBComponents.Item(i).Type.ToString() == "vbext_ct_StdModule")
                    {
                        String firstLine = xlWorkbook.VBProject.VBComponents.Item(i).CodeModule.Lines[1, 1];
                        String funName = Regex.Replace(firstLine, "Sub ", "").Replace("(", "").Replace(")","").Trim();
                        if (funName == txtMacro.Text)
                            macroExit = true;
                    }
                if (macroExit)
                    xlApp.Run(txtMacro.Text);
                else
                    MessageBox.Show("Macro Name Not Exist");
            }
            xlWorkbook.Close();
            xlApp.Quit();
                
        }
        private void btnCompare_Click(object sender, RoutedEventArgs e)
        {
            if (txtExcelPath.Text != "" && File.Exists(txtExcelPath.Text))
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@txtExcelPath.Text);
                Excel._Worksheet sheetTrue = (Excel._Worksheet)xlWorkbook.Sheets[1];
                Excel._Worksheet sheetCheck = (Excel._Worksheet)xlWorkbook.Sheets[2];
                Excel.Range rangeTrue = (Excel.Range)sheetTrue.UsedRange;
                Excel.Range rangeCheck = (Excel.Range)sheetCheck.UsedRange;

                compare(rangeTrue, rangeCheck);
                xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();
                MessageBox.Show("Finish Check");
            }
        }
        private void compare (Excel.Range tSheet,Excel.Range cSheet)
        {
            for (int i = 1; i <= cSheet.Rows.Count; i++)
            {
                for (int j = 1; j <= cSheet.Columns.Count; j++)
                {
                    if (((Excel.Range)tSheet.Cells[i, j]).Value2 != ((Excel.Range)cSheet.Cells[i, j]).Value2)
                        ((Excel.Range)cSheet.Cells[i, j]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                }
            }
        }

    }
}
