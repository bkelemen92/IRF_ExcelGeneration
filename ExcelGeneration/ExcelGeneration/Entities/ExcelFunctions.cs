using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelGeneration.Entities
{
    public static class ExcelFunctions
    {
        static Excel.Application xlApp;
        static Excel.Workbook xlWb;
        static Excel.Worksheet xlWs;

        public static void CreateExcel()
        {
            try
            {
                xlApp = new Excel.Application();
                xlWb = xlApp.Workbooks.Add(Missing.Value);
                xlWs = xlWb.ActiveSheet;

                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex)
            {
                string errorMessage = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errorMessage, "Error");

                CloseExcelApp();

            }
        }


        private static void CloseExcelApp()
        {
            xlWb.Close(false);
            xlApp.Quit();
            xlWb = null;
            xlApp = null;
        }
    }
}
