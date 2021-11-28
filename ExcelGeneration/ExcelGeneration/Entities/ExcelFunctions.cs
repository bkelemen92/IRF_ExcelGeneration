using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Windows.Forms;
using ExcelGeneration.Data;
using System.Drawing;

namespace ExcelGeneration.Entities
{
    public static class ExcelFunctions
    {
        static Excel.Application xlApp;
        static Excel.Workbook xlWb;
        static Excel.Worksheet xlWs;

        static string[] headers = new string[] { "Kód", "Eladó", "Oldal", "Kerület", "Lift", "Szobák száma", "Alapterület (m2)", "Ár (mFt)", "Négyzetméter ár (Ft/m2)" };

        public static void CreateExcel(List<Flat> flats)
        {
            try
            {
                xlApp = new Excel.Application();
                xlWb = xlApp.Workbooks.Add(Missing.Value);
                xlWs = xlWb.ActiveSheet;

                CreateTable(flats);

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

        private static void CreateTable(List<Flat> flats)
        {
            
            object[,] values = new object[flats.Count, headers.Length];

            for (int i = 0; i < headers.Length; i++)
            {
                xlWs.Cells[1, i + 1] = headers[i];
            }

            int counter = 0;
            foreach (Flat f in flats)
            {
                values[counter, 0] = f.Code;
                values[counter, 1] = f.Vendor;
                values[counter, 2] = f.Side;
                values[counter, 3] = f.District;
                values[counter, 4] = f.Elevator == true ? "Van" : "Nincs";
                values[counter, 5] = f.NumberOfRooms;
                values[counter, 6] = f.FloorArea;
                values[counter, 7] = f.Price;
                values[counter, 8] = string.Format("={0}/{1}*1000000", GetCellInRange(counter + 2, 8), GetCellInRange(counter + 2, 7));
                counter++;
            }

            xlWs.get_Range(GetCellInRange(2, 1), GetCellInRange(1 + values.GetLength(0), values.GetLength(1))).Value2 = values;

            FormatTable(flats);
        }

        private static string GetCellInRange(int x, int y)
        {
            string ExcelCoordinate = "";
            int dividend = y;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                ExcelCoordinate = Convert.ToChar(65 + modulo).ToString() + ExcelCoordinate;
                dividend = (int)((dividend - modulo) / 26);
            }
            ExcelCoordinate += x.ToString();

            return ExcelCoordinate;
        }

        private static void FormatTable(List<Flat> flats)
        {
            Excel.Range headerRange = xlWs.get_Range(GetCellInRange(1, 1), GetCellInRange(1, headers.Length));
            Excel.Range valuesRange = xlWs.get_Range(GetCellInRange(2, 1), GetCellInRange(flats.Count + 1, headers.Length));
            Excel.Range firstColumnRange = xlWs.get_Range(GetCellInRange(2, 1), GetCellInRange(flats.Count + 1, 1));
            Excel.Range lastColumnRange = xlWs.get_Range(GetCellInRange(2, headers.Length), GetCellInRange(flats.Count + 1, headers.Length));

            //Header format
            headerRange.Font.Bold = true;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headerRange.EntireColumn.AutoFit();
            headerRange.RowHeight = 40;
            headerRange.Interior.Color = Color.LightBlue;
            headerRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            //Values format
            valuesRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            //First column format
            firstColumnRange.Font.Bold = true;
            firstColumnRange.Interior.Color = Color.LightYellow;

            //Last column format
            lastColumnRange.Interior.Color = Color.LightGreen;
            lastColumnRange.NumberFormat = "#,#.00 Ft";
        }

        public static void CloseExcelApp()
        {
            xlWb.Close(false);
            xlApp.Quit();
            xlWb = null;
            xlApp = null;
        }
    }
}
