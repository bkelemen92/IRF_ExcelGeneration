using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelGeneration.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using ExcelGeneration.Entities;

namespace ExcelGeneration
{
    public partial class Form1 : Form
    {
        List<Flat> Flats;
        RealEstateEntities context = new RealEstateEntities();
        //ExcelFunctions ExcelFn = new ExcelFunctions();

        public Form1()
        {
            InitializeComponent();
            LoadData();

            UiFormat.FormatLabelAsTitle(lbl_Magic);
            UiFormat.FormatTwoButtons(this, btn_Exit, btn_CloseExcelExit);

            ExcelFunctions.CreateExcel(Flats);

            btn_Exit.Click += Btn_Exit_Click;
            btn_CloseExcelExit.Click += Btn_CloseExcelExit_Click;
        }

        private void Btn_CloseExcelExit_Click(object sender, EventArgs e)
        {
            ExcelFunctions.CloseExcelApp();
            Application.Exit();
        }

        private void Btn_Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void LoadData()
        {
            Flats = context.Flats.ToList();
        }
    }
}
