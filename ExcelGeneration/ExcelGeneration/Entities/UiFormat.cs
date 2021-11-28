using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelGeneration.Entities
{
    class UiFormat
    {
        public static void FormatLabelAsTitle(Label label)
        {
            label.Text = Properties.Resources.Title;
            label.Font = new Font("Segoe UI", 55, FontStyle.Bold);
            label.AutoSize = false;
            label.Dock = DockStyle.Top;
            label.TextAlign = ContentAlignment.MiddleCenter;
            label.Height = 100;
        }

        public static void FormatTwoButtons(Form parent, Button btn1, Button btn2)
        {
            int btnWidth = 200;
            int btnHeight = 30;
            int btnSpacing = 12;

            btn1.Text = Properties.Resources.CloseForm;
            btn1.Width = btnWidth;
            btn1.Height = btnHeight;
            btn1.Left = (parent.Width / 2) - btnWidth - (btnSpacing);
            btn1.Top = 130;
            btn1.Anchor = AnchorStyles.Top;

            btn2.Text = Properties.Resources.CloseFormAndExcel;
            btn2.Width = btnWidth;
            btn2.Height = btnHeight;
            btn2.Left = (parent.Width / 2) + (btnSpacing / 2);
            btn2.Top = 130;
            btn2.Anchor = AnchorStyles.Top;
        }
    }
}
