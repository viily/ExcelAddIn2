using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn2
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            // Open Worksheet
            Excel.Worksheet activeworksheet = Globals.ThisAddIn.Application.ActiveSheet;  
            // Open Worksheet
            Excel.Range actCell = Globals.ThisAddIn.Application.ActiveCell;
            if (actCell.Value2 != null)
            {
                string sValue = actCell.Value2.ToString();
                string sText = actCell.Text;
                System.Windows.Forms.MessageBox.Show(sText);
            }

        }


    }
}
