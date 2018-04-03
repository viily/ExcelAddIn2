using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn2
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void AutoFill()
        {
            Excel.Range rng = this.Application.get_Range("B1");
            rng.AutoFill(this.Application.get_Range("B1", "B5"),
                Excel.XlAutoFillType.xlFillWeekdays);

            rng = this.Application.get_Range("C1");
            rng.AutoFill(this.Application.get_Range("C1", "C5"),
                Excel.XlAutoFillType.xlFillMonths);

            rng = this.Application.get_Range("D1", "D2");
            rng.AutoFill(this.Application.get_Range("D1", "D5"),
                Excel.XlAutoFillType.xlFillSeries);
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }

}
