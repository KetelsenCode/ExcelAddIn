using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace Basics
{
    public partial class RibbonMain
    {
        private void RibbonMain_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Btn1_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            //Gets the first worksheet in the active workbook
            Worksheet worksheet = (Worksheet)Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);
            //Inserts into a cell
            worksheet.Cells[1, 1] = eb1.Text;
        }
        private void Btn2_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            Worksheet worksheet = (Worksheet)Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);
            //Gets the cell - can also get a range of
            var cell = worksheet.Range["A1"];
            lb1.Label = cell.Text;
        }
    }
}
