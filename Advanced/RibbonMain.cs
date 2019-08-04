using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace Advanced
{
    public partial class RibbonMain
    {
        private void RibbonMain_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void Btn1_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            Workbook workbook = (Workbook)Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);
            workbook.SaveAs(@"C:\Users\Kent\Desktop\Coding\apps\ExcelAddIn\Basics\Advanced\test.csv", Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
            Worksheet worksheet = (Worksheet)Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);
            worksheet.SaveAs(@"C:\Users\Kent\Desktop\Coding\apps\ExcelAddIn\Basics\Advanced\test2.csv", Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);


        }

        private void Btn2_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            Worksheet worksheet = (Worksheet)Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);
            //Integer for number of rows
            int lastRow = worksheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, 
                            Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, 
                            System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            //At cell[lastRow+1,2] it writes =SUM(E1:Ex) which is Excel formula for calculate sum for E-column
            worksheet.Cells[lastRow + 1, 5] = "=SUM(E2:E" + lastRow + ")";
            //Does the same for column F
            worksheet.Cells[lastRow + 1, 6] = "=SUM(F2:F" + lastRow + ")";
        }

        private void Btn3_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            throw new System.NotImplementedException();
        }
    }
}
