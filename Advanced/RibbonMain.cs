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
        //Converts a workbook and worksheet to a CSV file
        private void Btn1_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            //Saves the workbook as a CSV file
            Workbook workbook = (Workbook)Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);
            workbook.SaveAs(@"C:\Users\Kent\Desktop\Coding\apps\ExcelAddIn\Basics\Advanced\test.csv", Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);

            //Saves the worksheet as a CSV file
            Worksheet worksheet = (Worksheet)Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);
            worksheet.SaveAs(@"C:\Users\Kent\Desktop\Coding\apps\ExcelAddIn\Basics\Advanced\test2.csv", Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);


        }
        //Calculates the sum of column E and F
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
        //Calculates the sum of a user chosen column
        private void Btn3_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            Worksheet worksheet = (Worksheet)Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);
            //Integer for number of rows
            int lastRow = worksheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            //Gets the userinput for which colum to calculate sum from
            var columnToPopulate = eb1.Text;
            //We have our headers so the first cell we'll calculate from is Users Choice + 2 i.e. G2
            var firstCell = columnToPopulate + "2";
            //Our last cell will be the chosen column and then the last row i.e. G16
            var lastCell = columnToPopulate + lastRow;
            //Here we populate i.e. Cell[17, G] will calculate formulate =SUM(G2:G16) 
            worksheet.Cells[lastRow + 1, columnToPopulate] = "=SUM(" + firstCell + ":" + lastCell + ")";
        }
    }
}
