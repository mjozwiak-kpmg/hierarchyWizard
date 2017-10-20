using System;
using Microsoft.Office.Interop.Excel;

namespace ExcelExporter
{
    public static class ExcelExporter
    {
        public static void WriteHierarchy(string[][] data, string savePath)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;
            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "Description";
                oSheet.Cells[1, 2] = "Bold";
                oSheet.Cells[1, 3] = "Italic";
                oSheet.Cells[1, 4] = "Balance Sheet";
                oSheet.Cells[1, 5] = "Note";
                oSheet.Cells[1, 6] = "Type";
                oSheet.Cells[1, 7] = "Parent";

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "B1").Font.Bold = true;
                oSheet.get_Range("A1", "B1").VerticalAlignment =
                    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                
                for(int i = 0; i < data.Length; i++)
                {
                    WriteColumn(oSheet, data[i], i + 1);
                }

                if(savePath != null)
                {
                    oXL.ActiveWorkbook.SaveAs(savePath);
                    oXL.Quit();
                }
            }
            catch (Exception ex)
            {

            }
        }

        private static void WriteColumn(_Worksheet oSheet, string[] array, int colIndex)
        {
            for (int i = 0; i < array.Length; i++)
            {
                oSheet.Cells[i + 2, colIndex] = array[i];
            }
        }
    }
}