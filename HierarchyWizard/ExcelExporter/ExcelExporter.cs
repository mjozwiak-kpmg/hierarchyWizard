using System;

namespace ExcelExporter
{
    public static class ExcelExporter
    {
        public static void WriteHierarchy(string[] descriptions, string[] types)
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
                oSheet.Cells[1, 2] = "Type";

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "B1").Font.Bold = true;
                oSheet.get_Range("A1", "B1").VerticalAlignment =
                    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                for(int i = 0; i<descriptions.Length; i++)
                {
                    oSheet.Cells[i+2, 1] = descriptions[i];
                }
                for (int i = 0; i < types.Length; i++)
                {
                    oSheet.Cells[i + 2, 2] = types[i];
                }
            }
            catch (Exception ex)
            {

            }
        }
    }
}