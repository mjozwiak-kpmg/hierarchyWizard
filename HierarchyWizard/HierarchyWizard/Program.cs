using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WordParser;

namespace HierarchyWizard
{
    class Program
    {

        //"Description";
        //"Bold";
        //"Italic";
        //"Balance Sheet";
        //"Note";
        //"Type";
        //"Parent";

        static void Main(string[] args)
        {
            string[] files = Directory.GetFiles("C:\\Users\\jsayer\\temp\\Hackathon - Word Docs", "*");
            foreach(var file in files)
            {
                var data = Parser.ParseFile(file);
                var nonEmptyLines = data.Where(l => !string.IsNullOrEmpty(l.Description));
                PrintDataToExcel(nonEmptyLines, GetSavePath(file));
            }
            //var lines = from datum in data
            //            select datum.ToLine();

            //IDocumentParser parser = new DummyParser();
            //var pages = parser.GetPages();
            //var classifier = ClassificationController.GetClassifier();
            //classifier.Classify(pages);
            
                //pages.BalanceSheetPage.Lines.Select(l => l.Description).ToArray(), 
                //pages.BalanceSheetPage.Lines.Select(l => l.Parent).ToArray());
        }

        private static string GetSavePath(string file)
        {
            var temp = Path.Combine("C:\\Users\\jsayer\\temp\\Hackathon - Excel Exports", Path.GetFileNameWithoutExtension(file));
            Path.ChangeExtension(temp, ".xlsx");
            return temp+".xlsx";
        }

        private static void PrintDataToExcel(IEnumerable<Line> data, string savePath = null)
        {
            ExcelExporter.ExcelExporter.WriteHierarchy(
                new string[][]{
                    data.Select(l => l.Description).ToArray(),
                    data.Select(l => l.IsBold.ToString()).ToArray(),
                    data.Select(l => l.IsItalic.ToString()).ToArray(),
                    data.Select(l => l.IsBalanceSheet.ToString()).ToArray(),
                    data.Select(l => l.HasNote.ToString()).ToArray(),
                    data.Select(l => "").ToArray(),
                    data.Select(l => l.Parent).ToArray(),
                                    }
                , savePath);
        }
    }
}
