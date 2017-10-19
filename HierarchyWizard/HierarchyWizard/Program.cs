using HierarchyWizard.Interfaces;
using System.Linq;
using WordParser;

namespace HierarchyWizard
{
    class Program
    {
        static void Main(string[] args)
        {
            var data = Parser.ParseFile("C:\\Users\\mjozwiak1\\Desktop\\DH Ventures Ltd.docx");

            IDocumentParser parser = new DummyParser();
            var pages = parser.GetPages();
            var classifier = ClassificationController.GetClassifier();
            classifier.Classify(pages);
            ExcelExporter.ExcelExporter.WriteHierarchy(
                pages.BalanceSheetPage.Lines.Select(l => l.Text).ToArray(), 
                pages.BalanceSheetPage.Lines.Select(l => l.Classification.ToString()).ToArray());
        }
    }
}
