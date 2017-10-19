using HierarchyWizard.enums;
using HierarchyWizard.Interfaces;

namespace HierarchyWizard
{
    public class BlindClassifier : IClassifier
    {
        public void Classify(PageBatch pages)
        {
            var lines = pages.ProfitLossPage.Lines;
            foreach(var line in lines)
            {
                line.Classification = Classification.Binding;
            }

            lines = pages.BalanceSheetPage.Lines;
            foreach (var line in lines)
            {
                line.Classification = Classification.Sigma;
            }
        }
    }
}