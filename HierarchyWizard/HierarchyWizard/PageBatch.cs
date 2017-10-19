namespace HierarchyWizard
{
    public class PageBatch
    {
        public WordPage ProfitLossPage { get; set; }
        public WordPage BalanceSheetPage { get; set; }

        public PageBatch()
        {
            ProfitLossPage = new WordPage();
            BalanceSheetPage = new WordPage();
        }

        public PageBatch(WordPage profitLoss, WordPage balanceSheet)
        {
            ProfitLossPage = profitLoss;
            BalanceSheetPage = balanceSheet;
        }
    }
}