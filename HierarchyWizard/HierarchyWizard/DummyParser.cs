using HierarchyWizard.Interfaces;
using System.Collections.Generic;
using HierarchyWizard.enums;
using WordParser;

namespace HierarchyWizard
{
    public class DummyParser : IDocumentParser
    {
        public DummyParser()
        {
        }

        public PageBatch GetPages()
        {
            return new PageBatch(GetPLPage(), GetBSPage());
        }

        private WordPage GetBSPage()
        {
            return new WordPage(new List<Line>()
            {
                //new Line("Fixed Assets", weight:FontWeight.Bold),
                //new Line("Investments"),
                //new Line("Tangible assets"),
                //new Line(""),
                //new Line(""),
                //new Line("Current assets", weight:FontWeight.Bold),
                //new Line("Stocks"),
                //new Line("Debtors"),
                //new Line("Current asset investments"),
                //new Line("Cash at bank and in hand"),
                //new Line(""),
                //new Line(""),
                //new Line(""),
                //new Line("Creditors: amounts falling due within one year"),
                //new Line(""),
                //new Line("Net current assets", weight:FontWeight.Bold),
                //new Line(""),
                //new Line("Total assets less current liabilities", weight:FontWeight.Bold),
                //new Line(""),
                //new Line("Provisions for liabilities", weight:FontWeight.Bold),
                //new Line("Deferred tax"),
                //new Line(""),
                //new Line(""),
                //new Line(""),
                //new Line("Net assets", weight:FontWeight.Bold),
                //new Line(""),
                //new Line("Capital and reserves", weight:FontWeight.Bold),
                //new Line("Called up share capital"),
                //new Line("Foreign exchange reserve"),
                //new Line("Profit and loss account"),
            });
        }

        private WordPage GetPLPage()
        {
            return new WordPage(new List<Line>()
            {
                //new Line("Turnover"),
                //new Line(""),
                //new Line("Other operating income"),
                //new Line(""),
                //new Line("Cost of sales"),
                //new Line(""),
                //new Line("Gross profit", weight:FontWeight.Bold),
                //new Line(""),
                //new Line("Administrative expenses"),
                //new Line(""),
                //new Line("Exceptional administrative expenses"),
                //new Line(""),
                //new Line("Operating profit", weight:FontWeight.Bold),
                //new Line(""),
                //new Line("Interest receivable and similar income"),
                //new Line(""),
                //new Line("Interest payable and similar expenses"),
                //new Line(""),
                //new Line("Profit before tax", weight:FontWeight.Bold),
                //new Line(""),
                //new Line("Tax on profit"),
                //new Line(""),
                //new Line(""),
                //new Line("Profit for the year", weight:FontWeight.Bold),
            });
        }
    }
}