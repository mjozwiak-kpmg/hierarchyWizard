using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace WordParser
{
    public class Parser
    {
        public static List<Line> ParseFile(string file)
        {
            string[] profitAndLoss = new[] { "interest", "turnover", "expenses", "expenditure", "income", "sales" };
            string[] balanceSheet = new[] { "liabilities", "capital", "reserves", "stock", "creditors", "debtors", "share" };

            var app = new Microsoft.Office.Interop.Word.Application();
            object filePath = file;
            object readonlyAccess = true;
            var doc = app.Documents.Open(ref filePath, ReadOnly:readonlyAccess);

            var balanceSheetRows = new List<Line>();
            var pAndlRows = new List<Line>();

            var profitAndLossTable = GetMostLikelyTable(doc, profitAndLoss);
            var balanceSheetTable = GetMostLikelyTable(doc, balanceSheet);

            var range = profitAndLossTable.Range;
            var rows = profitAndLossTable.Rows;
            var text = rows.ToString();

            for (int j = 0; j < rows.Count; j++)
            {
                var row = rows[j + 1];
                pAndlRows.Add(new Line(row, false));
            }

            range = balanceSheetTable.Range;
            rows = balanceSheetTable.Rows;

            for (int j = 0; j < rows.Count; j++)
            {
                var row = rows[j + 1];
                balanceSheetRows.Add(new Line(row, true));
            }

            return balanceSheetRows.Concat(pAndlRows).ToList();
        }

        private static Table GetMostLikelyTable(Document doc, string[] searchWords)
        {
            var maxCount = 0;
            Table currentTable = null;
            for (int i = 0; i < doc.Tables.Count; i++)
            {
                var table = doc.Tables[i + 1];
                var text = table.Range.Text.Trim().ToLowerInvariant();

                if (string.IsNullOrWhiteSpace(text))
                    continue;

                var currentCount = CountNumberOfOccurrences(text, searchWords);

                if (currentCount > maxCount)
                {
                    maxCount = currentCount;
                    currentTable = table;
                }
            }
            return currentTable;
        }

        private static int CountNumberOfOccurrences(string text, string[] searchWords)
        {
            var count = 0;
            foreach(var word in searchWords)
            {
                if (text.Contains(word)) count += 1;
            }
            return count;
        }
    }
}
