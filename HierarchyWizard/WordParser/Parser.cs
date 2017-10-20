using System.Collections.Generic;
using System.Linq;

namespace WordParser
{
    public class Parser
    {
        public static List<string> ParseFile(string file)
        {
            string[] profitAndLoss = new[] { "tax on profit", "financial year" };
            string[] balanceSheet = new[] { "fixed assets", "current assets", "net assets", "profit and loss" };

            var app = new Microsoft.Office.Interop.Word.Application();
            object filePath = file;
            object readonlyAccess = true;
            var doc = app.Documents.Open(ref filePath, ReadOnly:readonlyAccess);

            List<string> balanceSheetRows = new List<string>();
            List<string> pAndlRows = new List<string>();

            for (int i = 0; i < doc.Tables.Count; i++)
            {
                var table = doc.Tables[i + 1];
                var text = table.Range.Text.Trim().ToLowerInvariant();

                if (string.IsNullOrWhiteSpace(text))
                    continue;

                if (profitAndLoss.All(text.Contains))
                {
                    var range = table.Range;
                    var rows = table.Rows;

                    for (int j = 0; j < rows.Count; j++)
                    {
                        var row = rows[j + 1];
                        for (int x = 0; x < row.Cells.Count; x++)
                        {
                            var cell = row.Cells[x + 1];
                            var cellText = cell.Range.Text.Trim();
                            if (string.IsNullOrEmpty(cellText) == false)
                                pAndlRows.Add(row.Range.Text.Trim());
                        }


                        pAndlRows.Add(row.Range.Text.Trim());
                    }
                }

                if (balanceSheet.All(text.Contains))
                {
                    var range = table.Range;
                    var rows = table.Rows;

                    for (int j = 0; j < rows.Count; j++)
                    {
                        var row = rows[j + 1];
                        balanceSheetRows.Add(row.Range.Text.Trim());
                    }
                }
            }

            return balanceSheetRows.Concat(pAndlRows).ToList();
        }
    }
}
