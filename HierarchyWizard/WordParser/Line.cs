using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace WordParser
{
    public class Line
    {
        public string Description { get; set; }
        public bool IsBalanceSheet { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool HasNote { get; set; }
        public string BalanceOne { get; set; }
        public string BalanceTwo { get; set; }
        public bool IsEmpty { get; set; }
        public string Parent { get; set; }

        public Line(string description, bool isBalanceSheet, bool isBold = false, bool isItalic = false)
        {
            Description = description;
            IsBalanceSheet = isBalanceSheet;
            IsBold = isBold;
            IsItalic = isItalic;
            IsEmpty = string.IsNullOrEmpty(description);
            Parent = "";
        }

        public Line(Row row, bool isBalanceSheet)
        {
            var rowAsString = row.Range.Text.Trim();
            var cells = rowAsString.Split('\a');
            Description = cells[0].Trim();
            IsEmpty = string.IsNullOrEmpty(Description);
            IsBold = row.Range.Font.Bold == -1;
            HasNote = !string.IsNullOrEmpty(cells[1].Trim());
            IsBalanceSheet = isBalanceSheet;
            int i = 2;
            while (i<cells.Length && string.IsNullOrEmpty(cells[i]))
            {
                i+=1;
            }
            if (i < cells.Length)
            {
                BalanceOne = cells[i];
            }

            while (i < cells.Length && string.IsNullOrEmpty(cells[i]))
            {
                i += 1;
            }
            if (i < cells.Length)
            {
                BalanceTwo = cells[i];
            }
            Parent = "";
            
            
            // TODO: Italic
            
            var r = 2;


            var s = 3;


            var you = 7;
        }
    }
}
