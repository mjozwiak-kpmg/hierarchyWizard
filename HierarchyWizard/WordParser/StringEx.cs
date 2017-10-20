namespace WordParser
{
    public static class StringEx
    {
        public static Line ToLine(this string row, bool isBalanceSheet)
        {
            var cells = row.Split('\a');
            return new Line(cells[0], isBalanceSheet);
        }
    }
}
