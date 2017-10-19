using HierarchyWizard.enums;

namespace HierarchyWizard
{
    public class Line
    {
        public string Text { get; set; }
        public FontWeight Weight { get; set; }
        public FontStyle Style { get; set; }
        public bool IsEmpty { get; set; }
        public Classification Classification { get; set; }

        public Line(string text, FontWeight weight = FontWeight.Normal, FontStyle style = FontStyle.Normal)
        {
            Text = text;
            Weight = weight;
            Style = style;
            IsEmpty = string.IsNullOrEmpty(text);
            Classification = Classification.Unclassified;
        }
    }
}