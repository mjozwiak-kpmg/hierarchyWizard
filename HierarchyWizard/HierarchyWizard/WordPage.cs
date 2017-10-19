using System.Collections.Generic;

namespace HierarchyWizard
{
    public class WordPage
    {
        public List<Line> Lines { get; set; }

        public WordPage()
        {
            Lines = new List<Line>();
        }

        public WordPage(List<Line> lines)
        {
            Lines = lines;
        }
    }
}
