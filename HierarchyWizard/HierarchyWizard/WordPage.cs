using System.Collections.Generic;
using WordParser;

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
