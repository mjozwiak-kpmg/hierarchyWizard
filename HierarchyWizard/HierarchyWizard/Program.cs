using HierarchyWizard.Interfaces;

namespace HierarchyWizard
{
    class Program
    {
        static void Main(string[] args)
        {
            IDocumentParser pageServer = new DummyParser();
            var pages = pageServer.GetPages();
        }
    }
}
