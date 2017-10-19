using HierarchyWizard.Interfaces;

namespace HierarchyWizard
{
    public static class ClassificationController
    {
        public static IClassifier GetClassifier()
        {
            // TODO: Implement classification strategy
            return new BlindClassifier();
        }
    }
}