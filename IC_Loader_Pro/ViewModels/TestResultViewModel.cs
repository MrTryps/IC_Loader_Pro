using ArcGIS.Desktop.Framework.Contracts;
using BIS_Tools_DataModels_2025;
using IC_Rules_2025;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Media;

namespace IC_Loader_Pro.ViewModels
{
    public class TestResultViewModel : PropertyChangedBase
    {
        public string TestName { get; }
        public bool Passed { get; }
        public string Comments { get; }
        public string Icon { get; }

        // This property provides the CUMULATIVE action for both color and tooltip.
        public TestActionResponse FinalAction { get; }
        public bool IsExpanded { get; set; }

        public ObservableCollection<TestResultViewModel> SubResults { get; }
        public ObservableCollection<TestResultViewModel> RootResult => new ObservableCollection<TestResultViewModel> { this };

        public TestResultViewModel(IcTestResult model)
        {
            TestName = model.TestRule.Name;
            Passed = model.Passed;
            Comments = string.Join("; ", model.Comments);
            FinalAction = model.CumulativeAction.ResultAction;

            Icon = model.Passed ? "pack://application:,,,/ArcGIS.Desktop.Resources;component/Images/GenericCheckMark16.png" : "pack://application:,,,/ArcGIS.Desktop.Resources;component/Images/GenericError16.png";

            SubResults = new ObservableCollection<TestResultViewModel>(
                model.SubTestResults.Select(sr => new TestResultViewModel(sr))
            );
            IsExpanded = SubResults.Any() && CountDescendants(this) < 50;
        }
        private int CountDescendants(TestResultViewModel node)
        {
            int count = node.SubResults.Count;
            foreach (var child in node.SubResults)
            {
                count += CountDescendants(child);
            }
            return count;
        }
    }
}