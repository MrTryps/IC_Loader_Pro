using ArcGIS.Desktop.Framework.Contracts;
using IC_Rules_2025;
using System.Collections.ObjectModel;
using System.Linq;

namespace IC_Loader_Pro.ViewModels
{
    /// <summary>
    /// A ViewModel that wraps an IcTestResult to prepare it for display in a TreeView.
    /// </summary>
    public class TestResultViewModel : PropertyChangedBase
    {
        public string TestName { get; }
        public bool Passed { get; }
        public string Comments { get; }
        public string Icon { get; }

        /// <summary>
        /// A collection of child TestResultViewModels, which forms the tree structure.
        /// </summary>
        public ObservableCollection<TestResultViewModel> SubResults { get; }

        public TestResultViewModel(IcTestResult model)
        {
            TestName = model.TestRule.Name;
            Passed = model.Passed;
            Comments = string.Join("; ", model.Comments);

            // Determine which icon to show based on the pass/fail status
            Icon = Passed ? "pack://application:,,,/ArcGIS.Desktop.Resources;component/Images/GenericCheckMark16.png" : "pack://application:,,,/ArcGIS.Desktop.Resources;component/Images/GenericError16.png";

            // Recursively create ViewModels for all sub-test results
            SubResults = new ObservableCollection<TestResultViewModel>(
                model.SubTestResults.Select(sr => new TestResultViewModel(sr))
            );
        }
    }
}