using BIS_Tools_DataModels_2025;
using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace IC_Loader_Pro.ValueConverters
{
    public class TestActionToBackgroundBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is TestActionResponse action)
            {
                switch (action)
                {
                    case TestActionResponse.Pass:
                        return new SolidColorBrush(Colors.Green) { Opacity = 0.1 };

                    case TestActionResponse.Note:
                    case TestActionResponse.Warn:
                    case TestActionResponse.Manual:
                        return new SolidColorBrush(Colors.DarkOrange) { Opacity = 0.1 };

                    case TestActionResponse.Fail:
                    // IncompleteTest is also a failure, but less severe. Keeping it red.
                    case TestActionResponse.IncompleteTest:
                        return new SolidColorBrush(Colors.Red) { Opacity = 0.1 };

                    default:
                        return Brushes.Transparent;
                }
            }
            return Brushes.Transparent;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}