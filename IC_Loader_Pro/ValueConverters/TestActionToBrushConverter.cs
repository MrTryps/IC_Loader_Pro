using BIS_Tools_DataModels_2025;
using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace IC_Loader_Pro.ValueConverters
{
    public class TestActionToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is TestActionResponse action)
            {
                switch (action)
                {
                    case TestActionResponse.Pass:
                        return Brushes.Green;
                    case TestActionResponse.Note:
                    case TestActionResponse.Warn:
                        return Brushes.DarkOrange;
                    case TestActionResponse.Fail:
                    case TestActionResponse.Manual:
                        return Brushes.Red;
                    default:
                        return Brushes.Black;
                }
            }
            return Brushes.Black;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}