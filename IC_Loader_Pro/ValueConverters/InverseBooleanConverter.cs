// In IC_Loader_Pro/ValueConverters/InverseBooleanConverter.cs

using System;
using System.Globalization;
using System.Windows.Data;

namespace IC_Loader_Pro.ValueConverters
{
    /// <summary>
    /// A simple value converter that inverts a boolean value.
    /// True becomes False, and False becomes True.
    /// </summary>
    public class InverseBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return !boolValue;
            }
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return !boolValue;
            }
            return false;
        }
    }
}