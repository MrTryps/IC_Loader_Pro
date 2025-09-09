using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace IC_Loader_Pro.ValueConverters
{
    /// <summary>
    /// Converts a list of strings to a single, semicolon-separated string and back.
    /// </summary>
    public class StringListConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is IEnumerable<string> list)
            {
                // Join the list into a single string for the TextBox
                return string.Join("; ", list);
            }
            return string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // When the user edits the TextBox, split the string back into a list
            if (value is string str)
            {
                return str.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                          .Select(s => s.Trim())
                          .ToList();
            }
            return new List<string>();
        }
    }
}