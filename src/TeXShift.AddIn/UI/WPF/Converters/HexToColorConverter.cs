using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace TeXShift.AddIn.UI.WPF.Converters
{
    /// <summary>
    /// Converts between hex color strings (e.g., "#FF0000") and WPF Color.
    /// </summary>
    public class HexToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string hex)
            {
                return HexColorParser.ParseOrDefault(hex, Colors.White);
            }
            return Colors.White;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Color color)
            {
                return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
            }
            return "#FFFFFF";
        }
    }
}
