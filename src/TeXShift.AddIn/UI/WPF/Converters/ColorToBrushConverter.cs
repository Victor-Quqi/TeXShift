using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace TeXShift.AddIn.UI.WPF.Converters
{
    /// <summary>
    /// Converts between hex color strings and SolidColorBrush for binding to Background/Foreground.
    /// </summary>
    public class ColorToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string hex)
            {
                return new SolidColorBrush(HexColorParser.ParseOrDefault(hex, Colors.White));
            }
            else if (value is Color color)
            {
                return new SolidColorBrush(color);
            }
            return Brushes.White;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is SolidColorBrush brush)
            {
                var color = brush.Color;
                return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
            }
            return "#FFFFFF";
        }
    }
}
