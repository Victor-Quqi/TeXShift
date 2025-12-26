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
                try
                {
                    hex = hex.TrimStart('#');
                    if (hex.Length == 6)
                    {
                        var color = Color.FromRgb(
                            System.Convert.ToByte(hex.Substring(0, 2), 16),
                            System.Convert.ToByte(hex.Substring(2, 2), 16),
                            System.Convert.ToByte(hex.Substring(4, 2), 16));
                        return new SolidColorBrush(color);
                    }
                    if (hex.Length == 8)
                    {
                        var color = Color.FromArgb(
                            System.Convert.ToByte(hex.Substring(0, 2), 16),
                            System.Convert.ToByte(hex.Substring(2, 2), 16),
                            System.Convert.ToByte(hex.Substring(4, 2), 16),
                            System.Convert.ToByte(hex.Substring(6, 2), 16));
                        return new SolidColorBrush(color);
                    }
                }
                catch
                {
                    // Return default on parse failure
                }
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
