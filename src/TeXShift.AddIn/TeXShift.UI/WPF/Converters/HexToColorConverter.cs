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
                try
                {
                    hex = hex.TrimStart('#');
                    if (hex.Length == 6)
                    {
                        return Color.FromRgb(
                            System.Convert.ToByte(hex.Substring(0, 2), 16),
                            System.Convert.ToByte(hex.Substring(2, 2), 16),
                            System.Convert.ToByte(hex.Substring(4, 2), 16));
                    }
                    if (hex.Length == 8)
                    {
                        return Color.FromArgb(
                            System.Convert.ToByte(hex.Substring(0, 2), 16),
                            System.Convert.ToByte(hex.Substring(2, 2), 16),
                            System.Convert.ToByte(hex.Substring(4, 2), 16),
                            System.Convert.ToByte(hex.Substring(6, 2), 16));
                    }
                }
                catch
                {
                    // Return default on parse failure
                }
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
