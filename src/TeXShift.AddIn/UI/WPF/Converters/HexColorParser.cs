using System.Globalization;
using System.Windows.Media;

namespace TeXShift.AddIn.UI.WPF.Converters
{
    internal static class HexColorParser
    {
        public static Color ParseOrDefault(string hex, Color defaultColor)
        {
            return TryParse(hex, out var color) ? color : defaultColor;
        }

        public static bool TryParse(string hex, out Color color)
        {
            color = default(Color);
            hex = hex?.TrimStart('#');

            if (hex == null || (hex.Length != 6 && hex.Length != 8))
            {
                return false;
            }

            if (!byte.TryParse(hex.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var first)
                || !byte.TryParse(hex.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var second)
                || !byte.TryParse(hex.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var third))
            {
                return false;
            }

            if (hex.Length == 6)
            {
                color = Color.FromRgb(first, second, third);
                return true;
            }

            if (!byte.TryParse(hex.Substring(6, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var fourth))
            {
                return false;
            }

            color = Color.FromArgb(first, second, third, fourth);
            return true;
        }
    }
}
