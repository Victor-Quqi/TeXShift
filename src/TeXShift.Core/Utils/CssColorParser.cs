using System;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;

namespace TeXShift.Core.Utils
{
    internal static class CssColorParser
    {
        private static readonly Regex StyleAttributeRegex = new Regex(
            "(?:^|\\s)style\\s*=\\s*(\"(?<double>[^\"]*)\"|'(?<single>[^']*)'|(?<bare>[^\\s>]+))",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static bool TryGetColorFromAttributes(string attributes, out string normalizedHex)
        {
            normalizedHex = null;
            if (string.IsNullOrWhiteSpace(attributes))
            {
                return false;
            }

            var match = StyleAttributeRegex.Match(attributes);
            if (!match.Success)
            {
                return false;
            }

            string style = match.Groups["double"].Success
                ? match.Groups["double"].Value
                : match.Groups["single"].Success
                    ? match.Groups["single"].Value
                    : match.Groups["bare"].Value;

            return TryGetColorFromStyle(style, out normalizedHex);
        }

        public static bool TryGetColorFromStyle(string style, out string normalizedHex)
        {
            normalizedHex = null;
            if (string.IsNullOrWhiteSpace(style))
            {
                return false;
            }

            string colorValue = null;
            foreach (string declaration in style.Split(';'))
            {
                int colon = declaration.IndexOf(':');
                if (colon <= 0)
                {
                    continue;
                }

                string propertyName = declaration.Substring(0, colon).Trim();
                if (!string.Equals(propertyName, "color", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                colorValue = declaration.Substring(colon + 1).Trim();
                const string important = "!important";
                if (colorValue.EndsWith(important, StringComparison.OrdinalIgnoreCase))
                {
                    colorValue = colorValue.Substring(0, colorValue.Length - important.Length).TrimEnd();
                }
            }

            return TryNormalize(colorValue, out normalizedHex);
        }

        public static bool TryNormalize(string value, out string normalizedHex)
        {
            normalizedHex = null;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            string candidate = value.Trim();
            if (TryParseHex(candidate, out Color color) ||
                TryParseRgbFunction(candidate, out color) ||
                TryParseHslFunction(candidate, out color) ||
                TryParseNamedColor(candidate, out color))
            {
                normalizedHex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
                return true;
            }

            return false;
        }

        private static bool TryParseHex(string value, out Color color)
        {
            color = default(Color);
            if (string.IsNullOrEmpty(value) || value[0] != '#')
            {
                return false;
            }

            string hex = value.Substring(1);
            if (hex.Length == 3 || hex.Length == 4)
            {
                if (!TryParseHexByte(new string(hex[0], 2), out byte r) ||
                    !TryParseHexByte(new string(hex[1], 2), out byte g) ||
                    !TryParseHexByte(new string(hex[2], 2), out byte b))
                {
                    return false;
                }

                if (hex.Length == 4 && (!TryParseHexByte(new string(hex[3], 2), out byte a) || a != 255))
                {
                    return false;
                }

                color = Color.FromArgb(r, g, b);
                return true;
            }

            if (hex.Length == 6 || hex.Length == 8)
            {
                if (!TryParseHexByte(hex.Substring(0, 2), out byte r) ||
                    !TryParseHexByte(hex.Substring(2, 2), out byte g) ||
                    !TryParseHexByte(hex.Substring(4, 2), out byte b))
                {
                    return false;
                }

                if (hex.Length == 8 && (!TryParseHexByte(hex.Substring(6, 2), out byte a) || a != 255))
                {
                    return false;
                }

                color = Color.FromArgb(r, g, b);
                return true;
            }

            return false;
        }

        private static bool TryParseHexByte(string value, out byte result)
        {
            return byte.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out result);
        }

        private static bool TryParseRgbFunction(string value, out Color color)
        {
            color = default(Color);
            if (!TryGetFunctionBody(value, out string functionName, out string body) ||
                !(functionName == "rgb" || functionName == "rgba"))
            {
                return false;
            }

            if (!TrySplitColorFunctionArguments(body, out string[] components, out string alpha) ||
                components.Length != 3 ||
                !IsOpaqueAlpha(alpha) ||
                !TryParseRgbChannel(components[0], out byte r) ||
                !TryParseRgbChannel(components[1], out byte g) ||
                !TryParseRgbChannel(components[2], out byte b))
            {
                return false;
            }

            color = Color.FromArgb(r, g, b);
            return true;
        }

        private static bool TryParseRgbChannel(string value, out byte channel)
        {
            channel = 0;
            string candidate = (value ?? string.Empty).Trim();
            bool isPercent = candidate.EndsWith("%", StringComparison.Ordinal);
            if (isPercent)
            {
                candidate = candidate.Substring(0, candidate.Length - 1).TrimEnd();
            }

            if (!double.TryParse(candidate, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
            {
                return false;
            }

            double scaled = isPercent ? number * 255.0 / 100.0 : number;
            if (number < 0 || (isPercent ? number > 100 : number > 255))
            {
                return false;
            }

            channel = (byte)System.Math.Round(scaled, MidpointRounding.AwayFromZero);
            return true;
        }

        private static bool TryParseHslFunction(string value, out Color color)
        {
            color = default(Color);
            if (!TryGetFunctionBody(value, out string functionName, out string body) ||
                !(functionName == "hsl" || functionName == "hsla"))
            {
                return false;
            }

            if (!TrySplitColorFunctionArguments(body, out string[] components, out string alpha) ||
                components.Length != 3 ||
                !IsOpaqueAlpha(alpha) ||
                !TryParseHue(components[0], out double hue) ||
                !TryParsePercentage(components[1], out double saturation) ||
                !TryParsePercentage(components[2], out double lightness))
            {
                return false;
            }

            color = HslToColor(hue, saturation, lightness);
            return true;
        }

        private static bool TryParseHue(string value, out double degrees)
        {
            degrees = 0;
            string candidate = (value ?? string.Empty).Trim().ToLowerInvariant();
            double multiplier = 1.0;
            string suffix = null;

            foreach (string unit in new[] { "turn", "grad", "deg", "rad" })
            {
                if (candidate.EndsWith(unit, StringComparison.Ordinal))
                {
                    suffix = unit;
                    candidate = candidate.Substring(0, candidate.Length - unit.Length).TrimEnd();
                    break;
                }
            }

            if (!double.TryParse(candidate, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
            {
                return false;
            }

            switch (suffix)
            {
                case "turn":
                    multiplier = 360.0;
                    break;
                case "grad":
                    multiplier = 0.9;
                    break;
                case "rad":
                    multiplier = 180.0 / System.Math.PI;
                    break;
            }

            degrees = number * multiplier;
            degrees = ((degrees % 360.0) + 360.0) % 360.0;
            return true;
        }

        private static bool TryParsePercentage(string value, out double fraction)
        {
            fraction = 0;
            string candidate = (value ?? string.Empty).Trim();
            if (!candidate.EndsWith("%", StringComparison.Ordinal))
            {
                return false;
            }

            candidate = candidate.Substring(0, candidate.Length - 1).TrimEnd();
            if (!double.TryParse(candidate, NumberStyles.Float, CultureInfo.InvariantCulture, out double percent) ||
                percent < 0 || percent > 100)
            {
                return false;
            }

            fraction = percent / 100.0;
            return true;
        }

        private static Color HslToColor(double hueDegrees, double saturation, double lightness)
        {
            double hue = hueDegrees / 360.0;
            double r;
            double g;
            double b;

            if (saturation <= 0)
            {
                r = g = b = lightness;
            }
            else
            {
                double q = lightness < 0.5
                    ? lightness * (1.0 + saturation)
                    : lightness + saturation - lightness * saturation;
                double p = 2.0 * lightness - q;
                r = HueToRgb(p, q, hue + 1.0 / 3.0);
                g = HueToRgb(p, q, hue);
                b = HueToRgb(p, q, hue - 1.0 / 3.0);
            }

            return Color.FromArgb(ToByte(r), ToByte(g), ToByte(b));
        }

        private static double HueToRgb(double p, double q, double value)
        {
            if (value < 0) value += 1;
            if (value > 1) value -= 1;
            if (value < 1.0 / 6.0) return p + (q - p) * 6.0 * value;
            if (value < 1.0 / 2.0) return q;
            if (value < 2.0 / 3.0) return p + (q - p) * (2.0 / 3.0 - value) * 6.0;
            return p;
        }

        private static byte ToByte(double value)
        {
            double clamped = System.Math.Max(0, System.Math.Min(1, value));
            return (byte)System.Math.Round(clamped * 255.0, MidpointRounding.AwayFromZero);
        }

        private static bool TryGetFunctionBody(
            string value,
            out string functionName,
            out string body)
        {
            functionName = null;
            body = null;
            int open = value.IndexOf('(');
            if (open <= 0 || !value.EndsWith(")", StringComparison.Ordinal))
            {
                return false;
            }

            functionName = value.Substring(0, open).Trim().ToLowerInvariant();
            body = value.Substring(open + 1, value.Length - open - 2).Trim();
            return body.Length > 0;
        }

        private static bool TrySplitColorFunctionArguments(
            string body,
            out string[] components,
            out string alpha)
        {
            components = null;
            alpha = null;
            if (body.IndexOf(',') >= 0)
            {
                if (body.IndexOf('/') >= 0)
                {
                    return false;
                }

                string[] parts = body.Split(',');
                if (parts.Length != 3 && parts.Length != 4)
                {
                    return false;
                }

                components = new[] { parts[0].Trim(), parts[1].Trim(), parts[2].Trim() };
                alpha = parts.Length == 4 ? parts[3].Trim() : null;
                return true;
            }

            string channels = body;
            int slash = body.IndexOf('/');
            if (slash >= 0)
            {
                if (body.IndexOf('/', slash + 1) >= 0)
                {
                    return false;
                }

                channels = body.Substring(0, slash);
                alpha = body.Substring(slash + 1).Trim();
                if (alpha.Length == 0)
                {
                    return false;
                }
            }

            components = channels.Split(
                new[] { ' ', '\t', '\r', '\n' },
                StringSplitOptions.RemoveEmptyEntries);
            return components.Length == 3;
        }

        private static bool IsOpaqueAlpha(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return true;
            }

            string candidate = value.Trim();
            bool isPercent = candidate.EndsWith("%", StringComparison.Ordinal);
            if (isPercent)
            {
                candidate = candidate.Substring(0, candidate.Length - 1).TrimEnd();
            }

            if (!double.TryParse(candidate, NumberStyles.Float, CultureInfo.InvariantCulture, out double alpha))
            {
                return false;
            }

            return isPercent
                ? System.Math.Abs(alpha - 100.0) < 0.000001
                : System.Math.Abs(alpha - 1.0) < 0.000001;
        }

        private static bool TryParseNamedColor(string value, out Color color)
        {
            string name = value.Trim();
            switch (name.ToLowerInvariant())
            {
                case "grey": name = "gray"; break;
                case "dimgrey": name = "dimgray"; break;
                case "darkgrey": name = "darkgray"; break;
                case "lightgrey": name = "lightgray"; break;
                case "slategrey": name = "slategray"; break;
                case "darkslategrey": name = "darkslategray"; break;
                case "lightslategrey": name = "lightslategray"; break;
            }

            color = Color.FromName(name);
            return color.IsKnownColor && !color.IsSystemColor && color.A == 255;
        }
    }
}
