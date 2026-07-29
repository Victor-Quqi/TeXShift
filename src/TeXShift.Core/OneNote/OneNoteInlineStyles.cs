using System;

namespace TeXShift.Core.OneNote
{
    internal static class OneNoteInlineStyles
    {
        public const string HighlightColor = "#FFFF00";
        public const string HighlightCss = "background-color:" + HighlightColor;
        public const string UnderlineCss = "text-decoration:underline";
        public const string StrikeCss = "text-decoration:line-through";

        public static bool IsCanonicalHighlightColor(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var normalized = RemoveWhitespace(value).ToLowerInvariant();
            return normalized == "#ffff00"
                || normalized == "#ff0"
                || normalized == "yellow"
                || normalized == "rgb(255,255,0)"
                || normalized == "rgba(255,255,0,1)";
        }

        public static bool IsVisibleBackgroundColor(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var normalized = RemoveWhitespace(value).ToLowerInvariant();
            return normalized != "none"
                && normalized != "transparent"
                && normalized != "white"
                && normalized != "#fff"
                && normalized != "#ffffff"
                && normalized != "rgb(255,255,255)"
                && normalized != "rgba(255,255,255,0)";
        }

        private static string RemoveWhitespace(string value)
        {
            var chars = new char[value.Length];
            int length = 0;
            foreach (char ch in value)
            {
                if (!char.IsWhiteSpace(ch))
                {
                    chars[length++] = ch;
                }
            }

            return new string(chars, 0, length);
        }
    }
}
