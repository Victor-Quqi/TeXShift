using System;

namespace TeXShift.Core.Configuration
{
    /// <summary>
    /// Centralized font presets used by settings UIs.
    /// Keep this list in sync across WPF/WinForms dialogs.
    /// </summary>
    public static class FontPresets
    {
        public static readonly string[] MonospaceFonts =
        {
            "Consolas",
            "Courier New",
            "Source Code Pro",
            "Fira Code",
            "JetBrains Mono",
            "Cascadia Code",
            "Monaco",
            "Menlo"
        };

        public static bool ContainsMonospaceFont(string fontFamily)
        {
            if (string.IsNullOrWhiteSpace(fontFamily))
                return false;

            foreach (var font in MonospaceFonts)
            {
                if (string.Equals(font, fontFamily, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }
    }
}
