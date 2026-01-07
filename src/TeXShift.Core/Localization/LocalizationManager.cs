using System;
using System.Globalization;

namespace TeXShift.Core.Localization
{
    public static class LocalizationManager
    {
        private static CultureInfo _currentCulture;

        public static void Initialize(string languageCode = null)
        {
            if (string.IsNullOrEmpty(languageCode) || string.Equals(languageCode, "auto", StringComparison.OrdinalIgnoreCase))
            {
                var systemCulture = CultureInfo.CurrentUICulture;
                _currentCulture = systemCulture.Name.StartsWith("zh")
                    ? new CultureInfo("zh-CN")
                    : new CultureInfo("en-US");
            }
            else
            {
                try
                {
                    _currentCulture = new CultureInfo(languageCode);
                }
                catch (CultureNotFoundException)
                {
                    _currentCulture = new CultureInfo("en-US");
                }
            }

            Resources.Culture = _currentCulture;
        }

        public static string CurrentLanguage => _currentCulture?.Name ?? "en-US";
    }
}
