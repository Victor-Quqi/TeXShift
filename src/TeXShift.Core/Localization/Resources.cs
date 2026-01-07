using System.Globalization;
using System.Resources;

namespace TeXShift.Core.Localization
{
    public static class Resources
    {
        private static readonly ResourceManager ResourceManager =
            new ResourceManager("TeXShift.Core.Localization.Resources", typeof(Resources).Assembly);

        public static CultureInfo Culture { get; set; }

        public static string GetString(string name)
        {
            var value = ResourceManager.GetString(name, Culture);
            return string.IsNullOrEmpty(value) ? name : value;
        }
    }
}
