using System.Globalization;
using System.Resources;
using TeXShift.Core.Localization;

namespace TeXShift.AddIn.Localization
{
    public static class UIResources
    {
        private static readonly ResourceManager ResourceManager =
            new ResourceManager("TeXShift.AddIn.Localization.UIResources", typeof(UIResources).Assembly);

        public static CultureInfo Culture => Resources.Culture;

        public static string GetString(string name)
        {
            var value = ResourceManager.GetString(name, Culture);
            return string.IsNullOrEmpty(value) ? name : value;
        }
    }
}

