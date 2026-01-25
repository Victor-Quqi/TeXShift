using System.Collections.Generic;
using Microsoft.Office.Core;
using TeXShift.AddIn.Localization;

namespace TeXShift.AddIn
{
    /// <summary>
    /// Partial class containing Ribbon UI callback methods.
    /// </summary>
    public partial class Connect
    {
        #region Ribbon Label Callbacks

        private static readonly IReadOnlyDictionary<string, string> RibbonLabelKeys =
            new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase)
            {
                ["tabTeXShift"] = "Ribbon_Tab_TeXShift",

                ["grpConvertTools"] = "Ribbon_Group_ConvertTools",
                ["grpDebugTools"] = "Ribbon_Group_DebugTools",
                ["grpSettings"] = "Ribbon_Group_Settings",

                ["btnSilentConvert"] = "Ribbon_Button_Convert",
                ["btnReverseConvert"] = "Ribbon_Button_ReverseConvert",
                ["btnDebugConvert"] = "Ribbon_Button_DebugConvert",
                ["btnDebugReverseConvert"] = "Ribbon_Button_DebugReverseConvert",
                ["btnDebugXml"] = "Ribbon_Button_SavePageXml",
                ["btnDebugSelectionXml"] = "Ribbon_Button_SaveSelectionXml",
                ["btnSettings"] = "Ribbon_Button_Settings"
            };

        public string GetTabLabel(IRibbonControl control) => GetRibbonLabel(control?.Id);
        public string GetGroupLabel(IRibbonControl control) => GetRibbonLabel(control?.Id);
        public string GetButtonLabel(IRibbonControl control) => GetRibbonLabel(control?.Id);

        private static string GetRibbonLabel(string controlId)
        {
            if (string.IsNullOrWhiteSpace(controlId))
            {
                return string.Empty;
            }

            if (RibbonLabelKeys.TryGetValue(controlId, out var resourceKey))
            {
                return UIResources.GetString(resourceKey);
            }

            return controlId;
        }

        #endregion

        #region Ribbon Visibility Callbacks

        /// <summary>
        /// Ribbon callback: Returns whether the debug tools group should be visible.
        /// </summary>
        public bool GetDebugGroupVisible(IRibbonControl control)
        {
            return _appSettings?.Debug?.ShowDebugButtons ?? false;
        }

        #endregion
    }
}
