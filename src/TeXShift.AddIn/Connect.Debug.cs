using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using TeXShift.AddIn.Localization;
using TeXShift.Core.Localization;
using TeXShift.Core.Logging;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace TeXShift.AddIn
{
    /// <summary>
    /// Partial class containing debug tools functionality.
    /// </summary>
    public partial class Connect
    {
        #region Debug Button Handlers

        /// <summary>
        /// Debug button: Shows and saves the selected content's XML structure only.
        /// </summary>
        public void OnDebugSelectionXmlButtonClick(IRibbonControl control)
        {
            try
            {
                if (TryShowInitializationError())
                {
                    return;
                }
                OnDebugSelectionXmlButtonClickAsync();
            }
            catch (Exception ex)
            {
                ShowTopMostMessageBox(
                    string.Format(UIResources.GetString("Debug_Exception"), ex),
                    UIResources.GetString("Debug_ExceptionTitle"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private async void OnDebugSelectionXmlButtonClickAsync()
        {
            try
            {
                var reader = _serviceContainer.CreateContentReader(_oneNoteApp);
                var result = await reader.ExtractContentAsync();

                if (!result.IsSuccess)
                {
                    ShowTopMostMessageBox(result.ErrorMessage, UIResources.GetString("Debug_Title"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (result.OriginalXmlNode == null)
                {
                    ShowTopMostMessageBox(UIResources.GetString("Debug_NoXmlNode"), UIResources.GetString("Debug_Title"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var logger = _serviceContainer.CreateDebugLogger(_appSettings?.Debug?.DebugOutputPath);
                logger.StartSession(DebugSessionKind.SelectionXml);

                // Use OriginalXmlNodes for multi-selection, fall back to OriginalXmlNode for single/cursor mode
                var (savedPath, formattedXml) = result.OriginalXmlNodes != null && result.OriginalXmlNodes.Count > 1
                    ? await logger.LogSelectionXmlAsync(result.OriginalXmlNodes)
                    : await logger.LogSelectionXmlAsync(result.OriginalXmlNode);

                int nodeCount = result.OriginalXmlNodes?.Count ?? 1;

                // Show in dialog
                string caption = string.Format(
                    UIResources.GetString("Debug_SelectionXmlCaption"),
                    GetDetectionModeLabel(result.Mode),
                    Path.GetFileName(savedPath));
                ShowTextInScrollableMessageBox(formattedXml, caption);

                // Show success message
                var savedMessage = string.Format(
                    UIResources.GetString("Debug_SelectionXmlSaved"),
                    savedPath,
                    GetDetectionModeLabel(result.Mode),
                    nodeCount > 1
                        ? string.Format(UIResources.GetString("Debug_SelectionXmlNodeCount"), nodeCount)
                        : result.OriginalXmlNode.Name.LocalName);

                if (result.TargetObjectIds != null && result.TargetObjectIds.Count > 0)
                {
                    savedMessage += $"{Environment.NewLine}{UIResources.GetString("Debug_SelectionXmlObjectIdsLabel")}: {string.Join(", ", result.TargetObjectIds)}";
                }

                ShowTopMostMessageBox(
                    savedMessage,
                    UIResources.GetString("Debug_SaveSuccess"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowTopMostMessageBox(
                    string.Format(UIResources.GetString("Debug_Exception"), ex.ToString()),
                    UIResources.GetString("Debug_ExceptionTitle"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Debug button: Shows and saves the raw OneNote XML structure for entire page.
        /// </summary>
        public void OnDebugXmlButtonClick(IRibbonControl control)
        {
            try
            {
                if (TryShowInitializationError())
                {
                    return;
                }
                OnDebugXmlButtonClickAsync();
            }
            catch (Exception ex)
            {
                ShowTopMostMessageBox(
                    string.Format(UIResources.GetString("Debug_Exception"), ex),
                    UIResources.GetString("Debug_ExceptionTitle"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private async void OnDebugXmlButtonClickAsync()
        {
            try
            {
                var (pageId, xmlContent) = await Task.Run(() =>
                {
                    OneNote.Windows windows = null;
                    OneNote.Window window = null;
                    try
                    {
                        windows = _oneNoteApp.Windows;
                        window = windows.CurrentWindow;
                        string id = window?.CurrentPageId;
                        if (string.IsNullOrEmpty(id)) return (null, null);
                        _oneNoteApp.GetPageContent(id, out string xml, OneNote.PageInfo.piAll, OneNote.XMLSchema.xs2013);
                        return (id, xml);
                    }
                    finally
                    {
                        SafeReleaseComObject(window);
                        SafeReleaseComObject(windows);
                    }
                });

                if (string.IsNullOrEmpty(pageId) || string.IsNullOrEmpty(xmlContent))
                {
                    ShowTopMostMessageBox(UIResources.GetString("Debug_NoPageContent"), UIResources.GetString("Debug_Title"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var logger = _serviceContainer.CreateDebugLogger(_appSettings?.Debug?.DebugOutputPath);
                logger.StartSession(DebugSessionKind.PageXml);
                string savedPath = await logger.LogPageXmlAsync(xmlContent);
                string formattedXml = System.Xml.Linq.XDocument.Parse(xmlContent).ToString();

                // Show in dialog
                string caption = string.Format(UIResources.GetString("Debug_PageXmlCaption"), Path.GetFileName(savedPath));
                ShowTextInScrollableMessageBox(formattedXml, caption);

                // Show success message
                ShowTopMostMessageBox(
                    string.Format(
                        UIResources.GetString("Debug_PageXmlSaved"),
                        savedPath,
                        new FileInfo(savedPath).Length / 1024.0),
                    UIResources.GetString("Debug_SaveSuccess"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowTopMostMessageBox(
                    string.Format(UIResources.GetString("Debug_Exception"), ex.ToString()),
                    UIResources.GetString("Debug_ExceptionTitle"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        #endregion
    }
}
