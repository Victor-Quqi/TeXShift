using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace TeXShift
{
    /// <summary>
    /// OneNote COM Add-in entry point.
    /// </summary>
    [ComVisible(true)]
    [Guid("1EE8F914-ECBD-4709-92C0-E770851C4966")]
    [ProgId("TeXShift.Connect")]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private object oneNoteApplication;
        private IRibbonUI ribbon;

        /// <summary>
        /// Called when the add-in is connected to OneNote.
        /// </summary>
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            oneNoteApplication = Application;
        }

        /// <summary>
        /// Called when the add-in is disconnected from OneNote.
        /// </summary>
        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            oneNoteApplication = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Called when the add-in is loaded on startup.
        /// </summary>
        public void OnStartupComplete(ref Array custom)
        {
        }

        /// <summary>
        /// Called when OneNote is shutting down.
        /// </summary>
        public void OnBeginShutdown(ref Array custom)
        {
        }

        /// <summary>
        /// Called when add-ins are updated.
        /// </summary>
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /// <summary>
        /// Returns the XML for the custom Ribbon UI.
        /// </summary>
        public string GetCustomUI(string RibbonID)
        {
            return GetResourceText("TeXShift.Ribbon.xml");
        }

        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnConvertButtonClick(IRibbonControl control)
        {
            MessageBox.Show("“一键转换”按钮被点击！", "TeXShift");
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            foreach (string name in resourceNames)
            {
                if (string.Compare(resourceName, name, System.StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(name)))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
    }
}
