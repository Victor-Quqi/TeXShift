using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using TeXShift.Core;
using OneNote = Microsoft.Office.Interop.OneNote;
 
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
         private OneNote.Application _oneNoteApp;
         private IRibbonUI ribbon;
 
         /// <summary>
         /// Called when the add-in is connected to OneNote.
         /// </summary>
         public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
         {
             _oneNoteApp = (OneNote.Application)Application;
         }
 
         /// <summary>
         /// Called when the add-in is disconnected from OneNote.
         /// </summary>
         public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
         {
             _oneNoteApp = null;
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
            try
            {
                var reader = new ContentReader(_oneNoteApp);
                var result = reader.ExtractContent();

                if (result.IsSuccess)
                {
                    ShowTextInScrollableMessageBox(result.ExtractedText, $"检测到: {result.ModeAsString()}");
                }
                else
                {
                    MessageBox.Show(result.ErrorMessage, "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("插件发生未知错误：\n" + ex.ToString(), "插件异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Helper function: Creates a form with a scrollbar to display a large amount of text.
        /// </summary>
        private void ShowTextInScrollableMessageBox(string text, string caption)
        {
            Form form = new Form
            {
                Text = caption,
                Size = new System.Drawing.Size(600, 400),
                StartPosition = FormStartPosition.CenterParent
            };
            TextBox textBox = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 10),
                Text = text
            };
            form.Controls.Add(textBox);
            form.ShowDialog();
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
