using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.OneNote;

namespace TeXShift
{
    [Guid("1EE8F914-ECBD-4709-92C0-E770851C4966")]
    [ProgId("TeXShift.Connect")]
    [ComVisible(true)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {

    }
}