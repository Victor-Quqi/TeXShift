using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using TeXShift.Core.Localization;

namespace TeXShift.Core.Errors
{
    public static class ErrorMessages
    {
        private static readonly Dictionary<int, string> OneNoteErrorMessageKeys = new Dictionary<int, string>
        {
            { unchecked((int)0x80042000), "Error_OneNote_MalformedXml" },
            { unchecked((int)0x80042001), "Error_OneNote_InvalidXml" },
            { unchecked((int)0x80042005), "Error_OneNote_PageNotFound" },
            { unchecked((int)0x8004200b), "Error_OneNote_SectionReadOnly" },
            { unchecked((int)0x8004200c), "Error_OneNote_PageReadOnly" },
            { unchecked((int)0x80042010), "Error_OneNote_PageModified" },
            { unchecked((int)0x80042013), "Error_OneNote_NoActiveSelection" },
            { unchecked((int)0x80042014), "Error_OneNote_ObjectNotFound" },
            { unchecked((int)0x8004201b), "Error_OneNote_SectionEncrypted" },
            { unchecked((int)0x8004201d), "Error_OneNote_NotYetSynchronized" },
            { unchecked((int)0x80042023), "Error_OneNote_TimeOut" },
            { unchecked((int)0x80042030), "Error_OneNote_ModalDialog" }
        };

        public static string GetUserFriendlyMessage(Exception ex)
        {
            if (ex == null)
                return Resources.GetString("Error_GenericUnexpected");

            if (ex is TeXShiftException texShiftException)
                return string.IsNullOrWhiteSpace(texShiftException.UserMessage)
                    ? Resources.GetString("Error_GenericUnexpected")
                    : texShiftException.UserMessage;

            if (ex is COMException comEx && OneNoteErrorMessageKeys.TryGetValue(comEx.HResult, out var key))
                return Resources.GetString(key);

            if (ex is COMException)
                return Resources.GetString("Error_ComCommunication");

            if (ex is TimeoutException)
                return Resources.GetString("Error_Timeout");

            if (ex is IOException)
                return Resources.GetString("Error_Io");

            return Resources.GetString("Error_GenericUnexpected");
        }
    }
}
