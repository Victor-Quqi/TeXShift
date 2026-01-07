using System;

namespace TeXShift.Core.Errors
{
    public class MermaidConversionException : TeXShiftException
    {
        public MermaidConversionException(string userMessage, string technicalMessage, Exception inner = null)
            : base("TSE005", userMessage, technicalMessage, inner)
        {
        }
    }
}

