using System;

namespace TeXShift.Core.Errors
{
    public class TeXShiftException : Exception
    {
        public string ErrorCode { get; }
        public string UserMessage { get; }

        public TeXShiftException(string errorCode, string userMessage, string technicalMessage, Exception inner = null)
            : base(technicalMessage, inner)
        {
            ErrorCode = errorCode;
            UserMessage = userMessage;
        }
    }

    public class ContentReadException : TeXShiftException
    {
        public ContentReadException(string userMessage, string technicalMessage, Exception inner = null)
            : base("TSE001", userMessage, technicalMessage, inner)
        {
        }
    }

    public class ContentWriteException : TeXShiftException
    {
        public ContentWriteException(string userMessage, string technicalMessage, Exception inner = null)
            : base("TSE002", userMessage, technicalMessage, inner)
        {
        }
    }

    public class MathConversionException : TeXShiftException
    {
        public MathConversionException(string userMessage, string technicalMessage, Exception inner = null)
            : base("TSE003", userMessage, technicalMessage, inner)
        {
        }
    }

    public class ImageLoadException : TeXShiftException
    {
        public ImageLoadException(string userMessage, string technicalMessage, Exception inner = null)
            : base("TSE004", userMessage, technicalMessage, inner)
        {
        }
    }

    public class MermaidConversionException : TeXShiftException
    {
        public MermaidConversionException(string userMessage, string technicalMessage, Exception inner = null)
            : base("TSE005", userMessage, technicalMessage, inner)
        {
        }
    }
}
