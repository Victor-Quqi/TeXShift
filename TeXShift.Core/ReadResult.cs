using System.Xml.Linq;

namespace TeXShift.Core
{
    /// <summary>
    /// Defines the detected user selection mode in OneNote.
    /// </summary>
    public enum DetectionMode
    {
        /// <summary>
        /// No selection was detected.
        /// </summary>
        None,
        /// <summary>
        /// A cursor is placed within a text outline (no text selected).
        /// </summary>
        Cursor,
        /// <summary>
        /// A specific range of text is selected.
        /// </summary>
        Selection,
        /// <summary>
        /// An error occurred during detection.
        /// </summary>
        Error
    }

    /// <summary>
    /// Encapsulates the result of a content read operation from OneNote.
    /// </summary>
    public class ReadResult
    {
        public bool IsSuccess { get; set; }
        public string ExtractedText { get; set; }
        public DetectionMode Mode { get; set; }
        public string ErrorMessage { get; set; }

        /// <summary>
        /// The OneNote page ID where the content was read from.
        /// </summary>
        public string PageId { get; set; }

        /// <summary>
        /// The ObjectID of the node to be replaced (Outline in Cursor mode, OE in Selection mode).
        /// </summary>
        public string TargetObjectId { get; set; }

        /// <summary>
        /// The original XML node (used for preserving attributes during replacement).
        /// </summary>
        public XElement OriginalXmlNode { get; set; }

        public string ModeAsString()
        {
            switch (Mode)
            {
                case DetectionMode.Cursor:
                    return "光标模式 (操作整个文本框)";
                case DetectionMode.Selection:
                    return "选区模式 (只操作选中的文字)";
                default:
                    return "未知模式";
            }
        }
    }
}