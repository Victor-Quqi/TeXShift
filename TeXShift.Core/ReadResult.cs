using System.Collections.Generic;
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
        /// The ObjectIDs of the nodes to be replaced.
        /// In Cursor mode, this will contain one Outline ID.
        /// In Selection mode, this can contain multiple OE IDs for multi-line selections.
        /// </summary>
        public List<string> TargetObjectIds { get; set; } = new List<string>();

        /// <summary>
        /// The primary original XML node (used for preserving attributes during replacement).
        /// In multi-line selection, this is typically the first selected node.
        /// </summary>
        public XElement OriginalXmlNode { get; set; }

        /// <summary>
        /// All original XML nodes involved in the selection.
        /// </summary>
        public List<XElement> OriginalXmlNodes { get; set; } = new List<XElement>();

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