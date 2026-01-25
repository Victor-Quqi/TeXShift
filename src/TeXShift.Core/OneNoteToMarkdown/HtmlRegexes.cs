using System.Text.RegularExpressions;

namespace TeXShift.Core.OneNoteToMarkdown
{
    /// <summary>
    /// Shared regex patterns for parsing OneNote inline HTML fragments.
    /// </summary>
    internal static class HtmlRegexes
    {
        internal static readonly Regex OuterSpan = new Regex(
            "^\\s*<span\\b(?<attrs>[^>]*)>(?<inner>.*)</span>\\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

        internal static readonly Regex StyleAttr = new Regex(
            "\\bstyle\\s*=\\s*(\"(?<v>[^\"]*)\"|'(?<v>[^']*)'|(?<v>[^\\s>]+))",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }
}

