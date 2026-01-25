using System;
using System.Collections.Generic;
using System.Linq;

namespace TeXShift.Core.Utils
{
    internal static class MarkdownTableExtractor
    {
        internal sealed class MarkdownTableBlock
        {
            public string Text { get; }
            public string Key { get; }

            public MarkdownTableBlock(string text, string key)
            {
                Text = text ?? string.Empty;
                Key = key ?? string.Empty;
            }
        }

        internal static List<MarkdownTableBlock> ExtractMarkdownTables(string markdown)
        {
            var results = new List<MarkdownTableBlock>();
            if (string.IsNullOrEmpty(markdown))
            {
                return results;
            }

            var lines = NormalizeLines(markdown);
            bool inCodeFence = false;
            string fenceToken = null;

            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i] ?? string.Empty;
                var trimmed = line.TrimStart();

                if (IsFenceLine(trimmed, out var token))
                {
                    if (!inCodeFence)
                    {
                        inCodeFence = true;
                        fenceToken = token;
                        continue;
                    }

                    if (string.Equals(fenceToken, token, StringComparison.Ordinal))
                    {
                        inCodeFence = false;
                        fenceToken = null;
                        continue;
                    }
                }

                if (inCodeFence)
                {
                    continue;
                }

                if (!LooksLikeTableRow(line) || i + 1 >= lines.Length)
                {
                    continue;
                }

                var sep = lines[i + 1] ?? string.Empty;
                if (!IsMarkdownTableSeparatorLine(sep))
                {
                    continue;
                }

                int start = i;
                int end = i + 2;
                while (end < lines.Length && LooksLikeTableRow(lines[end]))
                {
                    end++;
                }

                var blockLines = new List<string>();
                for (int j = start; j < end; j++)
                {
                    blockLines.Add(lines[j] ?? string.Empty);
                }

                var blockText = string.Join("\r\n", blockLines).TrimEnd();
                if (TryBuildMarkdownTableKey(blockLines, out var key))
                {
                    results.Add(new MarkdownTableBlock(blockText, key));
                }

                i = end - 1;
            }

            return results;
        }

        internal static bool TryExtractSingleMarkdownTableKey(string markdown, out string key)
        {
            key = null;
            var blocks = ExtractMarkdownTables(markdown ?? string.Empty);
            if (blocks.Count != 1)
            {
                return false;
            }

            // Only accept when the rendered output is effectively just this table (ignoring surrounding whitespace).
            var normalizedWhole = NormalizeLineEndings(markdown).Trim();
            var normalizedBlock = NormalizeLineEndings(blocks[0].Text).Trim();
            if (!string.Equals(normalizedWhole, normalizedBlock, StringComparison.Ordinal))
            {
                return false;
            }

            key = blocks[0].Key;
            return !string.IsNullOrEmpty(key);
        }

        private static bool TryBuildMarkdownTableKey(List<string> blockLines, out string key)
        {
            key = null;
            if (blockLines == null || blockLines.Count < 2)
            {
                return false;
            }

            var header = SplitTableRow(blockLines[0]);
            if (header.Count == 0)
            {
                return false;
            }

            var rows = new List<List<string>>();
            rows.Add(header);

            for (int i = 2; i < blockLines.Count; i++)
            {
                var row = SplitTableRow(blockLines[i]);
                rows.Add(row);
            }

            int columnCount = rows.Max(r => r.Count);
            for (int i = 0; i < rows.Count; i++)
            {
                while (rows[i].Count < columnCount)
                {
                    rows[i].Add(string.Empty);
                }
            }

            var rowKeys = rows.Select(r => string.Join("\u001F", r.Select(NormalizeTableCellKey)));
            key = string.Join("\u001E", rowKeys);
            return true;
        }

        private static string NormalizeTableCellKey(string cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }

            var trimmed = cell.Trim();
            return UnescapeMarkdownPipe(trimmed);
        }

        private static string UnescapeMarkdownPipe(string text)
        {
            if (string.IsNullOrEmpty(text) || text.IndexOf('\\') < 0)
            {
                return text ?? string.Empty;
            }

            var sb = new System.Text.StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (ch == '\\' && i + 1 < text.Length)
                {
                    char next = text[i + 1];
                    if (next == '|' || next == '\\')
                    {
                        sb.Append(next);
                        i++;
                        continue;
                    }
                }

                sb.Append(ch);
            }

            return sb.ToString();
        }

        private static List<string> SplitTableRow(string line)
        {
            var results = new List<string>();
            if (string.IsNullOrWhiteSpace(line))
            {
                return results;
            }

            var s = line.Trim();

            // Strip optional leading/trailing pipes.
            if (s.StartsWith("|", StringComparison.Ordinal))
            {
                s = s.Substring(1);
            }
            if (s.EndsWith("|", StringComparison.Ordinal))
            {
                s = s.Substring(0, s.Length - 1);
            }

            var sb = new System.Text.StringBuilder();
            bool escaping = false;
            foreach (var ch in s)
            {
                if (escaping)
                {
                    sb.Append(ch);
                    escaping = false;
                    continue;
                }

                if (ch == '\\')
                {
                    sb.Append(ch);
                    escaping = true;
                    continue;
                }

                if (ch == '|')
                {
                    results.Add(sb.ToString().Trim());
                    sb.Clear();
                    continue;
                }

                sb.Append(ch);
            }

            results.Add(sb.ToString().Trim());
            return results;
        }

        private static bool LooksLikeTableRow(string line)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                return false;
            }

            return line.IndexOf('|') >= 0;
        }

        private static bool IsMarkdownTableSeparatorLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                return false;
            }

            bool hasDash = false;
            foreach (var ch in line)
            {
                if (ch == '-')
                {
                    hasDash = true;
                    continue;
                }

                if (ch == '|' || ch == ':' || ch == ' ' || ch == '\t')
                {
                    continue;
                }

                return false;
            }

            return hasDash;
        }

        private static string[] NormalizeLines(string text)
        {
            return (text ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace("\r", "\n")
                .Split(new[] { '\n' }, StringSplitOptions.None);
        }

        private static bool IsFenceLine(string trimmedLine, out string token)
        {
            token = null;
            if (string.IsNullOrEmpty(trimmedLine))
            {
                return false;
            }

            if (trimmedLine.StartsWith("```", StringComparison.Ordinal))
            {
                token = "```";
                return true;
            }

            if (trimmedLine.StartsWith("~~~", StringComparison.Ordinal))
            {
                token = "~~~";
                return true;
            }

            return false;
        }

        private static string NormalizeLineEndings(string text)
        {
            return (text ?? string.Empty).Replace("\r\n", "\n").Replace("\r", "\n");
        }
    }
}

