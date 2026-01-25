using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace TeXShift.Core.OneNoteMeta
{
    internal static class TeXShiftMetaReader
    {
        internal sealed class TeXShiftMetaReadResult
        {
            public bool HasTeXShiftMeta { get; set; }
            public bool IsValid { get; set; }
            public string Source { get; set; }
            public string Mode { get; set; }
            public string FailureReason { get; set; }
        }

        public static TeXShiftMetaReadResult ReadOutline(XElement outline)
        {
            var result = new TeXShiftMetaReadResult();
            if (outline == null)
            {
                return result;
            }

            var ns = outline.Name.Namespace;
            var allMeta = outline.Elements(ns + "Meta").ToList();
            if (allMeta.Count == 0)
            {
                return result;
            }

            var texshiftMeta = allMeta.Where(IsTeXShiftMeta).ToList();
            if (texshiftMeta.Count == 0)
            {
                return result;
            }

            result.HasTeXShiftMeta = true;

            string schema = GetMetaContent(texshiftMeta, TeXShiftMetaKeys.Schema);
            if (!string.Equals(schema, TeXShiftMetaKeys.SchemaVersion, StringComparison.Ordinal))
            {
                result.FailureReason = "TeXShift meta schema mismatch.";
                return result;
            }

            result.Mode = GetMetaContent(texshiftMeta, TeXShiftMetaKeys.Mode);

            string encoding = GetMetaContent(texshiftMeta, TeXShiftMetaKeys.SourceEncoding);
            if (!string.Equals(encoding, TeXShiftMetaKeys.EncodingPlainV1, StringComparison.Ordinal))
            {
                result.FailureReason = "TeXShift meta source encoding not supported.";
                return result;
            }

            string encodedSource = ReadSourceChunks(texshiftMeta, out string chunkError);
            if (chunkError != null)
            {
                result.FailureReason = chunkError;
                return result;
            }

            string decodedSource = DecodePlainV1(encodedSource);

            string sigVersion = GetMetaContent(texshiftMeta, TeXShiftMetaKeys.SigVersion);
            if (!string.Equals(sigVersion, TeXShiftMetaKeys.SigVersionValue, StringComparison.Ordinal))
            {
                result.FailureReason = "TeXShift meta signature version not supported.";
                return result;
            }

            string sig = GetMetaContent(texshiftMeta, TeXShiftMetaKeys.Sig);
            if (string.IsNullOrWhiteSpace(sig))
            {
                result.FailureReason = "TeXShift meta signature missing.";
                return result;
            }

            string computedSig = TeXShiftMetaWriter.ComputeSignature(outline);
            if (!string.Equals(sig, computedSig, StringComparison.Ordinal))
            {
                result.FailureReason = "TeXShift meta signature mismatch.";
                return result;
            }

            result.IsValid = true;
            result.Source = decodedSource ?? string.Empty;
            return result;
        }

        private static bool IsTeXShiftMeta(XElement meta)
        {
            var name = (string)meta.Attribute("name");
            return !string.IsNullOrEmpty(name)
                && name.StartsWith(TeXShiftMetaKeys.Prefix, StringComparison.OrdinalIgnoreCase);
        }

        private static string GetMetaContent(IEnumerable<XElement> metas, string name)
        {
            var meta = metas.FirstOrDefault(m =>
                string.Equals((string)m.Attribute("name"), name, StringComparison.OrdinalIgnoreCase));
            return (string)meta?.Attribute("content");
        }

        private static string ReadSourceChunks(IEnumerable<XElement> metas, out string error)
        {
            error = null;
            var chunks = new SortedDictionary<int, string>();

            foreach (var meta in metas)
            {
                var name = (string)meta.Attribute("name");
                if (string.IsNullOrEmpty(name))
                {
                    continue;
                }

                if (!name.StartsWith(TeXShiftMetaKeys.SourceChunkPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string indexText = name.Substring(TeXShiftMetaKeys.SourceChunkPrefix.Length);
                if (!int.TryParse(indexText, out int index))
                {
                    continue;
                }

                string content = (string)meta.Attribute("content") ?? string.Empty;
                chunks[index] = content;
            }

            if (chunks.Count == 0)
            {
                return string.Empty;
            }

            int maxIndex = chunks.Keys.Max();
            for (int i = 0; i <= maxIndex; i++)
            {
                if (!chunks.ContainsKey(i))
                {
                    error = "TeXShift meta source chunks are incomplete.";
                    return null;
                }
            }

            var builder = new StringBuilder();
            for (int i = 0; i <= maxIndex; i++)
            {
                builder.Append(chunks[i]);
            }

            return builder.ToString();
        }

        private static string DecodePlainV1(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            var builder = new StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (ch == '\\' && i + 1 < text.Length)
                {
                    char next = text[i + 1];
                    if (next == 'n')
                    {
                        builder.Append('\n');
                        i++;
                        continue;
                    }
                    if (next == 'r')
                    {
                        builder.Append('\r');
                        i++;
                        continue;
                    }
                    if (next == '\\')
                    {
                        builder.Append('\\');
                        i++;
                        continue;
                    }
                }

                builder.Append(ch);
            }

            return builder.ToString();
        }
    }
}
