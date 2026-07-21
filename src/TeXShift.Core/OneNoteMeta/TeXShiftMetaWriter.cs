using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using TeXShift.Core.Utils;

namespace TeXShift.Core.OneNoteMeta
{
    internal static class TeXShiftMetaWriter
    {
        public static void WriteSourceMeta(XElement outline, string sourceMarkdown, string mode = TeXShiftMetaKeys.ModeRender)
        {
            if (outline == null)
            {
                throw new ArgumentNullException(nameof(outline));
            }

            string signature = ComputeSignature(outline);
            WriteSourceMeta(outline, sourceMarkdown, mode, signature);
        }

        public static void WriteSourceMeta(XElement outline, string sourceMarkdown, string mode, string signature)
        {
            if (outline == null)
            {
                throw new ArgumentNullException(nameof(outline));
            }

            var ns = outline.Name.Namespace;
            RemoveExistingTeXShiftMeta(outline, ns);

            var encodedSource = EncodePlainV1(sourceMarkdown ?? string.Empty);
            var chunks = Chunk(encodedSource, TeXShiftMetaKeys.MaxChunkLength);

            var metaElements = new List<XElement>
            {
                CreateMeta(ns, TeXShiftMetaKeys.Schema, TeXShiftMetaKeys.SchemaVersion),
                CreateMeta(ns, TeXShiftMetaKeys.Mode, string.IsNullOrWhiteSpace(mode) ? TeXShiftMetaKeys.ModeRender : mode),
                CreateMeta(ns, TeXShiftMetaKeys.SourceEncoding, TeXShiftMetaKeys.EncodingPlainV1)
            };

            for (int i = 0; i < chunks.Count; i++)
            {
                metaElements.Add(CreateMeta(ns, $"{TeXShiftMetaKeys.SourceChunkPrefix}{i}", chunks[i]));
            }

            metaElements.Add(CreateMeta(ns, TeXShiftMetaKeys.SigVersion, TeXShiftMetaKeys.SigVersionValue));
            metaElements.Add(CreateMeta(ns, TeXShiftMetaKeys.Sig, signature ?? string.Empty));

            InsertMetaElements(outline, metaElements);
        }

        internal static string ComputeSignature(XElement outline)
        {
            if (outline == null)
            {
                return string.Empty;
            }

            return ComputeSemanticSignature(outline);
        }

        private static readonly Regex MathMlConditionalCommentRegex = new Regex(
            "<!--\\[if\\s+mathML\\]>(.*?)<!\\[endif\\]-->",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex HtmlTagRegex = new Regex(
            "<[^>]+>",
            RegexOptions.Compiled);

        private static readonly Regex BrTagRegex = new Regex(
            "<br\\b[^>]*>",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // OneNote may insert visible stray quote characters around embedded MathML objects.
        // Those quotes are not part of the user's content, but we must NOT remove legitimate quotes
        // inside code blocks (syntax highlighting can wrap quotes in standalone <span> nodes).
        // Only strip quote-only spans that are immediately adjacent to a MathML conditional comment.
        private static readonly Regex StrayQuoteSpanBeforeMathMlRegex = new Regex(
            "<span\\b[^>]*>\\s*(?:&quot;|\")\\s*</span>\\s*(?=<!--\\[if\\s+mathML\\]>)",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex StrayQuoteSpanAfterMathMlRegex = new Regex(
            "<!\\[endif\\]-->\\s*<span\\b[^>]*>\\s*(?:&quot;|\")\\s*</span>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex CollapseWhitespaceRegex = new Regex(
            "\\s+",
            RegexOptions.Compiled);

        private static readonly Regex MathMlNbspRegex = new Regex(
            "&nbsp;",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static string ComputeSemanticSignature(XElement outline)
        {
            if (outline == null)
            {
                return string.Empty;
            }

            // Goal: make signature stable across OneNote's XML reformatting (e.g., span/style rewrites,
            // OCRData injection for images, base64 line wrapping, selection/cursor empty <one:T> nodes).
            //
            // We hash a semantic token stream:
            // - Text nodes: merge consecutive <one:T> nodes within the same <one:OE>, then strip HTML tags and
            //   normalize whitespace, while preserving mathML conditional comments via an inner hash marker.
            // - Images: use a stable semantic marker based on attributes (OneNote may re-encode image bytes).
            var tokens = new List<string>(capacity: 256);

            // OneNote can split a single paragraph into multiple <one:T> nodes (and even inject empty
            // <one:T selected="all"><![CDATA[]]></one:T> nodes). It can also split/merge <one:OE> blocks
            // while keeping the visible text identical (caret placement, rich object boundaries, etc.).
            //
            // To avoid false mismatches, we merge all consecutive text runs into larger segments and only
            // break on non-text content (e.g., images). We still insert a soft separator between different
            // <one:OE> owners to avoid concatenating words across paragraphs.
            var mergedRichText = new StringBuilder(capacity: 512);
            XElement currentTextOwner = null; // current <one:OE>

            Action flushMergedText = () =>
            {
                if (mergedRichText.Length == 0)
                {
                    return;
                }

                string token = NormalizeRichTextToToken(mergedRichText.ToString());
                mergedRichText.Clear();
                currentTextOwner = null;

                if (!string.IsNullOrWhiteSpace(token))
                {
                    tokens.Add("T:" + token);
                }
            };

            foreach (var element in outline.Descendants())
            {
                if (element == null)
                {
                    continue;
                }

                var localName = element.Name.LocalName;
                if (ShouldSkipElementV2(localName))
                {
                    continue;
                }

                if (string.Equals(localName, "T", StringComparison.OrdinalIgnoreCase))
                {
                    var owner = element.Parent;
                    if (owner != null && !string.Equals(owner.Name.LocalName, "OE", StringComparison.OrdinalIgnoreCase))
                    {
                        owner = null;
                    }

                    if (!object.ReferenceEquals(owner, currentTextOwner))
                    {
                        // Soft separator between paragraphs/blocks. This stays stable after NormalizeRichTextToToken.
                        if (mergedRichText.Length > 0)
                        {
                            mergedRichText.Append('\n');
                        }
                        currentTextOwner = owner;
                    }

                    mergedRichText.Append(element.Value);
                    continue;
                }

                if (string.Equals(localName, "Image", StringComparison.OrdinalIgnoreCase))
                {
                    flushMergedText();

                    // OneNote may rewrite embedded image bytes (e.g., re-encode PNGs), which would make a raw
                    // base64 hash unstable. Use a semantic marker based on stable attributes instead.
                    tokens.Add("IMG:" + GetImageToken(element));
                }
            }

            flushMergedText();

            string payload = string.Join("\n", tokens);
            return Sha256Hex(payload);
        }

        private static bool ShouldSkipElementV2(string localName)
        {
            if (string.IsNullOrWhiteSpace(localName))
            {
                return true;
            }

            // Non-content / layout / OneNote-generated elements.
            return string.Equals(localName, "Meta", StringComparison.OrdinalIgnoreCase)
                || string.Equals(localName, "Position", StringComparison.OrdinalIgnoreCase)
                || string.Equals(localName, "Size", StringComparison.OrdinalIgnoreCase)
                || string.Equals(localName, "Indents", StringComparison.OrdinalIgnoreCase)
                || string.Equals(localName, "Indent", StringComparison.OrdinalIgnoreCase)
                || string.Equals(localName, "Data", StringComparison.OrdinalIgnoreCase)
                || string.Equals(localName, "OCRData", StringComparison.OrdinalIgnoreCase)
                || string.Equals(localName, "OCRText", StringComparison.OrdinalIgnoreCase)
                || string.Equals(localName, "OCRToken", StringComparison.OrdinalIgnoreCase);
        }

        private static string GetImageToken(XElement imageElement)
        {
            if (imageElement == null)
            {
                return string.Empty;
            }

            string alt = ((string)imageElement.Attribute("alt") ?? string.Empty).Trim();

            // OneNote may drop/rename Image attributes (e.g., 'format="png"') after UpdatePageContent.
            // For signature stability, only keep semantic hints that tend to survive (like our own 'alt' markers).
            if (string.IsNullOrEmpty(alt))
            {
                return "image";
            }

            return "image:" + alt;
        }

        private static string NormalizeRichTextToToken(string html)
        {
            if (string.IsNullOrEmpty(html))
            {
                return string.Empty;
            }

            // Preserve math objects by hashing the mathML payload into a stable marker, but avoid
            // signature churn due to OneNote rewriting MathML (e.g., different codepoints for the same symbol).
            string withMathMarkers = MathMlConditionalCommentRegex.Replace(html, match =>
            {
                string inner = match.Groups[1].Value ?? string.Empty;
                string innerHash = HashMathMl(inner);
                return $"__MATHML_{innerHash}__";
            });

            // OneNote may rewrite literal newlines into <br /> tags and vice versa. Treat <br> as a soft line break
            // so the signature remains stable across this rewrite.
            withMathMarkers = BrTagRegex.Replace(withMathMarkers, "\n");

            // OneNote may insert stray quote spans around embedded math objects (visible in the editor as '"').
            // Remove only quote-only spans that are adjacent to MathML conditional comments to avoid
            // corrupting signatures for code blocks (where quotes are meaningful content).
            withMathMarkers = StrayQuoteSpanBeforeMathMlRegex.Replace(withMathMarkers, string.Empty);
            withMathMarkers = StrayQuoteSpanAfterMathMlRegex.Replace(withMathMarkers, "<![endif]-->");

            // Strip HTML tags, decode entities, and normalize whitespace.
            string text = HtmlTagRegex.Replace(withMathMarkers, string.Empty);
            text = OneNoteHtmlEntityDecoder.Decode(text);
            text = text.Replace("\u200B", string.Empty); // ZWSP is an internal OneNote/math artifact.
            text = CollapseWhitespaceRegex.Replace(text, " ").Trim();
            return text;
        }

        private static string HashMathMl(string mathMlXml)
        {
            if (string.IsNullOrWhiteSpace(mathMlXml))
            {
                return Sha256Hex(string.Empty);
            }

            try
            {
                // OneNote sometimes injects HTML-only entities (e.g., &nbsp;) into the MathML payload,
                // which is not valid XML. Replace them with numeric equivalents so XElement.Parse works.
                string sanitized = MathMlNbspRegex.Replace(mathMlXml, "&#xA0;");

                // Canonicalize MathML to reduce false mismatches caused by OneNote rewriting
                // equivalent content (e.g., π vs 𝜋, entity forms vs literal characters, extra mrow wrappers).
                var root = XElement.Parse(sanitized, LoadOptions.None);

                // OneNote may rewrite some fences into <mfenced open="|" close="|"> forms (or vice versa).
                // XElement.Value does not include mfenced open/close attributes, so we render a minimal
                // text representation that includes mfenced fences for signature stability.
                string rendered = RenderMathMlTextForSignature(root).Normalize(NormalizationForm.FormKC);
                rendered = NormalizeMathMlRenderedText(rendered);
                rendered = rendered.Replace("\u200B", string.Empty); // ZWSP is an internal OneNote/math artifact.
                rendered = CollapseWhitespaceRegex.Replace(rendered, string.Empty);
                return Sha256Hex(rendered);
            }
            catch
            {
                // Best-effort fallback: still deterministic, but less stable across OneNote rewrites.
                string normalized = CollapseWhitespaceRegex.Replace(mathMlXml, string.Empty);
                return Sha256Hex(normalized);
            }
        }

        private static string RenderMathMlTextForSignature(XElement root)
        {
            if (root == null)
            {
                return string.Empty;
            }

            var sb = new StringBuilder(capacity: 64);
            AppendMathMlTextForSignature(root, sb);
            return sb.ToString();
        }

        private static void AppendMathMlTextForSignature(XElement element, StringBuilder sb)
        {
            if (element == null)
            {
                return;
            }

            if (sb == null)
            {
                throw new ArgumentNullException(nameof(sb));
            }

            // mfenced fences are often encoded as attributes (open/close), which are not part of XElement.Value.
            // To keep signatures stable across OneNote rewrites (mfenced <-> explicit mo tokens), include them.
            if (string.Equals(element.Name.LocalName, "mfenced", StringComparison.OrdinalIgnoreCase))
            {
                string open = (string)element.Attribute("open");
                string close = (string)element.Attribute("close");

                // MathML defaults apply only when the attribute is absent.
                // An explicitly empty value represents a one-sided fence.
                if (open == null)
                {
                    open = "(";
                }
                if (close == null)
                {
                    close = ")";
                }

                sb.Append(open);
                foreach (var node in element.Nodes())
                {
                    if (node is XText text)
                    {
                        sb.Append(text.Value);
                    }
                    else if (node is XElement child)
                    {
                        AppendMathMlTextForSignature(child, sb);
                    }
                }
                sb.Append(close);
                return;
            }

            foreach (var node in element.Nodes())
            {
                if (node is XText text)
                {
                    sb.Append(text.Value);
                }
                else if (node is XElement child)
                {
                    AppendMathMlTextForSignature(child, sb);
                }
            }
        }

        private static string NormalizeMathMlRenderedText(string rendered)
        {
            if (string.IsNullOrEmpty(rendered))
            {
                return string.Empty;
            }

            // OneNote rewrites some accent glyphs in embedded MathML when serializing/parsing page content.
            // For signature stability, normalize equivalent forms to a single representation.
            // - U+02D9 (DOT ABOVE) is sometimes rewritten to '.' inside <mover accent="true">.
            // - U+0307 (COMBINING DOT ABOVE) is another common representation for dot accents.
            rendered = rendered.Replace('\u02D9', '.');
            rendered = rendered.Replace('\u0307', '.');
            return rendered;
        }

        private static string Sha256Hex(string input)
        {
            if (input == null)
            {
                input = string.Empty;
            }

            using (var sha = SHA256.Create())
            {
                var bytes = Encoding.UTF8.GetBytes(input);
                var hash = sha.ComputeHash(bytes);
                var builder = new StringBuilder(hash.Length * 2);
                foreach (var b in hash)
                {
                    builder.Append(b.ToString("x2"));
                }
                return builder.ToString();
            }
        }

        private static XElement CreateMeta(XNamespace ns, string name, string content)
        {
            return new XElement(ns + "Meta",
                new XAttribute("name", name ?? string.Empty),
                new XAttribute("content", content ?? string.Empty));
        }

        private static void RemoveExistingTeXShiftMeta(XElement outline, XNamespace ns)
        {
            var metas = outline.Elements(ns + "Meta")
                .Where(meta => IsTeXShiftMeta(meta))
                .ToList();

            foreach (var meta in metas)
            {
                meta.Remove();
            }
        }

        private static bool IsTeXShiftMeta(XElement meta)
        {
            var name = (string)meta.Attribute("name");
            return !string.IsNullOrEmpty(name)
                && name.StartsWith(TeXShiftMetaKeys.Prefix, StringComparison.OrdinalIgnoreCase);
        }

        private static void InsertMetaElements(XElement outline, List<XElement> metaElements)
        {
            if (metaElements == null || metaElements.Count == 0)
            {
                return;
            }

            // Outline extends PageObject: Position?, Size?, Meta*, Indents?, OEChildren+ (sequence matters).
            // Insert after existing Meta (or after Size/Position if no Meta), and before Indents/OEChildren.
            XElement insertAfter = outline.Elements()
                .Where(e =>
                    string.Equals(e.Name.LocalName, "Position", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(e.Name.LocalName, "Size", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(e.Name.LocalName, "Meta", StringComparison.OrdinalIgnoreCase))
                .LastOrDefault();

            if (insertAfter != null)
            {
                foreach (var meta in metaElements)
                {
                    insertAfter.AddAfterSelf(meta);
                    insertAfter = meta;
                }
                return;
            }

            outline.AddFirst(metaElements);
        }

        private static string EncodePlainV1(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            var builder = new StringBuilder(text.Length + 16);
            foreach (var ch in text)
            {
                switch (ch)
                {
                    case '\\':
                        builder.Append("\\\\");
                        break;
                    case '\n':
                        builder.Append("\\n");
                        break;
                    case '\r':
                        builder.Append("\\r");
                        break;
                    default:
                        builder.Append(ch);
                        break;
                }
            }

            return builder.ToString();
        }

        private static List<string> Chunk(string text, int maxChunkLength)
        {
            var chunks = new List<string>();
            if (string.IsNullOrEmpty(text))
            {
                chunks.Add(string.Empty);
                return chunks;
            }

            int index = 0;
            while (index < text.Length)
            {
                int length = System.Math.Min(maxChunkLength, text.Length - index);
                chunks.Add(text.Substring(index, length));
                index += length;
            }

            return chunks;
        }
    }
}
