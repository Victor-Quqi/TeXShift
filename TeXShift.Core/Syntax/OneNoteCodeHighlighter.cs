using ColorCode;
using ColorCode.Compilation;
using ColorCode.Compilation.Languages;
using ColorCode.Parsing;
using System;
using System.Collections.Generic;
using System.Text;
using TeXShift.Core.Configuration;
using TeXShift.Core.Utils;

namespace TeXShift.Core.Syntax
{
    /// <summary>
    /// Syntax highlighter that outputs OneNote-compatible inline styles.
    /// Inherits from CodeColorizerBase to use ColorCode's parsing infrastructure.
    /// </summary>
    internal class OneNoteCodeHighlighter : CodeColorizerBase, ISyntaxHighlighter
    {
        private readonly OneNoteStyleConfig.CodeBlockConfig _config;
        private readonly Dictionary<string, string> _scopeColors;
        private StringBuilder _outputBuffer;

        public OneNoteCodeHighlighter(OneNoteStyleConfig.CodeBlockConfig config)
            : base(null, null)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
            _scopeColors = CreateGitHubDarkTheme();
        }

        public string HighlightLine(string line, string language)
        {
            if (string.IsNullOrEmpty(line))
            {
                return string.Empty;
            }

            if (!_config.EnableSyntaxHighlight || string.IsNullOrEmpty(language))
            {
                return HtmlEscaper.Escape(line);
            }

            var lang = FindLanguage(language);
            if (lang == null)
            {
                return HtmlEscaper.Escape(line);
            }

            try
            {
                _outputBuffer = new StringBuilder();
                languageParser.Parse(line, lang, (parsedCode, scopes) => Write(parsedCode, scopes));
                return _outputBuffer.ToString();
            }
            catch
            {
                return HtmlEscaper.Escape(line);
            }
        }

        protected override void Write(string parsedSourceCode, IList<Scope> scopes)
        {
            if (scopes == null || scopes.Count == 0)
            {
                _outputBuffer.Append(HtmlEscaper.Escape(parsedSourceCode));
                return;
            }

            var styleInsertions = new List<TextInsertion>();
            foreach (var scope in scopes)
            {
                GetStyleInsertions(parsedSourceCode, scope, styleInsertions);
            }

            styleInsertions.Sort((a, b) => a.Index.CompareTo(b.Index));

            int offset = 0;
            foreach (var insertion in styleInsertions)
            {
                var text = parsedSourceCode.Substring(offset, insertion.Index - offset);
                _outputBuffer.Append(HtmlEscaper.Escape(text));
                _outputBuffer.Append(insertion.Text);
                offset = insertion.Index;
            }

            // Write remaining text
            if (offset < parsedSourceCode.Length)
            {
                _outputBuffer.Append(HtmlEscaper.Escape(parsedSourceCode.Substring(offset)));
            }
        }

        private void GetStyleInsertions(string parsedSourceCode, Scope scope, List<TextInsertion> styleInsertions)
        {
            var color = GetColorForScope(scope.Name);

            // Opening tag
            styleInsertions.Add(new TextInsertion
            {
                Index = scope.Index,
                Text = color != _config.DefaultTextColor
                    ? $"<span style='color:{color}'>"
                    : ""
            });

            // Process children recursively
            if (scope.Children != null)
            {
                foreach (var child in scope.Children)
                {
                    GetStyleInsertions(parsedSourceCode, child, styleInsertions);
                }
            }

            // Closing tag
            styleInsertions.Add(new TextInsertion
            {
                Index = scope.Index + scope.Length,
                Text = color != _config.DefaultTextColor ? "</span>" : ""
            });
        }

        private string GetColorForScope(string scopeName)
        {
            if (_scopeColors.TryGetValue(scopeName, out var color))
            {
                return color;
            }
            return _config.DefaultTextColor;
        }

        public bool IsLanguageSupported(string language)
        {
            return FindLanguage(language) != null;
        }

        private ILanguage FindLanguage(string language)
        {
            if (string.IsNullOrEmpty(language))
            {
                return null;
            }

            // Try direct lookup first
            var lang = Languages.FindById(language);
            if (lang != null)
            {
                return lang;
            }

            // Try common aliases
            var lowerLang = language.ToLowerInvariant();
            switch (lowerLang)
            {
                case "js":
                case "javascript":
                    return Languages.JavaScript;
                case "ts":
                case "typescript":
                    return Languages.Typescript;
                case "py":
                case "python":
                    return Languages.Python;
                case "cs":
                case "c#":
                case "csharp":
                    return Languages.CSharp;
                case "cpp":
                case "c++":
                case "c":
                    return Languages.Cpp;
                case "sh":
                case "bash":
                case "shell":
                case "powershell":
                    return Languages.PowerShell;
                case "md":
                case "markdown":
                    return Languages.Markdown;
                case "xml":
                    return Languages.Xml;
                case "html":
                case "htm":
                    return Languages.Html;
                case "css":
                    return Languages.Css;
                case "sql":
                    return Languages.Sql;
                case "java":
                    return Languages.Java;
                case "php":
                    return Languages.Php;
                case "vb":
                case "vbnet":
                    return Languages.VbDotNet;
                case "fs":
                case "fsharp":
                    return Languages.FSharp;
                case "hs":
                case "haskell":
                    return Languages.Haskell;
                case "fortran":
                    return Languages.Fortran;
                case "matlab":
                    return Languages.MATLAB;
                // JSON uses JavaScript highlighting (JSON is a subset of JS object literals)
                case "json":
                    return Languages.JavaScript;
                // Languages without ColorCode support - return null for plain text display
                // Go, Rust, Kotlin, Swift, Ruby, Lua, R, Scala, etc.
                default:
                    return null;
            }
        }

        private Dictionary<string, string> CreateGitHubDarkTheme()
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // Keywords
                { "Keyword", "#FF7B72" },

                // Strings
                { "String", "#A5D6FF" },

                // Comments
                { "Comment", "#8B949E" },

                // Numbers
                { "Number", "#79C0FF" },

                // Built-in functions (Python: print, len, abs, etc.)
                { "Intrinsic", "#FFA657" },

                // Types and classes
                { "Type", "#FFA657" },
                { "TypeVariable", "#FFA657" },
                { "NameSpace", "#FFA657" },
                { "ClassName", "#FFA657" },

                // HTML/XML
                { "HtmlComment", "#8B949E" },
                { "HtmlTagDelimiter", "#7EE787" },
                { "HtmlElementName", "#7EE787" },
                { "HtmlAttributeName", "#79C0FF" },
                { "HtmlAttributeValue", "#A5D6FF" },
                { "HtmlOperator", "#C9D1D9" },
                { "HtmlEntity", "#79C0FF" },
                { "XmlDelimiter", "#7EE787" },
                { "XmlName", "#7EE787" },
                { "XmlAttribute", "#79C0FF" },
                { "XmlAttributeQuotes", "#A5D6FF" },
                { "XmlAttributeValue", "#A5D6FF" },
                { "XmlCDataSection", "#C9D1D9" },
                { "XmlComment", "#8B949E" },

                // CSS
                { "CssPropertyName", "#79C0FF" },
                { "CssPropertyValue", "#A5D6FF" },
                { "CssSelector", "#7EE787" },

                // SQL
                { "SqlSystemFunction", "#D2A8FF" },

                // Preprocessor
                { "PreprocessorKeyword", "#FF7B72" },

                // Others
                { "Operator", "#FF7B72" },
                { "Punctuation", "#C9D1D9" },
                { "PlainText", "#C9D1D9" }
            };
        }

        private class TextInsertion
        {
            public int Index { get; set; }
            public string Text { get; set; }
        }
    }
}
