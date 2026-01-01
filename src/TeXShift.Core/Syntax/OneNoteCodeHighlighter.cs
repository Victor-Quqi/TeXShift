using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TeXShift.Core.Configuration;
using TeXShift.Core.Utils;
using TextMateSharp.Grammars;
using TextMateSharp.Registry;
using TextMateSharp.Themes;

namespace TeXShift.Core.Syntax
{
    /// <summary>
    /// Syntax highlighter using TextMateSharp that outputs OneNote-compatible inline styles.
    /// Uses VS Code grammars for accurate syntax highlighting across 40+ languages.
    /// </summary>
    internal class OneNoteCodeHighlighter : ISyntaxHighlighter
    {
        private readonly OneNoteStyleConfig.CodeBlockConfig _config;
        private readonly Registry _registry;
        private readonly RegistryOptions _registryOptions;
        private readonly Theme _theme;

        // Cache loaded grammars for performance
        private readonly Dictionary<string, IGrammar> _grammarCache;

        // Language aliases for unsupported languages (fallback to similar syntax)
        private static readonly Dictionary<string, string> LanguageAliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "kotlin", "java" },
            { "kt", "java" },
            { "scala", "java" },
            { "tsx", "typescriptbasics" },
            { "jsx", "javascript" },
            { "sh", "shellscript" },
            { "bash", "shellscript" },
            { "zsh", "shellscript" },
            { "ps1", "powershell" },
            { "yml", "yaml" },
            { "md", "markdownbasics" },
            { "cs", "csharp" },
            { "fs", "fsharp" },
            { "ts", "typescriptbasics" },
            { "js", "javascript" },
            { "py", "python" },
            { "rb", "ruby" },
            { "rs", "rust" },
        };

        public OneNoteCodeHighlighter(OneNoteStyleConfig.CodeBlockConfig config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
            _grammarCache = new Dictionary<string, IGrammar>(StringComparer.OrdinalIgnoreCase);

            // Use built-in Dark Plus theme (similar to GitHub Dark)
            _registryOptions = new RegistryOptions(ThemeName.DarkPlus);
            _registry = new Registry(_registryOptions);
            _theme = _registry.GetTheme();
        }

        public IList<string> HighlightBlock(IList<string> lines, string language)
        {
            if (lines == null || lines.Count == 0)
            {
                return new List<string>();
            }

            // No highlighting: just escape HTML
            if (!_config.EnableSyntaxHighlight || string.IsNullOrEmpty(language))
            {
                return lines.Select(l => HtmlEscaper.Escape(l ?? "")).ToList();
            }

            var grammar = GetOrLoadGrammar(language);
            if (grammar == null)
            {
                return lines.Select(l => HtmlEscaper.Escape(l ?? "")).ToList();
            }

            var result = new List<string>(lines.Count);
            IStateStack ruleStack = null;  // Cross-line state for multi-line comments/strings

            foreach (var line in lines)
            {
                if (string.IsNullOrEmpty(line))
                {
                    result.Add(string.Empty);
                    continue;
                }

                try
                {
                    var tokenResult = grammar.TokenizeLine(line, ruleStack, TimeSpan.MaxValue);
                    ruleStack = tokenResult.RuleStack;  // Preserve state for next line
                    result.Add(TokenizeResultToHtml(line, tokenResult));
                }
                catch
                {
                    result.Add(HtmlEscaper.Escape(line));
                }
            }

            return result;
        }

        private string TokenizeResultToHtml(string line, ITokenizeLineResult tokenResult)
        {
            var output = new StringBuilder();

            foreach (var token in tokenResult.Tokens)
            {
                int start = System.Math.Min(token.StartIndex, line.Length);
                int end = System.Math.Min(token.EndIndex, line.Length);

                if (start >= end)
                {
                    continue;  // Skip empty tokens
                }

                string text = line.Substring(start, end - start);
                string color = GetTokenColor(token.Scopes);

                if (!string.IsNullOrEmpty(color) && !color.Equals(_config.DefaultTextColor, StringComparison.OrdinalIgnoreCase))
                {
                    output.Append($"<span style='color:{color}'>");
                    output.Append(HtmlEscaper.Escape(text));
                    output.Append("</span>");
                }
                else
                {
                    output.Append(HtmlEscaper.Escape(text));
                }
            }

            return output.ToString();
        }

        private string GetTokenColor(IReadOnlyList<string> scopes)
        {
            // Match scopes in reverse order (most specific first)
            for (int i = scopes.Count - 1; i >= 0; i--)
            {
                var rules = _theme.Match(new string[] { scopes[i] });
                foreach (var rule in rules)
                {
                    if (rule.foreground != 0)
                    {
                        return _theme.GetColor(rule.foreground);
                    }
                }
            }
            return null;
        }

        public bool IsLanguageSupported(string language)
        {
            return GetOrLoadGrammar(language) != null;
        }

        private IGrammar GetOrLoadGrammar(string language)
        {
            if (string.IsNullOrEmpty(language))
            {
                return null;
            }

            if (_grammarCache.TryGetValue(language, out var cached))
            {
                return cached;
            }

            // Try alias first if direct language not found
            string effectiveLanguage = language;
            if (LanguageAliases.TryGetValue(language, out var alias))
            {
                effectiveLanguage = alias;
            }

            // RegistryOptions handles language ID -> scope mapping internally
            string scopeName = _registryOptions.GetScopeByLanguageId(effectiveLanguage);

            if (string.IsNullOrEmpty(scopeName))
            {
                _grammarCache[language] = null;
                return null;
            }

            var grammar = _registry.LoadGrammar(scopeName);
            _grammarCache[language] = grammar;
            return grammar;
        }
    }
}
