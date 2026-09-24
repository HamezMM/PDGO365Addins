using System.Collections.Generic;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace BracketFields
{
    /// <summary>
    /// Finds every unique <c>[bracketed]</c> field in a document, in order of first appearance.
    /// </summary>
    internal static class BracketFieldScanner
    {
        /// <summary>
        /// Word's Find text is capped at 255 characters, so a longer token could never be
        /// replaced. Such long bracketed runs are almost certainly prose, not fields.
        /// </summary>
        public const int MaxTokenLength = 255;

        private const int ContextChars = 40;

        // "[" + one or more chars that are not brackets or a paragraph/line/cell break + "]".
        // Word range text uses \r (paragraph), \v (manual line break), \a (table cell end),
        // \f (page/section break). Excluding them keeps a field on one line and stops a stray
        // "[" from swallowing everything up to the next "]" pages later.
        private static readonly Regex FieldPattern =
            new Regex(@"\[[^\[\]\r\n\v\a\f]+\]", RegexOptions.Compiled);

        public static List<BracketField> Scan(Word.Document doc)
        {
            var fields = new List<BracketField>();
            var byToken = new Dictionary<string, BracketField>();

            foreach (Word.Range story in DocumentStories.Enumerate(doc))
            {
                string text = story.Text;
                if (string.IsNullOrEmpty(text)) continue;

                foreach (Match match in FieldPattern.Matches(text))
                {
                    string token = match.Value;
                    if (token.Length > MaxTokenLength) continue;
                    if (string.IsNullOrWhiteSpace(token.Substring(1, token.Length - 2))) continue;

                    if (!byToken.TryGetValue(token, out BracketField field))
                    {
                        field = new BracketField(token, ContextAround(text, match));
                        byToken.Add(token, field);
                        fields.Add(field);
                    }

                    field.Occurrences++;
                }
            }

            return fields;
        }

        private static string ContextAround(string text, Match match)
        {
            int start = System.Math.Max(0, match.Index - ContextChars);
            int end = System.Math.Min(text.Length, match.Index + match.Length + ContextChars);

            string snippet = text.Substring(start, end - start);
            snippet = Regex.Replace(snippet, @"[\r\n\v\a\f\t]+", " ").Trim();

            return (start > 0 ? "…" : "") + snippet + (end < text.Length ? "…" : "");
        }
    }
}
