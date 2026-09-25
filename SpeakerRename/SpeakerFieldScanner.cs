using System.Collections.Generic;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace SpeakerRename
{
    /// <summary>
    /// Finds every unique <c>Speaker N</c> label in a document, ordered by speaker number.
    /// </summary>
    internal static class SpeakerFieldScanner
    {
        /// <summary>
        /// Word's Find text is capped at 255 characters, so a longer token could never be
        /// replaced. A "Speaker N" label is never remotely this long in practice; the guard
        /// just keeps scanner and replacer consistent about what can actually be fixed.
        /// </summary>
        public const int MaxTokenLength = 255;

        private const int ContextChars = 40;

        // "Speaker" + one or more spaces/tabs + one or more digits, on word boundaries at
        // both ends. \b after the digits keeps "Speaker 10" from matching inside "Speaker 10a"
        // or "Speaker 100"; requiring \s+ (not just " ") keeps "Speaker1" (no space) from
        // matching, since that isn't the label format transcripts use. Matching is
        // case-sensitive, same policy as BracketFields: "Speaker 1" and "speaker 1" would be
        // treated as different text if the latter ever appeared.
        private static readonly Regex SpeakerPattern =
            new Regex(@"\bSpeaker[ \t]+(\d+)\b", RegexOptions.Compiled);

        public static List<SpeakerField> Scan(Word.Document doc)
        {
            var fields = new List<SpeakerField>();
            var byNumber = new Dictionary<int, SpeakerField>();

            foreach (Word.Range story in DocumentStories.Enumerate(doc))
            {
                string text = story.Text;
                if (string.IsNullOrEmpty(text)) continue;

                foreach (Match match in SpeakerPattern.Matches(text))
                {
                    string token = match.Value;
                    if (token.Length > MaxTokenLength) continue;

                    int number = int.Parse(match.Groups[1].Value);

                    if (!byNumber.TryGetValue(number, out SpeakerField field))
                    {
                        field = new SpeakerField(number, token, ContextAround(text, match));
                        byNumber.Add(number, field);
                        fields.Add(field);
                    }
                    else if (!field.Tokens.Contains(token))
                    {
                        field.Tokens.Add(token);
                    }

                    field.Occurrences++;
                }
            }

            // Order by speaker number rather than first appearance: users think in terms of
            // "Speaker 1, 2, 3…", and a transcript's speakers usually first appear roughly in
            // that order anyway, so this is rarely a surprise and always predictable.
            fields.Sort((a, b) => a.Number.CompareTo(b.Number));
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
