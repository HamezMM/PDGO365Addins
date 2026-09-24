using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace BracketFields
{
    /// <summary>Result of a fill: how many fields got a value and how many occurrences changed.</summary>
    internal sealed class FillResult
    {
        public int FieldsFilled { get; set; }
        public int FieldsSkipped { get; set; }
        public int OccurrencesReplaced { get; set; }
    }

    /// <summary>
    /// Replaces every occurrence of each filled <see cref="BracketField"/> token — brackets
    /// included — with its value, across all document stories, as a single Undo step.
    /// </summary>
    internal static class BracketFieldReplacer
    {
        private const string UndoName = "Fill Bracket Fields";

        public static FillResult Replace(Word.Application app, Word.Document doc, IReadOnlyList<BracketField> fields)
        {
            if (doc.ProtectionType != Word.WdProtectionType.wdNoProtection)
            {
                throw new InvalidOperationException(
                    "This document is protected against editing. Unprotect it (Review ▸ Restrict Editing) and try again.");
            }

            var result = new FillResult();
            List<BracketField> filled = fields.Where(f => f.HasValue).ToList();
            result.FieldsFilled = filled.Count;
            result.FieldsSkipped = fields.Count - filled.Count;
            if (filled.Count == 0) return result;

            // Materialize once: the story list doesn't change as text is replaced.
            List<Word.Range> stories = DocumentStories.Enumerate(doc).ToList();

            Word.UndoRecord undo = app.UndoRecord;
            bool screenUpdating = app.ScreenUpdating;
            undo.StartCustomRecord(UndoName);
            app.ScreenUpdating = false;
            try
            {
                foreach (BracketField field in filled)
                {
                    string value = NormalizeValue(field.Value);
                    foreach (Word.Range story in stories)
                    {
                        result.OccurrencesReplaced += ReplaceInStory(story, field.Token, value);
                    }
                }
            }
            finally
            {
                app.ScreenUpdating = screenUpdating;
                undo.EndCustomRecord();
            }

            return result;
        }

        /// <summary>
        /// Find/replace one token inside one story. Uses a Find loop and sets
        /// <c>Range.Text</c> instead of <c>Find.Execute(Replace: wdReplaceAll)</c> because
        /// Word caps replacement text at 255 characters and treats <c>^</c> specially in it.
        /// The inserted text takes the formatting of the token it replaces.
        /// </summary>
        private static int ReplaceInStory(Word.Range story, string token, string value)
        {
            int count = 0;
            Word.Range range = story.Duplicate;
            Word.Find find = range.Find;

            find.ClearFormatting();
            find.Text = EscapeFindText(token);
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            find.Format = false;
            find.MatchCase = true;
            find.MatchWholeWord = false;
            find.MatchWildcards = false;
            find.MatchSoundsLike = false;
            find.MatchAllWordForms = false;

            while (find.Execute())
            {
                range.Text = value;
                count++;
                // Continue searching after the inserted text, so a value that itself
                // contains the token can't loop forever.
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }

            return count;
        }

        /// <summary>In non-wildcard Find text only <c>^</c> is special; <c>^^</c> matches a literal caret.</summary>
        private static string EscapeFindText(string text) => text.Replace("^", "^^");

        /// <summary>Word paragraphs end in a bare CR; a CRLF would insert an extra empty paragraph.</summary>
        private static string NormalizeValue(string value) => value.Replace("\r\n", "\r").Replace("\n", "\r");
    }
}
