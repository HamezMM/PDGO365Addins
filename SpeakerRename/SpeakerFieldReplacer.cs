using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace SpeakerRename
{
    /// <summary>Result of a rename: how many speakers got a name and how many occurrences changed.</summary>
    internal sealed class RenameResult
    {
        public int SpeakersRenamed { get; set; }
        public int SpeakersSkipped { get; set; }
        public int OccurrencesReplaced { get; set; }
    }

    /// <summary>
    /// Replaces every occurrence of each named <see cref="SpeakerField"/>'s label — every
    /// whitespace variant in <see cref="SpeakerField.Tokens"/> — with its name, across all
    /// document stories, as a single Undo step.
    /// </summary>
    internal static class SpeakerFieldReplacer
    {
        private const string UndoName = "Rename Speakers";

        public static RenameResult Replace(Word.Application app, Word.Document doc, IReadOnlyList<SpeakerField> fields)
        {
            if (doc.ProtectionType != Word.WdProtectionType.wdNoProtection)
            {
                throw new InvalidOperationException(
                    "This document is protected against editing. Unprotect it (Review ▸ Restrict Editing) and try again.");
            }

            var result = new RenameResult();
            List<SpeakerField> renamed = fields.Where(f => f.HasValue).ToList();
            result.SpeakersRenamed = renamed.Count;
            result.SpeakersSkipped = fields.Count - renamed.Count;
            if (renamed.Count == 0) return result;

            // Materialize once: the story list doesn't change as text is replaced.
            List<Word.Range> stories = DocumentStories.Enumerate(doc).ToList();

            Word.UndoRecord undo = app.UndoRecord;
            bool screenUpdating = app.ScreenUpdating;
            undo.StartCustomRecord(UndoName);
            app.ScreenUpdating = false;
            try
            {
                foreach (SpeakerField field in renamed)
                {
                    string value = NormalizeValue(field.Value);
                    foreach (string token in field.Tokens)
                    {
                        foreach (Word.Range story in stories)
                        {
                            result.OccurrencesReplaced += ReplaceInStory(story, token, value);
                        }
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
        /// Find/replace one label variant inside one story. Uses a Find loop and sets
        /// <c>Range.Text</c> instead of <c>Find.Execute(Replace: wdReplaceAll)</c> because
        /// Word caps replacement text at 255 characters and treats <c>^</c> specially in it.
        /// The inserted text takes the formatting of the label it replaces.
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
                // Continue searching after the inserted text, so a name that itself
                // contains "Speaker N" can't loop forever.
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
