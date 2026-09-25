using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace SpeakerRename
{
    /// <summary>
    /// Enumerates every text story in a document: main body, footnotes, endnotes, comments,
    /// text boxes, and each section's headers/footers — plus text boxes/shapes that live
    /// inside headers and footers, which <c>StoryRanges</c> alone does not reach.
    /// Scanner and replacer both use this so they always see the same text.
    /// </summary>
    internal static class DocumentStories
    {
        public static IEnumerable<Word.Range> Enumerate(Word.Document doc)
        {
            // Word quirk: header/footer stories (and shapes inside them) are not always
            // present in StoryRanges until a header has been touched once this session.
            TouchPrimaryHeader(doc);

            foreach (Word.Range first in doc.StoryRanges)
            {
                // Each StoryRanges entry is the first story of its type; NextStoryRange
                // walks the rest (e.g. the headers of later sections, further text boxes).
                Word.Range story = first;
                while (story != null)
                {
                    yield return story;

                    if (IsHeaderFooter(story.StoryType))
                    {
                        foreach (Word.Range shapeText in ShapeTextRanges(story))
                            yield return shapeText;
                    }

                    story = story.NextStoryRange;
                }
            }
        }

        private static void TouchPrimaryHeader(Word.Document doc)
        {
            Word.Sections sections = doc.Sections;
            Word.Section section = sections[1];
            Word.HeadersFooters headers = section.Headers;
            Word.HeaderFooter header = headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
            Word.Range range = header.Range;
            _ = range.StoryType;
        }

        private static IEnumerable<Word.Range> ShapeTextRanges(Word.Range headerFooterStory)
        {
            Word.ShapeRange shapes = headerFooterStory.ShapeRange;
            if (shapes == null || shapes.Count == 0) yield break;

            foreach (Word.Shape shape in shapes)
            {
                Word.TextFrame frame;
                try
                {
                    // Pictures, lines, and group shapes have no usable text frame.
                    frame = shape.TextFrame;
                    if (frame == null || frame.HasText == 0) continue;
                }
                catch
                {
                    continue;
                }

                yield return frame.TextRange;
            }
        }

        private static bool IsHeaderFooter(Word.WdStoryType type)
        {
            switch (type)
            {
                case Word.WdStoryType.wdEvenPagesHeaderStory:
                case Word.WdStoryType.wdPrimaryHeaderStory:
                case Word.WdStoryType.wdFirstPageHeaderStory:
                case Word.WdStoryType.wdEvenPagesFooterStory:
                case Word.WdStoryType.wdPrimaryFooterStory:
                case Word.WdStoryType.wdFirstPageFooterStory:
                    return true;
                default:
                    return false;
            }
        }
    }
}
