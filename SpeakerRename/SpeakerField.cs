using System.Collections.Generic;

namespace SpeakerRename
{
    /// <summary>
    /// One unique speaker number found in a document, e.g. the <c>1</c> in <c>Speaker 1</c>.
    /// Two occurrences are the same speaker only if the number matches. The literal text can
    /// still vary in whitespace (e.g. <c>Speaker  1</c> with two spaces), so every distinct
    /// raw string seen for that number is kept in <see cref="Tokens"/> so all variants get
    /// replaced, not just the first one encountered.
    /// </summary>
    internal sealed class SpeakerField
    {
        public SpeakerField(int number, string token, string context)
        {
            Number = number;
            Label = "Speaker " + number;
            Tokens = new List<string> { token };
            Context = context;
        }

        /// <summary>The number after "Speaker", e.g. <c>1</c>.</summary>
        public int Number { get; }

        /// <summary>Canonical display text, e.g. <c>Speaker 1</c>.</summary>
        public string Label { get; }

        /// <summary>Every distinct exact-text form of this speaker's label seen in the document.</summary>
        public List<string> Tokens { get; }

        /// <summary>How many times any variant of this speaker's label appears across all document stories.</summary>
        public int Occurrences { get; set; }

        /// <summary>A short snippet of text around the first occurrence, to help the user identify the speaker.</summary>
        public string Context { get; }

        /// <summary>The name the user entered. Null or empty = leave this speaker's label as-is.</summary>
        public string Value { get; set; }

        public bool HasValue => !string.IsNullOrEmpty(Value);

        /// <summary>List-box display text; the tick marks speakers that already have a name.</summary>
        public override string ToString() => (HasValue ? "✓  " : "     ") + Label;
    }
}
