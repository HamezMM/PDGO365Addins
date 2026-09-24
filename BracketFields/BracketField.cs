namespace BracketFields
{
    /// <summary>
    /// One unique bracketed parameter found in a document, e.g. <c>[Client Name]</c>.
    /// Two occurrences are the same field only if their text matches exactly (case-sensitive).
    /// </summary>
    internal sealed class BracketField
    {
        public BracketField(string token, string context)
        {
            Token = token;
            Context = context;
        }

        /// <summary>The full text as it appears in the document, brackets included: <c>[Client Name]</c>.</summary>
        public string Token { get; }

        /// <summary>The text between the brackets: <c>Client Name</c>.</summary>
        public string Name => Token.Substring(1, Token.Length - 2);

        /// <summary>How many times the token appears across all document stories.</summary>
        public int Occurrences { get; set; }

        /// <summary>A short snippet of text around the first occurrence, to help the user identify the field.</summary>
        public string Context { get; }

        /// <summary>The replacement the user entered. Null or empty = leave the field untouched.</summary>
        public string Value { get; set; }

        public bool HasValue => !string.IsNullOrEmpty(Value);

        /// <summary>List-box display text; the tick marks fields that already have a value.</summary>
        public override string ToString() => (HasValue ? "✓  " : "     ") + Name;
    }
}
