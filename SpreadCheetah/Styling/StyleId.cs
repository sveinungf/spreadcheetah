namespace SpreadCheetah.Styling
{
    /// <summary>
    /// Identifier for style lookup. Style IDs are created with <see cref="Spreadsheet.AddStyle(Style)"/>.
    /// </summary>
    public sealed class StyleId
    {
        /// <summary>
        /// The unique style identifier.
        /// </summary>
        public int Id { get; }

        internal StyleId(int id) => Id = id;
    }
}
