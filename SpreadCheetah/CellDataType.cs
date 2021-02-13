namespace SpreadCheetah
{
    /// <summary>
    /// Open XML data type for worksheet cell.
    /// </summary>
    public enum CellDataType
    {
        /// <summary>
        /// String type serialized directly in the worksheet XML.
        /// </summary>
        InlineString,

        /// <summary>
        /// Number type used for both integers and floating point numbers.
        /// </summary>
        Number,

        /// <summary>
        /// Logical type with value true or false.
        /// </summary>
        Boolean
    }
}
