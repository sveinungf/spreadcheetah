namespace SpreadCheetah
{
    /// <summary>
    /// Values that indicate whether to emphasize speed or archive size when generating the spreadsheet archive.
    /// </summary>
    public enum SpreadCheetahCompressionLevel
    {
        /// <summary>
        /// Smaller archive size, but can take longer to generate.
        /// </summary>
        Optimal = 0,

        /// <summary>
        /// Generates faster, but can result in a larger archive size.
        /// </summary>
        Fastest = 1
    }
}
