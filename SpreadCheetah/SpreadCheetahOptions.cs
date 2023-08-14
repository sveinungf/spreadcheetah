using SpreadCheetah.Styling;

namespace SpreadCheetah;

/// <summary>
/// Provides options to be used when creating a spreadsheet with <see cref="Spreadsheet"/>.
/// </summary>
public class SpreadCheetahOptions
{
    /// <summary>The default buffer size.</summary>
    public static readonly int DefaultBufferSize = 65536;

    /// <summary>The minimum allowed buffer size.</summary>
    public static readonly int MinimumBufferSize = 512;

    /// <summary>The default compression level.</summary>
    public static readonly SpreadCheetahCompressionLevel DefaultCompressionLevel = SpreadCheetahCompressionLevel.Fastest;

    /// <summary>The default default number format used for <see cref="DateTime"/> cells, which is <see cref="NumberFormats.DateTimeSortable"/>.</summary>
    public static readonly NumberFormat DefaultDefaultDateTimeFormat = NumberFormat.Custom(NumberFormats.DateTimeSortable);

    /// <summary>
    /// Compression level to use when generating the spreadsheet archive. Defaults to <see cref="SpreadCheetahCompressionLevel.Fastest"/>.
    /// </summary>
    public SpreadCheetahCompressionLevel CompressionLevel { get; set; } = DefaultCompressionLevel;

    /// <summary>
    /// The buffer size in number of bytes. The default size is 65536. The minimum allowed size is 512.
    /// </summary>
    public int BufferSize
    {
        get => _bufferSize;
        set => _bufferSize = value < MinimumBufferSize
            ? throw new ArgumentOutOfRangeException(nameof(value), value, "Buffer size must be at least " + MinimumBufferSize)
            : value;
    }

    private int _bufferSize = DefaultBufferSize;

    /// <summary>
    /// The default number format used for <see cref="DateTime"/> cells. Defaults to <see cref="NumberFormats.DateTimeSortable"/>.
    /// </summary>
#pragma warning disable S1133 // Deprecated code should be removed - This is required for backwards binary compatibility
    [Obsolete($"Use {nameof(SpreadCheetahOptions)}.{nameof(DefaultDateTimeFormat)} instead")]
    public string? DefaultDateTimeNumberFormat {
        get => DefaultDateTimeFormat?.CustomFormat;
        set => DefaultDateTimeFormat = value == null ? null : NumberFormat.FromLegacyString(value);
    }
#pragma warning restore S1133 // Deprecated code should be removed

    /// <summary>
    /// The default number format used for <see cref="DateTime"/> cells. Defaults to <see cref="NumberFormats.DateTimeSortable"/>.
    /// </summary>
    public NumberFormat? DefaultDateTimeFormat { get; set; } = DefaultDefaultDateTimeFormat;

    /// <summary>
    /// Write the explicit cell reference attribute for each cell in the resulting spreadsheet.
    /// The attribute is optional according to the Open XML specification, but it is required for Microsoft's Spreadsheet Compare tool to do a comparison.
    /// Writing the attribute will slightly affect the performance of writing cells, and will also increase the resulting file size.
    /// Defaults to <c>false</c>.
    /// </summary>
    public bool WriteCellReferenceAttributes { get; set; }
}
