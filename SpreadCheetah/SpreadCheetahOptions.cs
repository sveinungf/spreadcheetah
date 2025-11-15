using SpreadCheetah.Helpers;
using SpreadCheetah.Metadata;
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

    /// <summary>The initial default number format used for <see cref="DateTime"/> cells, which is <see cref="NumberFormats.DateTimeSortable"/>.</summary>
    internal static readonly NumberFormat InitialDefaultDateTimeFormat = NumberFormat.Custom(NumberFormats.DateTimeSortable);

    /// <summary>
    /// Compression level to use when generating the spreadsheet archive. Defaults to <see cref="SpreadCheetahCompressionLevel.Fastest"/>.
    /// </summary>
    public SpreadCheetahCompressionLevel CompressionLevel { get; set; } = DefaultCompressionLevel;

    /// <summary>
    /// The buffer size in number of bytes. The default size is 65536. The minimum allowed size is 512.
    /// </summary>
    public int BufferSize { get; set => field = Guard.SufficientBufferSize(value); } = DefaultBufferSize;

    /// <summary>
    /// The default number format used for <see cref="DateTime"/> cells. Defaults to <see cref="NumberFormats.DateTimeSortable"/>.
    /// </summary>
    [Obsolete($"Use {nameof(SpreadCheetahOptions)}.{nameof(DefaultDateTimeFormat)} instead")]
    public string? DefaultDateTimeNumberFormat
    {
        get => DefaultDateTimeFormat?.CustomFormat;
        set => DefaultDateTimeFormat = value == null ? null : NumberFormat.FromLegacyString(value);
    }

    /// <summary>
    /// The default number format used for <see cref="DateTime"/> cells. Defaults to <see cref="NumberFormats.DateTimeSortable"/>.
    /// </summary>
    public NumberFormat? DefaultDateTimeFormat { get; set; } = InitialDefaultDateTimeFormat;

    /// <summary>
    /// The default font for all worksheets. When not set, the default is Calibri with size 11.
    /// </summary>
    public DefaultFont? DefaultFont { get; set; }

    /// <summary>
    /// Write the explicit cell reference attribute for each cell in the resulting spreadsheet.
    /// The attribute is optional according to the Open XML specification, but it can be required for some readers such as:
    /// <list type="bullet">
    ///   <item>Microsoft Excel ADO OleDb provider.</item>
    ///   <item>Microsoft Spreadsheet Compare tool.</item>
    /// </list>
    /// Writing the attribute will slightly affect the performance of writing cells, and will also increase the resulting file size.
    /// Defaults to <see langword="false"/>.
    /// </summary>
    public bool WriteCellReferenceAttributes { get; set; }

    /// <summary>
    /// Document properties, such as author and title. This generates the <c>docProps/app.xml</c> and <c>docProps/core.xml</c> files
    /// inside the XLSX file. By default, these files are included for better compatibility with other spreadsheet applications.
    /// The files can be excluded by setting this property to <see langword="null"/>.
    /// </summary>
    public DocumentProperties? DocumentProperties { get; set; } = new();
}
