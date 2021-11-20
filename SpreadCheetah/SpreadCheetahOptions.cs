namespace SpreadCheetah;

/// <summary>
/// Provides options to be used when creating a spreadsheet with <see cref="Spreadsheet"/>.
/// </summary>
public class SpreadCheetahOptions

{
    /// <summary>
    /// The minimum allowed buffer size.
    /// </summary>
    public static readonly int MinimumBufferSize = 512;

    /// <summary>
    /// Compression level to use when generating the spreadsheet archive. Defaults to <see cref="SpreadCheetahCompressionLevel.Fastest"/>.
    /// </summary>
    public SpreadCheetahCompressionLevel CompressionLevel { get; set; } = SpreadCheetahCompressionLevel.Fastest;

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

    private int _bufferSize = 65536;
}
