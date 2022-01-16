using SpreadCheetah.Helpers;

namespace SpreadCheetah;

internal sealed class SpreadsheetBuffer
{
    private readonly byte[] _buffer;
    private int _index;

    public SpreadsheetBuffer(byte[] buffer)
    {
        _buffer = buffer;
    }

    public Span<byte> GetNextSpan() => _buffer.AsSpan(_index);
    public int GetRemainingBuffer() => _buffer.Length - _index;
    public void Advance(int bytes) => _index += bytes;
    public bool Empty => _index == 0;

    public async ValueTask WriteAsciiStringAsync(string value, Stream stream, CancellationToken token)
    {
        // When value is ASCII, the number of bytes equals the length of the string
        if (value.Length > GetRemainingBuffer())
            await FlushToStreamAsync(stream, token).ConfigureAwait(false);

        _index += Utf8Helper.GetBytes(value, GetNextSpan());
    }

    public ValueTask FlushToStreamAsync(Stream stream, CancellationToken token)
    {
        var index = _index;
        _index = 0;
#if NETSTANDARD2_0
        return new ValueTask(stream.WriteAsync(_buffer, 0, index, token));
#else
        return stream.WriteAsync(_buffer.AsMemory(0, index), token);
#endif
    }
}
