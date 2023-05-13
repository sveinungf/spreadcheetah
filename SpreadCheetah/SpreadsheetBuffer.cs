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

    public Span<byte> GetSpan() => _buffer.AsSpan(_index);
    public int FreeCapacity => _buffer.Length - _index;
    public void Advance(int bytes) => _index += bytes;

    public bool WriteLongString(ReadOnlySpan<char> value, ref int valueIndex)
    {
        var bytesWritten = 0;
        var result = SpanHelper.TryWriteLongString(value, ref valueIndex, GetSpan(), ref bytesWritten);
        _index += bytesWritten;
        return result;
    }

#if NETSTANDARD2_0
    public bool WriteLongString(string? value, ref int valueIndex) => WriteLongString(value.AsSpan(), ref valueIndex);
#endif

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
