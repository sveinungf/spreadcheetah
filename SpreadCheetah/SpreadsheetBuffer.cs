using SpreadCheetah.Helpers;
using System.Diagnostics;
using System.Text;

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

    public async ValueTask WriteAsciiStringAsync(string value, Stream stream, CancellationToken token)
    {
        // When value is ASCII, the number of bytes equals the length of the string
        if (value.Length > GetRemainingBuffer())
            await FlushToStreamAsync(stream, token).ConfigureAwait(false);

        _index += Utf8Helper.GetBytes(value, GetNextSpan());
    }

    public async ValueTask WriteStringAsync(StringBuilder sb, Stream stream, CancellationToken token)
    {
        var remaining = GetRemainingBuffer();
        var value = sb.ToString();

        // Try with an approximate cell value length
        var bytesNeeded = value.Length * Utf8Helper.MaxBytePerChar;
        if (bytesNeeded > remaining)
        {
            // Try with a more accurate cell value length
            bytesNeeded = Utf8Helper.GetByteCount(value);
        }

        if (bytesNeeded > remaining)
            await FlushToStreamAsync(stream, token).ConfigureAwait(false);

        // Write whole value if it fits in the buffer
        if (bytesNeeded <= _buffer.Length)
        {
            _index += Utf8Helper.GetBytes(value, GetNextSpan());
            return;
        }

        // Otherwise, write value piece by piece
        var valueIndex = 0;
        while (!WriteLongString(value.AsSpan(), ref valueIndex))
        {
            await FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }
    }

    public bool WriteLongString(ReadOnlySpan<char> value, ref int valueIndex)
    {
        var remainingBuffer = GetRemainingBuffer();
        var maxCharCount = remainingBuffer / Utf8Helper.MaxBytePerChar;
        var remainingLength = value.Length - valueIndex;
        var lastIteration = remainingLength <= maxCharCount;
        var length = Math.Min(remainingLength, maxCharCount);
        _index += Utf8Helper.GetBytes(value.Slice(valueIndex, length), GetNextSpan());
        valueIndex += length;
        return lastIteration;
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
