using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using System.Buffers;
using System.Buffers.Text;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace SpreadCheetah;

internal sealed class SpreadsheetBuffer(int bufferSize) : IDisposable
{
    private readonly byte[] _buffer = ArrayPool<byte>.Shared.Rent(bufferSize);
    private int _index;

    public void Dispose() => ArrayPool<byte>.Shared.Return(_buffer, true);
    public Span<byte> GetSpan() => _buffer.AsSpan(_index);
    private Span<byte> GetSpan(int start) => _buffer.AsSpan(_index + start);
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

    public bool TryWrite(scoped ReadOnlySpan<byte> utf8Value)
    {
        if (utf8Value.TryCopyTo(GetSpan()))
        {
            Advance(utf8Value.Length);
            return true;
        }

        return false;
    }

    public bool TryWrite([InterpolatedStringHandlerArgument("")] ref TryWriteInterpolatedStringHandler handler)
    {
        var pos = handler._pos;
        if (pos != 0)
        {
            Advance(pos);
            return true;
        }

        return false;
    }

    [InterpolatedStringHandler]
#pragma warning disable CS9113 // Parameter is unread.
    public ref struct TryWriteInterpolatedStringHandler(int literalLength, int formattedCount, SpreadsheetBuffer buffer)
#pragma warning restore CS9113 // Parameter is unread.
    {
        internal int _pos;

        private readonly Span<byte> GetSpan() => buffer.GetSpan(_pos);

        [ExcludeFromCodeCoverage]
        public bool AppendLiteral(string value)
        {
            Debug.Fail("Use ReadOnlySpan<byte> instead of string literals");

            if (value is not null && Utf8Helper.TryGetBytes(value.AsSpan(), GetSpan(), out var bytesWritten))
            {
                _pos += bytesWritten;
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(int value)
        {
            if (Utf8Formatter.TryFormat(value, GetSpan(), out var bytesWritten))
            {
                _pos += bytesWritten;
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(uint value)
        {
            if (Utf8Formatter.TryFormat(value, GetSpan(), out var bytesWritten))
            {
                _pos += bytesWritten;
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(float value)
        {
            if (Utf8Formatter.TryFormat(value, GetSpan(), out var bytesWritten))
            {
                _pos += bytesWritten;
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(double value)
        {
            if (Utf8Formatter.TryFormat(value, GetSpan(), out var bytesWritten))
            {
                _pos += bytesWritten;
                return true;
            }

            return Fail();
        }

        [ExcludeFromCodeCoverage]
        public bool AppendFormatted<T>(T value)
        {
            Debug.Fail("Create non-generic overloads to avoid allocations when running on .NET Framework");

            string? s = value is IFormattable f
                ? f.ToString(null, CultureInfo.InvariantCulture)
                : value?.ToString();

            return AppendFormatted(s);
        }

        public bool AppendFormatted(string? value) => AppendFormatted(value.AsSpan());

        public bool AppendFormatted(scoped ReadOnlySpan<char> value)
        {
            if (value.IsEmpty)
                return true;

            if (XmlUtility.TryXmlEncodeToUtf8(value, GetSpan(), out var bytesWritten))
            {
                _pos += bytesWritten;
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(scoped ReadOnlySpan<byte> utf8Value)
        {
            if (utf8Value.TryCopyTo(GetSpan()))
            {
                _pos += utf8Value.Length;
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(CellWriterState state)
        {
            var bytes = GetSpan();
            var bytesWritten = 0;

            if (!"<c r=\""u8.TryCopyTo(bytes, ref bytesWritten)) return Fail();
            if (!SpanHelper.TryWriteCellReference(state.Column + 1, state.NextRowIndex - 1, bytes, ref bytesWritten)) return Fail();

            _pos += bytesWritten;
            return true;
        }

        private bool Fail()
        {
            _pos = 0;
            return false;
        }
    }
}
