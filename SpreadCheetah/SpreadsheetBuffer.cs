using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using System.Buffers.Text;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.CompilerServices;

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

    public bool TryWrite([InterpolatedStringHandlerArgument("")] ref TryWriteInterpolatedStringHandler handler)
    {
        if (handler._success)
        {
            Advance(handler._pos);
            return true;
        }

        return false;
    }

    [InterpolatedStringHandler]
    public ref struct TryWriteInterpolatedStringHandler
    {
        private readonly SpreadsheetBuffer _buffer;
        internal int _pos;
        internal bool _success;

        public TryWriteInterpolatedStringHandler(int literalLength, int formattedCount, SpreadsheetBuffer buffer)
        {
            _ = literalLength;
            _ = formattedCount;
            _buffer = buffer;
            _success = true;
        }

        private readonly Span<byte> GetSpan() => _buffer._buffer.AsSpan(_buffer._index + _pos);

        // TODO: Remove or throw? Literals should rather use ReadOnlySpan<byte>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool AppendLiteral(string value)
        {
            if (value is not null)
            {
                var dest = GetSpan();
#if NET8_0_OR_GREATER
                if (System.Text.Unicode.Utf8.TryWrite(dest, CultureInfo.InvariantCulture, $"{value}", out var bytesWritten))
#else
                if (Utf8Helper.TryGetBytes(value, dest, out var bytesWritten))
#endif
                {
                    _pos += bytesWritten;
                    return true;
                }
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

        public bool AppendFormatted<T>(T value)
        {
            Debug.Fail("Create non-generic overloads to avoid allocations when running on .NET Framework");

            string? s = value is IFormattable f
                ? f.ToString(null, CultureInfo.InvariantCulture)
                : value?.ToString();

            return AppendFormatted(s);
        }

        public bool AppendFormatted(string? value)
        {
            if (Utf8Helper.TryGetBytes(value, GetSpan(), out int bytesWritten))
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
            _success = false;
            return false;
        }
    }
}
