using SpreadCheetah.CellReferences;
using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.MetadataXml.Attributes;
using System.Buffers;
using System.Buffers.Text;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace SpreadCheetah;

internal sealed class SpreadsheetBuffer(int bufferSize) : IDisposable
{
    private readonly byte[] _buffer = ArrayPool<byte>.Shared.Rent(bufferSize);

    public void Dispose() => ArrayPool<byte>.Shared.Return(_buffer, true);
    private Span<byte> GetSpan() => _buffer.AsSpan(Index);
    private Span<byte> GetSpan(int start) => _buffer.AsSpan(Index + start);
    public int Index { get; private set; }
    public void Advance(int bytes) => Index += bytes;

    public bool WriteLongString(ReadOnlySpan<char> value, ref int valueIndex)
    {
        var source = value.Slice(valueIndex);
        var result = XmlUtility.TryXmlEncodeToUtf8(source, GetSpan(), out var charsRead, out var bytesWritten);
        valueIndex += charsRead;
        Index += bytesWritten;
        return result;
    }

    public ValueTask FlushToStreamAsync(Stream stream, CancellationToken token)
    {
        var index = Index;
        Index = 0;
        return stream.WriteAsync(_buffer.AsMemory(0, index), token);
    }

    public bool TryWrite(scoped ReadOnlySpan<byte> utf8Value)
    {
        Debug.Assert(utf8Value.Length <= SpreadCheetahOptions.MinimumBufferSize);
        if (utf8Value.TryCopyTo(GetSpan()))
        {
            Advance(utf8Value.Length);
            return true;
        }

        return false;
    }

    public bool TryWrite([InterpolatedStringHandlerArgument("")] ref TryWriteInterpolatedStringHandler handler)
    {
        Advance(handler._pos);
        return handler._isSuccess;
    }

    public bool TryWrite2(
#pragma warning disable RCS1163, IDE0060 // Unused parameter
        BufferWriteProgress start,
#pragma warning restore RCS1163, IDE0060 // Unused parameter
        out BufferWriteProgress written,
        [InterpolatedStringHandlerArgument("", nameof(start))] ref ResumableTryWriteInterpolatedStringHandler handler)
    {
        written = handler.GetProgress();
        Advance(handler._pos);
        return handler._isSuccess;
    }

    [InterpolatedStringHandler]
#pragma warning disable CS9113 // Parameter is unread.
    public ref struct ResumableTryWriteInterpolatedStringHandler(
        int literalLength,
        int formattedCount,
        SpreadsheetBuffer buffer,
        BufferWriteProgress start)
#pragma warning restore CS9113 // Parameter is unread.
    {
        private readonly int _startingStep = start.Step;
        private int _step;
        private int _index = start.Index;
        internal int _pos;
        internal bool _isSuccess = true;

        public readonly BufferWriteProgress GetProgress() => new()
        {
            Step = _step - 1,
            Index = _index
        };

        private readonly Span<byte> GetSpan() => buffer.GetSpan(_pos);

        [ExcludeFromCodeCoverage]
        public readonly bool AppendLiteral(string value)
        {
            _ = _pos;
            _ = value;
            throw new InvalidOperationException("Use ReadOnlySpan<byte> instead of string literals");
        }

        public bool AppendFormatted(int value)
        {
            if (_step++ < _startingStep)
                return true;

            return _isSuccess = Formatter.TryFormat(value, GetSpan(), ref _pos);
        }

        [ExcludeFromCodeCoverage]
        public bool AppendFormatted<T>(T value)
        {
            Debug.Fail("Create non-generic overloads to avoid allocations when running on .NET Framework");

            var s = value is IFormattable f
                ? f.ToString(null, CultureInfo.InvariantCulture)
                : value?.ToString();

            return AppendFormatted(s);
        }

        public bool AppendFormatted(string? value) => AppendFormatted(value.AsSpan());

        public bool AppendFormatted(scoped ReadOnlySpan<char> value)
        {
            if (_step++ < _startingStep)
                return true;

            var remaining = value.Slice(_index);

            if (remaining.IsEmpty)
                return true;

            var destination = GetSpan();
            if (destination.Length <= remaining.Length)
                return Fail();

            if (XmlUtility.TryXmlEncodeToUtf8(remaining, destination, out var charsRead, out var bytesWritten))
            {
                _pos += bytesWritten;
                return true;
            }

            if (charsRead > 0)
            {
                _pos += bytesWritten;
                _index += charsRead;
            }

            return Fail();
        }

        public bool AppendFormatted(scoped ReadOnlySpan<byte> utf8Value)
        {
            if (_step++ < _startingStep)
                return true;

            if (utf8Value.TryCopyTo(GetSpan()))
            {
                _pos += utf8Value.Length;
                return true;
            }

            return Fail();
        }

        private bool Fail()
        {
            _isSuccess = false;
            return false;
        }
    }

    [InterpolatedStringHandler]
#pragma warning disable CS9113 // Parameter is unread.
    public ref struct TryWriteInterpolatedStringHandler(int literalLength, int formattedCount, SpreadsheetBuffer buffer)
#pragma warning restore CS9113 // Parameter is unread.
    {
        internal int _pos;
        internal bool _isSuccess = true;

        private readonly Span<byte> GetSpan() => buffer.GetSpan(_pos);

        [ExcludeFromCodeCoverage]
        public readonly bool AppendLiteral(string value)
        {
            _ = _pos;
            _ = value;
            throw new InvalidOperationException("Use ReadOnlySpan<byte> instead of string literals");
        }

        /// <summary>
        /// Writes '1' for true and '0' for false.
        /// </summary>
        public bool AppendFormatted(bool value)
        {
            var destination = GetSpan();
            if (destination.Length > 0)
            {
                destination[0] = (byte)('0' + (value ? 1 : 0)); // Branchless on .NET 8+
                _pos++;
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(int value) => Formatter.TryFormat(value, GetSpan(), ref _pos) || Fail();

        public bool AppendFormatted(uint value)
        {
#if NET8_0_OR_GREATER
            if (value.TryFormat(GetSpan(), out var bytesWritten, provider: NumberFormatInfo.InvariantInfo))
#else
            if (Utf8Formatter.TryFormat(value, GetSpan(), out var bytesWritten))
#endif
            {
                _pos += bytesWritten;
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(long value)
        {
#if !NET8_0_OR_GREATER
            return AppendFormatted((double)value);
#else
            ReadOnlySpan<char> format = (ulong)(value + 9999999999999999L) < 19999999999999999L
                ? []
                : "0.################E+00";

            if (value.TryFormat(GetSpan(), out var bytesWritten, format, NumberFormatInfo.InvariantInfo))
            {
                _pos += bytesWritten;
                return true;
            }

            return Fail();
#endif
        }

        public bool AppendFormatted(ushort value)
        {
#if NET8_0_OR_GREATER
            if (value.TryFormat(GetSpan(), out var bytesWritten, provider: NumberFormatInfo.InvariantInfo))
#else
            if (Utf8Formatter.TryFormat(value, GetSpan(), out var bytesWritten))
#endif
            {
                _pos += bytesWritten;
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(float value)
        {
#if NET8_0_OR_GREATER
            if (value.TryFormat(GetSpan(), out var bytesWritten, provider: NumberFormatInfo.InvariantInfo))
#else
            if (Utf8Formatter.TryFormat(value, GetSpan(), out var bytesWritten))
#endif
            {
                _pos += bytesWritten;
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(double value)
        {
#if NET8_0_OR_GREATER
            if (value.TryFormat(GetSpan(), out var bytesWritten, provider: NumberFormatInfo.InvariantInfo))
#else
            if (Utf8Formatter.TryFormat(value, GetSpan(), out var bytesWritten))
#endif
            {
                _pos += bytesWritten;
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted((double, StandardFormat) value)
        {
            if (Utf8Formatter.TryFormat(value.Item1, GetSpan(), out var bytesWritten, value.Item2))
            {
                _pos += bytesWritten;
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(Color color)
        {
            var span = GetSpan();
            if (span.Length >= 8)
            {
                var format = new StandardFormat('X', 2);
                Utf8Formatter.TryFormat(color.A, span, out _, format);
                span = span.Slice(2);
                Utf8Formatter.TryFormat(color.R, span, out _, format);
                span = span.Slice(2);
                Utf8Formatter.TryFormat(color.G, span, out _, format);
                span = span.Slice(2);
                Utf8Formatter.TryFormat(color.B, span, out _, format);
                _pos += 8;
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(DateTime dateTime)
        {
            var span = GetSpan();
            if (Utf8Formatter.TryFormat(dateTime, span, out _, new StandardFormat('O')))
            {
                span[19] = (byte)'Z';
                _pos += 20;
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(OADate oaDate)
        {
            if (oaDate.TryFormat(GetSpan(), out var bytesWritten))
            {
                _pos += bytesWritten;
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(SimpleSingleCellReference reference)
        {
            if (!SpreadsheetUtility.TryGetColumnNameUtf8(reference.Column, GetSpan(), out var nameLength))
                return Fail();

            _pos += nameLength;

            if (!AppendFormatted(reference.Row))
                return Fail();

            return true;
        }

        public bool AppendFormatted(BooleanAttribute attribute)
        {
            if (attribute.Value is not { } value)
                return true;

            if (!AppendFormatted(" "u8))
                return Fail();

            if (!AppendFormatted(attribute.AttributeName))
                return Fail();

            if (!AppendFormatted("=\""u8))
                return Fail();

            if (!AppendFormatted(value))
                return Fail();

            if (!AppendFormatted("\""u8))
                return Fail();

            return true;
        }

        public bool AppendFormatted(IntAttribute attribute)
        {
            if (attribute.Value is not { } value)
                return true;

            if (!AppendFormatted(" "u8))
                return Fail();

            if (!AppendFormatted(attribute.AttributeName))
                return Fail();

            if (!AppendFormatted("=\""u8))
                return Fail();

            if (!AppendFormatted(value))
                return Fail();

            if (!AppendFormatted("\""u8))
                return Fail();

            return true;
        }

        public bool AppendFormatted(DoubleAttribute attribute)
        {
            if (attribute.Value is not { } value)
                return true;

            if (!AppendFormatted(" "u8))
                return Fail();

            if (!AppendFormatted(attribute.AttributeName))
                return Fail();

            if (!AppendFormatted("=\""u8))
                return Fail();

            if (!AppendFormatted(value))
                return Fail();

            if (!AppendFormatted("\""u8))
                return Fail();

            return true;
        }

        public bool AppendFormatted(SpanByteAttribute attribute)
        {
            if (attribute.Value.IsEmpty)
                return true;

            if (!AppendFormatted(" "u8))
                return Fail();

            if (!AppendFormatted(attribute.AttributeName))
                return Fail();

            if (!AppendFormatted("=\""u8))
                return Fail();

            if (!AppendFormatted(attribute.Value))
                return Fail();

            if (!AppendFormatted("\""u8))
                return Fail();

            return true;
        }

        public bool AppendFormatted(SimpleSingleCellReferenceAttribute attribute)
        {
            if (attribute.Value is not { } value)
                return true;

            if (!AppendFormatted(" "u8))
                return Fail();

            if (!AppendFormatted(attribute.AttributeName))
                return Fail();

            if (!AppendFormatted("=\""u8))
                return Fail();

            if (!AppendFormatted(value))
                return Fail();

            if (!AppendFormatted("\""u8))
                return Fail();

            return true;
        }

        public bool AppendFormatted(ColorAttribute attribute)
        {
            if (attribute.Value is not { } value)
                return true;

            if (!AppendFormatted(" "u8))
                return Fail();

            if (!AppendFormatted(attribute.AttributeName))
                return Fail();

            if (!AppendFormatted("=\""u8))
                return Fail();

            if (!AppendFormatted(value))
                return Fail();

            if (!AppendFormatted("\""u8))
                return Fail();

            return true;
        }

        [ExcludeFromCodeCoverage]
        public bool AppendFormatted<T>(T value)
        {
            Debug.Fail("Create non-generic overloads to avoid allocations when running on .NET Framework");

            var s = value is IFormattable f
                ? f.ToString(null, CultureInfo.InvariantCulture)
                : value?.ToString();

            return AppendFormatted(s);
        }

        public bool AppendFormatted(string? value) => AppendFormatted(value.AsSpan());

        public bool AppendFormatted(scoped ReadOnlySpan<char> value)
        {
            if (value.IsEmpty)
                return true;

            var destination = GetSpan();
            if (destination.Length > value.Length &&
                XmlUtility.TryXmlEncodeToUtf8(value, destination, out _, out var bytesWritten))
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
            if (!AppendFormatted("<c r=\""u8))
                return Fail();

            var reference = new SimpleSingleCellReference((ushort)(state.Column + 1), state.NextRowIndex - 1);
            if (!AppendFormatted(reference))
                return Fail();

            return true;
        }

        private bool Fail()
        {
            _pos = 0;
            _isSuccess = false;
            return false;
        }
    }
}

file static class Formatter
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryFormat(int value, Span<byte> destination, ref int pos)
    {
#if NET8_0_OR_GREATER
        var success = value.TryFormat(destination, out var bytesWritten, provider: NumberFormatInfo.InvariantInfo);
#else
        var success = Utf8Formatter.TryFormat(value, destination, out var bytesWritten);
#endif
        pos += bytesWritten;
        return success;
    }
}