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
    public Span<byte> GetSpan() => _buffer.AsSpan(Index);
    private Span<byte> GetSpan(int start) => _buffer.AsSpan(Index + start);
    public int Index { get; private set; }
    public void Advance(int bytes) => Index += bytes;

    public bool WriteLongString(ReadOnlySpan<char> value, ref int valueIndex)
    {
        var bytesWritten = 0;
        var result = SpanHelper.TryWriteLongString(value, ref valueIndex, GetSpan(), ref bytesWritten);
        Index += bytesWritten;
        return result;
    }

#if NETSTANDARD2_0
    public bool WriteLongString(string? value, ref int valueIndex) => WriteLongString(value.AsSpan(), ref valueIndex);
#endif

    public ValueTask FlushToStreamAsync(Stream stream, CancellationToken token)
    {
        var index = Index;
        Index = 0;
#if NETSTANDARD2_0
        return new ValueTask(stream.WriteAsync(_buffer, 0, index, token));
#else
        return stream.WriteAsync(_buffer.AsMemory(0, index), token);
#endif
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
        var pos = handler._pos;
        if (pos != 0)
        {
            Advance(pos);
            return true;
        }

        return false;
    }

    public bool TryWrite(int bytesToSkip, out int bytesWritten, [InterpolatedStringHandlerArgument("", "bytesToSkip")] ref TryWriteInterpolatedStringHandler handler)
    {
        var (pos, isSuccess) = handler;
        if (pos != 0)
        {
            Advance(pos);
            bytesWritten = pos;
        }
        else
        {
            bytesWritten = 0;
        }
        
        return isSuccess;
    }

    [InterpolatedStringHandler]
#pragma warning disable CS9113 // Parameter is unread.
    public ref struct TryWriteInterpolatedStringHandler(int literalLength, int formattedCount, SpreadsheetBuffer buffer)
#pragma warning restore CS9113 // Parameter is unread.
    {
        internal int _pos;
        internal int _bytesToSkip;
        internal bool _keepPosOnFail; // for backward compatibility
        internal bool _isSuccess;

        public TryWriteInterpolatedStringHandler(int literalLength, int formattedCount, SpreadsheetBuffer buffer, int bytesToSkip)
            : this(literalLength, formattedCount, buffer)
        {
            _bytesToSkip = bytesToSkip;
            _keepPosOnFail = true;
        }

        public void Deconstruct(out int pos, out bool isSuccess)
        {
            pos = _pos;
            isSuccess = _isSuccess;
        }

        private readonly Span<byte> GetSpan() => buffer.GetSpan(_pos);

        private void Advance(int bytes)
        {
            if (_bytesToSkip >= bytes)
            {
                // Ideally, _bytesToSkip will never be negative. We can check for this.
                _bytesToSkip -= bytes;
            }
            else
            {
                _pos += bytes;
            }
        }

        [ExcludeFromCodeCoverage]
        public bool AppendLiteral(string value)
        {
            Debug.Fail("Use ReadOnlySpan<byte> instead of string literals");

            if (value is not null && Utf8Helper.TryGetBytes(value.AsSpan(), GetSpan(), out var bytesWritten))
            {
                Advance(bytesWritten);
                return Success();
            }

            return Fail();
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
                Advance(1);
                return Success();
            }

            return Fail();
        }

        public bool AppendFormatted(int value)
        {
#if NET8_0_OR_GREATER
            if (value.TryFormat(GetSpan(), out var bytesWritten, provider: NumberFormatInfo.InvariantInfo))
#else
            if (Utf8Formatter.TryFormat(value, GetSpan(), out var bytesWritten))
#endif
            {
                Advance(bytesWritten);
                return Success();
            }

            return Fail();
        }

        public bool AppendFormatted(uint value)
        {
#if NET8_0_OR_GREATER
            if (value.TryFormat(GetSpan(), out var bytesWritten, provider: NumberFormatInfo.InvariantInfo))
#else
            if (Utf8Formatter.TryFormat(value, GetSpan(), out var bytesWritten))
#endif
            {
                Advance(bytesWritten);
                return Success();
            }

            return Fail();
        }

        public bool AppendFormatted(ushort value)
        {
#if NET8_0_OR_GREATER
            if (value.TryFormat(GetSpan(), out var bytesWritten, provider: NumberFormatInfo.InvariantInfo))
#else
            if (Utf8Formatter.TryFormat(value, GetSpan(), out var bytesWritten))
#endif
            {
                Advance(bytesWritten);
                return Success();
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
                Advance(bytesWritten);
                return Success();
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
                Advance(bytesWritten);
                return Success();
            }

            return Fail();
        }

        public bool AppendFormatted((double, StandardFormat) value)
        {
            if (Utf8Formatter.TryFormat(value.Item1, GetSpan(), out var bytesWritten, value.Item2))
            {
                Advance(bytesWritten);
                return Success();
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
                Advance(8);
                return Success();
            }

            return Fail();
        }

        public bool AppendFormatted(DateTime dateTime)
        {
            var span = GetSpan();
            if (Utf8Formatter.TryFormat(dateTime, span, out _, new StandardFormat('O')))
            {
                span[19] = (byte)'Z';
                Advance(20);
                return Success();
            }

            return Fail();
        }

        public bool AppendFormatted(OADate oaDate)
        {
            if (oaDate.TryFormat(GetSpan(), out var bytesWritten))
            {
                Advance(bytesWritten);
                return Success();
            }

            return Fail();
        }

        public bool AppendFormatted(SimpleSingleCellReference reference)
        {
            var bytesWritten = 0;
            if (SpanHelper.TryWriteCellReference(reference.Column, reference.Row, GetSpan(), ref bytesWritten))
            {
                Advance(bytesWritten);
                return Success();
            }

            return Fail();
        }

        public bool AppendFormatted(IntAttribute attribute)
        {
            if (attribute.Value is not { } value)
                return Success();

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
            
            return Success();
        }

        public bool AppendFormatted(BooleanAttribute attribute)
        {
            if (attribute.Value is not { } value)
                return Success();

            var span = GetSpan();
            span[0] = (byte)' ';
            Advance(1);

            if (!AppendFormatted(attribute.AttributeName))
                return Fail();

            span = GetSpan();
            span[0] = (byte)'=';
            span[1] = (byte)'"';
            Advance(2);

            if (!AppendFormatted(value))
                return Fail();

            span = GetSpan();
            span[0] = (byte)'"';
            Advance(1);

            return Success();
        }

        public bool AppendFormatted(SpanByteAttribute attribute)
        {
            if (attribute.Value.IsEmpty)
                return Success();

            var span = GetSpan();
            span[0] = (byte)' ';
            Advance(1);

            if (!AppendFormatted(attribute.AttributeName))
                return Fail();

            span = GetSpan();
            span[0] = (byte)'=';
            span[1] = (byte)'"';
            Advance(2);

            if (!AppendFormatted(attribute.Value))
                return Fail();

            span = GetSpan();
            span[0] = (byte)'"';
            Advance(1);

            return Success();
        }
        
        public bool AppendFormatted(SimpleSingleCellReferenceAttribute attribute)
        {
            if (attribute.Value is not { } value)
                return true;

            var span = GetSpan();
            span[0] = (byte)' ';
            Advance(1);

            if (!AppendFormatted(attribute.AttributeName))
                return Fail();

            span = GetSpan();
            span[0] = (byte)'=';
            span[1] = (byte)'"';
            Advance(2);

            if (!AppendFormatted(value))
                return Fail();

            span = GetSpan();
            span[0] = (byte)'"';
            Advance(1);
            
            return Success();
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
                return Success();

            if (XmlUtility.TryXmlEncodeToUtf8(value, GetSpan(), out var bytesWritten))
            {
                Advance(bytesWritten);
                return Success();
            }

            return Fail();
        }

        public bool AppendFormatted(scoped ReadOnlySpan<byte> utf8Value)
        {
            if (utf8Value.TryCopyTo(GetSpan()))
            {
                Advance(utf8Value.Length);
                return Success();
            }

            return Fail();
        }

        public bool AppendFormatted(CellWriterState state)
        {
            var bytes = GetSpan();
            var bytesWritten = 0;

            if (!"<c r=\""u8.TryCopyTo(bytes, ref bytesWritten))
            {
                Advance(bytesWritten);
                return Fail();
            }
            if (!SpanHelper.TryWriteCellReference(state.Column + 1, state.NextRowIndex - 1, bytes, ref bytesWritten))
            {
                Advance(bytesWritten);
                return Fail();
            }

            Advance(bytesWritten);
            return Success();
        }

        private bool Fail()
        {
            if (!_keepPosOnFail)
            {
                _pos = 0;
            }

            return _isSuccess = false;
        }

        private bool Success()
        {
            return _isSuccess = true;
        }
    }
}
