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
        Advance(handler.Written);
        return handler._isSuccess;
    }

    public bool TryWrite(
#pragma warning disable RCS1163, IDE0060 // Unused parameter
        BufferWriteProgress start,
#pragma warning restore RCS1163, IDE0060 // Unused parameter
        out BufferWriteProgress written,
        [InterpolatedStringHandlerArgument("", nameof(start))] ref ResumableTryWriteInterpolatedStringHandler handler)
    {
        written = handler.GetProgress();
        Advance(handler.Written);
        return handler._isSuccess;
    }

    [InterpolatedStringHandler]
    public ref struct TryWriteInterpolatedStringHandler
    {
        private readonly int _initialLength;
        private Span<byte> _destination;
        internal bool _isSuccess = true;

        public readonly int Written => _isSuccess ? _initialLength - _destination.Length : 0;

        public TryWriteInterpolatedStringHandler(int _, int _2, SpreadsheetBuffer buffer)
        {
            _destination = buffer.GetSpan();
            _initialLength = _destination.Length;
        }

        [ExcludeFromCodeCoverage]
        public readonly bool AppendLiteral(string value)
        {
            _ = _isSuccess;
            _ = value;
            throw new InvalidOperationException("Use ReadOnlySpan<byte> instead of string literals");
        }

        /// <summary>
        /// Writes '1' for true and '0' for false.
        /// </summary>
        public bool AppendFormatted(bool value)
        {
            if (_destination.Length > 0)
            {
                _destination[0] = (byte)('0' + (value ? 1 : 0)); // Branchless on .NET 8+
                _destination = _destination[1..];
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(int value) => Formatter.TryFormat(value, ref _destination) || Fail();
        public bool AppendFormatted(scoped ReadOnlySpan<byte> value) => Formatter.TryFormat(value, ref _destination) || Fail();

        public bool AppendFormatted(uint value)
        {
#if NET8_0_OR_GREATER
            if (value.TryFormat(_destination, out var bytesWritten, provider: NumberFormatInfo.InvariantInfo))
#else
            if (Utf8Formatter.TryFormat(value, _destination, out var bytesWritten))
#endif
            {
                _destination = _destination[bytesWritten..];
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

            if (value.TryFormat(_destination, out var bytesWritten, format, NumberFormatInfo.InvariantInfo))
            {
                _destination = _destination[bytesWritten..];
                return true;
            }

            return Fail();
#endif
        }

        public bool AppendFormatted(ushort value)
        {
#if NET8_0_OR_GREATER
            if (value.TryFormat(_destination, out var bytesWritten, provider: NumberFormatInfo.InvariantInfo))
#else
            if (Utf8Formatter.TryFormat(value, _destination, out var bytesWritten))
#endif
            {
                _destination = _destination[bytesWritten..];
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(float value)
        {
#if NET8_0_OR_GREATER
            if (value.TryFormat(_destination, out var bytesWritten, provider: NumberFormatInfo.InvariantInfo))
#else
            if (Utf8Formatter.TryFormat(value, _destination, out var bytesWritten))
#endif
            {
                _destination = _destination[bytesWritten..];
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(double value)
        {
#if NET8_0_OR_GREATER
            if (value.TryFormat(_destination, out var bytesWritten, provider: NumberFormatInfo.InvariantInfo))
#else
            if (Utf8Formatter.TryFormat(value, _destination, out var bytesWritten))
#endif
            {
                _destination = _destination[bytesWritten..];
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted((double, StandardFormat) value)
        {
            if (Utf8Formatter.TryFormat(value.Item1, _destination, out var bytesWritten, value.Item2))
            {
                _destination = _destination[bytesWritten..];
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(Color color)
        {
            if (_destination.Length >= 8)
            {
                var format = new StandardFormat('X', 2);
                Utf8Formatter.TryFormat(color.A, _destination, out _, format);
                _destination = _destination[2..];
                Utf8Formatter.TryFormat(color.R, _destination, out _, format);
                _destination = _destination[2..];
                Utf8Formatter.TryFormat(color.G, _destination, out _, format);
                _destination = _destination[2..];
                Utf8Formatter.TryFormat(color.B, _destination, out _, format);
                _destination = _destination[2..];
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(DateTime dateTime)
        {
            if (Utf8Formatter.TryFormat(dateTime, _destination, out _, new StandardFormat('O')))
            {
                _destination[19] = (byte)'Z';
                _destination = _destination[20..];
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(OADate oaDate)
        {
            if (oaDate.TryFormat(_destination, out var bytesWritten))
            {
                _destination = _destination[bytesWritten..];
                return true;
            }

            return Fail();
        }

        public bool AppendFormatted(SimpleSingleCellReference reference)
        {
            if (!SpreadsheetUtility.TryGetColumnNameUtf8(reference.Column, _destination, out var nameLength))
                return Fail();

            _destination = _destination[nameLength..];

            if (!AppendFormatted(reference.Row))
                return Fail();

            return true;
        }

        public bool AppendFormatted(BooleanAttribute attribute)
        {
            return Formatter.TryFormat(attribute, ref _destination) || Fail();
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

        public bool AppendFormatted(string? value) => AppendFormatted(value.AsSpan());

        public bool AppendFormatted(scoped ReadOnlySpan<char> value)
        {
            if (value.IsEmpty)
                return true;

            if (_destination.Length > value.Length &&
                XmlUtility.TryXmlEncodeToUtf8(value, _destination, out _, out var bytesWritten))
            {
                _destination = _destination[bytesWritten..];
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
            _isSuccess = false;
            return false;
        }
    }

    [InterpolatedStringHandler]
#pragma warning disable CS9113 // Parameter is unread.
    public ref struct ResumableTryWriteInterpolatedStringHandler
#pragma warning restore CS9113 // Parameter is unread.
    {
        private readonly int _initialLength;
        private readonly int _startingStep;
        private int _step;
        private int _index;
        internal bool _isSuccess = true;
        private Span<byte> _destination;

        public readonly int Written => _initialLength - _destination.Length;

        public ResumableTryWriteInterpolatedStringHandler(int _, int _2, SpreadsheetBuffer buffer, BufferWriteProgress start)
        {
            _destination = buffer.GetSpan();
            _initialLength = _destination.Length;
            _startingStep = start.Step;
            _index = start.Index;
        }

        public readonly BufferWriteProgress GetProgress() => new()
        {
            Step = _step - 1,
            Index = _index
        };

        [ExcludeFromCodeCoverage]
        public readonly bool AppendLiteral(string value)
        {
            _ = _isSuccess;
            _ = value;
            throw new InvalidOperationException("Use ReadOnlySpan<byte> instead of string literals");
        }

        public bool AppendFormatted(int value)
        {
            if (_step++ < _startingStep)
                return true;

            return _isSuccess = Formatter.TryFormat(value, ref _destination);
        }

        public bool AppendFormatted(scoped ReadOnlySpan<byte> value)
        {
            if (_step++ < _startingStep)
                return true;

            return _isSuccess = Formatter.TryFormat(value, ref _destination);
        }

        public bool AppendFormatted(BooleanAttribute attribute)
        {
            if (_step++ < _startingStep)
                return true;

            return _isSuccess = Formatter.TryFormat(attribute, ref _destination);
        }

        public bool AppendFormatted(SimpleSingleCellReference reference)
        {
            if (_step++ < _startingStep)
                return true;

            if (!SpreadsheetUtility.TryGetColumnNameUtf8(reference.Column, _destination, out var nameLength))
                return _isSuccess = false;

            var span = _destination[nameLength..];
            if (!Utf8Formatter.TryFormat(reference.Row, span, out var rowLength))
                return _isSuccess = false;

            _destination = span[rowLength..];
            return true;
        }

        public bool AppendFormatted(string? value) => AppendFormatted(value.AsSpan());

        public bool AppendFormatted(scoped ReadOnlySpan<char> value)
        {
            if (_step++ < _startingStep)
                return true;

            var remaining = value[_index..];
            _isSuccess = XmlUtility.TryXmlEncodeToUtf8(remaining, _destination, out var charsRead, out var bytesWritten);
            _destination = _destination[bytesWritten..];
            _index += charsRead;
            return _isSuccess;
        }
    }
}

file static class Formatter
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryFormat(int value, ref Span<byte> destination)
    {
#if NET8_0_OR_GREATER
        var success = value.TryFormat(destination, out var bytesWritten, provider: NumberFormatInfo.InvariantInfo);
#else
        var success = Utf8Formatter.TryFormat(value, destination, out var bytesWritten);
#endif
        destination = destination[bytesWritten..];
        return success;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryFormat(scoped ReadOnlySpan<byte> utf8Value, ref Span<byte> destination)
    {
        if (utf8Value.TryCopyTo(destination))
        {
            destination = destination[utf8Value.Length..];
            return true;
        }
        return false;
    }

    public static bool TryFormat(BooleanAttribute attribute, ref Span<byte> destination)
    {
        if (attribute.Value is not { } value)
            return true;

        var name = attribute.AttributeName;
        if (destination.Length < name.Length + 5)
            return false;

        destination[0] = (byte)' ';
        name.TryCopyTo(destination[1..]);
        destination[name.Length + 1] = (byte)'=';
        destination[name.Length + 2] = (byte)'"';
        destination[name.Length + 3] = (byte)(value ? '1' : '0');
        destination[name.Length + 4] = (byte)'"';

        destination = destination[(name.Length + 5)..];
        return true;
    }
}