using SpreadCheetah.Helpers;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadCheetah
{
    internal sealed class SpreadsheetBuffer
    {
        private readonly byte[] _buffer;

        public int Index { get; set; }

        public SpreadsheetBuffer(byte[] buffer)
        {
            _buffer = buffer;
        }

        public Span<byte> GetNextSpan() => _buffer.AsSpan(Index);
        public int GetRemainingBuffer() => _buffer.Length - Index;

        public async ValueTask WriteAsciiStringAsync(string value, Stream stream, CancellationToken token)
        {
            if (value.Length > GetRemainingBuffer())
                await FlushToStreamAsync(stream, token).ConfigureAwait(false);

            Index += Utf8Helper.GetBytes(value, GetNextSpan());
        }

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
    }
}
