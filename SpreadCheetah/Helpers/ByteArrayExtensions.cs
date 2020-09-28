#if !NETSTANDARD2_0
using System;
#endif
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadCheetah.Helpers
{
    internal static class ByteArrayExtensions
    {
        public static ValueTask FlushToStreamAsync(this byte[] buffer, Stream stream, ref int bufferIndex, CancellationToken token)
        {
            var index = bufferIndex;
            bufferIndex = 0;
#if NETSTANDARD2_0
            return new ValueTask(stream.WriteAsync(buffer, 0, index, token));
#else
            return stream.WriteAsync(buffer.AsMemory(0, index), token);
#endif
        }
    }
}
