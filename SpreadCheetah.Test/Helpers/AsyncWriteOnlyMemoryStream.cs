namespace SpreadCheetah.Test.Helpers
{
    internal class AsyncWriteOnlyMemoryStream : WriteOnlyMemoryStream
    {
        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotImplementedException();
        }

#if NETCOREAPP
        public override void Write(ReadOnlySpan<byte> source)
        {
            throw new NotImplementedException();
        }
#endif

        public override IAsyncResult BeginWrite(byte[] buffer, int offset, int count, AsyncCallback? callback, object? state)
        {
            throw new NotImplementedException();
        }

        public override void WriteByte(byte value)
        {
            throw new NotImplementedException();
        }

        public override void WriteTo(Stream stream)
        {
            throw new NotImplementedException();
        }

        public override void EndWrite(IAsyncResult asyncResult)
        {
            throw new NotImplementedException();
        }

        public override void Flush()
        {
            throw new NotImplementedException();
        }
    }
}
