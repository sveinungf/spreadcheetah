namespace SpreadCheetah.Test.Helpers;

internal class WriteOnlyMemoryStream : MemoryStream
{
    public override bool CanRead => false;
    public override bool CanSeek => false;

    public override long Position
    {
        get => base.Position;
        set => throw new NotImplementedException();
    }

    public override IAsyncResult BeginRead(byte[] buffer, int offset, int count, AsyncCallback? callback, object? state)
    {
        throw new NotImplementedException();
    }

#if NETCOREAPP
        public override void CopyTo(Stream destination, int bufferSize)
        {
            throw new NotImplementedException();
        }
#endif

    public override Task CopyToAsync(Stream destination, int bufferSize, CancellationToken cancellationToken)
    {
        throw new NotImplementedException();
    }

    public override int EndRead(IAsyncResult asyncResult)
    {
        throw new NotImplementedException();
    }

    public override byte[] GetBuffer()
    {
        throw new NotImplementedException();
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        throw new NotImplementedException();
    }

#if NETCOREAPP
        public override int Read(Span<byte> destination)
        {
            throw new NotImplementedException();
        }

        public override ValueTask<int> ReadAsync(Memory<byte> destination, CancellationToken cancellationToken = new CancellationToken())
        {
            throw new NotImplementedException();
        }
#endif

    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
    {
        throw new NotImplementedException();
    }

    public override int ReadByte()
    {
        throw new NotImplementedException();
    }

    public override long Seek(long offset, SeekOrigin loc)
    {
        throw new NotImplementedException();
    }

    public override bool TryGetBuffer(out ArraySegment<byte> buffer)
    {
        throw new NotImplementedException();
    }

    public override byte[] ToArray()
    {
        throw new NotImplementedException();
    }
}
