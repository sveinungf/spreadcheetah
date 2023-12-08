namespace SpreadCheetah.Test.Helpers;

internal class WriteOnlyMemoryStream : MemoryStream
{
    public override bool CanRead => false;
    public override bool CanSeek => false;

    public override long Position
    {
        get => base.Position;
        set => throw new InvalidOperationException();
    }

    public override IAsyncResult BeginRead(byte[] buffer, int offset, int count, AsyncCallback? callback, object? state)
    {
        throw new InvalidOperationException();
    }

#if NETCOREAPP
    public override void CopyTo(Stream destination, int bufferSize)
    {
        throw new InvalidOperationException();
    }
#endif

    public override Task CopyToAsync(Stream destination, int bufferSize, CancellationToken cancellationToken)
    {
        throw new InvalidOperationException();
    }

    public override int EndRead(IAsyncResult asyncResult)
    {
        throw new InvalidOperationException();
    }

    public override byte[] GetBuffer()
    {
        throw new InvalidOperationException();
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        throw new InvalidOperationException();
    }

#if NETCOREAPP
    public override int Read(Span<byte> destination)
    {
        throw new InvalidOperationException();
    }

    public override ValueTask<int> ReadAsync(Memory<byte> destination, CancellationToken cancellationToken = new CancellationToken())
    {
        throw new InvalidOperationException();
    }
#endif

    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
    {
        throw new InvalidOperationException();
    }

    public override int ReadByte()
    {
        throw new InvalidOperationException();
    }

    public override long Seek(long offset, SeekOrigin loc)
    {
        throw new InvalidOperationException();
    }

    public override bool TryGetBuffer(out ArraySegment<byte> buffer)
    {
        throw new InvalidOperationException();
    }

    public override byte[] ToArray()
    {
        throw new InvalidOperationException();
    }
}
