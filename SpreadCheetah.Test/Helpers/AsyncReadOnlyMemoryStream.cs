namespace SpreadCheetah.Test.Helpers;

internal class AsyncReadOnlyMemoryStream : Stream
{
    private readonly Stream _backingStream;

    public override bool CanRead => true;
    public override bool CanSeek => false;
    public override bool CanWrite => false;

    public override long Position
    {
        get => throw new InvalidOperationException();
        set => throw new InvalidOperationException();
    }

    public override long Length => throw new InvalidOperationException();

    public AsyncReadOnlyMemoryStream(Stream backingStream)
    {
        _backingStream = backingStream;
    }

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        if (disposing)
            _backingStream.Dispose();
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        throw new InvalidOperationException();
    }

#if NETCOREAPP
    public override int Read(Span<byte> buffer)
    {
        throw new InvalidOperationException();
    }

    public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
    {
        return _backingStream.ReadAsync(buffer, cancellationToken);
    }
#endif

    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
    {
        return _backingStream.ReadAsync(buffer, offset, count, cancellationToken);
    }

    public override int ReadByte()
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
        return _backingStream.CopyToAsync(destination, bufferSize, cancellationToken);
    }

    public override IAsyncResult BeginWrite(byte[] buffer, int offset, int count, AsyncCallback? callback, object? state)
    {
        throw new InvalidOperationException();
    }

    public override void EndWrite(IAsyncResult asyncResult)
    {
        throw new InvalidOperationException();
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        throw new InvalidOperationException();
    }

#if NETCOREAPP
    public override void Write(ReadOnlySpan<byte> buffer)
    {
        throw new InvalidOperationException();
    }

    public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
    {
        throw new InvalidOperationException();
    }
#endif

    public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
    {
        throw new InvalidOperationException();
    }

    public override long Seek(long offset, SeekOrigin loc)
    {
        throw new InvalidOperationException();
    }

    public override void WriteByte(byte value)
    {
        throw new InvalidOperationException();
    }

    public override void Flush()
    {
        throw new InvalidOperationException();
    }

    public override Task FlushAsync(CancellationToken cancellationToken)
    {
        throw new InvalidOperationException();
    }

    public override void SetLength(long value)
    {
        throw new InvalidOperationException();
    }
}
