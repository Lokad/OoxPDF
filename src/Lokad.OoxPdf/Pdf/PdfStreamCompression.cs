using System.IO.Compression;
using System.Runtime.ExceptionServices;

namespace Lokad.OoxPdf.Pdf;

internal static class PdfStreamCompression
{
    // The caller owns the aggregate reservation, including any earlier payload.
    // Success retains only the result; failure restores the earlier payload.
    public static byte[] Compress(ReadOnlySpan<byte> bytes, IMemoryReservation? storage,
        long retainedBytes, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        using var output = new CompressionBuffer(storage, retainedBytes, cancellationToken);
        using (var zlib = new ZLibStream(output, CompressionLevel.SmallestSize, leaveOpen: true))
        {
            zlib.Write(bytes);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return output.Complete();
    }

    private sealed class CompressionBuffer(IMemoryReservation? storage, long retainedBytes,
        CancellationToken cancellationToken) : Stream
    {
        private MemoryStream? buffer = new();
        private ExceptionDispatchInfo? writeFailure;
        private MemoryStream Active => buffer ?? throw new ObjectDisposedException(nameof(CompressionBuffer));
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => buffer is not null;
        public override long Length => Active.Length;
        public override long Position { get => Active.Position; set => throw new NotSupportedException(); }

        public override void Write(ReadOnlySpan<byte> bytes)
        {
            // Deflater disposal can flush after a failed write. Preserve the
            // original failure without another admission request or data write.
            writeFailure?.Throw();
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                MemoryStream output = Active;
                int required = checked((int)(output.Length + bytes.Length));
                int oldCapacity = output.Capacity;
                if (required > oldCapacity)
                {
                    int capacity = checked((int)Math.Min(int.MaxValue,
                        Math.Max(required, Math.Max(256L, 2L * oldCapacity))));
                    OoxMemory.Resize(storage, checked(retainedBytes + oldCapacity + capacity));
                    cancellationToken.ThrowIfCancellationRequested();
                    output.Capacity = capacity;
                    OoxMemory.Resize(storage, checked(retainedBytes + capacity));
                }
                cancellationToken.ThrowIfCancellationRequested();
                output.Write(bytes);
            }
            catch (Exception ex)
            {
                writeFailure = ExceptionDispatchInfo.Capture(ex);
                throw;
            }
        }

        public override void Write(byte[] bytes, int offset, int count) => Write(bytes.AsSpan(offset, count));
        public override void WriteByte(byte value) => Write([value]);
        public override void Flush()
        {
            writeFailure?.Throw();
            cancellationToken.ThrowIfCancellationRequested();
            Active.Flush();
        }

        public byte[] Complete()
        {
            MemoryStream output = Active;
            int length = checked((int)output.Length);
            OoxMemory.Resize(storage, checked(retainedBytes + output.Capacity + length));
            cancellationToken.ThrowIfCancellationRequested();
            byte[] result = output.ToArray();
            cancellationToken.ThrowIfCancellationRequested();
            output.Dispose();
            buffer = null;
            OoxMemory.Resize(storage, checked(retainedBytes + length));
            return result;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && buffer is { } output)
            {
                output.Dispose();
                buffer = null;
                OoxMemory.Resize(storage, retainedBytes);
            }
            base.Dispose(disposing);
        }

        public override int Read(byte[] bytes, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }
}
