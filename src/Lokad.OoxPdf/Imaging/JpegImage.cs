namespace Lokad.OoxPdf.Imaging;

internal sealed class JpegImage
{
    private static readonly int[] ZigZag =
    [
        0, 1, 8, 16, 9, 2, 3, 10,
        17, 24, 32, 25, 18, 11, 4, 5,
        12, 19, 26, 33, 40, 48, 41, 34,
        27, 20, 13, 6, 7, 14, 21, 28,
        35, 42, 49, 56, 57, 50, 43, 36,
        29, 22, 15, 23, 30, 37, 44, 51,
        58, 59, 52, 45, 38, 31, 39, 46,
        53, 60, 61, 54, 47, 55, 62, 63
    ];

    private readonly record struct HuffmanEntry(int Code, int Length, int Symbol);

    private sealed class Component
    {
        public int Id;
        public int Horizontal;
        public int Vertical;
        public int QuantizationTableId;
        public int DcTableId;
        public int AcTableId;
        public int DcPredictor;
        public int SampleWidth;
        public int SampleHeight;
        public byte[] Samples = [];
    }

    private JpegImage(int width, int height, byte[] rgb)
    {
        Width = width;
        Height = height;
        Rgb = rgb;
    }

    public int Width { get; }

    public int Height { get; }

    public byte[] Rgb { get; }

    // Read keeps release-on-return lifetime for header probes, tests, and callers
    // without downstream transforms. Production image paths use ReadOwned and hold
    // the reservation across crop/recolor/compress work (R02).
    public static JpegImage Read(byte[] bytes, CancellationToken cancellationToken = default)
    {
        using DecodedPixels owned = ReadOwned(bytes, cancellationToken);
        return new JpegImage(owned.Width, owned.Height, owned.Rgb);
    }
    public static DecodedPixels ReadOwned(byte[] bytes, CancellationToken cancellationToken = default)
    {
        var decoder = new Decoder(bytes, cancellationToken);
        (JpegImage image, OoxConversionBudget.LiveReservation? reservation) = decoder.Decode();
        return new DecodedPixels(image.Width, image.Height, image.Rgb, null, reservation);
    }

    private sealed class Decoder
    {
        private readonly byte[] bytes;
        private readonly CancellationToken cancellationToken;
        private readonly int[][] quantizationTables = new int[4][];
        private readonly HuffmanEntry[][] dcTables = new HuffmanEntry[4][];
        private readonly HuffmanEntry[][] acTables = new HuffmanEntry[4][];
        private readonly List<Component> components = [];
        private int width;
        private int height;
        private int maxHorizontal = 1;
        private int maxVertical = 1;

        public Decoder(byte[] bytes, CancellationToken cancellationToken = default)
        {
            this.bytes = bytes;
            this.cancellationToken = cancellationToken;
        }

        // R02: the scan reservation (sample planes plus output pixels) spans the
        // entropy decode and RGB construction, then transfers to the caller via
        // ReadOwned. It releases here only on failure; the SOS-less path keeps the
        // legacy failure behavior below.
        public (JpegImage Image, OoxConversionBudget.LiveReservation? Reservation) Decode()
        {
            if (bytes.Length < 4 || bytes[0] != 0xFF || bytes[1] != 0xD8)
            {
                throw new InvalidDataException("Data is not a JPEG image.");
            }
            int offset = 2;
            OoxConversionBudget.LiveReservation? scanReservation = null;
            while (offset < bytes.Length)
            {
                byte marker = ReadMarker(ref offset);
                if (marker == 0xD9)
                {
                    break;
                }
                if (marker == 0xDA)
                {
                    DecodeScan(ref offset, out scanReservation);
                    break;
                }
                int length = ReadSegmentLength(offset);
                ReadOnlySpan<byte> segment = bytes.AsSpan(offset + 2, length - 2);
                switch (marker)
                {
                    case 0xDB:
                        ReadQuantizationTables(segment);
                        break;
                    case 0xC0:
                        ReadStartOfFrame(segment);
                        break;
                    case 0xC4:
                        ReadHuffmanTables(segment);
                        break;
                }
                offset += length;
            }
            if (width <= 0 || height <= 0 || components.Count is not (1 or 3))
            {
                scanReservation?.Dispose();
                throw new InvalidDataException("JPEG frame metadata is incomplete.");
            }
            ImagePixelBudget.Check(width, height, "JPEG");
            try
            {
                return (new JpegImage(width, height, BuildRgb()), scanReservation);
            }
            catch
            {
                scanReservation?.Dispose();
                throw;
            }
        }

        private byte ReadMarker(ref int offset)
        {
            while (offset < bytes.Length && bytes[offset] != 0xFF)
            {
                offset++;
            }

            while (offset < bytes.Length && bytes[offset] == 0xFF)
            {
                offset++;
            }

            if (offset >= bytes.Length)
            {
                throw new InvalidDataException("JPEG marker is truncated.");
            }

            return bytes[offset++];
        }

        private int ReadSegmentLength(int offset)
        {
            if (offset + 2 > bytes.Length)
            {
                throw new InvalidDataException("JPEG segment length is truncated.");
            }

            int length = (bytes[offset] << 8) | bytes[offset + 1];
            if (length < 2 || offset + length > bytes.Length)
            {
                throw new InvalidDataException("JPEG segment length is invalid.");
            }

            return length;
        }

        private void ReadQuantizationTables(ReadOnlySpan<byte> segment)
        {
            int offset = 0;
            while (offset < segment.Length)
            {
                byte spec = segment[offset++];
                int precision = spec >> 4;
                int tableId = spec & 0x0F;
                if (precision != 0 || tableId >= quantizationTables.Length || offset + 64 > segment.Length)
                {
                    throw new NotSupportedException("Only 8-bit JPEG quantization tables are supported.");
                }

                int[] table = new int[64];
                for (int i = 0; i < 64; i++)
                {
                    table[ZigZag[i]] = segment[offset++];
                }

                quantizationTables[tableId] = table;
            }
        }

        private void ReadStartOfFrame(ReadOnlySpan<byte> segment)
        {
            if (segment.Length < 6 || segment[0] != 8)
            {
                throw new NotSupportedException("Only 8-bit baseline JPEG frames are supported.");
            }

            height = (segment[1] << 8) | segment[2];
            width = (segment[3] << 8) | segment[4];
            int count = segment[5];
            if (count is not (1 or 3) || segment.Length < 6 + count * 3)
            {
                throw new NotSupportedException($"Unsupported JPEG component count {count}.");
            }

            components.Clear();
            maxHorizontal = 1;
            maxVertical = 1;
            int offset = 6;
            for (int i = 0; i < count; i++)
            {
                int sampling = segment[offset + 1];
                var component = new Component
                {
                    Id = segment[offset],
                    Horizontal = sampling >> 4,
                    Vertical = sampling & 0x0F,
                    QuantizationTableId = segment[offset + 2]
                };
                if (component.Horizontal is < 1 or > 4 || component.Vertical is < 1 or > 4)
                {
                    throw new NotSupportedException("Unsupported JPEG sampling factors.");
                }

                maxHorizontal = Math.Max(maxHorizontal, component.Horizontal);
                maxVertical = Math.Max(maxVertical, component.Vertical);
                components.Add(component);
                offset += 3;
            }

            ImagePixelBudget.Check(width, height, "JPEG");
        }

        private void ReadHuffmanTables(ReadOnlySpan<byte> segment)
        {
            int offset = 0;
            while (offset < segment.Length)
            {
                byte spec = segment[offset++];
                int tableClass = spec >> 4;
                int tableId = spec & 0x0F;
                if (tableId >= 4 || offset + 16 > segment.Length)
                {
                    throw new InvalidDataException("JPEG Huffman table is invalid.");
                }

                byte[] counts = segment.Slice(offset, 16).ToArray();
                offset += 16;
                int symbolCount = 0;
                for (int i = 0; i < counts.Length; i++)
                {
                    symbolCount += counts[i];
                }

                if (offset + symbolCount > segment.Length)
                {
                    throw new InvalidDataException("JPEG Huffman symbols are truncated.");
                }

                var entries = new List<HuffmanEntry>(symbolCount);
                int code = 0;
                for (int length = 1; length <= 16; length++)
                {
                    for (int i = 0; i < counts[length - 1]; i++)
                    {
                        entries.Add(new HuffmanEntry(code, length, segment[offset++]));
                        code++;
                    }

                    code <<= 1;
                }

                if (tableClass == 0)
                {
                    dcTables[tableId] = entries.ToArray();
                }
                else if (tableClass == 1)
                {
                    acTables[tableId] = entries.ToArray();
                }
                else
                {
                    throw new InvalidDataException("JPEG Huffman table class is invalid.");
                }
            }
        }

        private void DecodeScan(ref int offset, out OoxConversionBudget.LiveReservation? reservation)
        {
            reservation = null;
            int length = ReadSegmentLength(offset);
            ReadOnlySpan<byte> segment = bytes.AsSpan(offset + 2, length - 2);
            if (segment.Length < 4)
            {
                throw new InvalidDataException("JPEG scan header is truncated.");
            }
            int selectorCount = segment[0];
            if (selectorCount != components.Count || segment.Length < 1 + selectorCount * 2 + 3)
            {
                throw new NotSupportedException("Only single-scan baseline JPEG images are supported.");
            }
            int segmentOffset = 1;
            for (int i = 0; i < selectorCount; i++)
            {
                int componentId = segment[segmentOffset++];
                Component component = components.FirstOrDefault(item => item.Id == componentId) ??
                    throw new InvalidDataException("JPEG scan references an unknown component.");
                int tableSpec = segment[segmentOffset++];
                component.DcTableId = tableSpec >> 4;
                component.AcTableId = tableSpec & 0x0F;
            }
            offset += length;
            if (width <= 0 || height <= 0 || components.Count is not (1 or 3))
            {
                throw new InvalidDataException("JPEG frame metadata is incomplete.");
            }
            ImagePixelBudget.Check(width, height, "JPEG");
            // R02: reserve sample planes plus output pixels together before allocating
            // either; the reservation transfers to the caller through Decode and releases
            // only after crop/recolor/compress work consumes the pixels.
            long planeBytes = MeasureComponentPlanes();
            long liveEstimate = checked(planeBytes + checked((long)width * height * 3L));
            OoxConversionBudget.LiveReservation? live = OoxConversionBudget.Current?.ReserveLiveImageBytes(liveEstimate);
            try
            {
                AllocateComponentBuffers();
                var reader = new EntropyReader(bytes, offset);
                int mcuColumns = (width + maxHorizontal * 8 - 1) / (maxHorizontal * 8);
                int mcuRows = (height + maxVertical * 8 - 1) / (maxVertical * 8);
                for (int mcuY = 0; mcuY < mcuRows; mcuY++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    for (int mcuX = 0; mcuX < mcuColumns; mcuX++)
                    {
                        foreach (Component component in components)
                        {
                            for (int v = 0; v < component.Vertical; v++)
                            {
                                for (int h = 0; h < component.Horizontal; h++)
                                {
                                    DecodeBlock(reader, component, (mcuX * component.Horizontal + h) * 8, (mcuY * component.Vertical + v) * 8);
                                }
                            }
                        }
                    }
                }
                reservation = live;
            }
            catch
            {
                live?.Dispose();
                throw;
            }
        }

        private long MeasureComponentPlanes()
        {
            cancellationToken.ThrowIfCancellationRequested();
            int mcuUnitX;
            int mcuUnitY;
            try
            {
                mcuUnitX = checked(maxHorizontal * 8);
                mcuUnitY = checked(maxVertical * 8);
            }
            catch (OverflowException ex)
            {
                throw new InvalidDataException("JPEG sampling factors overflow.", ex);
            }
            if (mcuUnitX <= 0 || mcuUnitY <= 0)
            {
                throw new InvalidDataException("JPEG sampling factors are invalid.");
            }
            int mcuColumns;
            int mcuRows;
            try
            {
                mcuColumns = checked((checked(width + mcuUnitX - 1) / mcuUnitX));
                mcuRows = checked((checked(height + mcuUnitY - 1) / mcuUnitY));
            }
            catch (OverflowException ex)
            {
                throw new InvalidDataException("JPEG dimensions overflow.", ex);
            }
            long totalBytes = 0;
            foreach (Component component in components)
            {
                int sampleWidth;
                int sampleHeight;
                try
                {
                    sampleWidth = checked(checked(mcuColumns * component.Horizontal) * 8);
                    sampleHeight = checked(checked(mcuRows * component.Vertical) * 8);
                }
                catch (OverflowException ex)
                {
                    throw new InvalidDataException("JPEG component dimensions overflow.", ex);
                }
                long planeBytes = checked((long)sampleWidth * sampleHeight);
                if (planeBytes > int.MaxValue)
                {
                    throw new InvalidDataException("JPEG component plane is too large.");
                }
                component.SampleWidth = sampleWidth;
                component.SampleHeight = sampleHeight;
                component.DcPredictor = 0;
                totalBytes = checked(totalBytes + planeBytes);
            }
            return totalBytes;

        }

        private void AllocateComponentBuffers()
        {
            foreach (Component component in components)
            {
                component.Samples = new byte[checked(component.SampleWidth * component.SampleHeight)];
            }
        }

        private void DecodeBlock(EntropyReader reader, Component component, int blockX, int blockY)
        {
            int[] quantization = quantizationTables[component.QuantizationTableId] ??
                throw new InvalidDataException("JPEG scan references a missing quantization table.");
            HuffmanEntry[] dcTable = dcTables[component.DcTableId] ??
                throw new InvalidDataException("JPEG scan references a missing DC Huffman table.");
            HuffmanEntry[] acTable = acTables[component.AcTableId] ??
                throw new InvalidDataException("JPEG scan references a missing AC Huffman table.");

            var coefficients = new int[64];
            int dcSize = DecodeHuffmanSymbol(reader, dcTable);
            component.DcPredictor += ReceiveExtended(reader, dcSize);
            coefficients[0] = component.DcPredictor * quantization[0];

            int index = 1;
            while (index < 64)
            {
                int symbol = DecodeHuffmanSymbol(reader, acTable);
                if (symbol == 0)
                {
                    break;
                }

                if (symbol == 0xF0)
                {
                    index += 16;
                    continue;
                }

                int zeroRun = symbol >> 4;
                int size = symbol & 0x0F;
                index += zeroRun;
                if (index >= 64)
                {
                    break;
                }

                int coefficientIndex = ZigZag[index++];
                coefficients[coefficientIndex] = ReceiveExtended(reader, size) * quantization[coefficientIndex];
            }

            WriteBlockSamples(component, blockX, blockY, coefficients);
        }

        private static int DecodeHuffmanSymbol(EntropyReader reader, HuffmanEntry[] table)
        {
            int code = 0;
            for (int length = 1; length <= 16; length++)
            {
                code = (code << 1) | reader.ReadBit();
                foreach (HuffmanEntry entry in table)
                {
                    if (entry.Length == length && entry.Code == code)
                    {
                        return entry.Symbol;
                    }
                }
            }

            throw new InvalidDataException("JPEG Huffman code was not found.");
        }

        private static int ReceiveExtended(EntropyReader reader, int size)
        {
            if (size == 0)
            {
                return 0;
            }

            int value = 0;
            for (int i = 0; i < size; i++)
            {
                value = (value << 1) | reader.ReadBit();
            }

            int threshold = 1 << (size - 1);
            return value < threshold
                ? value - ((1 << size) - 1)
                : value;
        }

        private static void WriteBlockSamples(Component component, int blockX, int blockY, int[] coefficients)
        {
            for (int y = 0; y < 8; y++)
            {
                int targetY = blockY + y;
                if (targetY >= component.SampleHeight)
                {
                    continue;
                }

                for (int x = 0; x < 8; x++)
                {
                    int targetX = blockX + x;
                    if (targetX >= component.SampleWidth)
                    {
                        continue;
                    }

                    component.Samples[targetY * component.SampleWidth + targetX] = ToByte(InverseDct(x, y) + 128d);
                }
            }

            double InverseDct(int x, int y)
            {
                double sum = 0d;
                for (int v = 0; v < 8; v++)
                {
                    double cv = v == 0 ? Math.Sqrt(0.5d) : 1d;
                    double yCos = Math.Cos((2 * y + 1) * v * Math.PI / 16d);
                    for (int u = 0; u < 8; u++)
                    {
                        double cu = u == 0 ? Math.Sqrt(0.5d) : 1d;
                        sum += cu * cv * coefficients[v * 8 + u] *
                            Math.Cos((2 * x + 1) * u * Math.PI / 16d) *
                            yCos;
                    }
                }

                return sum / 4d;
            }
        }

        private byte[] BuildRgb()
        {
            long rgbBytes = checked((long)width * height * 3L);
            if (rgbBytes > int.MaxValue)
            {
                throw new InvalidDataException("JPEG pixel plane is too large.");
            }

            var rgb = new byte[(int)rgbBytes];
            if (components.Count == 1)
            {
                Component gray = components[0];
                int target = 0;
                for (int y = 0; y < height; y++)
                {
                    if ((y & 255) == 0)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                    }

                    for (int x = 0; x < width; x++)
                    {
                        byte value = Sample(gray, x, y);
                        rgb[target++] = value;
                        rgb[target++] = value;
                        rgb[target++] = value;
                    }
                }

                return rgb;
            }

            Component yComponent = components[0];
            Component cbComponent = components[1];
            Component crComponent = components[2];
            int offset = 0;
            for (int y = 0; y < height; y++)
            {
                if ((y & 255) == 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                }

                for (int x = 0; x < width; x++)
                {
                    double luma = Sample(yComponent, x, y);
                    double cb = Sample(cbComponent, x, y) - 128d;
                    double cr = Sample(crComponent, x, y) - 128d;
                    rgb[offset++] = ToByte(luma + 1.402d * cr);
                    rgb[offset++] = ToByte(luma - 0.344136d * cb - 0.714136d * cr);
                    rgb[offset++] = ToByte(luma + 1.772d * cb);
                }
            }

            return rgb;
        }

        private byte Sample(Component component, int x, int y)
        {
            int sampleX = Math.Min(component.SampleWidth - 1, x * component.Horizontal / maxHorizontal);
            int sampleY = Math.Min(component.SampleHeight - 1, y * component.Vertical / maxVertical);
            return component.Samples[sampleY * component.SampleWidth + sampleX];
        }
    }

    private sealed class EntropyReader
    {
        private readonly byte[] bytes;
        private int offset;
        private int bitBuffer;
        private int bitCount;

        public EntropyReader(byte[] bytes, int offset)
        {
            this.bytes = bytes;
            this.offset = offset;
        }

        public int ReadBit()
        {
            if (bitCount == 0)
            {
                bitBuffer = ReadEntropyByte();
                bitCount = 8;
            }

            bitCount--;
            return (bitBuffer >> bitCount) & 1;
        }

        private int ReadEntropyByte()
        {
            if (offset >= bytes.Length)
            {
                throw new InvalidDataException("JPEG entropy stream is truncated.");
            }

            int value = bytes[offset++];
            if (value != 0xFF)
            {
                return value;
            }

            while (offset < bytes.Length && bytes[offset] == 0xFF)
            {
                offset++;
            }

            if (offset >= bytes.Length)
            {
                throw new InvalidDataException("JPEG entropy marker is truncated.");
            }

            int marker = bytes[offset++];
            if (marker == 0x00)
            {
                return 0xFF;
            }

            if (marker is >= 0xD0 and <= 0xD7)
            {
                return ReadEntropyByte();
            }

            throw new InvalidDataException($"Unexpected JPEG entropy marker 0xFF{marker:X2}.");
        }
    }

    private static byte ToByte(double value)
    {
        return (byte)Math.Clamp((int)Math.Round(value), 0, 255);
    }
}
