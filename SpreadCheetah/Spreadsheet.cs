using SpreadCheetah.MetadataXml;
using SpreadCheetah.Styling;
using SpreadCheetah.Worksheets;
using System;
using System.Buffers;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadCheetah
{
    public sealed class Spreadsheet : IDisposable
    {
        // Invalid worksheet name characters in Excel
        private const string InvalidSheetNameCharString = "/\\*?[]";
        private static readonly char[] InvalidSheetNameChars = InvalidSheetNameCharString.ToCharArray();

        private readonly List<string> _worksheetNames = new List<string>();
        private readonly List<string> _worksheetPaths = new List<string>();
        private readonly ZipArchive _archive;
        private readonly CompressionLevel _compressionLevel;
        private readonly SpreadsheetBuffer _buffer;
        private readonly byte[] _arrayPoolBuffer;
        private List<Style>? _styles;
        private Worksheet? _worksheet;
        private bool _disposed;

        private Spreadsheet(ZipArchive archive, CompressionLevel compressionLevel, int bufferSize)
        {
            _archive = archive;
            _compressionLevel = compressionLevel;
            _arrayPoolBuffer = ArrayPool<byte>.Shared.Rent(bufferSize);
            _buffer = new SpreadsheetBuffer(_arrayPoolBuffer);
        }

        public static async ValueTask<Spreadsheet> CreateNewAsync(
            Stream stream,
            SpreadCheetahOptions? options = null,
            CancellationToken cancellationToken = default)
        {
            var archive = new ZipArchive(stream, ZipArchiveMode.Create, true);
            options ??= new SpreadCheetahOptions();
            var spreadsheet = new Spreadsheet(archive, GetCompressionLevel(options.CompressionLevel), options.BufferSize);
            await spreadsheet.InitializeAsync(cancellationToken).ConfigureAwait(false);
            return spreadsheet;
        }

        private static CompressionLevel GetCompressionLevel(SpreadCheetahCompressionLevel level)
        {
            return level == SpreadCheetahCompressionLevel.Optimal ? CompressionLevel.Optimal : CompressionLevel.Fastest;
        }

        private ValueTask InitializeAsync(CancellationToken token)
        {
            return RelationshipsXml.WriteAsync(_archive, _compressionLevel, _buffer, token);
        }

        public ValueTask StartWorksheetAsync(string name, WorksheetOptions? options = null, CancellationToken token = default)
        {
            if (name is null)
                throw new ArgumentNullException(nameof(name));

            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("The name can not be empty or consist only of whitespace.", nameof(name));

            if (name.Length > 31)
                throw new ArgumentException("The name can not be more than 31 characters.", nameof(name));

            if (name.StartsWith("\'", StringComparison.Ordinal) || name.EndsWith("\'", StringComparison.Ordinal))
                throw new ArgumentException("The name can not start or end with a single quote.", nameof(name));

            if (name.IndexOfAny(InvalidSheetNameChars) != -1)
                throw new ArgumentException("The name can not contain any of the following characters: " + InvalidSheetNameCharString, nameof(name));

            if (_worksheetNames.Contains(name))
                throw new ArgumentException("A worksheet with the given name already exists.", nameof(name));

            return StartWorksheetInternalAsync(name, options, token);
        }

        private async ValueTask StartWorksheetInternalAsync(string name, WorksheetOptions? options, CancellationToken token)
        {
            await FinishAndDisposeWorksheetAsync(token).ConfigureAwait(false);

            var path = $"xl/worksheets/sheet{_worksheetNames.Count + 1}.xml";
            var entry = _archive.CreateEntry(path, _compressionLevel);
            var entryStream = entry.Open();
            _worksheet = new Worksheet(entryStream, _buffer);
            await _worksheet.WriteHeadAsync(options, token).ConfigureAwait(false);
            _worksheetNames.Add(name);
            _worksheetPaths.Add(path);
        }

        public ValueTask AddRowAsync(IList<Cell> cells, CancellationToken token = default)
        {
            EnsureCanAddRows(cells);
            return _worksheet!.TryAddRow(cells, out var currentIndex)
                ? default
                : _worksheet.AddRowAsync(cells, currentIndex, token);
        }

        public ValueTask AddRowAsync(IList<StyledCell> cells, CancellationToken token = default)
        {
            EnsureCanAddRows(cells);
            return _worksheet!.TryAddRow(cells, out var currentIndex)
                ? default
                : _worksheet.AddRowAsync(cells, currentIndex, token);
        }

        private void EnsureCanAddRows<T>(IList<T> cells)
        {
            if (cells is null)
                throw new ArgumentNullException(nameof(cells));
            if (_worksheet is null)
                throw new SpreadCheetahException("Can't add rows when there is not an active worksheet.");
        }

        public StyleId AddStyle(Style style)
        {
            if (style is null)
                throw new ArgumentNullException(nameof(style));

            if (_styles is null)
                _styles = new List<Style> { style };
            else
                _styles.Add(style);

            return new StyleId(_styles.Count);
        }

        private async ValueTask FinishAndDisposeWorksheetAsync(CancellationToken token)
        {
            if (_worksheet is null) return;
            await _worksheet.FinishAsync(token).ConfigureAwait(false);
            await _worksheet.DisposeAsync().ConfigureAwait(false);
            _worksheet = null;
        }

        public ValueTask FinishAsync(CancellationToken token = default)
        {
            if (_worksheetNames.Count == 0)
                throw new SpreadCheetahException("Spreadsheet must contain at least one worksheet.");

            return FinishInternalAsync(token);
        }

        private async ValueTask FinishInternalAsync(CancellationToken token)
        {
            await FinishAndDisposeWorksheetAsync(token).ConfigureAwait(false);

            var hasStyles = _styles != null;
            await ContentTypesXml.WriteAsync(_archive, _compressionLevel, _buffer, _worksheetPaths, hasStyles, token).ConfigureAwait(false);
            await WorkbookRelsXml.WriteAsync(_archive, _compressionLevel, _buffer, _worksheetPaths, hasStyles, token).ConfigureAwait(false);
            await WorkbookXml.WriteAsync(_archive, _compressionLevel, _buffer, _worksheetNames, token).ConfigureAwait(false);

            if (_styles != null)
                await StylesXml.WriteAsync(_archive, _compressionLevel, _buffer, _styles, token).ConfigureAwait(false);
        }

        public async ValueTask DisposeAsync()
        {
            if (_disposed) return;
            _disposed = true;

            if (_worksheet != null)
                await _worksheet.DisposeAsync().ConfigureAwait(false);

            ArrayPool<byte>.Shared.Return(_arrayPoolBuffer, true);
            _archive?.Dispose();
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _worksheet?.Dispose();
            ArrayPool<byte>.Shared.Return(_arrayPoolBuffer, true);
            _archive?.Dispose();
        }
    }
}
