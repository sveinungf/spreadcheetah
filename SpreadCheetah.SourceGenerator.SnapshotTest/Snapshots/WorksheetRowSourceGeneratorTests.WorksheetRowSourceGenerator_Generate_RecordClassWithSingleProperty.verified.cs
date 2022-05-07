//HintName: MyNamespace.MyGenRowContext.g.cs
// <auto-generated />
#nullable enable
using SpreadCheetah;
using SpreadCheetah.SourceGeneration;
using System;
using System.Buffers;
using System.Threading;
using System.Threading.Tasks;

namespace MyNamespace
{
    public partial class MyGenRowContext
    {
        private static MyGenRowContext? _default;
        public static MyGenRowContext Default => _default ??= new MyGenRowContext();

        public MyGenRowContext()
        {
        }

        private WorksheetRowTypeInfo<SpreadCheetah.SourceGenerator.SnapshotTest.Models.RecordClassWithSingleProperty>? _RecordClassWithSingleProperty;
        public WorksheetRowTypeInfo<SpreadCheetah.SourceGenerator.SnapshotTest.Models.RecordClassWithSingleProperty> RecordClassWithSingleProperty => _RecordClassWithSingleProperty ??= WorksheetRowMetadataServices.CreateObjectInfo<SpreadCheetah.SourceGenerator.SnapshotTest.Models.RecordClassWithSingleProperty>(AddAsRowAsync);

        private static ValueTask AddAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, SpreadCheetah.SourceGenerator.SnapshotTest.Models.RecordClassWithSingleProperty? obj, CancellationToken token)
        {
            if (spreadsheet is null)
                throw new ArgumentNullException(nameof(spreadsheet));
            if (obj is null)
                return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);
            return AddAsRowInternalAsync(spreadsheet, obj, token);
        }

        private static async ValueTask AddAsRowInternalAsync(SpreadCheetah.Spreadsheet spreadsheet, SpreadCheetah.SourceGenerator.SnapshotTest.Models.RecordClassWithSingleProperty obj, CancellationToken token)
        {
            var cells = ArrayPool<DataCell>.Shared.Rent(1);
            try
            {
                cells[0] = new DataCell(obj.Value);
                await spreadsheet.AddRowAsync(cells.AsMemory(0, 1), token).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<DataCell>.Shared.Return(cells, true);
            }
        }
    }
}