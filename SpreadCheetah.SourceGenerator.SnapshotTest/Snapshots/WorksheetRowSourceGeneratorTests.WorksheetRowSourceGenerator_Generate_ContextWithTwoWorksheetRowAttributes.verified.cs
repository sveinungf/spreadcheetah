//HintName: MyNamespace.MyGenRowContext.g.cs
// <auto-generated />
#nullable enable
using SpreadCheetah;
using SpreadCheetah.SourceGeneration;
using System.Buffers;
using System.Threading;
using System.Threading.Tasks;

namespace MyNamespace
{
    public partial class MyGenRowContext
    {
        private static MyGenRowContext? _default;
        public static MyGenRowContext Default => _default ??= new();

        public MyGenRowContext()
        {
        }

        private WorksheetRowTypeInfo<SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithSingleProperty>? _ClassWithSingleProperty;
        public WorksheetRowTypeInfo<SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithSingleProperty> ClassWithSingleProperty => _ClassWithSingleProperty ??= WorksheetRowMetadataServices.CreateObjectInfo<SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithSingleProperty>(AddAsRowAsync);

        private static ValueTask AddAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithSingleProperty? obj, CancellationToken token)
        {
            if (spreadsheet is null)
                throw new ArgumentNullException(nameof(spreadsheet));
            if (obj is null)
                return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);
            return AddAsRowInternalAsync(spreadsheet, obj, token);
        }

        private static async ValueTask AddAsRowInternalAsync(SpreadCheetah.Spreadsheet spreadsheet, SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithSingleProperty obj, CancellationToken token)
        {
            var cells = ArrayPool<DataCell>.Shared.Rent(1);
            try
            {
                cells[0] = new DataCell(obj.Name);
                await spreadsheet.AddRowAsync(cells.AsMemory(0, 1), token).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<DataCell>.Shared.Return(cells, true);
            }
        }

        private WorksheetRowTypeInfo<SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithMultipleProperties>? _ClassWithMultipleProperties;
        public WorksheetRowTypeInfo<SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithMultipleProperties> ClassWithMultipleProperties => _ClassWithMultipleProperties ??= WorksheetRowMetadataServices.CreateObjectInfo<SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithMultipleProperties>(AddAsRowAsync);

        private static ValueTask AddAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithMultipleProperties? obj, CancellationToken token)
        {
            if (spreadsheet is null)
                throw new ArgumentNullException(nameof(spreadsheet));
            if (obj is null)
                return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);
            return AddAsRowInternalAsync(spreadsheet, obj, token);
        }

        private static async ValueTask AddAsRowInternalAsync(SpreadCheetah.Spreadsheet spreadsheet, SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithMultipleProperties obj, CancellationToken token)
        {
            var cells = ArrayPool<DataCell>.Shared.Rent(6);
            try
            {
                cells[0] = new DataCell(obj.FirstName);
                cells[1] = new DataCell(obj.MiddleName);
                cells[2] = new DataCell(obj.LastName);
                cells[3] = new DataCell(obj.Age);
                cells[4] = new DataCell(obj.Employed);
                cells[5] = new DataCell(obj.Score);
                await spreadsheet.AddRowAsync(cells.AsMemory(0, 6), token).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<DataCell>.Shared.Return(cells, true);
            }
        }
    }
}