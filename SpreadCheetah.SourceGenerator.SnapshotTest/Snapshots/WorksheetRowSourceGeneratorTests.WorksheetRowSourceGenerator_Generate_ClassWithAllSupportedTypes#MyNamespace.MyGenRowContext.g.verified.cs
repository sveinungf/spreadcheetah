﻿//HintName: MyNamespace.MyGenRowContext.g.cs
// <auto-generated />
#nullable enable
using SpreadCheetah;
using SpreadCheetah.SourceGeneration;
using System;
using System.Buffers;
using System.Collections.Generic;
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

        private WorksheetRowTypeInfo<SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithAllSupportedTypes>? _ClassWithAllSupportedTypes;
        public WorksheetRowTypeInfo<SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithAllSupportedTypes> ClassWithAllSupportedTypes => _ClassWithAllSupportedTypes ??= WorksheetRowMetadataServices.CreateObjectInfo<SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithAllSupportedTypes>(AddAsRowAsync, AddRangeAsRowsAsync);

        private static ValueTask AddAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithAllSupportedTypes? obj, CancellationToken token)
        {
            if (spreadsheet is null)
                throw new ArgumentNullException(nameof(spreadsheet));
            if (obj is null)
                return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);
            return AddAsRowInternalAsync(spreadsheet, obj, token);
        }

        private static ValueTask AddRangeAsRowsAsync(SpreadCheetah.Spreadsheet spreadsheet, IEnumerable<SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithAllSupportedTypes?> objs, CancellationToken token)
        {
            if (spreadsheet is null)
                throw new ArgumentNullException(nameof(spreadsheet));
            if (objs is null)
                throw new ArgumentNullException(nameof(objs));
            return AddRangeAsRowsInternalAsync(spreadsheet, objs, token);
        }

        private static async ValueTask AddAsRowInternalAsync(SpreadCheetah.Spreadsheet spreadsheet, SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithAllSupportedTypes obj, CancellationToken token)
        {
            var cells = ArrayPool<DataCell>.Shared.Rent(16);
            try
            {
                await AddCellsAsRowAsync(spreadsheet, obj, cells, token).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<DataCell>.Shared.Return(cells, true);
            }
        }

        private static async ValueTask AddRangeAsRowsInternalAsync(SpreadCheetah.Spreadsheet spreadsheet, IEnumerable<SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithAllSupportedTypes?> objs, CancellationToken token)
        {
            var cells = ArrayPool<DataCell>.Shared.Rent(16);
            try
            {
                await AddEnumerableAsRowsAsync(spreadsheet, objs, cells, token).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<DataCell>.Shared.Return(cells, true);
            }
        }

        private static async ValueTask AddEnumerableAsRowsAsync(SpreadCheetah.Spreadsheet spreadsheet, IEnumerable<SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithAllSupportedTypes?> objs, DataCell[] cells, CancellationToken token)
        {
            foreach (var obj in objs)
            {
                await AddCellsAsRowAsync(spreadsheet, obj, cells, token).ConfigureAwait(false);
            }
        }

        private static ValueTask AddCellsAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithAllSupportedTypes? obj, DataCell[] cells, CancellationToken token)
        {
            if (obj is null)
                return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);

            cells[0] = new DataCell(obj.StringValue);
            cells[1] = new DataCell(obj.NullableStringValue);
            cells[2] = new DataCell(obj.IntValue);
            cells[3] = new DataCell(obj.NullableIntValue);
            cells[4] = new DataCell(obj.LongValue);
            cells[5] = new DataCell(obj.NullableLongValue);
            cells[6] = new DataCell(obj.FloatValue);
            cells[7] = new DataCell(obj.NullableFloatValue);
            cells[8] = new DataCell(obj.DoubleValue);
            cells[9] = new DataCell(obj.NullableDoubleValue);
            cells[10] = new DataCell(obj.DecimalValue);
            cells[11] = new DataCell(obj.NullableDecimalValue);
            cells[12] = new DataCell(obj.DateTimeValue);
            cells[13] = new DataCell(obj.NullableDateTimeValue);
            cells[14] = new DataCell(obj.BoolValue);
            cells[15] = new DataCell(obj.NullableBoolValue);
            return spreadsheet.AddRowAsync(cells.AsMemory(0, 16), token);
        }
    }
}