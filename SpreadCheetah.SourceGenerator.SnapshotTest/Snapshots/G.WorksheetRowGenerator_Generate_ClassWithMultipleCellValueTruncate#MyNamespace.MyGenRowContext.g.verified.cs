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

        private WorksheetRowTypeInfo<MyNamespace.ClassWithMultipleCellValueTruncate>? _ClassWithMultipleCellValueTruncate;
        public WorksheetRowTypeInfo<MyNamespace.ClassWithMultipleCellValueTruncate> ClassWithMultipleCellValueTruncate => _ClassWithMultipleCellValueTruncate
            ??= WorksheetRowMetadataServices.CreateObjectInfo<MyNamespace.ClassWithMultipleCellValueTruncate>(AddHeaderRow0Async, AddAsRowAsync, AddRangeAsRowsAsync);

        private static async ValueTask AddHeaderRow0Async(SpreadCheetah.Spreadsheet spreadsheet, SpreadCheetah.Styling.StyleId? styleId, CancellationToken token)
        {
            var cells = ArrayPool<StyledCell>.Shared.Rent(4);
            try
            {
                cells[0] = new StyledCell("FirstName", styleId);
                cells[1] = new StyledCell("LastName", styleId);
                cells[2] = new StyledCell("YearOfBirth", styleId);
                cells[3] = new StyledCell("Initials", styleId);
                await spreadsheet.AddRowAsync(cells.AsMemory(0, 4), token).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<StyledCell>.Shared.Return(cells, true);
            }
        }

        private static ValueTask AddAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, MyNamespace.ClassWithMultipleCellValueTruncate? obj, CancellationToken token)
        {
            if (spreadsheet is null)
                throw new ArgumentNullException(nameof(spreadsheet));
            if (obj is null)
                return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);
            return AddAsRowInternalAsync(spreadsheet, obj, token);
        }

        private static ValueTask AddRangeAsRowsAsync(SpreadCheetah.Spreadsheet spreadsheet, IEnumerable<MyNamespace.ClassWithMultipleCellValueTruncate?> objs, CancellationToken token)
        {
            if (spreadsheet is null)
                throw new ArgumentNullException(nameof(spreadsheet));
            if (objs is null)
                throw new ArgumentNullException(nameof(objs));
            return AddRangeAsRowsInternalAsync(spreadsheet, objs, token);
        }

        private static async ValueTask AddAsRowInternalAsync(SpreadCheetah.Spreadsheet spreadsheet, MyNamespace.ClassWithMultipleCellValueTruncate obj, CancellationToken token)
        {
            var cells = ArrayPool<DataCell>.Shared.Rent(4);
            try
            {
                await AddCellsAsRowAsync(spreadsheet, obj, cells, token).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<DataCell>.Shared.Return(cells, true);
            }
        }

        private static async ValueTask AddRangeAsRowsInternalAsync(SpreadCheetah.Spreadsheet spreadsheet, IEnumerable<MyNamespace.ClassWithMultipleCellValueTruncate?> objs, CancellationToken token)
        {
            var cells = ArrayPool<DataCell>.Shared.Rent(4);
            try
            {
                await AddEnumerableAsRowsAsync(spreadsheet, objs, cells, token).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<DataCell>.Shared.Return(cells, true);
            }
        }

        private static async ValueTask AddEnumerableAsRowsAsync(SpreadCheetah.Spreadsheet spreadsheet, IEnumerable<MyNamespace.ClassWithMultipleCellValueTruncate?> objs, DataCell[] cells, CancellationToken token)
        {
            foreach (var obj in objs)
            {
                await AddCellsAsRowAsync(spreadsheet, obj, cells, token).ConfigureAwait(false);
            }
        }

        private static ValueTask AddCellsAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, MyNamespace.ClassWithMultipleCellValueTruncate? obj, DataCell[] cells, CancellationToken token)
        {
            if (obj is null)
                return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);

            var p0 = obj.FirstName;
            cells[0] = p0 is null || p0.Length <= 20 ? new DataCell(p0) : new DataCell(p0.AsMemory(0, 20));
            var p1 = obj.LastName;
            cells[1] = p1 is null || p1.Length <= 20 ? new DataCell(p1) : new DataCell(p1.AsMemory(0, 20));
            cells[2] = new DataCell(obj.YearOfBirth);
            var p3 = obj.Initials;
            cells[3] = p3 is null || p3.Length <= 5 ? new DataCell(p3) : new DataCell(p3.AsMemory(0, 5));
            return spreadsheet.AddRowAsync(cells.AsMemory(0, 4), token);
        }
    }
}
