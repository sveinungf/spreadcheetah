﻿//HintName: MyNamespace.MyGenRowContext.g.cs
// <auto-generated />
#nullable enable
using SpreadCheetah;
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.Styling;
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

        private WorksheetRowTypeInfo<MyNamespace.ClassWithColumnOrderForSomeProperties>? _ClassWithColumnOrderForSomeProperties;
        public WorksheetRowTypeInfo<MyNamespace.ClassWithColumnOrderForSomeProperties> ClassWithColumnOrderForSomeProperties => _ClassWithColumnOrderForSomeProperties
            ??= WorksheetRowMetadataServices.CreateObjectInfo<MyNamespace.ClassWithColumnOrderForSomeProperties>(
                AddHeaderRow0Async, AddAsRowAsync, AddRangeAsRowsAsync, null);

        private static async ValueTask AddHeaderRow0Async(SpreadCheetah.Spreadsheet spreadsheet, SpreadCheetah.Styling.StyleId? styleId, CancellationToken token)
        {
            var headerNames = ArrayPool<string?>.Shared.Rent(6);
            try
            {
                headerNames[0] = "MiddleName";
                headerNames[1] = "FirstName";
                headerNames[2] = "Score";
                headerNames[3] = "LastName";
                headerNames[4] = "Employed";
                headerNames[5] = "Age";
                await spreadsheet.AddHeaderRowAsync(headerNames.AsMemory(0, 6)!, styleId, token).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<string?>.Shared.Return(headerNames, true);
            }
        }

        private static ValueTask AddAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, MyNamespace.ClassWithColumnOrderForSomeProperties? obj, CancellationToken token)
        {
            if (spreadsheet is null)
                throw new ArgumentNullException(nameof(spreadsheet));
            if (obj is null)
                return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);
            return AddAsRowInternalAsync(spreadsheet, obj, token);
        }

        private static ValueTask AddRangeAsRowsAsync(SpreadCheetah.Spreadsheet spreadsheet,
            IEnumerable<MyNamespace.ClassWithColumnOrderForSomeProperties?> objs,
            CancellationToken token)
        {
            if (spreadsheet is null)
                throw new ArgumentNullException(nameof(spreadsheet));
            if (objs is null)
                throw new ArgumentNullException(nameof(objs));
            return AddRangeAsRowsInternalAsync(spreadsheet, objs, token);
        }

        private static async ValueTask AddAsRowInternalAsync(SpreadCheetah.Spreadsheet spreadsheet,
            MyNamespace.ClassWithColumnOrderForSomeProperties obj,
            CancellationToken token)
        {
            var cells = ArrayPool<DataCell>.Shared.Rent(6);
            try
            {
                var styleIds = Array.Empty<StyleId>();
                await AddCellsAsRowAsync(spreadsheet, obj, cells, styleIds, token).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<DataCell>.Shared.Return(cells, true);
            }
        }

        private static async ValueTask AddRangeAsRowsInternalAsync(SpreadCheetah.Spreadsheet spreadsheet,
            IEnumerable<MyNamespace.ClassWithColumnOrderForSomeProperties?> objs,
            CancellationToken token)
        {
            var cells = ArrayPool<DataCell>.Shared.Rent(6);
            try
            {
                var styleIds = Array.Empty<StyleId>();
                foreach (var obj in objs)
                {
                    await AddCellsAsRowAsync(spreadsheet, obj, cells, styleIds, token).ConfigureAwait(false);
                }
            }
            finally
            {
                ArrayPool<DataCell>.Shared.Return(cells, true);
            }
        }

        private static ValueTask AddCellsAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet,
            MyNamespace.ClassWithColumnOrderForSomeProperties? obj,
            DataCell[] cells, IReadOnlyList<StyleId> styleIds, CancellationToken token)
        {
            if (obj is null)
                return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);

            cells[0] = new DataCell(obj.MiddleName);
            cells[1] = new DataCell(obj.FirstName);
            cells[2] = new DataCell(obj.Score);
            cells[3] = new DataCell(obj.LastName);
            cells[4] = new DataCell(obj.Employed);
            cells[5] = new DataCell(obj.Age);
            return spreadsheet.AddRowAsync(cells.AsMemory(0, 6), token);
        }

        private static DataCell ConstructTruncatedDataCell(string? value, int truncateLength)
        {
            return value is null || value.Length <= truncateLength
                ? new DataCell(value)
                : new DataCell(value.AsMemory(0, truncateLength));
        }
    }
}
