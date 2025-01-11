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
        private static readonly MyNamespace.StringValueConverter _valueConverter0 = new MyNamespace.StringValueConverter(); 

        private WorksheetRowTypeInfo<MyNamespace.Person>? _Person;
        public WorksheetRowTypeInfo<MyNamespace.Person> Person => _Person
            ??= WorksheetRowMetadataServices.CreateObjectInfo<MyNamespace.Person>(
                AddHeaderRow0Async, AddAsRowAsync, AddRangeAsRowsAsync, null);

        private static async ValueTask AddHeaderRow0Async(SpreadCheetah.Spreadsheet spreadsheet, SpreadCheetah.Styling.StyleId? styleId, CancellationToken token)
        {
            var headerNames = ArrayPool<string?>.Shared.Rent(1);
            try
            {
                headerNames[0] = "Name";
                await spreadsheet.AddHeaderRowAsync(headerNames.AsMemory(0, 1)!, styleId, token).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<string?>.Shared.Return(headerNames, true);
            }
        }

        private static ValueTask AddAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, MyNamespace.Person? obj, CancellationToken token)
        {
            if (spreadsheet is null)
                throw new ArgumentNullException(nameof(spreadsheet));
            if (obj is null)
                return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);
            return AddAsRowInternalAsync(spreadsheet, obj, token);
        }

        private static ValueTask AddRangeAsRowsAsync(SpreadCheetah.Spreadsheet spreadsheet,
            IEnumerable<MyNamespace.Person?> objs,
            CancellationToken token)
        {
            if (spreadsheet is null)
                throw new ArgumentNullException(nameof(spreadsheet));
            if (objs is null)
                throw new ArgumentNullException(nameof(objs));
            return AddRangeAsRowsInternalAsync(spreadsheet, objs, token);
        }

        private static async ValueTask AddAsRowInternalAsync(SpreadCheetah.Spreadsheet spreadsheet,
            MyNamespace.Person obj,
            CancellationToken token)
        {
            var cells = ArrayPool<DataCell>.Shared.Rent(1);
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
            IEnumerable<MyNamespace.Person?> objs,
            CancellationToken token)
        {
            var cells = ArrayPool<DataCell>.Shared.Rent(1);
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
            MyNamespace.Person? obj,
            DataCell[] cells, IReadOnlyList<StyleId> styleIds, CancellationToken token)
        {
            if (obj is null)
                return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);

            cells[0] = _valueConverter0.ConvertToDataCell(obj.Name);
            return spreadsheet.AddRowAsync(cells.AsMemory(0, 1), token);
        }

        private WorksheetRowTypeInfo<MyNamespace.Car>? _Car;
        public WorksheetRowTypeInfo<MyNamespace.Car> Car => _Car
            ??= WorksheetRowMetadataServices.CreateObjectInfo<MyNamespace.Car>(
                AddHeaderRow1Async, AddAsRowAsync, AddRangeAsRowsAsync, null);

        private static async ValueTask AddHeaderRow1Async(SpreadCheetah.Spreadsheet spreadsheet, SpreadCheetah.Styling.StyleId? styleId, CancellationToken token)
        {
            var headerNames = ArrayPool<string?>.Shared.Rent(1);
            try
            {
                headerNames[0] = "Model";
                await spreadsheet.AddHeaderRowAsync(headerNames.AsMemory(0, 1)!, styleId, token).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<string?>.Shared.Return(headerNames, true);
            }
        }

        private static ValueTask AddAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, MyNamespace.Car? obj, CancellationToken token)
        {
            if (spreadsheet is null)
                throw new ArgumentNullException(nameof(spreadsheet));
            if (obj is null)
                return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);
            return AddAsRowInternalAsync(spreadsheet, obj, token);
        }

        private static ValueTask AddRangeAsRowsAsync(SpreadCheetah.Spreadsheet spreadsheet,
            IEnumerable<MyNamespace.Car?> objs,
            CancellationToken token)
        {
            if (spreadsheet is null)
                throw new ArgumentNullException(nameof(spreadsheet));
            if (objs is null)
                throw new ArgumentNullException(nameof(objs));
            return AddRangeAsRowsInternalAsync(spreadsheet, objs, token);
        }

        private static async ValueTask AddAsRowInternalAsync(SpreadCheetah.Spreadsheet spreadsheet,
            MyNamespace.Car obj,
            CancellationToken token)
        {
            var cells = ArrayPool<DataCell>.Shared.Rent(1);
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
            IEnumerable<MyNamespace.Car?> objs,
            CancellationToken token)
        {
            var cells = ArrayPool<DataCell>.Shared.Rent(1);
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
            MyNamespace.Car? obj,
            DataCell[] cells, IReadOnlyList<StyleId> styleIds, CancellationToken token)
        {
            if (obj is null)
                return spreadsheet.AddRowAsync(ReadOnlyMemory<DataCell>.Empty, token);

            cells[0] = _valueConverter0.ConvertToDataCell(obj.Model);
            return spreadsheet.AddRowAsync(cells.AsMemory(0, 1), token);
        }

        private static DataCell ConstructTruncatedDataCell(string? value, int truncateLength)
        {
            return value is null || value.Length <= truncateLength
                ? new DataCell(value)
                : new DataCell(value.AsMemory(0, truncateLength));
        }
    }
}
