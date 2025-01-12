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
        private static readonly MyNamespace.ObjectConverter _valueConverter0 = new MyNamespace.ObjectConverter(); 

        private WorksheetRowTypeInfo<MyNamespace.ClassWithConverterOnComplexProperty>? _ClassWithConverterOnComplexProperty;
        public WorksheetRowTypeInfo<MyNamespace.ClassWithConverterOnComplexProperty> ClassWithConverterOnComplexProperty => _ClassWithConverterOnComplexProperty
            ??= WorksheetRowMetadataServices.CreateObjectInfo<MyNamespace.ClassWithConverterOnComplexProperty>(
                AddHeaderRow0Async, AddAsRowAsync, AddRangeAsRowsAsync, null, CreateWorksheetRowDependencyInfo0);

        private static WorksheetRowDependencyInfo CreateWorksheetRowDependencyInfo0(Spreadsheet spreadsheet)
        {
            var styleIds = new[]
            {
                spreadsheet.GetStyleId("object style"),
            };
            return new WorksheetRowDependencyInfo(styleIds);
        }

        private static async ValueTask AddHeaderRow0Async(SpreadCheetah.Spreadsheet spreadsheet, SpreadCheetah.Styling.StyleId? styleId, CancellationToken token)
        {
            var headerNames = ArrayPool<string?>.Shared.Rent(2);
            try
            {
                headerNames[0] = "Property1";
                headerNames[1] = "Property2";
                await spreadsheet.AddHeaderRowAsync(headerNames.AsMemory(0, 2)!, styleId, token).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<string?>.Shared.Return(headerNames, true);
            }
        }

        private static ValueTask AddAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet, MyNamespace.ClassWithConverterOnComplexProperty? obj, CancellationToken token)
        {
            if (spreadsheet is null)
                throw new ArgumentNullException(nameof(spreadsheet));
            if (obj is null)
                return spreadsheet.AddRowAsync(ReadOnlyMemory<StyledCell>.Empty, token);
            return AddAsRowInternalAsync(spreadsheet, obj, token);
        }

        private static ValueTask AddRangeAsRowsAsync(SpreadCheetah.Spreadsheet spreadsheet,
            IEnumerable<MyNamespace.ClassWithConverterOnComplexProperty?> objs,
            CancellationToken token)
        {
            if (spreadsheet is null)
                throw new ArgumentNullException(nameof(spreadsheet));
            if (objs is null)
                throw new ArgumentNullException(nameof(objs));
            return AddRangeAsRowsInternalAsync(spreadsheet, objs, token);
        }

        private static async ValueTask AddAsRowInternalAsync(SpreadCheetah.Spreadsheet spreadsheet,
            MyNamespace.ClassWithConverterOnComplexProperty obj,
            CancellationToken token)
        {
            var cells = ArrayPool<StyledCell>.Shared.Rent(2);
            try
            {
                var worksheetRowDependencyInfo = spreadsheet.GetOrCreateWorksheetRowDependencyInfo(Default.ClassWithConverterOnComplexProperty);
                var styleIds = worksheetRowDependencyInfo.StyleIds;
                await AddCellsAsRowAsync(spreadsheet, obj, cells, styleIds, token).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<StyledCell>.Shared.Return(cells, true);
            }
        }

        private static async ValueTask AddRangeAsRowsInternalAsync(SpreadCheetah.Spreadsheet spreadsheet,
            IEnumerable<MyNamespace.ClassWithConverterOnComplexProperty?> objs,
            CancellationToken token)
        {
            var cells = ArrayPool<StyledCell>.Shared.Rent(2);
            try
            {
                var worksheetRowDependencyInfo = spreadsheet.GetOrCreateWorksheetRowDependencyInfo(Default.ClassWithConverterOnComplexProperty);
                var styleIds = worksheetRowDependencyInfo.StyleIds;
                foreach (var obj in objs)
                {
                    await AddCellsAsRowAsync(spreadsheet, obj, cells, styleIds, token).ConfigureAwait(false);
                }
            }
            finally
            {
                ArrayPool<StyledCell>.Shared.Return(cells, true);
            }
        }

        private static ValueTask AddCellsAsRowAsync(SpreadCheetah.Spreadsheet spreadsheet,
            MyNamespace.ClassWithConverterOnComplexProperty? obj,
            StyledCell[] cells, IReadOnlyList<StyleId> styleIds, CancellationToken token)
        {
            if (obj is null)
                return spreadsheet.AddRowAsync(ReadOnlyMemory<StyledCell>.Empty, token);

            cells[0] = new StyledCell(_valueConverter0.ConvertToDataCell(obj.Property1), null);
            cells[1] = new StyledCell(_valueConverter0.ConvertToDataCell(obj.Property2), styleIds[0]);
            return spreadsheet.AddRowAsync(cells.AsMemory(0, 2), token);
        }

        private static DataCell ConstructTruncatedDataCell(string? value, int truncateLength)
        {
            return value is null || value.Length <= truncateLength
                ? new DataCell(value)
                : new DataCell(value.AsMemory(0, truncateLength));
        }
    }
}
