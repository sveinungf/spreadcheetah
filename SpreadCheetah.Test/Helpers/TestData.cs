using System;
using System.Collections.Generic;
using System.Linq;

namespace SpreadCheetah.Test.Helpers
{
    internal static class TestData
    {
        private static readonly Type[] CellTypesArray = new[] { typeof(Cell), typeof(DataCell), typeof(StyledCell) };

        public static IEnumerable<object?[]> CellTypes() => CellTypesArray.Select(x => new object[] { x });

        public static IEnumerable<object?[]> CombineWithCellTypes(params object?[] values)
        {
            return values.SelectMany(_ => CellTypesArray, (value, type) => new object?[] { value, type });
        }

        public static IEnumerable<object?[]> CombineWithCellTypes(params (object?, object?)[] values)
        {
            return values.SelectMany(_ => CellTypesArray, (value, type) => new object?[] { value.Item1, value.Item2, type });
        }
    }
}
