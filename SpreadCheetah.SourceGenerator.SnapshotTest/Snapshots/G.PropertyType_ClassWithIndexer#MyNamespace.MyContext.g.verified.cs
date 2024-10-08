﻿//HintName: MyNamespace.MyContext.g.cs
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
    public partial class MyContext
    {
        private static MyContext? _default;
        public static MyContext Default => _default ??= new MyContext();

        public MyContext()
        {
        }

        private WorksheetRowTypeInfo<MyNamespace.MyClass>? _MyClass;
        public WorksheetRowTypeInfo<MyNamespace.MyClass> MyClass => _MyClass
            ??= EmptyWorksheetRowContext.CreateTypeInfo<MyNamespace.MyClass>();

        private static DataCell ConstructTruncatedDataCell(string? value, int truncateLength)
        {
            return value is null || value.Length <= truncateLength
                ? new DataCell(value)
                : new DataCell(value.AsMemory(0, truncateLength));
        }
    }
}
