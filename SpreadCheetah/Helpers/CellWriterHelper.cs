using System;

namespace SpreadCheetah.Helpers
{
    internal static partial class CellWriterHelper
    {
        public static readonly int RowStartMaxByteCount = RowStart().Length + SpreadsheetConstants.RowIndexMaxDigits + RowStartEndTag().Length;

        [StringLiteral.Utf8("<row r=\"")]
        public static partial ReadOnlySpan<byte> RowStart();

        [StringLiteral.Utf8("\">")]
        public static partial ReadOnlySpan<byte> RowStartEndTag();

        [StringLiteral.Utf8("</row>")]
        public static partial ReadOnlySpan<byte> RowEnd();
    }
}
