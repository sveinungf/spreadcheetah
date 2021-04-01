using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadCheetah.Helpers
{
    internal static class FormulaCellSpanHelper
    {
        // TODO: Cell start with style (part 1)
        // <c t="str" s="
        private static ReadOnlySpan<byte> StyledStringFormulaCellStart => new[]
        {
            (byte)'<', (byte)'c', (byte)' ', (byte)'t', (byte)'=', (byte)'"', (byte)'s',
            (byte)'t', (byte)'r', (byte)'"', (byte)' ', (byte)'s', (byte)'=', (byte)'"'
        };

        // TODO: Cell start with style (part 2)
        // "><f>

        // TODO: Cell start without style
        // <c t="str"><f>

        // TODO: After formula, with cached value
        // </f><v>

        // TODO: End with cached value
        // </v></c>

        // TODO: Try with skipping the <v> element when no cached value.
        // TODO: After formula, without cached value
        // </f></c>



        public static readonly int MaxCellElementLength = 0;
            //<c t="str" s="1">
			//	<f>UPPER(B1)</f>
			//	<v>TEST</v>
			//</c>
    }
}
