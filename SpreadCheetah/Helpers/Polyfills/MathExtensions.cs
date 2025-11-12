#if !NET6_0_OR_GREATER
namespace System;

internal static class MathExtensions
{
    extension(Math)
    {
        public static (ulong Quotient, ulong Remainder) DivRem(ulong left, ulong right)
        {
            var quotient = left / right;
            return (quotient, left - (quotient * right));
        }
    }
}
#endif
