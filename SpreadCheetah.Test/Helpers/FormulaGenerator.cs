using System;
using System.Text;
using Xunit;

namespace SpreadCheetah.Test.Helpers
{
    internal static class FormulaGenerator
    {
        public static string Generate(int length)
        {
            var sb = new StringBuilder("CONCAT(A1; ");
            var remaining = length - sb.Length - 1;
            if (remaining < 3)
                throw new ArgumentException("Length is too small", nameof(length));

            var value = new string('a', 100);
            while (remaining > 0)
            {
                if (remaining > value.Length + 3)
                {
                    var before = sb.Length;
                    sb.Append('"');
                    sb.Append(value);
                    sb.Append('"');
                    sb.Append("; ");
                    var after = sb.Length;
                    remaining -= after - before;
                    continue;
                }

                sb.Append('"');
                sb.Append(new string('b', remaining - 2));
                sb.Append('"');
                remaining = 0;
            }

            sb.Append(')');

            Assert.Equal(length, sb.Length);
            return sb.ToString();
        }
    }
}
