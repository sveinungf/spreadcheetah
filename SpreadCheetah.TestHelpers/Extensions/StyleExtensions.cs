using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Implementations;
using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Extensions;

public static class StyleExtensions
{
#pragma warning disable CA1034 // Nested types should not be visible (false positive, should be resolved in .NET 11)
    extension(Style style)
#pragma warning restore CA1034 // Nested types should not be visible
    {
        public IStyle ToIStyle() => ClosedXmlStyle.Create(style);
    }
}
