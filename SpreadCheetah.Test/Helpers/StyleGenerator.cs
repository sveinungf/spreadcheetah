using Bogus;
using SpreadCheetah.Styling;
using System.Drawing;
using Font = SpreadCheetah.Styling.Font;

namespace SpreadCheetah.Test.Helpers;

internal static class StyleGenerator
{
    public static ICollection<Style> Generate(int count)
    {
        var testAlignments = new Faker<Alignment>()
            .StrictMode(true)
            .RuleFor(x => x.Horizontal, f => f.Random.Enum<HorizontalAlignment>())
            .RuleFor(x => x.Indent, f => f.Random.Number(255))
            .RuleFor(x => x.Vertical, f => f.Random.Enum<VerticalAlignment>())
            .RuleFor(x => x.WrapText, f => f.Random.Bool());

        var testEdgeBorders = new Faker<EdgeBorder>()
            .StrictMode(true)
            .RuleFor(x => x.BorderStyle, f => f.Random.Enum<BorderStyle>())
            .RuleFor(x => x.Color, f => f.Random.Color().OrNull(f, .1f));

        var testDiagonalBorders = new Faker<DiagonalBorder>()
            .StrictMode(true)
            .RuleFor(x => x.BorderStyle, f => f.Random.Enum<BorderStyle>())
            .RuleFor(x => x.Color, f => f.Random.Color().OrNull(f, .1f))
            .RuleFor(x => x.Type, f => f.Random.Enum<DiagonalBorderType>());

        var testBorders = new Faker<Border>()
            .StrictMode(true)
            .RuleFor(x => x.Bottom, _ => testEdgeBorders.Generate())
            .RuleFor(x => x.Diagonal, _ => testDiagonalBorders.Generate())
            .RuleFor(x => x.Left, _ => testEdgeBorders.Generate())
            .RuleFor(x => x.Right, _ => testEdgeBorders.Generate())
            .RuleFor(x => x.Top, _ => testEdgeBorders.Generate());

        var testFills = new Faker<Fill>()
            .StrictMode(true)
            .RuleFor(x => x.Color, f => f.Random.Color().OrNull(f, .1f));

        var testFonts = new Faker<Font>()
            .StrictMode(true)
            .RuleFor(x => x.Bold, f => f.Random.Bool())
            .RuleFor(x => x.Color, f => f.Random.Color().OrNull(f, .1f))
            .RuleFor(x => x.Italic, f => f.Random.Bool())
            .RuleFor(x => x.Name, f => f.Commerce.ProductName().OrNull(f, .1f))
            .RuleFor(x => x.Size, f => f.Random.Double(1, 72))
            .RuleFor(x => x.Strikethrough, f => f.Random.Bool());

        var styles = new Faker<Style>()
            .StrictMode(true)
            .RuleFor(x => x.Alignment, _ => testAlignments.Generate())
            .RuleFor(x => x.Border, _ => testBorders.Generate())
            .RuleFor(x => x.Fill, _ => testFills.Generate())
            .RuleFor(x => x.Font, _ => testFonts.Generate())
            .RuleFor(x => x.NumberFormat, f => f.Random.String2(2, 255, "dhmsy0.,;-#%?/[] ").OrNull(f, .1f));

        return styles.Generate(count);
    }
}
