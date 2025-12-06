using Bogus;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Extensions;
using Font = SpreadCheetah.Styling.Font;

namespace SpreadCheetah.TestHelpers.TestData;

public static class StyleGenerator
{
    public static ICollection<Style> Generate(int count)
    {
        var testAlignments = CreateFaker<Alignment>()
            .StrictMode(true)
            .RuleFor(x => x.Horizontal, f => f.Random.Enum<HorizontalAlignment>())
            .RuleFor(x => x.Indent, f => f.Random.Number(255))
            .RuleFor(x => x.Vertical, f => f.Random.Enum<VerticalAlignment>())
            .RuleFor(x => x.WrapText, f => f.Random.Bool());

        var testEdgeBorders = CreateFaker<EdgeBorder>()
            .StrictMode(true)
            .RuleFor(x => x.BorderStyle, f => f.Random.Enum<BorderStyle>())
            .RuleFor(x => x.Color, f => f.Random.Color().OrNull(f, .1f));

        var testDiagonalBorders = CreateFaker<DiagonalBorder>()
            .StrictMode(true)
            .RuleFor(x => x.BorderStyle, f => f.Random.Enum<BorderStyle>())
            .RuleFor(x => x.Color, f => f.Random.Color().OrNull(f, .1f))
            .RuleFor(x => x.Type, f => f.Random.Enum<DiagonalBorderType>());

        var testBorders = CreateFaker<Border>()
            .StrictMode(true)
            .RuleFor(x => x.Bottom, _ => testEdgeBorders.Generate())
            .RuleFor(x => x.Diagonal, _ => testDiagonalBorders.Generate())
            .RuleFor(x => x.Left, _ => testEdgeBorders.Generate())
            .RuleFor(x => x.Right, _ => testEdgeBorders.Generate())
            .RuleFor(x => x.Top, _ => testEdgeBorders.Generate());

        var testFills = CreateFaker<Fill>()
            .StrictMode(true)
            .RuleFor(x => x.Color, f => f.Random.Color().OrNull(f, .1f));

        var testFonts = CreateFaker<Font>()
            .StrictMode(true)
            .Ignore("ActualSize")
            .RuleFor(x => x.Bold, f => f.Random.Bool())
            .RuleFor(x => x.Color, f => f.Random.Color().OrNull(f, .1f))
            .RuleFor(x => x.Italic, f => f.Random.Bool())
            .RuleFor(x => x.Name, f => f.Commerce.ProductName().OrNull(f, .1f))
            .RuleFor(x => x.Size, f => f.Random.Number(10000, 72000) / 1000.0)
            .RuleFor(x => x.Strikethrough, f => f.Random.Bool())
            .RuleFor(x => x.Underline, f => f.Random.Enum<Underline>());

        var styles = CreateFaker<Style>()
            .StrictMode(true)
            .RuleFor(x => x.Alignment, _ => testAlignments.Generate())
            .RuleFor(x => x.Border, _ => testBorders.Generate())
            .RuleFor(x => x.Fill, _ => testFills.Generate())
            .RuleFor(x => x.Font, _ => testFonts.Generate())
            .RuleFor(x => x.Format, f => f.Random.NumberFormat().OrNull(f, .1f))
#pragma warning disable CS0618 // Type or member is obsolete - Being ignored
            .Ignore(x => x.NumberFormat);
#pragma warning restore CS0618 // Type or member is obsolete

        return styles.Generate(count);
    }

    private static Faker<T> CreateFaker<T>() where T : class
    {
        return new Faker<T>()
            .UseSeed(42);
    }
}
