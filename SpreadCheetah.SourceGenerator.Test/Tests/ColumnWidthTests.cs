using System.Reflection;
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Models.ColumnWidth;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class ColumnWidthTests
{
    [Fact]
    public void ColumnWidth_ClassWithColumnWidth_CanReadOrder()
    {
        // Arrange
        var nameProperty = typeof(ClassWithColumnWidth).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithColumnWidth.Name), StringComparison.Ordinal));

        // Act
        var nameColWidthAttr = nameProperty?.GetCustomAttribute<ColumnWidthAttribute>();

        // Assert
        Assert.NotNull(nameProperty);
        Assert.NotNull(nameColWidthAttr);
        Assert.Equal(20, nameColWidthAttr.Width);
    }
}
