using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Helpers;
using SpreadCheetah.SourceGenerator.Test.Models.ColumnWidth;
using System.Reflection;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class ColumnWidthTests
{
    [Fact]
    public void ColumnWidth_ClassWithColumnWidth_CanReadOrder()
    {
        // Arrange
        var properties = typeof(ClassWithColumnWidth).ToPropertyDictionary();
        var nameProperty = properties[nameof(ClassWithColumnWidth.Name)];

        // Act
        var attribute = nameProperty.GetCustomAttribute<ColumnWidthAttribute>();

        // Assert
        Assert.NotNull(attribute);
        Assert.Equal(20, attribute.Width);
    }
}
