using System.Reflection;
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Models.CellValueTruncation;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class CellValueTruncateTests
{
    [Fact]
    public void CellValueTruncate_ClassWithSingleAccessProperty_CanReadLength()
    {
        // Arrange
        var property = typeof(ClassWithSingleAccessProperty).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithSingleAccessProperty.Value), StringComparison.Ordinal));

        // Act
        var cellValueTruncateAttr = property?.GetCustomAttribute<CellValueTruncateAttribute>();

        // Assert
        Assert.NotNull(property);
        Assert.NotNull(cellValueTruncateAttr);
        Assert.Equal(1, cellValueTruncateAttr.Length);
    }

    [Fact]
    public void CellValueTruncate_ClassWithTruncation_CanReadLength()
    {
        // Arrange
        var property = typeof(ClassWithTruncation).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithTruncation.Value), StringComparison.Ordinal));

        // Act
        var cellValueTruncateAttr = property?.GetCustomAttribute<CellValueTruncateAttribute>();

        // Assert
        Assert.NotNull(property);
        Assert.NotNull(cellValueTruncateAttr);
        Assert.Equal(15, cellValueTruncateAttr.Length);
    }
}
