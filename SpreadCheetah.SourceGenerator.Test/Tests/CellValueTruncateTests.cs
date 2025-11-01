using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Helpers;
using SpreadCheetah.SourceGenerator.Test.Models.CellValueTruncation;
using System.Reflection;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class CellValueTruncateTests
{
    [Fact]
    public void CellValueTruncate_ClassWithSingleAccessProperty_CanReadLength()
    {
        // Arrange
        var properties = typeof(ClassWithSingleAccessProperty).ToPropertyDictionary();
        var property = properties[nameof(ClassWithSingleAccessProperty.Value)];

        // Act
        var attribute = property.GetCustomAttribute<CellValueTruncateAttribute>();

        // Assert
        Assert.NotNull(attribute);
        Assert.Equal(1, attribute.Length);
    }
}
