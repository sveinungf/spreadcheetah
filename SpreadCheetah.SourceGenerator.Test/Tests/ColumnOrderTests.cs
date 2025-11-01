using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Helpers;
using SpreadCheetah.SourceGenerator.Test.Models.ColumnOrdering;
using System.Reflection;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class ColumnOrderTests
{
    [Fact]
    public void ColumnOrder_ClassWithPropertyReferenceColumnHeaders_CanReadOrder()
    {
        // Arrange
        var properties = typeof(ClassWithColumnOrdering).ToPropertyDictionary();
        var firstNameProperty = properties[nameof(ClassWithColumnOrdering.FirstName)];
        var lastNameProperty = properties[nameof(ClassWithColumnOrdering.LastName)];
        var gpaProperty = properties[nameof(ClassWithColumnOrdering.Gpa)];
        var ageProperty = properties[nameof(ClassWithColumnOrdering.Age)];

        // Act
        var firstNameColOrderAttr = firstNameProperty.GetCustomAttribute<ColumnOrderAttribute>();
        var lastNameColOrderAttr = lastNameProperty.GetCustomAttribute<ColumnOrderAttribute>();
        var gpaColOrderAttr = gpaProperty.GetCustomAttribute<ColumnOrderAttribute>();
        var ageColOrderAttr = ageProperty.GetCustomAttribute<ColumnOrderAttribute>();

        // Assert
        Assert.NotNull(firstNameColOrderAttr);
        Assert.Equal(2, firstNameColOrderAttr.Order);

        Assert.NotNull(lastNameColOrderAttr);
        Assert.Equal(1, lastNameColOrderAttr.Order);

        Assert.Null(gpaColOrderAttr);

        Assert.NotNull(ageColOrderAttr);
        Assert.Equal(3, ageColOrderAttr.Order);
    }
}
