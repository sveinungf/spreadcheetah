using System.Reflection;
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Models.ColumnOrdering;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class ColumnOrderTests
{
    [Fact]
    public void ColumnOrder_ClassWithPropertyReferenceColumnHeaders_CanReadOrder()
    {
        // Arrange
        var publicProperties = typeof(ClassWithColumnOrdering).GetProperties(BindingFlags.Public | BindingFlags.Instance);
        var firstNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithColumnOrdering.FirstName), StringComparison.Ordinal));
        var lastNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithColumnOrdering.LastName), StringComparison.Ordinal));
        var gpaProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithColumnOrdering.Gpa), StringComparison.Ordinal));
        var ageProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithColumnOrdering.Age), StringComparison.Ordinal));

        // Act
        var firstNameColOrderAttr = firstNameProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var lastNameColOrderAttr = lastNameProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var gpaColOrderAttr = gpaProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var ageColOrderAttr = ageProperty?.GetCustomAttribute<ColumnOrderAttribute>();

        // Assert
        Assert.NotNull(firstNameProperty);
        Assert.NotNull(firstNameColOrderAttr);
        Assert.Equal(2, firstNameColOrderAttr.Order);

        Assert.NotNull(lastNameProperty);
        Assert.NotNull(lastNameColOrderAttr);
        Assert.Equal(1, lastNameColOrderAttr.Order);

        Assert.NotNull(gpaProperty);
        Assert.Null(gpaColOrderAttr);

        Assert.NotNull(ageProperty);
        Assert.NotNull(ageColOrderAttr);
        Assert.Equal(3, ageColOrderAttr.Order);
    }

    [Fact]
    public void ColumnOrder_ReadOnlyRecordStructWithColumnOrdering_CanReadOrder()
    {
        // Arrange
        var publicProperties = typeof(ReadOnlyRecordStructWithColumnOrdering).GetProperties(BindingFlags.Public | BindingFlags.Instance);
        var firstNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ReadOnlyRecordStructWithColumnOrdering.FirstName), StringComparison.Ordinal));
        var lastNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ReadOnlyRecordStructWithColumnOrdering.LastName), StringComparison.Ordinal));
        var gpaProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ReadOnlyRecordStructWithColumnOrdering.Gpa), StringComparison.Ordinal));
        var ageProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ReadOnlyRecordStructWithColumnOrdering.Age), StringComparison.Ordinal));

        // Act
        var firstNameColOrderAttr = firstNameProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var lastNameColOrderAttr = lastNameProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var gpaColOrderAttr = gpaProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var ageColOrderAttr = ageProperty?.GetCustomAttribute<ColumnOrderAttribute>();

        // Assert
        Assert.NotNull(firstNameProperty);
        Assert.NotNull(firstNameColOrderAttr);
        Assert.Equal(2, firstNameColOrderAttr.Order);

        Assert.NotNull(lastNameProperty);
        Assert.NotNull(lastNameColOrderAttr);
        Assert.Equal(1, lastNameColOrderAttr.Order);

        Assert.NotNull(gpaProperty);
        Assert.Null(gpaColOrderAttr);

        Assert.NotNull(ageProperty);
        Assert.NotNull(ageColOrderAttr);
        Assert.Equal(3, ageColOrderAttr.Order);
    }

    [Fact]
    public void ColumnOrder_ReadOnlyStructWithColumnOrdering_CanReadOrder()
    {
        // Arrange
        var publicProperties = typeof(ReadOnlyStructWithColumnOrdering).GetProperties(BindingFlags.Public | BindingFlags.Instance);
        var firstNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ReadOnlyStructWithColumnOrdering.FirstName), StringComparison.Ordinal));
        var lastNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ReadOnlyStructWithColumnOrdering.LastName), StringComparison.Ordinal));
        var gpaProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ReadOnlyStructWithColumnOrdering.Gpa), StringComparison.Ordinal));
        var ageProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ReadOnlyStructWithColumnOrdering.Age), StringComparison.Ordinal));

        // Act
        var firstNameColOrderAttr = firstNameProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var lastNameColOrderAttr = lastNameProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var gpaColOrderAttr = gpaProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var ageColOrderAttr = ageProperty?.GetCustomAttribute<ColumnOrderAttribute>();

        // Assert
        Assert.NotNull(firstNameProperty);
        Assert.NotNull(firstNameColOrderAttr);
        Assert.Equal(2, firstNameColOrderAttr.Order);

        Assert.NotNull(lastNameProperty);
        Assert.NotNull(lastNameColOrderAttr);
        Assert.Equal(1, lastNameColOrderAttr.Order);

        Assert.NotNull(gpaProperty);
        Assert.Null(gpaColOrderAttr);

        Assert.NotNull(ageProperty);
        Assert.NotNull(ageColOrderAttr);
        Assert.Equal(3, ageColOrderAttr.Order);
    }

    [Fact]
    public void ColumnOrder_RecordClassWithColumnOrdering_CanReadOrder()
    {
        // Arrange
        var publicProperties = typeof(RecordClassWithColumnOrdering).GetProperties(BindingFlags.Public | BindingFlags.Instance);
        var firstNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(RecordClassWithColumnOrdering.FirstName), StringComparison.Ordinal));
        var lastNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(RecordClassWithColumnOrdering.LastName), StringComparison.Ordinal));
        var gpaProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(RecordClassWithColumnOrdering.Gpa), StringComparison.Ordinal));
        var ageProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(RecordClassWithColumnOrdering.Age), StringComparison.Ordinal));

        // Act
        var firstNameColOrderAttr = firstNameProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var lastNameColOrderAttr = lastNameProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var gpaColOrderAttr = gpaProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var ageColOrderAttr = ageProperty?.GetCustomAttribute<ColumnOrderAttribute>();

        // Assert
        Assert.NotNull(firstNameProperty);
        Assert.NotNull(firstNameColOrderAttr);
        Assert.Equal(2, firstNameColOrderAttr.Order);

        Assert.NotNull(lastNameProperty);
        Assert.NotNull(lastNameColOrderAttr);
        Assert.Equal(1, lastNameColOrderAttr.Order);

        Assert.NotNull(gpaProperty);
        Assert.Null(gpaColOrderAttr);

        Assert.NotNull(ageProperty);
        Assert.NotNull(ageColOrderAttr);
        Assert.Equal(3, ageColOrderAttr.Order);
    }

    [Fact]
    public void ColumnOrder_RecordStructWithColumnOrdering_CanReadOrder()
    {
        // Arrange
        var publicProperties = typeof(RecordStructWithColumnOrdering).GetProperties(BindingFlags.Public | BindingFlags.Instance);
        var firstNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(RecordStructWithColumnOrdering.FirstName), StringComparison.Ordinal));
        var lastNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(RecordStructWithColumnOrdering.LastName), StringComparison.Ordinal));
        var gpaProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(RecordStructWithColumnOrdering.Gpa), StringComparison.Ordinal));
        var ageProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(RecordStructWithColumnOrdering.Age), StringComparison.Ordinal));

        // Act
        var firstNameColOrderAttr = firstNameProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var lastNameColOrderAttr = lastNameProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var gpaColOrderAttr = gpaProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var ageColOrderAttr = ageProperty?.GetCustomAttribute<ColumnOrderAttribute>();

        // Assert
        Assert.NotNull(firstNameProperty);
        Assert.NotNull(firstNameColOrderAttr);
        Assert.Equal(2, firstNameColOrderAttr.Order);

        Assert.NotNull(lastNameProperty);
        Assert.NotNull(lastNameColOrderAttr);
        Assert.Equal(1, lastNameColOrderAttr.Order);

        Assert.NotNull(gpaProperty);
        Assert.Null(gpaColOrderAttr);

        Assert.NotNull(ageProperty);
        Assert.NotNull(ageColOrderAttr);
        Assert.Equal(3, ageColOrderAttr.Order);
    }

    [Fact]
    public void ColumnOrder_StructWithColumnOrdering_CanReadOrder()
    {
        // Arrange
        var publicProperties = typeof(StructWithColumnOrdering).GetProperties(BindingFlags.Public | BindingFlags.Instance);
        var firstNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(StructWithColumnOrdering.FirstName), StringComparison.Ordinal));
        var lastNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(StructWithColumnOrdering.LastName), StringComparison.Ordinal));
        var gpaProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(StructWithColumnOrdering.Gpa), StringComparison.Ordinal));
        var ageProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(StructWithColumnOrdering.Age), StringComparison.Ordinal));

        // Act
        var firstNameColOrderAttr = firstNameProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var lastNameColOrderAttr = lastNameProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var gpaColOrderAttr = gpaProperty?.GetCustomAttribute<ColumnOrderAttribute>();
        var ageColOrderAttr = ageProperty?.GetCustomAttribute<ColumnOrderAttribute>();

        // Assert
        Assert.NotNull(firstNameProperty);
        Assert.NotNull(firstNameColOrderAttr);
        Assert.Equal(2, firstNameColOrderAttr.Order);

        Assert.NotNull(lastNameProperty);
        Assert.NotNull(lastNameColOrderAttr);
        Assert.Equal(1, lastNameColOrderAttr.Order);

        Assert.NotNull(gpaProperty);
        Assert.Null(gpaColOrderAttr);

        Assert.NotNull(ageProperty);
        Assert.NotNull(ageColOrderAttr);
        Assert.Equal(3, ageColOrderAttr.Order);
    }
}
