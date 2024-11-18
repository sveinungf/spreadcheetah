using System.Reflection;
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Models.ColumnHeader;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class ColumnHeaderTests
{
    [Fact]
    public void ColumnHeader_ClassWithPropertyReferenceColumnHeaders_CanReadTypeAndPropertyName()
    {
        // Arrange
        var publicProperties = typeof(ClassWithPropertyReferenceColumnHeaders).GetProperties(BindingFlags.Public | BindingFlags.Instance);
        var firstNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithPropertyReferenceColumnHeaders.FirstName), StringComparison.Ordinal));
        var lastNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithPropertyReferenceColumnHeaders.LastName), StringComparison.Ordinal));
        var nationalityProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithPropertyReferenceColumnHeaders.Nationality), StringComparison.Ordinal));
        var addressLine1Property = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithPropertyReferenceColumnHeaders.AddressLine1), StringComparison.Ordinal));
        var addressLine2Property = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithPropertyReferenceColumnHeaders.AddressLine2), StringComparison.Ordinal));
        var ageProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithPropertyReferenceColumnHeaders.Age), StringComparison.Ordinal));

        // Act
        var firstNameColHeaderAttr = firstNameProperty?.GetCustomAttribute<ColumnHeaderAttribute>();
        var lastNameColHeaderAttr = lastNameProperty?.GetCustomAttribute<ColumnHeaderAttribute>();
        var nationalityColHeaderAttr = nationalityProperty?.GetCustomAttribute<ColumnHeaderAttribute>();
        var addressLine1ColHeaderAttr = addressLine1Property?.GetCustomAttribute<ColumnHeaderAttribute>();
        var addressLine2ColHeaderAttr = addressLine2Property?.GetCustomAttribute<ColumnHeaderAttribute>();
        var ageColHeaderAttr = ageProperty?.GetCustomAttribute<ColumnHeaderAttribute>();

        // Assert
        Assert.NotNull(firstNameProperty);
        Assert.NotNull(firstNameColHeaderAttr);
        Assert.Equal(typeof(ColumnHeaderResources), firstNameColHeaderAttr.Type);
        Assert.Equal(nameof(ColumnHeaderResources.Header_FirstName), firstNameColHeaderAttr.PropertyName);

        Assert.NotNull(lastNameProperty);
        Assert.NotNull(lastNameColHeaderAttr);
        Assert.Equal(typeof(ColumnHeaderResources), lastNameColHeaderAttr.Type);
        Assert.Equal(nameof(ColumnHeaderResources.Header_LastName), lastNameColHeaderAttr.PropertyName);

        Assert.NotNull(nationalityProperty);
        Assert.NotNull(nationalityColHeaderAttr);
        Assert.Equal(typeof(ColumnHeaders), nationalityColHeaderAttr.Type);
        Assert.Equal(nameof(ColumnHeaders.HeaderNationality), nationalityColHeaderAttr.PropertyName);

        Assert.NotNull(addressLine1Property);
        Assert.NotNull(addressLine1ColHeaderAttr);
        Assert.Equal(typeof(ColumnHeaders), addressLine1ColHeaderAttr.Type);
        Assert.Equal(nameof(ColumnHeaders.HeaderAddressLine1), addressLine1ColHeaderAttr.PropertyName);

        Assert.NotNull(addressLine2Property);
        Assert.NotNull(addressLine2ColHeaderAttr);
        Assert.Equal(typeof(ColumnHeaders), addressLine2ColHeaderAttr.Type);
        Assert.Equal(nameof(ColumnHeaders.HeaderAddressLine2), addressLine2ColHeaderAttr.PropertyName);

        Assert.NotNull(ageProperty);
        Assert.NotNull(ageColHeaderAttr);
        Assert.Equal(typeof(ColumnHeaders), ageColHeaderAttr.Type);
        Assert.Equal(nameof(ColumnHeaders.HeaderAge), ageColHeaderAttr.PropertyName);
    }

    [Fact]
    public void ColumnHeader_ClassWithSpecialCharacterColumnHeaders_CanReadName()
    {
        // Arrange
        var publicProperties = typeof(ClassWithSpecialCharacterColumnHeaders).GetProperties(BindingFlags.Public | BindingFlags.Instance);
        var firstNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithSpecialCharacterColumnHeaders.FirstName), StringComparison.Ordinal));
        var lastNameProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithSpecialCharacterColumnHeaders.LastName), StringComparison.Ordinal));
        var nationalityProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithSpecialCharacterColumnHeaders.Nationality), StringComparison.Ordinal));
        var addressLine1Property = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithSpecialCharacterColumnHeaders.AddressLine1), StringComparison.Ordinal));
        var addressLine2Property = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithSpecialCharacterColumnHeaders.AddressLine2), StringComparison.Ordinal));
        var ageProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithSpecialCharacterColumnHeaders.Age), StringComparison.Ordinal));
        var noteProperty = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithSpecialCharacterColumnHeaders.Note), StringComparison.Ordinal));
        var note2Property = publicProperties.SingleOrDefault(p => string.Equals(p.Name, nameof(ClassWithSpecialCharacterColumnHeaders.Note2), StringComparison.Ordinal));

        // Act
        var firstNameColHeaderAttr = firstNameProperty?.GetCustomAttribute<ColumnHeaderAttribute>();
        var lastNameColHeaderAttr = lastNameProperty?.GetCustomAttribute<ColumnHeaderAttribute>();
        var nationalityColHeaderAttr = nationalityProperty?.GetCustomAttribute<ColumnHeaderAttribute>();
        var addressLine1ColHeaderAttr = addressLine1Property?.GetCustomAttribute<ColumnHeaderAttribute>();
        var addressLine2ColHeaderAttr = addressLine2Property?.GetCustomAttribute<ColumnHeaderAttribute>();
        var ageColHeaderAttr = ageProperty?.GetCustomAttribute<ColumnHeaderAttribute>();
        var noteColHeaderAttr = noteProperty?.GetCustomAttribute<ColumnHeaderAttribute>();
        var note2ColHeaderAttr = note2Property?.GetCustomAttribute<ColumnHeaderAttribute>();

        // Assert
        Assert.NotNull(firstNameProperty);
        Assert.NotNull(firstNameColHeaderAttr);
        Assert.Equal("First name", firstNameColHeaderAttr.Name);

        Assert.NotNull(lastNameProperty);
        Assert.NotNull(lastNameColHeaderAttr);
        Assert.Equal("", lastNameColHeaderAttr.Name);

        Assert.NotNull(nationalityProperty);
        Assert.NotNull(nationalityColHeaderAttr);
        Assert.Equal("Nationality (escaped characters \", \', \\)", nationalityColHeaderAttr.Name);

        Assert.NotNull(addressLine1Property);
        Assert.NotNull(addressLine1ColHeaderAttr);
        Assert.Equal("Address line 1 (escaped characters \r\n, \t)", addressLine1ColHeaderAttr.Name);

        Assert.NotNull(addressLine2Property);
        Assert.NotNull(addressLine2ColHeaderAttr);
        Assert.Equal(@"Address line 2 (verbatim
string: "", \)", addressLine2ColHeaderAttr.Name);

        Assert.NotNull(ageProperty);
        Assert.NotNull(ageColHeaderAttr);
        Assert.Equal("""
        Age (
            raw
            string
            literal
        )
    """, ageColHeaderAttr.Name);

        Assert.NotNull(noteProperty);
        Assert.NotNull(noteColHeaderAttr);
        Assert.Equal("Note (unicode escape sequence 🌉, \ud83d\udc4d, \xE7)", noteColHeaderAttr.Name);

        Assert.NotNull(note2Property);
        Assert.NotNull(note2ColHeaderAttr);
        Assert.Equal($"Note 2 (constant interpolated string: This is a constant)", note2ColHeaderAttr.Name);
    }
}
