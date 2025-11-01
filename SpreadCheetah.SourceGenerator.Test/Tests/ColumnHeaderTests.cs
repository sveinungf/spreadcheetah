using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Helpers;
using SpreadCheetah.SourceGenerator.Test.Models.ColumnHeader;
using System.Reflection;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class ColumnHeaderTests
{
    [Fact]
    public void ColumnHeader_ClassWithPropertyReferenceColumnHeaders_CanReadTypeAndPropertyName()
    {
        // Arrange
        var properties = typeof(ClassWithPropertyReferenceColumnHeaders).ToPropertyDictionary();

        var firstNameProperty = properties[nameof(ClassWithPropertyReferenceColumnHeaders.FirstName)];
        var lastNameProperty = properties[nameof(ClassWithPropertyReferenceColumnHeaders.LastName)];
        var nationalityProperty = properties[nameof(ClassWithPropertyReferenceColumnHeaders.Nationality)];
        var addressLine1Property = properties[nameof(ClassWithPropertyReferenceColumnHeaders.AddressLine1)];
        var addressLine2Property = properties[nameof(ClassWithPropertyReferenceColumnHeaders.AddressLine2)];
        var ageProperty = properties[nameof(ClassWithPropertyReferenceColumnHeaders.Age)];

        // Act
        var firstNameColHeaderAttr = firstNameProperty.GetCustomAttribute<ColumnHeaderAttribute>();
        var lastNameColHeaderAttr = lastNameProperty.GetCustomAttribute<ColumnHeaderAttribute>();
        var nationalityColHeaderAttr = nationalityProperty.GetCustomAttribute<ColumnHeaderAttribute>();
        var addressLine1ColHeaderAttr = addressLine1Property.GetCustomAttribute<ColumnHeaderAttribute>();
        var addressLine2ColHeaderAttr = addressLine2Property.GetCustomAttribute<ColumnHeaderAttribute>();
        var ageColHeaderAttr = ageProperty.GetCustomAttribute<ColumnHeaderAttribute>();

        // Assert
        Assert.NotNull(firstNameColHeaderAttr);
        Assert.Equal(typeof(ColumnHeaderResources), firstNameColHeaderAttr.Type);
        Assert.Equal(nameof(ColumnHeaderResources.Header_FirstName), firstNameColHeaderAttr.PropertyName);

        Assert.NotNull(lastNameColHeaderAttr);
        Assert.Equal(typeof(ColumnHeaderResources), lastNameColHeaderAttr.Type);
        Assert.Equal(nameof(ColumnHeaderResources.Header_LastName), lastNameColHeaderAttr.PropertyName);

        Assert.NotNull(nationalityColHeaderAttr);
        Assert.Equal(typeof(ColumnHeaders), nationalityColHeaderAttr.Type);
        Assert.Equal(nameof(ColumnHeaders.HeaderNationality), nationalityColHeaderAttr.PropertyName);

        Assert.NotNull(addressLine1ColHeaderAttr);
        Assert.Equal(typeof(ColumnHeaders), addressLine1ColHeaderAttr.Type);
        Assert.Equal(nameof(ColumnHeaders.HeaderAddressLine1), addressLine1ColHeaderAttr.PropertyName);

        Assert.NotNull(addressLine2ColHeaderAttr);
        Assert.Equal(typeof(ColumnHeaders), addressLine2ColHeaderAttr.Type);
        Assert.Equal(nameof(ColumnHeaders.HeaderAddressLine2), addressLine2ColHeaderAttr.PropertyName);

        Assert.NotNull(ageColHeaderAttr);
        Assert.Equal(typeof(ColumnHeaders), ageColHeaderAttr.Type);
        Assert.Equal(nameof(ColumnHeaders.HeaderAge), ageColHeaderAttr.PropertyName);
    }

    [Fact]
    public void ColumnHeader_ClassWithSpecialCharacterColumnHeaders_CanReadName()
    {
        // Arrange
        var properties = typeof(ClassWithSpecialCharacterColumnHeaders).ToPropertyDictionary();

        var firstNameProperty = properties[nameof(ClassWithSpecialCharacterColumnHeaders.FirstName)];
        var lastNameProperty = properties[nameof(ClassWithSpecialCharacterColumnHeaders.LastName)];
        var nationalityProperty = properties[nameof(ClassWithSpecialCharacterColumnHeaders.Nationality)];
        var addressLine1Property = properties[nameof(ClassWithSpecialCharacterColumnHeaders.AddressLine1)];
        var addressLine2Property = properties[nameof(ClassWithSpecialCharacterColumnHeaders.AddressLine2)];
        var ageProperty = properties[nameof(ClassWithSpecialCharacterColumnHeaders.Age)];
        var noteProperty = properties[nameof(ClassWithSpecialCharacterColumnHeaders.Note)];
        var note2Property = properties[nameof(ClassWithSpecialCharacterColumnHeaders.Note2)];

        // Act
        var firstNameColHeaderAttr = firstNameProperty.GetCustomAttribute<ColumnHeaderAttribute>();
        var lastNameColHeaderAttr = lastNameProperty.GetCustomAttribute<ColumnHeaderAttribute>();
        var nationalityColHeaderAttr = nationalityProperty.GetCustomAttribute<ColumnHeaderAttribute>();
        var addressLine1ColHeaderAttr = addressLine1Property.GetCustomAttribute<ColumnHeaderAttribute>();
        var addressLine2ColHeaderAttr = addressLine2Property.GetCustomAttribute<ColumnHeaderAttribute>();
        var ageColHeaderAttr = ageProperty.GetCustomAttribute<ColumnHeaderAttribute>();
        var noteColHeaderAttr = noteProperty.GetCustomAttribute<ColumnHeaderAttribute>();
        var note2ColHeaderAttr = note2Property.GetCustomAttribute<ColumnHeaderAttribute>();

        // Assert
        Assert.NotNull(firstNameColHeaderAttr);
        Assert.Equal("First name", firstNameColHeaderAttr.Name);

        Assert.NotNull(lastNameColHeaderAttr);
        Assert.Equal("", lastNameColHeaderAttr.Name);

        Assert.NotNull(nationalityColHeaderAttr);
        Assert.Equal("Nationality (escaped characters \", \', \\)", nationalityColHeaderAttr.Name);

        Assert.NotNull(addressLine1ColHeaderAttr);
        Assert.Equal("Address line 1 (escaped characters \r\n, \t)", addressLine1ColHeaderAttr.Name);

        Assert.NotNull(addressLine2ColHeaderAttr);
        Assert.Equal(@"Address line 2 (verbatim
string: "", \)", addressLine2ColHeaderAttr.Name);

        Assert.NotNull(ageColHeaderAttr);
        Assert.Equal("""
        Age (
            raw
            string
            literal
        )
    """, ageColHeaderAttr.Name);

        Assert.NotNull(noteColHeaderAttr);
        Assert.Equal("Note (unicode escape sequence 🌉, \ud83d\udc4d, \xE7)", noteColHeaderAttr.Name);

        Assert.NotNull(note2ColHeaderAttr);
        Assert.Equal($"Note 2 (constant interpolated string: This is a constant)", note2ColHeaderAttr.Name);
    }
}
