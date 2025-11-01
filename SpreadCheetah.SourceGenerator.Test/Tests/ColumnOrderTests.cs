using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Helpers;
using SpreadCheetah.SourceGenerator.Test.Models;
using SpreadCheetah.SourceGenerator.Test.Models.ColumnOrdering;
using SpreadCheetah.SourceGenerator.Test.Models.Contexts;
using SpreadCheetah.TestHelpers.Assertions;
using System.Reflection;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class ColumnOrderTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Theory, CombinatorialData]
    public async Task ColumnHeader_ObjectWithColumnOrdering(ObjectType type)
    {
        // Arrange
        var ctx = ColumnOrderingContext.Default;

        using var stream = new MemoryStream();
        await using var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await s.StartWorksheetAsync("Sheet", token: Token);

        // Act
        var task = type switch
        {
            ObjectType.Class => s.AddHeaderRowAsync(ctx.ClassWithColumnOrdering, token: Token),
            ObjectType.RecordClass => s.AddHeaderRowAsync(ctx.RecordClassWithColumnOrdering, token: Token),
            ObjectType.Struct => s.AddHeaderRowAsync(ctx.StructWithColumnOrdering, token: Token),
            ObjectType.RecordStruct => s.AddHeaderRowAsync(ctx.RecordStructWithColumnOrdering, token: Token),
            ObjectType.ReadOnlyStruct => s.AddHeaderRowAsync(ctx.ReadOnlyStructWithColumnOrdering, token: Token),
            ObjectType.ReadOnlyRecordStruct => s.AddHeaderRowAsync(ctx.ReadOnlyRecordStructWithColumnOrdering, token: Token),
            _ => throw new NotImplementedException()
        };

        await task;
        await s.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal("LastName", sheet["A1"].StringValue);
        Assert.Equal("FirstName", sheet["B1"].StringValue);
        Assert.Equal("Age", sheet["C1"].StringValue);
        Assert.Equal("Gpa", sheet["D1"].StringValue);
        Assert.Equal(4, sheet.CellCount);
    }

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
