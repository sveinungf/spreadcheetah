#if !NET5_0_OR_GREATER
using ArchUnitNET.Domain;
using ArchUnitNET.Loader;
using ArchUnitNET.xUnit;
using SpreadCheetah.Test.Helpers;
using static ArchUnitNET.Fluent.ArchRuleDefinition;

namespace SpreadCheetah.Test.Tests;

public class ArchitectureTests
{
    private static readonly Architecture Architecture = new ArchLoader().LoadAssemblies(typeof(Spreadsheet).Assembly).Build();

    /// <summary>
    /// Structs with a default constructor leads to a compilation error on UWP: https://github.com/sveinungf/spreadcheetah/issues/58
    /// </summary>
    [Fact]
    public void Architecture_NoDefaultConstructorForStructs()
    {
        // Arrange
        var rule = Types().Should().FollowCustomCondition(new NoDefaultConstructorForStructsCondition());

        // Act & Assert
        rule.Check(Architecture);
    }
}
#endif