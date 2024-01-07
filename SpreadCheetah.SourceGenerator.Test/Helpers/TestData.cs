using SpreadCheetah.SourceGenerator.Test.Helpers.Backporting;
using SpreadCheetah.SourceGenerator.Test.Models;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Helpers;

internal static class TestData
{
    public static TheoryData<ObjectType> ObjectTypes => EnumHelper.GetValues<ObjectType>().ToTheoryData();
}
