namespace SpreadCheetah.SourceGenerator.Test.Models;

public record CustomType(string Name);

public record RecordClassWithCustomType(CustomType CustomType, string AnotherValue);
