namespace SpreadCheetah.SourceGenerator.Test.Models;

public record CustomType(string Name);

public record RecordWithCustomType(CustomType CustomType, string AnotherValue);
