namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class VerifyChecksTests
{
    [Fact]
    public Task Verify_Conventions() => VerifyChecks.Run();
}