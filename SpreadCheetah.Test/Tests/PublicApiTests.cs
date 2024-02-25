using PublicApiGenerator;

namespace SpreadCheetah.Test.Tests;

public class PublicApiTests
{
    [Fact]
    public Task PublicApi_Generate()
    {
        // Act
        var publicApi = typeof(Spreadsheet).Assembly.GeneratePublicApi();

        // Assert
        var settings = new VerifySettings();
        settings.UniqueForTargetFrameworkAndVersion();
        return Verify(publicApi, settings);
    }
}
