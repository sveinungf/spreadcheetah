using PublicApiGenerator;
using SpreadCheetah.Helpers;

namespace SpreadCheetah.Test.Tests;

public class PublicApiTests
{
    [Fact]
    public Task PublicApi_Generate()
    {
        // Arrange
        var options = new ApiGeneratorOptions
        {
#if DEBUG
            ExcludeTypes = [typeof(XmlUtility)]
#endif
        };

        // Act
        var publicApi = typeof(Spreadsheet).Assembly.GeneratePublicApi(options);

        // Assert
        var settings = new VerifySettings();
        settings.UniqueForTargetFrameworkAndVersion();
        return Verify(publicApi, settings);
    }
}
