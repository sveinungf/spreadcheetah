namespace SpreadCheetah.Test.Helpers;

internal static class EmbeddedResources
{
    public static Stream GetStream(string filename)
    {
        var assembly = typeof(EmbeddedResources).Assembly;
        var resourceNames = assembly.GetManifestResourceNames()
            .Where(x => x.EndsWith(filename, StringComparison.OrdinalIgnoreCase))
            .ToList();

        var stream = resourceNames switch
        {
            [var resourceName] => assembly.GetManifestResourceStream(resourceName),
            [] => throw new ArgumentException("Could not find embedded resource.", nameof(filename)),
            _ => throw new ArgumentException("Found multiple embedded resources with the same name.", nameof(filename))
        };

        return stream ?? throw new InvalidOperationException("Embedded resource stream was null.");
    }
}
