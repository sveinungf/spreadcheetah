using SpreadCheetah.Helpers;

namespace SpreadCheetah;

/// <summary>
/// Represents a formula for a worksheet cell.
/// </summary>
public readonly record struct Formula
{
    internal string FormulaText { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="Formula"/> struct with a formula text.
    /// The formula text should be without a starting equal sign (=).
    /// </summary>
    public Formula(string? formulaText)
    {
        FormulaText = formulaText ?? "";
    }

    public static Formula Hyperlink(Uri uri)
    {
        var absoluteUri = EnsureValidAbsoluteUri(uri);
        var xmlEncodedUri = XmlUtility.XmlEncode(absoluteUri);
        return new Formula($"""HYPERLINK("{xmlEncodedUri}")""");
    }

    public static Formula Hyperlink(Uri uri, string friendlyName)
    {
        var absoluteUri = EnsureValidAbsoluteUri(uri);

        // TODO: Use throw helpers
        if (friendlyName is null)
            throw new ArgumentNullException(nameof(friendlyName));
        if (friendlyName.Length > 255)
            throw new ArgumentException("The URI friendly name can not exceed 255 characters.", nameof(uri));

        var xmlEncodedUri = XmlUtility.XmlEncode(absoluteUri);
        var xmlEncodedFriendlyName = XmlUtility.XmlEncode(friendlyName);
        return new Formula($"""HYPERLINK("{xmlEncodedUri}","{xmlEncodedFriendlyName}")""");
    }

    private static string EnsureValidAbsoluteUri(Uri uri)
    {
        // TODO: Use throw helpers
        if (uri is null)
            throw new ArgumentNullException(nameof(uri));
        if (!uri.IsAbsoluteUri)
            throw new ArgumentException("The URI must be an absolute URI.", nameof(uri)); // TODO: Covered by IsWellFormedOriginalString?
        if (!uri.IsWellFormedOriginalString())
            throw new ArgumentException("The URI is not well formed.", nameof(uri));

        var absoluteUri = uri.AbsoluteUri;
        if (absoluteUri.Length > 255)
            throw new ArgumentException("The URI can not exceed 255 characters.", nameof(uri));

        return absoluteUri;
    }
}
