using SpreadCheetah.Helpers;

#if !NET6_0_OR_GREATER
using ArgumentNullException = SpreadCheetah.Helpers.Backporting.ArgumentNullExceptionBackport;
#endif

namespace SpreadCheetah.Formulas;

internal static class HyperlinkFormula
{
    public static Formula From(Uri uri)
    {
        var absoluteUri = EnsureValidAbsoluteUri(uri);
        return new Formula($"""HYPERLINK("{absoluteUri}")""");
    }

    public static Formula From(Uri uri, string friendlyName)
    {
        ArgumentNullException.ThrowIfNull(friendlyName);
        if (friendlyName.Length > 255)
            ThrowHelper.HyperlinkFriendlyNameTooLong(nameof(friendlyName));

        var absoluteUri = EnsureValidAbsoluteUri(uri);
        var escapedFriendlyName = EscapeDoubleQuotes(friendlyName);
        return new Formula($"""HYPERLINK("{absoluteUri}","{escapedFriendlyName}")""");
    }

    private static string EnsureValidAbsoluteUri(Uri uri)
    {
        ArgumentNullException.ThrowIfNull(uri);
        if (!uri.IsAbsoluteUri)
            ThrowHelper.UriMustBeAbsolute(nameof(uri));
        if (!uri.IsWellFormedOriginalString())
            ThrowHelper.UriMustBeWellFormed(nameof(uri));

        var absoluteUri = uri.AbsoluteUri;
        if (absoluteUri.Length > 255)
            ThrowHelper.UriTooLong(nameof(uri));

        return absoluteUri;
    }

    private static string EscapeDoubleQuotes(string value)
    {
#pragma warning disable CA1307 // No StringComparison overload for .NET Standard 2.0
        return value.Replace("\"", "\"\"");
#pragma warning restore CA1307 // No StringComparison overload for .NET Standard 2.0
    }
}
