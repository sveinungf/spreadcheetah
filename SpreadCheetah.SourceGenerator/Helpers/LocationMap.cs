using Microsoft.CodeAnalysis;
using SpreadCheetah.SourceGenerator.Models;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal static class LocationMap
{
    public static Location ToLocation(this LocationInfo info)
    {
        return Location.Create(info.FilePath, info.TextSpan, info.LineSpan);
    }

    public static LocationInfo? ToLocationInfo(this Location location)
    {
        if (location.SourceTree is null)
            return null;

        return new LocationInfo(location.SourceTree.FilePath, location.SourceSpan, location.GetLineSpan().Span);
    }
}
