using Microsoft.CodeAnalysis;
using SpreadCheetah.SourceGenerator.Models;

namespace SpreadCheetah.SourceGenerator.Helpers;

// TODO: Remove?
internal static class LocationMap
{
    public static Location ToLocation(this LocationInfo info)
    {
        return Location.Create(info.FilePath, info.TextSpan, info.LineSpan);
    }
}
