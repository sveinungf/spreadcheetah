using Microsoft.CodeAnalysis.Text;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record LocationInfo(
    string FilePath,
    TextSpan TextSpan,
    LinePositionSpan LineSpan);