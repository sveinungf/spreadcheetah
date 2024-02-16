using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record DiagnosticInfo(DiagnosticDescriptor Descriptor, LocationInfo? Location);
