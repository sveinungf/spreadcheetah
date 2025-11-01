using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Models.CellValueTruncation;

namespace SpreadCheetah.SourceGenerator.Test.Models.Contexts;

[WorksheetRow(typeof(ClassWithTruncation))]
[WorksheetRow(typeof(ClassWithSingleAccessProperty))]
public partial class TruncationContext : WorksheetRowContext;
