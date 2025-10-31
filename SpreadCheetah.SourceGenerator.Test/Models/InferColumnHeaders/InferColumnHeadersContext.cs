using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.InferColumnHeaders;

[WorksheetRow(typeof(ClassWithMultipleProperties))]
public partial class InferColumnHeadersContext : WorksheetRowContext;