using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.InferColumnHeaders;

[WorksheetRow(typeof(DerivedClassWithInferFromBaseClass))]
[WorksheetRow(typeof(DerivedClassWithInferFromBaseClassButNoInheritColumns))]
[WorksheetRow(typeof(ClassWithMultipleProperties))]
[WorksheetRow(typeof(DerivedClassWithInfer))]
[WorksheetRow(typeof(DerivedClassWithInferAlsoFromBaseClass))]
public partial class InferColumnHeadersContext : WorksheetRowContext;