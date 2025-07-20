using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.Formulas;

[WorksheetRow(typeof(ClassWithFormula))]
[WorksheetRow(typeof(ClassWithNullableFormula))]
[WorksheetRow(typeof(ClassWithStyledFormulas))]
[WorksheetRow(typeof(ClassWithUri))]
public partial class FormulaContext : WorksheetRowContext;