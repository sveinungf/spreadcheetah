using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.Formulas;

[WorksheetRow(typeof(ClassWithFormula))]
[WorksheetRow(typeof(ClassWithNullableFormula))]
[WorksheetRow(typeof(ClassWithStyledFormulas))]
public partial class FormulaContext : WorksheetRowContext;