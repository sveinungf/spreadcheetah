using SpreadCheetah.SourceGeneration;

namespace SpreadCheetah.SourceGenerator.Test.Models.Formulas;

[WorksheetRow(typeof(ClassWithFormula))]
public partial class FormulaContext : WorksheetRowContext;