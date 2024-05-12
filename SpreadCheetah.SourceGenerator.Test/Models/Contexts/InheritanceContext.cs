using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Models.Inheritance;

namespace SpreadCheetah.SourceGenerator.Test.Models.Contexts;

[WorksheetRow(typeof(ClassDog))]
[WorksheetRow(typeof(RecordClassDog))]
public partial class InheritanceContext : WorksheetRowContext;