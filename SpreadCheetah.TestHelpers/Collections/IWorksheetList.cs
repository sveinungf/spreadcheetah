using SpreadCheetah.TestHelpers.Assertions;

namespace SpreadCheetah.TestHelpers.Collections;

public interface IWorksheetList : IReadOnlyList<ISpreadsheetAssertSheet>, IDisposable;
