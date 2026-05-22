using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Collections;

public interface IWorksheetList : IReadOnlyList<IWorksheet>, IDisposable;
