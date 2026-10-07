## Projects are multi-targeted — only build and test .NET 10

The projects target multiple frameworks (`netstandard2.0;net8.0;net9.0;net10.0`, plus `net472`/`net48` for the test and benchmark projects). In this environment only `net10.0` can be built and executed — the other runtimes are not installed (and `net472`/`net48` would need `mono`). Always pin the TFM when building or running tests, e.g. `-f net10.0`. Do not attempt to build or test the other TFMs.

## Build the test projects in Debug, not Release

`SpreadCheetah/Assembly.cs` only emits `InternalsVisibleTo("SpreadCheetah.Test" / "SpreadCheetah.TestHelpers" / "SpreadCheetah.Benchmark")` under `#if !RELEASE`.

Consequence: building `SpreadCheetah.Test` / `SpreadCheetah.TestHelpers` with `-c Release` fails. These are NOT pre-existing or library bugs — they are purely the missing internals in Release. Use Debug for anything test-related:

```bash
dotnet build SpreadCheetah.Test/SpreadCheetah.Test.csproj -c Debug
```

The main `SpreadCheetah` library itself builds fine in both configurations.

## Running the test suite

`dotnet test` works via Microsoft.Testing.Platform (the runner is declared in `global.json`). This SDK also requires the `--project` flag (a positional project path is rejected):

```bash
dotnet test --project SpreadCheetah.Test/SpreadCheetah.Test.csproj -c Debug -f net10.0
```

A full run takes ~15s. Look for the `total:` / `failed:` summary lines.

### `dotnet test` reports the project "is using VSTest test runner"

That means the NuGet package build files are not being imported — usually because the `obj/` folders are stale copies from a Windows machine.

Delete the `obj/` of every project in the restore graph (including `SpreadCheetah.SourceGenerator`) and restore fresh:

```bash
rm -rf SpreadCheetah*/obj
dotnet restore SpreadCheetah.Test/SpreadCheetah.Test.csproj
```

## Known baseline test failures (25)

On an untouched checkout, exactly 25 tests fail. All 25 are the `With\tValid\r\nControlCharacters` cases and are environment-related, not caused by your change:

- `SpreadsheetRowTests.Spreadsheet_AddRow_CellWithStringValue` (12)
- `SpreadsheetRowTests.Spreadsheet_AddRow_CellWithReadOnlyMemoryOfCharValue` (12)
- `SpreadsheetTableTests.Spreadsheet_Table_ValidHeaderName` (1)

If you see these 25 plus a small number of new failures, only the new ones are yours. To confirm a failure is pre-existing, `git stash -u`, rebuild, rerun, `git stash pop`.

## `PublicApiTests.PublicApi_Generate` fails whenever public API changes

The test snapshots the full public API per TFM into
`SpreadCheetah.Test/Tests/PublicApiTests.PublicApi_Generate.{Net4_7,DotNet8_0,DotNet9_0,DotNet10_0}.verified.txt`.
Adding any public type or member makes the test fail and writes a `*.received.*` file next to it (received files are gitignored). If your change intentionally changes the public API, regenerate by diffing received vs verified (the received file contains the new correct content).

## `SpreadCheetah/CompatibilitySuppressions.xml` is build noise

The pack step (`GeneratePackageOnBuild`) rewrites this file on Release builds with SDK-drift noise (a ~240-line diff). Never include it in your change set. If it shows up dirty after a build:

```bash
git checkout -- SpreadCheetah/CompatibilitySuppressions.xml
```

## NuGet cache corruption

If a build fails with `NETSDK1064: Package X, version Y was not found` even though the package exists under `~/.nuget/packages/x/`, the cache entry is corrupt. Delete it and re-restore:

```bash
rm -rf ~/.nuget/packages/<package-id>
dotnet restore <project>.csproj
```

## Referencing SpreadCheetah from standalone projects

The global NuGet cache also contains the published `SpreadCheetah` package (older version, no latest APIs), so a `PackageReference` in a scratch/verification project will silently resolve to the wrong assembly. Reference the built DLL directly instead:

```xml
<Reference Include="SpreadCheetah">
  <HintPath>/f/_git/spreadcheetah/SpreadCheetah/bin/Debug/net10.0/SpreadCheetah.dll</HintPath>
</Reference>
```
