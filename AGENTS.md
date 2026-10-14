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

The runner is Microsoft.Testing.Platform (declared in `global.json`). **`dotnet test` is broken in this environment** — it fails with "is using VSTest test runner" because SDK 10.0.111 writes `*.csproj.nuget.g.targets` with backslash paths that fail `Exists()` on Linux. Cleaning `obj/` does NOT fix it (pre-existing, unrelated to your change). `dotnet build` is unaffected.

So build, then run the test exe directly (`SpreadCheetah.Test/bin/Debug/net10.0/SpreadCheetah.Test.dll`):

```bash
dotnet build SpreadCheetah.Test/SpreadCheetah.Test.csproj -c Debug -f net10.0
dotnet SpreadCheetah.Test/bin/Debug/net10.0/SpreadCheetah.Test.dll
```

A full run takes ~15s; look for the `total:` / `failed:` lines.

### Filtering (run a subset)

There is no `--filter-method` flag — it silently runs zero tests. Use `--filter` with the `FullyQualifiedName` property (`~` = contains, `=` = exact):

```bash
dotnet SpreadCheetah.Test/bin/Debug/net10.0/SpreadCheetah.Test.dll --filter "FullyQualifiedName~NonFiniteValue"
```

If a filter matches nothing, the test was likely renamed — list real names with `dotnet <exe> --list-tests`. (`--trep` publishes results; it does not filter.)

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

A `NETSDK1064: Package X, version Y was not found` error (even though the package is under `~/.nuget/packages/`) means a corrupt cache entry. Delete it and re-restore.

**The cache directory is lowercase** — the error prints friendly casing (`DotNetProjectFile.Analyzers`) but the path is `~/.nuget/packages/dotnetprojectfile.analyzers`. `rm`-ing the casing from the error silently no-ops; lowercase the id first:

```bash
rm -rf ~/.nuget/packages/dotnetprojectfile.analyzers
dotnet restore SpreadCheetah.Test/SpreadCheetah.Test.csproj
```

Restore must target a project that references the package. Repeat for each distinct `NETSDK1064` (several packages can be corrupt at once). If it persists, the project's `obj/` is stale too — delete and restore again.

## Referencing SpreadCheetah from standalone projects

The global NuGet cache also contains the published `SpreadCheetah` package (older version, no latest APIs), so a `PackageReference` in a scratch/verification project will silently resolve to the wrong assembly. Reference the built DLL directly instead:

```xml
<Reference Include="SpreadCheetah">
  <HintPath>/f/_git/spreadcheetah/SpreadCheetah/bin/Debug/net10.0/SpreadCheetah.dll</HintPath>
</Reference>
```
