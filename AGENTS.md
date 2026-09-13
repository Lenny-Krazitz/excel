# Repository Guidelines

## Project Structure & Module Organization

- `src/FormulaNavigator.Core/`: formula parsing, reference parsing, and navigation-tree projection; targets `net48` and `net10.0` without Excel dependencies.
- `src/FormulaNavigator.AddIn/`: Excel-DNA entry points, the `Excel/` COM gateway, and `UI/` WPF windows. Layouts are constructed in C# in `*.Layout.cs`; no XAML compilation is required.
- `tests/FormulaNavigator.Core.Tests/`: executable test harness and XLL verifier. `tests/fixtures/Navigation.xlsx` supports manual Excel checks; root GIFs illustrate expected interactions.
- `tools/`: Docker build/export script and fixture generator. `artifacts/` contains exported deliverables. See `BUILD_AND_INSTALL.md` for installation and troubleshooting.

## Build, Test, and Development Commands

Run from the repository root:

- `bash tools/build-in-docker.sh`: restores dependencies, runs tests, builds the x64 add-in, verifies its PE format, and exports `artifacts/FormulaNavigator64.xll`. Supports Buildx and legacy Docker builds.
- `bash tools/build-in-docker.sh --network=host`: uses host networking when container access to NuGet fails.
- `dotnet run --project tests/FormulaNavigator.Core.Tests/FormulaNavigator.Core.Tests.csproj --configuration Release`: runs core tests inside the SDK build environment.
- `python3 tools/create_demo_workbook.py`: regenerates the deterministic Excel fixture using Python's standard library.

Keep Microsoft build tooling inside Docker; do not install it on the host automatically. Load the resulting XLL in Windows Excel 64-bit; it cannot run on Linux.

## Coding Style & Naming Conventions

Use C# 7.3, four-space indentation, Allman braces, explicit imports, PascalCase types/methods/properties, and camelCase parameters/locals. Follow each file's private-field convention. No formatter or linter is configured.

Keep parsing independent of COM, preserve original formula spans for highlighting, and keep Excel COM calls on its STA thread. Preserve modal keyboard handling and perform final navigation after dialogs close.

## Testing Guidelines

Tests use a custom console runner, not a test framework. Add PascalCase test methods with descriptive `Run("behavior", Method)` labels. Cover parser edge cases, source offsets, and reference projection when changing those behaviors. No coverage percentage is enforced.

For UI changes, check keyboard focus, Enter/Esc, highlighting, and positioning in Windows Excel using the fixture. Report explicitly when Windows verification is unavailable.

## Commit & Pull Request Guidelines

History contains `init` and `up`; no formal convention is established. Prefer concise imperative summaries, such as `Fix formula reference highlighting`.

Keep changes focused. Describe the problem, resulting behavior, validation, and limitations. Link relevant issues and include screenshots for visible UI changes. Update installation instructions when requirements or commands change.
