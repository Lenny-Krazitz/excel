# syntax=docker/dockerfile:1
# Build on Linux; the resulting XLL is loaded only by Windows Excel 64-bit.
FROM mcr.microsoft.com/dotnet/sdk:10.0 AS builder
ENV DOTNET_CLI_TELEMETRY_OPTOUT=1 \
    DOTNET_NOLOGO=1 \
    NUGET_XMLDOC_MODE=skip
WORKDIR /src

COPY global.json Directory.Build.props FormulaNavigator.sln ./
COPY src/FormulaNavigator.Core/FormulaNavigator.Core.csproj src/FormulaNavigator.Core/
COPY src/FormulaNavigator.AddIn/FormulaNavigator.AddIn.csproj src/FormulaNavigator.AddIn/
COPY tests/FormulaNavigator.Core.Tests/FormulaNavigator.Core.Tests.csproj tests/FormulaNavigator.Core.Tests/
RUN dotnet restore FormulaNavigator.sln

COPY src/ src/
COPY tests/ tests/
COPY BUILD_AND_INSTALL.md ./
RUN dotnet run --project tests/FormulaNavigator.Core.Tests/FormulaNavigator.Core.Tests.csproj --configuration Release --no-restore
RUN dotnet build src/FormulaNavigator.AddIn/FormulaNavigator.AddIn.csproj --configuration Release --no-restore
RUN dotnet run --project tests/FormulaNavigator.Core.Tests/FormulaNavigator.Core.Tests.csproj --configuration Release --no-build -- \
    --verify-xll /src/src/FormulaNavigator.AddIn/bin/Release/net48/publish/FormulaNavigator64.xll
RUN dotnet run --project tests/FormulaNavigator.Core.Tests/FormulaNavigator.Core.Tests.csproj --configuration Release --no-build -- \
    --verify-addin /src/src/FormulaNavigator.AddIn/bin/Release/net48/FormulaNavigator.AddIn.dll

# Build and export with: bash tools/build-in-docker.sh (Buildx is optional).
# With Buildx: docker buildx build --output type=local,dest=./artifacts .
# This stage contains distributable files, not a runnable Linux application.
FROM scratch AS artifact
COPY --from=builder /src/src/FormulaNavigator.AddIn/bin/Release/net48/publish/FormulaNavigator64.xll /FormulaNavigator64.xll
COPY --from=builder /src/tests/fixtures/Navigation.xlsx /Navigation.xlsx
COPY --from=builder /src/BUILD_AND_INSTALL.md /BUILD_AND_INSTALL.md
