param(
    [string]$Runtime = "win-x64",
    [string]$Configuration = "Release",
    [string]$OutputDir = "artifacts/publish-single/win-x64"
)

$ErrorActionPreference = "Stop"

Write-Host "[publish-single] Cleaning output: $OutputDir"
if (Test-Path $OutputDir) {
    Remove-Item -Recurse -Force $OutputDir
}

Write-Host "[publish-single] Publishing single-file executable"
dotnet publish .\win9xplorer.csproj `
    -c $Configuration `
    -r $Runtime `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -p:DebugSymbols=false `
    -p:DebugType=None `
    -o $OutputDir

Write-Host "[publish-single] Done"
