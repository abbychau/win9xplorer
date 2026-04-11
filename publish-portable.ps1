param(
    [string]$Runtime = "win-x64",
    [string]$Configuration = "Release",
    [string]$OutputDir = "artifacts/publish-portable/win-x64"
)

$ErrorActionPreference = "Stop"

Write-Host "[publish-portable] Cleaning output: $OutputDir"
if (Test-Path $OutputDir) {
    Remove-Item -Recurse -Force $OutputDir
}

Write-Host "[publish-portable] Publishing self-contained single-file"
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

Write-Host "[publish-portable] Done"
