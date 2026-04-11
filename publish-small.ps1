param(
    [string]$Runtime = "win-x64",
    [string]$Configuration = "Release",
    [string]$OutputDir = "artifacts/publish-fd/win-x64"
)

$ErrorActionPreference = "Stop"

Write-Host "[publish-small] Cleaning output: $OutputDir"
if (Test-Path $OutputDir) {
    Remove-Item -Recurse -Force $OutputDir
}

Write-Host "[publish-small] Publishing framework-dependent single-file"
dotnet publish .\win9xplorer.csproj `
    -c $Configuration `
    -r $Runtime `
    --self-contained false `
    -p:PublishSingleFile=true `
    -p:EnableCompressionInSingleFile=true `
    -p:DebugSymbols=false `
    -p:DebugType=None `
    -o $OutputDir

Write-Host "[publish-small] Done"
