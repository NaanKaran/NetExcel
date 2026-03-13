# ──────────────────────────────────────────────────────────────────────────────
# NetExcel Build Script (PowerShell — Windows / macOS / Linux)
# Prerequisites: .NET 10 SDK (https://dotnet.microsoft.com/download)
# ──────────────────────────────────────────────────────────────────────────────
$ErrorActionPreference = "Stop"

Write-Host "🔨 Restoring packages..." -ForegroundColor Cyan
dotnet restore NetExcel.sln

Write-Host "🏗️  Building solution..." -ForegroundColor Cyan
dotnet build NetExcel.sln -c Release --no-restore

Write-Host "🧪 Running tests..." -ForegroundColor Cyan
dotnet test tests/NetExcel.Tests/NetExcel.Tests.csproj `
    -c Release `
    --no-build `
    --verbosity normal `
    --collect:"XPlat Code Coverage"

Write-Host "📦 Packing NuGet packages..." -ForegroundColor Cyan
New-Item -ItemType Directory -Force -Path nupkgs | Out-Null

@(
    "src/NetExcel.Core",
    "src/NetExcel.DataFrame",
    "src/NetExcel.Excel",
    "src/NetExcel.Csv",
    "src/NetExcel.Streaming",
    "src/NetExcel.Formatting"
) | ForEach-Object {
    dotnet pack $_ -c Release --no-build -o nupkgs
}

Write-Host "✅ Build complete. Packages in ./nupkgs/" -ForegroundColor Green
Get-ChildItem nupkgs/*.nupkg | Format-Table Name, Length
