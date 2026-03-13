#!/usr/bin/env bash
# ──────────────────────────────────────────────────────────────────────────────
# NetExcel Build Script
# Prerequisites: .NET 10 SDK (https://dotnet.microsoft.com/download)
# ──────────────────────────────────────────────────────────────────────────────
set -e

echo "🔨 Restoring packages..."
dotnet restore NetExcel.sln

echo "🏗️  Building solution..."
dotnet build NetExcel.sln -c Release --no-restore

echo "🧪 Running tests..."
dotnet test tests/NetExcel.Tests/NetExcel.Tests.csproj \
    -c Release \
    --no-build \
    --verbosity normal \
    --collect:"XPlat Code Coverage"

echo "📦 Packing NuGet packages..."
mkdir -p nupkgs
for proj in src/NetExcel.Core src/NetExcel.DataFrame src/NetExcel.Excel \
            src/NetExcel.Csv src/NetExcel.Streaming src/NetExcel.Formatting; do
    dotnet pack "$proj" -c Release --no-build -o nupkgs
done

echo "✅ Build complete. Packages in ./nupkgs/"
ls -la nupkgs/*.nupkg
