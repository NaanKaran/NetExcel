# Publishing NetXLCsv to NuGet.org

Step-by-step guide to building, packing, and pushing all NetXLCsv packages.

---

## Prerequisites

| Tool | Version | Notes |
|---|---|---|
| .NET SDK | 10.0+ | `dotnet --version` |
| NuGet API Key | — | Get from https://www.nuget.org/account/apikeys |

---

## Step 1 — Get a NuGet API Key

1. Sign in at https://www.nuget.org (create a free account if needed).
2. Go to **Account → API Keys → Create**.
3. Set scope: **Push new packages and package versions**.
4. Copy the key — you only see it once.

Store it safely (environment variable, secret manager, not in source code):

```powershell
# PowerShell — set for current session
$env:NUGET_API_KEY = "your-api-key-here"

# Or persist in Windows environment variables
[System.Environment]::SetEnvironmentVariable("NUGET_API_KEY", "your-api-key-here", "User")
```

---

## Step 2 — Build in Release Mode

```bash
# From the repo root
dotnet build -c Release
```

All 6 sub-libraries + the umbrella meta-package must build with zero errors before packing.

---

## Step 3 — Run Tests

```bash
dotnet test -c Release
```

All 54 tests must pass. Never publish a failing build.

---

## Step 4 — Pack All Packages

```bash
# Create output folder
mkdir nupkgs

# Pack sub-packages first (umbrella depends on them)
dotnet pack src/NetExcel.Core       -c Release --no-build -o ./nupkgs
dotnet pack src/NetExcel.DataFrame  -c Release --no-build -o ./nupkgs
dotnet pack src/NetExcel.Excel      -c Release --no-build -o ./nupkgs
dotnet pack src/NetExcel.Csv        -c Release --no-build -o ./nupkgs
dotnet pack src/NetExcel.Streaming  -c Release --no-build -o ./nupkgs
dotnet pack src/NetExcel.Formatting -c Release --no-build -o ./nupkgs

# Pack the umbrella meta-package last
dotnet pack src/NetExcel           -c Release --no-build -o ./nupkgs
```

This produces `.nupkg` and `.snupkg` (symbols) files in `./nupkgs/`:

```
nupkgs/
├── NetXLCsv.Core.1.0.0.nupkg
├── NetXLCsv.Core.1.0.0.snupkg
├── NetXLCsv.DataFrame.1.0.0.nupkg
├── NetXLCsv.DataFrame.1.0.0.snupkg
├── NetXLCsv.Excel.1.0.0.nupkg
├── NetXLCsv.Excel.1.0.0.snupkg
├── NetXLCsv.Csv.1.0.0.nupkg
├── NetXLCsv.Csv.1.0.0.snupkg
├── NetXLCsv.Streaming.1.0.0.nupkg
├── NetXLCsv.Streaming.1.0.0.snupkg
├── NetXLCsv.Formatting.1.0.0.nupkg
├── NetXLCsv.Formatting.1.0.0.snupkg
├── NetXLCsv.1.0.0.nupkg           ← umbrella (users install this one)
└── NetXLCsv.1.0.0.snupkg
```

---

## Step 5 — Inspect Before Pushing (Optional but Recommended)

```bash
# NuGet Package Explorer (GUI) — Windows
winget install NuGetPackageExplorer

# Or inspect via CLI
dotnet tool install -g NuGetPackageExplorer
```

Check that:
- `icon.png` is embedded ✓
- All namespaces and types are correct ✓
- `README.md` is included ✓

---

## Step 6 — Push to NuGet.org

**Push sub-packages first** (the umbrella lists them as dependencies — NuGet.org validation requires them to exist):

```bash
# Push sub-packages
dotnet nuget push "./nupkgs/NetXLCsv.Core.1.0.0.nupkg"       --api-key %NUGET_API_KEY% --source https://api.nuget.org/v3/index.json
dotnet nuget push "./nupkgs/NetXLCsv.DataFrame.1.0.0.nupkg"  --api-key %NUGET_API_KEY% --source https://api.nuget.org/v3/index.json
dotnet nuget push "./nupkgs/NetXLCsv.Excel.1.0.0.nupkg"      --api-key %NUGET_API_KEY% --source https://api.nuget.org/v3/index.json
dotnet nuget push "./nupkgs/NetXLCsv.Csv.1.0.0.nupkg"        --api-key %NUGET_API_KEY% --source https://api.nuget.org/v3/index.json
dotnet nuget push "./nupkgs/NetXLCsv.Streaming.1.0.0.nupkg"  --api-key %NUGET_API_KEY% --source https://api.nuget.org/v3/index.json
dotnet nuget push "./nupkgs/NetXLCsv.Formatting.1.0.0.nupkg" --api-key %NUGET_API_KEY% --source https://api.nuget.org/v3/index.json

# Push umbrella meta-package last
dotnet nuget push "./nupkgs/NetXLCsv.1.0.0.nupkg" --api-key %NUGET_API_KEY% --source https://api.nuget.org/v3/index.json

# Push all symbol packages (.snupkg) for debugger support
dotnet nuget push "./nupkgs/*.snupkg" --api-key %NUGET_API_KEY% --source https://api.nuget.org/v3/index.json --skip-duplicate
```

PowerShell version (using `$env:NUGET_API_KEY`):

```powershell
$key    = $env:NUGET_API_KEY
$source = "https://api.nuget.org/v3/index.json"

@(
  "NetXLCsv.Core", "NetXLCsv.DataFrame", "NetXLCsv.Excel",
  "NetXLCsv.Csv", "NetXLCsv.Streaming", "NetXLCsv.Formatting",
  "NetXLCsv"
) | ForEach-Object {
  dotnet nuget push ".\nupkgs\$_.1.0.0.nupkg" --api-key $key --source $source
}

# Symbols
dotnet nuget push ".\nupkgs\*.snupkg" --api-key $key --source $source --skip-duplicate
```

---

## Step 7 — Verify on NuGet.org

After a few minutes (indexing takes ~5–10 min):

```bash
dotnet add package NetXLCsv
```

Check the listing at: https://www.nuget.org/packages/NetXLCsv

---

## Releasing a New Version

1. Bump `<Version>` in **all** sub-project `.csproj` files to the same number (e.g. `1.1.0`).
2. Repeat Steps 2–7.
3. Use `--skip-duplicate` flag to avoid errors if you accidentally re-push the same version.

---

## Semantic Versioning Rules

| Change Type | Version Bump | Example |
|---|---|---|
| Bug fix, no API change | Patch | `1.0.0` → `1.0.1` |
| New feature, backward-compatible | Minor | `1.0.0` → `1.1.0` |
| Breaking API change | Major | `1.0.0` → `2.0.0` |

---

## Private / GitHub Packages Feed (Alternative)

If you don't want to publish to NuGet.org publicly, you can use GitHub Packages:

```bash
dotnet nuget add source \
  --username YOUR_GITHUB_USERNAME \
  --password YOUR_GITHUB_PAT \
  --store-password-in-clear-text \
  --name github \
  "https://nuget.pkg.github.com/YOUR_GITHUB_ORG/index.json"

dotnet nuget push "./nupkgs/*.nupkg" --source github
```

This lets you keep packages private while still using `dotnet add package NetXLCsv` within your org.
