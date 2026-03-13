// Global usings for the test project.
// 'using Xunit' is not included in ImplicitUsings, so we add it here once
// rather than repeating it in every test file.
global using Xunit;

// NetDataFrame was moved to NetXLCsv.Core to break circular dependencies.
// The global alias in NetXLCsv.DataFrame is project-scoped and does not
// propagate to consumers, so we re-export it here for test convenience.
global using NetDataFrame = NetXLCsv.Core.NetDataFrame;

// 'DataFrame' is both the namespace (NetXLCsv.DataFrame) and the static façade class
// (NetXLCsv.DataFrame.DataFrame). The C# spec evaluates ancestor-namespace member lookup
// before alias resolution, so tests inside namespace NetXLCsv.* would always resolve
// 'DataFrame' as the namespace rather than the alias. The test namespace is therefore
// 'NetXLCsvTests' (no dot) so the ancestor lookup never finds NetXLCsv.DataFrame.
global using DataFrame = NetXLCsv.DataFrame.DataFrame;
