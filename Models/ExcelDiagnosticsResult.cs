namespace ExcelDoctor.Models;

public sealed record ExcelDiagnosticsResult
{
    public required string FileName { get; init; }
    public required long FileSizeBytes { get; init; }
    public required int HealthScore { get; init; }
    public required string Rank { get; init; }
    public required WorkbookStructureMetrics Structure { get; init; }
    public required WorkbookLogicMetrics Logic { get; init; }
    public required WorkbookOperationMetrics Operation { get; init; }
    public required VbaDiagnostics Vba { get; init; }
    public required IReadOnlyList<DiagnosticFinding> Findings { get; init; }
    public required IReadOnlyList<string> RecommendedActions { get; init; }
}

public sealed record WorkbookStructureMetrics(
    int TotalSheetCount,
    int HiddenSheetCount,
    int PivotTableCount);

public sealed record WorkbookLogicMetrics(
    int TotalFormulaCount,
    int LookupFunctionCount,
    int VolatileFunctionCount,
    int ExternalLinkCount,
    int DefinedNameCount);

public sealed record WorkbookOperationMetrics(
    int MergedCellCount,
    int ConditionalFormatRuleCount,
    int DataValidationCount,
    int ProtectedSheetCount);

public sealed record VbaDiagnostics(
    bool HasMacro,
    bool IsPasswordProtected,
    int SuspiciousKeywordCount,
    int DataAccessKeywordCount,
    int PersonalInfoPatternCount,
    IReadOnlyList<string> DetectedKeywords,
    IReadOnlyList<string> DetectedDependencies,
    IReadOnlyList<string> DetectedConnectionHints,
    IReadOnlyList<string> DetectedSqlTables);

public sealed record DiagnosticFinding(
    string Category,
    string Severity,
    string Title,
    string Detail);
