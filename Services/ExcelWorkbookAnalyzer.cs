using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDoctor.Models;
using Microsoft.AspNetCore.Components.Forms;

namespace ExcelDoctor.Services;

public sealed partial class ExcelWorkbookAnalyzer
{
    private const long MaxFileSize = 50 * 1024 * 1024;

    private static readonly string[] LookupFunctions =
    [
        "VLOOKUP(",
        "HLOOKUP(",
        "XLOOKUP(",
        "XMATCH(",
        "MATCH(",
        "INDEX("
    ];

    private static readonly string[] VolatileFunctions =
    [
        "INDIRECT(",
        "OFFSET(",
        "NOW(",
        "TODAY(",
        "RAND(",
        "RANDBETWEEN("
    ];

    private static readonly string[] SuspiciousVbaKeywords =
    [
        "Auto_Open",
        "Workbook_Open",
        "Shell",
        "WScript.Shell",
        "CreateObject",
        "PowerShell",
        "cmd.exe",
        "URLDownloadToFile",
        "FileSystemObject",
        "OpenTextFile",
        "CreateTextFile",
        "CopyFile",
        "MoveFile",
        "DeleteFile",
        "Kill "
    ];

    private static readonly string[] DataAccessVbaKeywords =
    [
        "ADODB",
        "DAO.",
        "ConnectionString",
        "Provider=",
        "Data Source=",
        "SQL Server",
        "Oracle",
        "Outlook.Application",
        "Word.Application",
        "XMLHTTP",
        "WinHttp"
    ];

    private static readonly string[] ConnectionHints =
    [
        "SQL Server",
        "SQLOLEDB",
        "MSOLEDBSQL",
        "Oracle",
        "OraOLEDB",
        "Microsoft.ACE.OLEDB",
        "Microsoft.Jet.OLEDB",
        "ODBC",
        "MySQL",
        "PostgreSQL"
    ];

    public async Task<ExcelDiagnosticsResult> AnalyzeAsync(IBrowserFile file, bool includeVbaDetails = false)
    {
        ValidateFile(file);

        await using var browserStream = file.OpenReadStream(MaxFileSize);
        using var workbookStream = new MemoryStream();
        await browserStream.CopyToAsync(workbookStream);
        workbookStream.Position = 0;

        using var document = SpreadsheetDocument.Open(workbookStream, false);
        var workbookPart = document.WorkbookPart
            ?? throw new InvalidOperationException("Excel ブックの構造を読み取れませんでした。");

        var structure = AnalyzeStructure(workbookPart);
        var logic = AnalyzeLogic(workbookPart);
        var operation = AnalyzeOperation(workbookPart);
        var vba = AnalyzeVba(workbookPart, includeVbaDetails);
        var findings = BuildFindings(file.Size, structure, logic, operation, vba);
        var score = CalculateHealthScore(file.Size, structure, logic, operation, vba);

        return new ExcelDiagnosticsResult
        {
            FileName = file.Name,
            FileSizeBytes = file.Size,
            HealthScore = score,
            Rank = ToRank(score),
            Structure = structure,
            Logic = logic,
            Operation = operation,
            Vba = vba,
            Findings = findings,
            RecommendedActions = BuildRecommendedActions(findings, logic, operation, vba)
        };
    }

    private static void ValidateFile(IBrowserFile file)
    {
        var extension = Path.GetExtension(file.Name);
        if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(".xlsx または .xlsm ファイルを選択してください。");
        }

        if (file.Size > MaxFileSize)
        {
            throw new InvalidOperationException("50MB 以下の Excel ファイルを選択してください。");
        }
    }

    private static WorkbookStructureMetrics AnalyzeStructure(WorkbookPart workbookPart)
    {
        var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList() ?? [];
        var hiddenSheetCount = sheets.Count(sheet =>
            sheet.State?.Value.Equals(SheetStateValues.Hidden) == true ||
            sheet.State?.Value.Equals(SheetStateValues.VeryHidden) == true);

        return new WorkbookStructureMetrics(
            sheets.Count,
            hiddenSheetCount,
            workbookPart.WorksheetParts.Sum(part => part.PivotTableParts.Count()));
    }

    private static WorkbookLogicMetrics AnalyzeLogic(WorkbookPart workbookPart)
    {
        var formulas = workbookPart.WorksheetParts
            .SelectMany(part => part.Worksheet.Descendants<CellFormula>())
            .Select(formula => formula.Text ?? string.Empty)
            .ToList();

        var upperFormulas = formulas.Select(formula => formula.ToUpperInvariant()).ToList();
        var definedNames = workbookPart.Workbook.DefinedNames?.Elements<DefinedName>().ToList() ?? [];
        var definedNameTexts = definedNames.Select(name => name.Text ?? string.Empty).ToList();

        var externalReferenceCount =
            upperFormulas.Count(formula => formula.Contains('[')) +
            definedNameTexts.Count(name => name.Contains('[')) +
            workbookPart.ExternalWorkbookParts.Count();

        return new WorkbookLogicMetrics(
            formulas.Count,
            CountFunctionUsage(upperFormulas, LookupFunctions),
            CountFunctionUsage(upperFormulas, VolatileFunctions),
            externalReferenceCount,
            definedNames.Count);
    }

    private static WorkbookOperationMetrics AnalyzeOperation(WorkbookPart workbookPart)
    {
        var worksheets = workbookPart.WorksheetParts.Select(part => part.Worksheet).ToList();

        return new WorkbookOperationMetrics(
            worksheets.Sum(sheet => sheet.Descendants<MergeCell>().Count()),
            worksheets.Sum(sheet => sheet.Descendants<ConditionalFormattingRule>().Count()),
            worksheets.Sum(sheet => sheet.Descendants<DataValidation>().Count()),
            worksheets.Count(sheet => sheet.Elements<SheetProtection>().Any()));
    }

    private static VbaDiagnostics AnalyzeVba(WorkbookPart workbookPart, bool includeDetails)
    {
        if (workbookPart.VbaProjectPart is null)
        {
            return new VbaDiagnostics(false, false, false, false, 0, 0, 0, 0, 0, [], [], [], [], [], [], []);
        }

        using var stream = workbookPart.VbaProjectPart.GetStream(FileMode.Open, FileAccess.Read);
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        var bytes = memory.ToArray();

        var binarySearchableText = BuildSearchableBinaryText(bytes);
        var isPasswordProtected =
            binarySearchableText.Contains("DPB=", StringComparison.OrdinalIgnoreCase) ||
            binarySearchableText.Contains("GC=", StringComparison.OrdinalIgnoreCase);

        if (!includeDetails)
        {
            return new VbaDiagnostics(true, false, isPasswordProtected, false, 0, 0, 0, 0, 0, [], [], [], [], [], [], []);
        }

        var sourceCodeParsed = false;
        var sourceLineCount = 0;
        var moduleDiagnostics = new List<VbaModuleDiagnostics>();
        var searchableText = binarySearchableText;

        try
        {
            var projectSource = VbaProjectSourceReader.Read(bytes);
            if (projectSource.Modules.Count > 0)
            {
                sourceCodeParsed = true;
                searchableText = string.Join('\n', projectSource.Modules.Select(module => module.SourceText));
                moduleDiagnostics = projectSource.Modules
                    .Select(module => AnalyzeVbaModule(module.Name, module.SourceText))
                    .ToList();
                sourceLineCount = moduleDiagnostics.Sum(module => module.LineCount);
            }
        }
        catch (InvalidDataException)
        {
            // Keep the previous binary-trace based behavior when source extraction is not available.
        }
        catch (ArgumentException)
        {
            // Malformed OLE/VBA projects should not block the workbook-level diagnosis.
        }
        catch (OverflowException)
        {
            // Corrupt sector chains can produce invalid offsets; fall back to binary traces.
        }
        catch (IndexOutOfRangeException)
        {
            // Corrupt directory records can produce invalid offsets; fall back to binary traces.
        }

        var signals = AnalyzeVbaSignals(searchableText);

        return new VbaDiagnostics(
            true,
            true,
            isPasswordProtected,
            sourceCodeParsed,
            moduleDiagnostics.Count,
            sourceLineCount,
            signals.DetectedSuspicious.Count,
            signals.DetectedDependencies.Count,
            signals.PersonalInfoPatternCount,
            signals.DetectedSuspicious,
            signals.DetectedDependencies,
            signals.DetectedObjectTargets,
            signals.DetectedFileTargets,
            signals.DetectedConnectionHints,
            signals.DetectedSqlTables,
            moduleDiagnostics);
    }

    private static VbaModuleDiagnostics AnalyzeVbaModule(string name, string sourceText)
    {
        var signals = AnalyzeVbaSignals(sourceText);

        return new VbaModuleDiagnostics(
            name,
            CountSourceLines(sourceText),
            signals.DetectedSuspicious.Count,
            signals.DetectedDependencies.Count,
            signals.PersonalInfoPatternCount,
            signals.DetectedSuspicious,
            signals.DetectedDependencies,
            signals.DetectedObjectTargets,
            signals.DetectedFileTargets,
            signals.DetectedConnectionHints,
            signals.DetectedSqlTables);
    }

    private static VbaSignals AnalyzeVbaSignals(string searchableText)
    {
        var detectedSuspicious = SuspiciousVbaKeywords
            .Where(keyword => searchableText.Contains(keyword, StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Order(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var detectedDependencies = DataAccessVbaKeywords
            .Where(keyword => searchableText.Contains(keyword, StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Order(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (SqlSelectStatementRegex().IsMatch(searchableText))
        {
            detectedDependencies.Add("SQL SELECT");
        }

        var personalInfoPatternCount =
            EmailRegex().Matches(searchableText).Count +
            UncOrIpPathRegex().Matches(searchableText).Count +
            UrlRegex().Matches(searchableText).Count;

        var detectedConnectionHints = ConnectionHints
            .Where(keyword => searchableText.Contains(keyword, StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Order(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var detectedSqlTables = SqlTableRegex().Matches(searchableText)
            .Select(match => match.Groups["table"].Value.Trim('[', ']', '"', '`'))
            .Where(table => !string.IsNullOrWhiteSpace(table))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Order(StringComparer.OrdinalIgnoreCase)
            .Take(30)
            .ToList();

        var detectedObjectTargets = CreateObjectTargetRegex().Matches(searchableText)
            .Select(match => match.Groups["target"].Value.Trim())
            .Where(target => !string.IsNullOrWhiteSpace(target))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Order(StringComparer.OrdinalIgnoreCase)
            .Take(30)
            .ToList();

        var detectedFileTargets = FileTargetRegex().Matches(searchableText)
            .Select(match => match.Groups["path"].Value.Trim())
            .Concat(FileSystemObjectTargetRegex().Matches(searchableText)
                .Select(match => FormatFileSystemObjectTarget(match.Groups["method"].Value, match.Groups["target"].Value)))
            .Concat(VbaOpenStatementTargetRegex().Matches(searchableText)
                .Select(match => $"Open: {match.Groups["target"].Value.Trim()}"))
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Order(StringComparer.OrdinalIgnoreCase)
            .Take(30)
            .ToList();

        return new VbaSignals(
            detectedSuspicious,
            detectedDependencies,
            detectedObjectTargets,
            detectedFileTargets,
            detectedConnectionHints,
            detectedSqlTables,
            personalInfoPatternCount);
    }

    private static int CountSourceLines(string sourceText)
    {
        if (string.IsNullOrEmpty(sourceText))
        {
            return 0;
        }

        return sourceText.Count(character => character == '\n') + 1;
    }

    private static string FormatFileSystemObjectTarget(string method, string target)
    {
        var cleanTarget = target.Trim().Trim('"', '\'');
        return string.IsNullOrWhiteSpace(cleanTarget)
            ? string.Empty
            : $"{method}: {cleanTarget}";
    }

    private static string BuildSearchableBinaryText(byte[] bytes)
    {
        var ascii = Encoding.Latin1.GetString(bytes);
        var unicode = Encoding.Unicode.GetString(bytes);
        return ascii + "\n" + unicode;
    }

    private static int CountFunctionUsage(IReadOnlyList<string> formulas, string[] functionNames)
    {
        return formulas.Sum(formula =>
            functionNames.Count(functionName => formula.Contains(functionName, StringComparison.Ordinal)));
    }

    private static List<DiagnosticFinding> BuildFindings(
        long fileSize,
        WorkbookStructureMetrics structure,
        WorkbookLogicMetrics logic,
        WorkbookOperationMetrics operation,
        VbaDiagnostics vba)
    {
        var findings = new List<DiagnosticFinding>();

        if (fileSize >= 10 * 1024 * 1024)
        {
            findings.Add(new("保守性", "High", "ファイルサイズが大きい", "ブックの肥大化により、開く・保存する・共有する操作が重くなる可能性があります。"));
        }

        if (structure.HiddenSheetCount > 0)
        {
            findings.Add(new("属人化", "Medium", "非表示シートがあります", "隠れた前提データや計算ロジックが存在する可能性があります。"));
        }

        if (logic.TotalFormulaCount >= 1000)
        {
            findings.Add(new("保守性", "High", "数式が多く含まれています", "計算負荷と変更影響範囲が大きく、仕様把握に時間がかかる状態です。"));
        }
        else if (logic.TotalFormulaCount >= 300)
        {
            findings.Add(new("保守性", "Medium", "数式がやや多い", "主要な計算ブロックを棚卸しして、ロジックの重複を確認してください。"));
        }

        if (logic.LookupFunctionCount >= 100)
        {
            findings.Add(new("保守性", "High", "参照関数が多用されています", "検索・参照系の数式が多く、再計算の重さや参照切れの原因になります。"));
        }

        if (logic.VolatileFunctionCount > 0)
        {
            findings.Add(new("保守性", "Medium", "揮発性関数があります", "INDIRECT/OFFSET などは再計算負荷と参照追跡の難しさを上げます。"));
        }

        if (logic.ExternalLinkCount > 0)
        {
            findings.Add(new("業務依存", "High", "外部リンクがあります", "別ブックやローカルパスに依存しており、他環境で壊れる可能性があります。"));
        }

        if (logic.DefinedNameCount >= 50)
        {
            findings.Add(new("属人化", "Medium", "名前定義が多い", "古い名前定義や参照切れが埋もれている可能性があります。"));
        }

        if (operation.MergedCellCount >= 50)
        {
            findings.Add(new("自動化難度", "Medium", "セル結合が多い", "データ取り込みや機械処理の前処理コストが高くなります。"));
        }

        if (operation.ConditionalFormatRuleCount >= 100 || operation.DataValidationCount >= 100)
        {
            findings.Add(new("運用", "Medium", "書式・入力規則が増殖しています", "コピペ運用で条件付き書式や入力規則が過剰に複製されている可能性があります。"));
        }

        if (operation.ProtectedSheetCount > 0)
        {
            findings.Add(new("属人化", "Medium", "保護シートがあります", "編集ロックされた範囲がブラックボックス化している可能性があります。"));
        }

        if (vba.HasMacro)
        {
            findings.Add(new("セキュリティ", "Medium", "VBA マクロを含みます", "マクロの処理内容と実行タイミングを確認してください。"));
        }

        if (vba.SourceCodeParsed && vba.SourceLineCount >= 1000)
        {
            findings.Add(new("保守性", "High", "VBA コード量が多い", "モジュール全体の行数が多く、仕様把握と改修影響の確認に時間がかかる状態です。"));
        }
        else if (vba.SourceCodeParsed && vba.SourceLineCount >= 300)
        {
            findings.Add(new("保守性", "Medium", "VBA コード量がやや多い", "主要なマクロ処理をモジュール単位で棚卸ししてください。"));
        }

        if (vba.SuspiciousKeywordCount > 0)
        {
            findings.Add(new("セキュリティ", "High", "注意が必要な VBA 構文があります", "自動実行、外部プログラム実行、ファイル操作に関連する語句を検出しました。"));
        }

        if (vba.DataAccessKeywordCount > 0)
        {
            findings.Add(new("業務依存", "High", "VBA 内に外部接続の痕跡があります", "DB、HTTP、他 Office 製品との連携に依存している可能性があります。"));
        }

        if (vba.PersonalInfoPatternCount > 0)
        {
            findings.Add(new("属人化", "High", "ハードコード情報の痕跡があります", "メールアドレス、URL、社内パス、IP アドレスのような固定値を検出しました。"));
        }

        if (vba.IsPasswordProtected)
        {
            findings.Add(new("属人化", "High", "VBA プロジェクト保護の痕跡があります", "マクロの中身を確認できない場合、保守担当者への依存が強くなります。"));
        }

        return findings;
    }

    private static IReadOnlyList<string> BuildRecommendedActions(
        IReadOnlyList<DiagnosticFinding> findings,
        WorkbookLogicMetrics logic,
        WorkbookOperationMetrics operation,
        VbaDiagnostics vba)
    {
        var actions = new List<string>();

        if (logic.ExternalLinkCount > 0)
        {
            actions.Add("外部リンク一覧を棚卸しし、共有フォルダ・DB・Web API など管理された参照先へ置き換える。");
        }

        if (logic.LookupFunctionCount >= 100 || logic.VolatileFunctionCount > 0)
        {
            actions.Add("VLOOKUP/INDIRECT/OFFSET の多い範囲を特定し、Power Query または DB 化で再計算負荷を下げる。");
        }

        if (operation.MergedCellCount >= 50)
        {
            actions.Add("入力シートと帳票シートを分離し、データ取り込み対象からセル結合を排除する。");
        }

        if (vba.HasMacro)
        {
            actions.Add("VBA の自動実行・外部接続・ファイル操作をレビューし、処理フローを仕様書化する。");
        }

        if (vba.SourceCodeParsed && vba.Modules.Count > 0)
        {
            actions.Add("検出語句のある VBA モジュールから優先的に確認し、DB 接続・ファイル操作・自動実行処理を分離する。");
        }

        if (findings.Count == 0)
        {
            actions.Add("現時点では大きなリスクは少ないため、シート構成と主要数式の概要をドキュメント化する。");
        }

        return actions;
    }

    private static int CalculateHealthScore(
        long fileSize,
        WorkbookStructureMetrics structure,
        WorkbookLogicMetrics logic,
        WorkbookOperationMetrics operation,
        VbaDiagnostics vba)
    {
        var penalty = 0;

        penalty += Math.Min(15, (int)(fileSize / 1024d / 1024d / 2));
        penalty += Math.Min(10, structure.HiddenSheetCount * 3);
        penalty += Math.Min(8, structure.PivotTableCount);
        penalty += Math.Min(20, logic.TotalFormulaCount / 120);
        penalty += Math.Min(12, logic.LookupFunctionCount / 20);
        penalty += Math.Min(10, logic.VolatileFunctionCount * 2);
        penalty += Math.Min(15, logic.ExternalLinkCount * 5);
        penalty += Math.Min(8, logic.DefinedNameCount / 10);
        penalty += Math.Min(10, operation.MergedCellCount / 20);
        penalty += Math.Min(8, (operation.ConditionalFormatRuleCount + operation.DataValidationCount) / 50);
        penalty += Math.Min(8, operation.ProtectedSheetCount * 2);

        if (vba.HasMacro)
        {
            penalty += 6;
        }

        penalty += Math.Min(10, vba.SourceLineCount / 250);
        penalty += Math.Min(15, vba.SuspiciousKeywordCount * 5);
        penalty += Math.Min(12, vba.DataAccessKeywordCount * 4);
        penalty += Math.Min(10, vba.PersonalInfoPatternCount * 2);
        penalty += vba.IsPasswordProtected ? 8 : 0;

        return Math.Clamp(100 - penalty, 0, 100);
    }

    private static string ToRank(int score) => score switch
    {
        >= 85 => "A",
        >= 70 => "B",
        >= 55 => "C",
        >= 40 => "D",
        _ => "E"
    };

    private sealed record VbaSignals(
        IReadOnlyList<string> DetectedSuspicious,
        IReadOnlyList<string> DetectedDependencies,
        IReadOnlyList<string> DetectedObjectTargets,
        IReadOnlyList<string> DetectedFileTargets,
        IReadOnlyList<string> DetectedConnectionHints,
        IReadOnlyList<string> DetectedSqlTables,
        int PersonalInfoPatternCount);

    [GeneratedRegex(@"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", RegexOptions.IgnoreCase)]
    private static partial Regex EmailRegex();

    [GeneratedRegex(@"(\\\\[A-Z0-9_.-]+\\[^\s""']+)|(\\\\\d{1,3}(\.\d{1,3}){3}\\[^\s""']+)", RegexOptions.IgnoreCase)]
    private static partial Regex UncOrIpPathRegex();

    [GeneratedRegex(@"https?://[^\s""']+", RegexOptions.IgnoreCase)]
    private static partial Regex UrlRegex();

    [GeneratedRegex(@"\b(?:FROM|JOIN)\s+(?<table>(?:\[?[A-Z0-9_.$#]+\]?\.){0,2}\[?[A-Z0-9_.$#]+\]?)", RegexOptions.IgnoreCase)]
    private static partial Regex SqlTableRegex();

    [GeneratedRegex(@"\bSELECT\b[\s\S]{1,2000}?\bFROM\b", RegexOptions.IgnoreCase)]
    private static partial Regex SqlSelectStatementRegex();

    [GeneratedRegex(@"\b(?:CreateObject|GetObject)\s*\(\s*[""'](?<target>[^""']+)[""']", RegexOptions.IgnoreCase)]
    private static partial Regex CreateObjectTargetRegex();

    [GeneratedRegex(@"[""'](?<path>(?:[A-Z]:\\|\\\\)[^""'\r\n]+|[^""'\r\n]+\.(?:xlsx|xlsm|xls|csv|txt|xml|json|accdb|mdb|sql|bat|ps1|vbs|exe|pdf))[""']", RegexOptions.IgnoreCase)]
    private static partial Regex FileTargetRegex();

    [GeneratedRegex(@"\.\s*(?<method>OpenTextFile|CreateTextFile|GetFile|GetFolder|FileExists|FolderExists|CopyFile|CopyFolder|MoveFile|MoveFolder|DeleteFile|DeleteFolder|BuildPath)\s*\(\s*(?<target>[^,\)\r\n]+)", RegexOptions.IgnoreCase)]
    private static partial Regex FileSystemObjectTargetRegex();

    [GeneratedRegex(@"^\s*Open\s+(?<target>.+?)\s+For\s+(?:Input|Output|Append|Binary|Random)\b", RegexOptions.IgnoreCase | RegexOptions.Multiline)]
    private static partial Regex VbaOpenStatementTargetRegex();
}
