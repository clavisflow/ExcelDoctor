window.excelDoctorReport = (() => {
    function escapeHtml(value) {
        return String(value ?? "")
            .replaceAll("&", "&amp;")
            .replaceAll("<", "&lt;")
            .replaceAll(">", "&gt;")
            .replaceAll('"', "&quot;")
            .replaceAll("'", "&#039;");
    }

    function list(items) {
        const values = Array.from(items ?? []).filter(Boolean);
        if (values.length === 0) {
            return '<p class="muted">なし</p>';
        }

        return `<ul>${values.map(item => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`;
    }

    function buildReportHtml(report) {
        const metrics = Array.from(report.metrics ?? [])
            .map(item => `
                <div class="metric">
                    <span>${escapeHtml(item.label)}</span>
                    <strong>${escapeHtml(item.value)}</strong>
                </div>`)
            .join("");

        const risks = Array.from(report.riskBars ?? [])
            .map(item => `
                <div class="risk">
                    <span>${escapeHtml(item.label)}</span>
                    <div><i style="width: ${Number(item.percent) || 0}%"></i></div>
                    <strong>${Number(item.percent) || 0}</strong>
                </div>`)
            .join("");

        const findings = Array.from(report.findings ?? [])
            .map(finding => `
                <article class="finding">
                    <div>
                        <span>${escapeHtml(finding.category)}</span>
                        <em>${escapeHtml(finding.severity)}</em>
                    </div>
                    <strong>${escapeHtml(finding.title)}</strong>
                    <p>${escapeHtml(finding.detail)}</p>
                </article>`)
            .join("") || '<p class="muted">大きなリスクは検出されませんでした。</p>';

        const actions = list(report.actions);
        const vba = report.vba ?? {};
        const modules = Array.from(vba.modules ?? [])
            .map(module => `
                <article class="module">
                    <div>
                        <strong>${escapeHtml(module.name)}</strong>
                        <span>${Number(module.lineCount) || 0} 行</span>
                    </div>
                    <p>${escapeHtml(module.summary)}</p>
                    <small>${escapeHtml(module.details)}</small>
                </article>`)
            .join("");

        return `<!doctype html>
<html lang="ja">
<head>
    <meta charset="utf-8">
    <title>Excel 健康診断レポート - ${escapeHtml(report.fileName)}</title>
    <style>
        @page { size: A4; margin: 14mm; }
        * { box-sizing: border-box; }
        body {
            margin: 0;
            color: #172033;
            background: #fff;
            font-family: "Yu Gothic UI", "Meiryo", Arial, sans-serif;
            line-height: 1.6;
        }
        main { max-width: 980px; margin: 0 auto; }
        h1, h2, h3, p { margin-top: 0; }
        h1 { margin-bottom: 6px; font-size: 28px; }
        h2 { margin: 26px 0 12px; border-bottom: 1px solid #d9e0ea; padding-bottom: 6px; font-size: 18px; }
        h3 { margin: 18px 0 8px; font-size: 14px; }
        ul { margin: 8px 0 0; padding-left: 1.2rem; }
        li { margin-bottom: 4px; overflow-wrap: anywhere; }
        .cover {
            border-bottom: 3px solid #111827;
            margin-bottom: 18px;
            padding-bottom: 16px;
        }
        .meta {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 10px;
            margin-top: 16px;
        }
        .box, .metric, .finding, .module {
            border: 1px solid #d9e0ea;
            border-radius: 8px;
            padding: 12px;
            background: #fff;
        }
        .box span, .metric span {
            display: block;
            color: #64748b;
            font-size: 12px;
            font-weight: 700;
        }
        .box strong, .metric strong {
            display: block;
            color: #111827;
            font-size: 24px;
            line-height: 1.25;
        }
        .summary {
            display: grid;
            grid-template-columns: 160px 1fr;
            gap: 18px;
            align-items: center;
            margin-top: 18px;
        }
        .score {
            display: grid;
            width: 136px;
            height: 136px;
            place-items: center;
            border: 12px solid #2563eb;
            border-radius: 50%;
        }
        .score strong { font-size: 34px; }
        .score span { color: #64748b; font-weight: 800; }
        .risk {
            display: grid;
            grid-template-columns: 110px 1fr 42px;
            align-items: center;
            gap: 10px;
            margin-bottom: 9px;
            font-size: 13px;
        }
        .risk div {
            height: 9px;
            border-radius: 999px;
            background: #e6ebf2;
            overflow: hidden;
        }
        .risk i {
            display: block;
            height: 100%;
            background: linear-gradient(90deg, #17803a, #b7791f, #b42318);
        }
        .metrics {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 10px;
        }
        .finding { margin-bottom: 8px; break-inside: avoid; }
        .finding div, .module div {
            display: flex;
            justify-content: space-between;
            gap: 12px;
            margin-bottom: 5px;
        }
        .finding span, .finding em, .module span, .module small, .muted {
            color: #64748b;
            font-size: 12px;
        }
        .finding em { font-style: normal; font-weight: 900; }
        .finding p, .module p { margin-bottom: 0; font-size: 13px; }
        .modules {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 10px;
        }
        .module { break-inside: avoid; }
        .module small { display: block; overflow-wrap: anywhere; }
    </style>
</head>
<body>
    <main>
        <section class="cover">
            <h1>Excel 健康診断レポート</h1>
            <p>${escapeHtml(report.assessmentTitle)}</p>
            <p class="muted">${escapeHtml(report.assessmentSummary)}</p>
            <div class="meta">
                <div class="box"><span>ファイル名</span><strong style="font-size: 14px;">${escapeHtml(report.fileName)}</strong></div>
                <div class="box"><span>サイズ</span><strong>${escapeHtml(report.fileSize)}</strong></div>
                <div class="box"><span>ランク</span><strong>${escapeHtml(report.rank)}</strong></div>
                <div class="box"><span>作成日時</span><strong style="font-size: 14px;">${escapeHtml(report.generatedAt)}</strong></div>
            </div>
        </section>
        <section class="summary">
            <div class="score"><div><strong>${Number(report.healthScore) || 0}</strong><br><span>Score</span></div></div>
            <div>
                <h2 style="margin-top: 0;">リスク分布</h2>
                ${risks}
                <p class="muted">${escapeHtml(report.scoreComment)}</p>
            </div>
        </section>
        <h2>診断メトリクス</h2>
        <section class="metrics">${metrics}</section>
        <h2>検出されたリスク</h2>
        ${findings}
        <h2>次に取るべきアクション</h2>
        ${actions}
        <h2>VBA マクロ診断</h2>
        <p>${escapeHtml(vba.note)}</p>
        <section class="metrics">
            <div class="metric"><span>解析方式</span><strong>${escapeHtml(vba.mode)}</strong></div>
            <div class="metric"><span>モジュール</span><strong>${Number(vba.moduleCount) || 0}</strong></div>
            <div class="metric"><span>VBA 行数</span><strong>${Number(vba.sourceLineCount) || 0}</strong></div>
            <div class="metric"><span>注意語句</span><strong>${Number(vba.suspiciousKeywordCount) || 0}</strong></div>
        </section>
        <h3>CreateObject / GetObject の対象</h3>
        ${list(vba.detectedObjectTargets)}
        <h3>対象ファイル・パス候補</h3>
        ${list(vba.detectedFileTargets)}
        <h3>接続先ヒント</h3>
        ${list(vba.detectedConnectionHints)}
        <h3>SQL テーブル候補</h3>
        ${list(vba.detectedSqlTables)}
        ${modules ? `<h3>モジュール別リスク</h3><section class="modules">${modules}</section>` : ""}
    </main>
</body>
</html>`;
    }

    function safeFileName(fileName) {
        return String(fileName ?? "excel")
            .replace(/\.[^.]+$/, "")
            .replace(/[\\/:*?"<>|]+/g, "_")
            .slice(0, 80);
    }

    function downloadHtmlReport(report) {
        const html = buildReportHtml(report);
        const blob = new Blob([html], { type: "text/html;charset=utf-8" });
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = `${safeFileName(report.fileName)}_ExcelDoctor_Report.html`;
        document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();
        URL.revokeObjectURL(url);
    }

    return {
        downloadHtmlReport
    };
})();
