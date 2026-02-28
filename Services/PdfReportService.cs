using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using ThinkReport.Models;

namespace ThinkReport.Services;

public sealed class PdfReportService : IPdfReportService
{
    private const string Blue   = "#1F497D";
    private const string Gray   = "#595959";
    private const string BgCell = "#EEF3F9";
    private const string Border = "#D0D0D0";

    public byte[] Generate(
        IncidentReportViewModel model,
        IReadOnlyList<(string FileName, byte[] Data)> images, byte[] logoBytes)
    {
        var generatedAt = DateTime.UtcNow.ToString("dd/MM/yyyy HH:mm:ss") + " UTC";
        var severity    = SeverityLabel(model.Severity);
        var bullets     = ParseBulletLines(model.RecommendedActions ?? string.Empty).ToList();

        return Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.A4);
                page.MarginHorizontal(2.2f, Unit.Centimetre);
                page.MarginVertical(2f, Unit.Centimetre);
                page.DefaultTextStyle(s => s.FontFamily("Calibri").FontSize(10.5f));

                page.Header().Column(h =>
                {
                    if (logoBytes != null && logoBytes.Length > 0)
                    {
                        h.Item().Height(40).AlignCenter().Image(logoBytes).FitHeight();;
                    }
                    h.Item().AlignCenter()
                        .Text("RELATÓRIO DE INCIDENTE DE SEGURANÇA")
                        .Bold().FontSize(17).FontColor(Blue);

                    h.Item().AlignCenter()
                        .Text("ThinkIT Security Operations Center")
                        .FontSize(10).FontColor(Gray);

                    h.Item().AlignCenter()
                        .Text($"Gerado em: {generatedAt}")
                        .FontSize(9).FontColor(Gray);

                    h.Item().PaddingTop(6).LineHorizontal(1.5f).LineColor(Blue);
                });

                page.Content().PaddingTop(10).Column(col =>
                {
                    SectionTitle(col, "Sumário Executivo");
                    col.Item().Text(model.ExecutiveSummary).FontSize(10.5f);
                    col.Item().PaddingVertical(6).LineHorizontal(0.75f).LineColor(Blue);

                    SectionTitle(col, "1. Informações do Alerta");
                    InfoTable(col,
                    [
                        ("ID",      model.AlertId),
                        ("Título",            model.Title),
                        ("Severidade",        severity),
                        ("Data / Hora (UTC)", model.IncidentDateTimeUtc.ToString("yyyy-MM-dd HH:mm:ss") + " UTC"),
                        ("Nº Chamado ITSM",   model.ItsmTicketNumber  ?? "N/A"),
                        ("Tática MITRE",      model.MitreTactic.Count > 0 ? string.Join(" | ", model.MitreTactic) : "N/A"),
                    ]);

                    SectionTitle(col, "2. Resumo do Evento");
                    col.Item().Text(model.EventSummary);

                    SectionTitle(col, "3. Detalhes Técnicos");
                    InfoTable(col,
                    [
                        ("Usuário",           model.AffectedUser  ?? "N/A"),
                        ("Endereço IP",       model.IpAddress     ?? "N/A"),
                        ("Hostname",          model.Host          ?? "N/A"),
                        ("Nome do Arquivo",   model.FileName      ?? "N/A"),
                        ("Hash (SHA1)",       model.Sha1Hash      ?? "N/A"),
                        ("Caminho (Path)",    model.FilePath      ?? "N/A"),
                        ("Assinatura",        model.FileSignature ?? "N/A"),
                    ]);

                    SectionTitle(col, "4. Avaliação e Ações");

                    if (model.SelectedSocAction == SocAction.Both)
                    {
                        FieldLabel(col, "Avaliação do SOC:");
                        col.Item().PaddingTop(2).Text(model.SocAssessment ?? "N/A");
                        col.Item().PaddingTop(6);
                        FieldLabel(col, "Ações Tomadas pelo SOC:");
                        col.Item().PaddingTop(2).Text(model.SocActionsTaken ?? "N/A");
                    }
                    else
                    {
                        FieldLabel(col, $"{model.SelectedSocAction.GetDisplayName()}:");
                        col.Item().PaddingTop(2).Text(model.SocAssessment ?? "N/A");
                    }

                    FieldLabel(col, "Ações Recomendadas:");
                    if (bullets.Count == 0)
                    {
                        col.Item().PaddingTop(2).Text("N/A");
                    }
                    else
                    {
                        col.Item().PaddingTop(2).Column(bCol =>
                        {
                            foreach (var (text, level) in bullets)
                            {
                                var marker  = level == 1 ? "•" : "◦";
                                var padding = level == 1 ? 6f : 18f;
                                bCol.Item().PaddingLeft(padding)
                                    .Text($"{marker}  {text}");
                            }
                        });
                    }

                    FieldLabel(col, "Observação Final:");
                    col.Item().PaddingTop(2).Text(model.FinalObservation ?? "N/A");

                    if (images.Count > 0)
                    {
                        col.Item().PageBreak();
                        SectionTitle(col, "Evidências do Evento");

                        uint idx = 1;
                        foreach (var (fileName, data) in images)
                        {
                            col.Item().PaddingTop(8)
                                .Text($"Figura {idx} — {fileName}")
                                .Italic().FontSize(9).FontColor(Gray);

                            col.Item().PaddingTop(4).Image(data).FitWidth();
                            col.Item().Height(12);
                            idx++;
                        }
                    }
                });

                page.Footer().AlignCenter().Text(t =>
                {
                    t.Span("CONFIDENCIAL — ThinkIT Security Operations Center  |  Pág. ")
                        .FontSize(8).FontColor(Colors.Grey.Medium);
                    t.CurrentPageNumber().FontSize(8);
                    t.Span(" de ").FontSize(8).FontColor(Colors.Grey.Medium);
                    t.TotalPages().FontSize(8);
                });
            });
        }).GeneratePdf();
    }

    private static void SectionTitle(ColumnDescriptor col, string title)
    {
        col.Item().PaddingTop(12).Text(title).Bold().FontSize(12).FontColor(Blue);
    }

    private static void FieldLabel(ColumnDescriptor col, string label)
    {
        col.Item().PaddingTop(8).Text(label).Bold();
    }

    private static void InfoTable(
        ColumnDescriptor col,
        IEnumerable<(string Label, string Value)> rows)
    {
        col.Item().PaddingTop(6).Table(table =>
        {
            table.ColumnsDefinition(c =>
            {
                c.RelativeColumn(2);
                c.RelativeColumn(3);
            });

            foreach (var (label, value) in rows)
            {
                table.Cell()
                    .Background(BgCell)
                    .BorderBottom(0.5f).BorderColor(Border)
                    .PaddingVertical(4).PaddingHorizontal(6)
                    .Text(label).Bold().FontSize(10);

                table.Cell()
                    .BorderBottom(0.5f).BorderColor(Border)
                    .PaddingVertical(4).PaddingHorizontal(6)
                    .Text(value ?? "N/A").FontSize(10);
            }
        });
    }

    private static string SeverityLabel(SeverityLevel s) => s switch
    {
        SeverityLevel.Informational => "Informacional",
        SeverityLevel.Low           => "Baixa",
        SeverityLevel.Medium        => "Média",
        SeverityLevel.High          => "Alta",
        SeverityLevel.Critical      => "Crítica",
        _                           => s.ToString()
    };

    private static IEnumerable<(string Text, int Level)> ParseBulletLines(string rawText)
    {
        foreach (var rawLine in rawText.Split('\n'))
        {
            var line = rawLine.TrimEnd('\r');
            if (string.IsNullOrWhiteSpace(line)) continue;

            if (System.Text.RegularExpressions.Regex.IsMatch(line, @"^\s{2,}"))
            {
                var t = line.TrimStart().TrimStart('-', '*', '•', '◦').Trim();
                if (!string.IsNullOrWhiteSpace(t)) yield return (t, 2);
            }
            else
            {
                var t = line.TrimStart().TrimStart('-', '*', '•').Trim();
                if (!string.IsNullOrWhiteSpace(t)) yield return (t, 1);
            }
        }
    }
}
