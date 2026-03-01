using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ThinkReport.Tools;

public static class TemplateGenerator
{
    private const string Font = "Calibri";

    public static void EnsureTemplateExists(string templatePath)
    {
        if (File.Exists(templatePath) && IsCurrentVersion(templatePath)) return;

        if (File.Exists(templatePath)) File.Delete(templatePath);

        Directory.CreateDirectory(Path.GetDirectoryName(templatePath)!);
        CreateTemplate(templatePath);
    }

    private static bool IsCurrentVersion(string path)
    {
        try
        {
            using var doc = WordprocessingDocument.Open(path, isEditable: false);
            var text = string.Concat(
                doc.MainDocumentPart!.Document.Body!
                   .Descendants<Text>().Select(t => t.Text));
            var takenIdx = text.IndexOf("{{SOC_ACTIONS_TAKEN_LABEL}}", StringComparison.Ordinal);
            var socIdx   = text.IndexOf("{{SOC_ACTION_LABEL}}",        StringComparison.Ordinal);
            var refsIdx  = text.IndexOf("{{REFERENCES}}",              StringComparison.Ordinal);
            return takenIdx >= 0 && socIdx >= 0 && takenIdx < socIdx && refsIdx >= 0;
        }
        catch { return false; }
    }

    public static void CreateTemplate(string outputPath)
    {
        using var doc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);

        var mainPart = doc.AddMainDocumentPart();

        AddDefaultStyles(mainPart);

        mainPart.Document = new Document(BuildBody());
        mainPart.Document.Save();

        var coreProps = doc.AddCoreFilePropertiesPart();
        using var xw = System.Xml.XmlWriter.Create(coreProps.GetStream(FileMode.Create));
        xw.WriteRaw("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                               xmlns:dc="http://purl.org/dc/elements/1.1/">
                <dc:title>Relatório de Incidente de Segurança</dc:title>
                <dc:creator>ThinkIT SOC</dc:creator>
            </cp:coreProperties>
            """);
    }

    private static void AddDefaultStyles(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(
                    new RunPropertiesBaseStyle(
                        SF(),
                        new FontSize { Val = "22" },
                        new FontSizeComplexScript { Val = "22" }
                    )
                ),
                new ParagraphPropertiesDefault(
                    new ParagraphPropertiesBaseStyle(
                        new SpacingBetweenLines { After = "0" }
                    )
                )
            )
        );
        stylesPart.Styles.Save();
    }

    private static Body BuildBody()
    {
        var body = new Body();

        body.AppendChild(HeadingPara("RELATÓRIO DE INCIDENTE DE SEGURANÇA", "32", "1F497D"));
        body.AppendChild(SubtitlePara("ThinkIT Security Operations Center"));
        body.AppendChild(SubtitlePara("Gerado em: {{GENERATED_DATE}}"));
        body.AppendChild(HorizontalRule());

        body.AppendChild(SectionTitle("Sumário Executivo"));
        body.AppendChild(ContentPara("{{EXECUTIVE_SUMMARY}}"));
        body.AppendChild(HorizontalRule());

        body.AppendChild(SectionTitle("1. INFORMAÇÕES DO ALERTA"));
        body.AppendChild(BuildInfoTable(
        [
            ("ID",      "{{ALERT_ID}}"),
            ("Título",            "{{TITLE}}"),
            ("Severidade",        "{{SEVERITY}}"),
            ("Data / Hora (UTC)", "{{DATETIME_UTC}}"),
            ("Nº Chamado ITSM",   "{{ITSM_TICKET}}"),
            ("Tática MITRE",      "{{MITRE_TACTIC}}"),
        ]));

        body.AppendChild(SectionTitle("2. RESUMO DO EVENTO"));
        body.AppendChild(ContentPara("{{EVENT_SUMMARY}}"));

        body.AppendChild(SectionTitle("3. DETALHES TÉCNICOS"));
        body.AppendChild(BuildInfoTable(
        [
            ("Usuário",           "{{USER}}"),
            ("Endereço IP",       "{{IP_ADDRESS}}"),
            ("Hostname",          "{{HOST}}"),
            ("Nome do Arquivo",   "{{FILE_NAME}}"),
            ("Hash (SHA1)",       "{{SHA1_HASH}}"),
            ("Caminho (Path)",    "{{FILE_PATH}}"),
            ("Assinatura",        "{{FILE_SIGNATURE}}"),
        ]));

        body.AppendChild(SectionTitle("4. AVALIAÇÃO E AÇÕES"));
        body.AppendChild(FieldLabel("{{SOC_ACTIONS_TAKEN_LABEL}}"));
        body.AppendChild(ContentPara("{{SOC_ACTIONS_TAKEN}}"));
        body.AppendChild(FieldLabel("{{SOC_ACTION_LABEL}}"));
        body.AppendChild(ContentPara("{{SOC_ASSESSMENT}}"));
        body.AppendChild(FieldLabel("Ações Recomendadas:"));
        body.AppendChild(ContentPara("{{RECOMMENDED_ACTIONS}}"));
        body.AppendChild(FieldLabel("Observação Final:"));
        body.AppendChild(ContentPara("{{FINAL_OBSERVATION}}"));

        body.AppendChild(SectionTitle("5. REFERÊNCIAS"));
        body.AppendChild(ContentPara("{{REFERENCES}}"));

        body.AppendChild(HorizontalRule());
        body.AppendChild(FooterNote(
            "Este documento é de uso CONFIDENCIAL e destinado exclusivamente ao cliente indicado. " +
            "ThinkIT Security Operations Center."));

        body.AppendChild(new SectionProperties(
            new PageSize   { Width = 11906, Height = 16838 },
            new PageMargin { Top = 1134, Right = 1134, Bottom = 1134, Left = 1134 }
        ));

        return body;
    }

    private static RunFonts SF() =>
        new() { Ascii = Font, HighAnsi = Font, ComplexScript = Font, EastAsia = Font };

    private static Paragraph HeadingPara(string text, string fontSize, string color)
    {
        var run = new Run(new Text(text));
        run.RunProperties = new RunProperties(
            SF(),
            new Bold(),
            new FontSize { Val = fontSize },
            new Color { Val = color });

        return new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "0", After = "120" }),
            run);
    }

    private static Paragraph SubtitlePara(string text) =>
        new(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "0", After = "80" }),
            new Run(
                new RunProperties(SF(), new Color { Val = "595959" }, new FontSize { Val = "20" }),
                new Text(text)));

    private static Paragraph SectionTitle(string text)
    {
        var run = new Run(new Text(text));
        run.RunProperties = new RunProperties(
            SF(),
            new Bold(),
            new FontSize { Val = "24" },
            new Color { Val = "1F497D" });

        return new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "240", After = "80" }),
            run);
    }

    private static Paragraph FieldLabel(string text) =>
        new(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "120", After = "40" }),
            new Run(
                new RunProperties(SF(), new Bold(), new FontSize { Val = "22" }),
                new Text(text)));

    private static Paragraph ContentPara(string text) =>
        new(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "40", After = "120" }),
            new Run(
                new RunProperties(SF(), new FontSize { Val = "22" }),
                new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

    private static Paragraph FooterNote(string text) =>
        new(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "80", After = "0" }),
            new Run(
                new RunProperties(SF(), new Italic(), new Color { Val = "808080" }, new FontSize { Val = "18" }),
                new Text(text)));

    private static Paragraph HorizontalRule()
    {
        var pBdr = new ParagraphBorders(
            new BottomBorder { Val = BorderValues.Single, Size = 6, Color = "1F497D" });

        return new Paragraph(
            new ParagraphProperties(pBdr, new SpacingBetweenLines { Before = "120", After = "120" }));
    }

    private static Table BuildInfoTable(IEnumerable<(string Label, string Placeholder)> rows)
    {
        var table = new Table();

        table.AppendChild(new TableProperties(
            new TableBorders(
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "D0D0D0" },
                new InsideVerticalBorder   { Val = BorderValues.Single, Size = 4, Color = "D0D0D0" },
                new TopBorder              { Val = BorderValues.Single, Size = 4, Color = "D0D0D0" },
                new BottomBorder           { Val = BorderValues.Single, Size = 4, Color = "D0D0D0" },
                new LeftBorder             { Val = BorderValues.Single, Size = 4, Color = "D0D0D0" },
                new RightBorder            { Val = BorderValues.Single, Size = 4, Color = "D0D0D0" }),
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }));

        foreach (var (label, placeholder) in rows)
        {
            var row = new TableRow();

            var labelCell = new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Width = "2000", Type = TableWidthUnitValues.Pct },
                    new Shading { Fill = "EEF3F9", Val = ShadingPatternValues.Clear }),
                new Paragraph(new Run(
                    new RunProperties(SF(), new Bold(), new FontSize { Val = "20" }),
                    new Text(label))));

            var valueCell = new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Width = "3000", Type = TableWidthUnitValues.Pct }),
                new Paragraph(new Run(
                    new RunProperties(SF(), new FontSize { Val = "20" }),
                    new Text(placeholder) { Space = SpaceProcessingModeValues.Preserve })));

            row.Append(labelCell, valueCell);
            table.AppendChild(row);
        }

        return table;
    }
}
