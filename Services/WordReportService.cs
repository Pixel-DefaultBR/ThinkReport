using System.Runtime.Versioning;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ThinkReport.Models;
using A      = DocumentFormat.OpenXml.Drawing;
using DW     = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC    = DocumentFormat.OpenXml.Drawing.Pictures;
using SysImg = System.Drawing.Image;

namespace ThinkReport.Services;

public sealed class WordReportService : IWordReportService
{
    private readonly string _templatePath;

    public WordReportService(IWebHostEnvironment env)
    {
        _templatePath = Path.Combine(
            env.WebRootPath, "templates", "incident_report_template.docx");
    }

    public byte[] Generate(
        IncidentReportViewModel model,
        IReadOnlyList<(string FileName, byte[] Data)> images)
    {
        if (!File.Exists(_templatePath))
            throw new FileNotFoundException(
                $"Template não encontrado em: {_templatePath}. " +
                "Coloque o arquivo em wwwroot/templates/incident_report_template.docx.");

        using var ms = new MemoryStream();
        using (var fs = File.OpenRead(_templatePath))
            fs.CopyTo(ms);
        ms.Position = 0;

        using (var doc = WordprocessingDocument.Open(ms, isEditable: true))
        {
            InsertBulletActions(doc, model.RecommendedActions ?? string.Empty);

            ReplaceAllPlaceholders(doc, BuildReplacements(model));

            if (images.Count > 0)
                AppendEvidenceImages(doc, images);

            doc.MainDocumentPart!.Document.Save();
        }

        return ms.ToArray();
    }

    private static Dictionary<string, string> BuildReplacements(IncidentReportViewModel m)
    {
        var severityLabel = m.Severity switch
        {
            SeverityLevel.Informational => "Informacional",
            SeverityLevel.Low           => "Baixa",
            SeverityLevel.Medium        => "Média",
            SeverityLevel.High          => "Alta",
            SeverityLevel.Critical      => "Crítica",
            _                           => m.Severity.ToString()
        };

        var actionLabel = m.SelectedSocAction.GetDisplayName() + ":";

        return new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["{{EXECUTIVE_SUMMARY}}"]  = m.ExecutiveSummary,
            ["{{ALERT_ID}}"]           = m.AlertId,
            ["{{TITLE}}"]              = m.Title,
            ["{{SEVERITY}}"]           = severityLabel,
            ["{{DATETIME_UTC}}"]       = m.IncidentDateTimeUtc.ToString("yyyy-MM-dd HH:mm:ss"),
            ["{{ITSM_TICKET}}"]        = m.ItsmTicketNumber      ?? "N/A",
            ["{{MITRE_TACTIC}}"]       = m.MitreTactic           ?? "N/A",
            ["{{EVENT_SUMMARY}}"]      = m.EventSummary,
            ["{{USER}}"]               = m.AffectedUser          ?? "N/A",
            ["{{IP_ADDRESS}}"]         = m.IpAddress             ?? "N/A",
            ["{{HOST}}"]               = m.Host                  ?? "N/A",
            ["{{FILE_NAME}}"]          = m.FileName              ?? "N/A",
            ["{{SHA1_HASH}}"]          = m.Sha1Hash              ?? "N/A",
            ["{{FILE_PATH}}"]          = m.FilePath              ?? "N/A",
            ["{{FILE_SIGNATURE}}"]     = m.FileSignature         ?? "N/A",
            ["{{SOC_ACTION_LABEL}}"]   = actionLabel,
            ["{{SOC_ASSESSMENT}}"]     = m.SocAssessment,
            ["{{FINAL_OBSERVATION}}"]  = m.FinalObservation ?? "N/A",
            ["{{GENERATED_DATE}}"]     = DateTime.UtcNow.ToString("dd/MM/yyyy HH:mm:ss") + " UTC",
        };
    }

    private static void ReplaceAllPlaceholders(
        WordprocessingDocument doc,
        Dictionary<string, string> replacements)
    {
        var mainPart = doc.MainDocumentPart!;

        foreach (var para in mainPart.Document.Body!.Descendants<Paragraph>())
            ProcessParagraph(para, replacements);

        foreach (var headerPart in mainPart.HeaderParts)
            foreach (var para in headerPart.Header.Descendants<Paragraph>())
                ProcessParagraph(para, replacements);

        foreach (var footerPart in mainPart.FooterParts)
            foreach (var para in footerPart.Footer.Descendants<Paragraph>())
                ProcessParagraph(para, replacements);
    }

    private static void ProcessParagraph(
        Paragraph para,
        Dictionary<string, string> replacements)
    {
        var runs = para.Elements<Run>().ToList();
        if (runs.Count == 0) return;

        var fullText = string.Concat(
            runs.SelectMany(r => r.Elements<Text>()).Select(t => t.Text));

        if (!replacements.Keys.Any(k => fullText.Contains(k, StringComparison.Ordinal)))
            return;

        var newText = fullText;
        foreach (var (key, value) in replacements)
            newText = newText.Replace(key, value ?? string.Empty, StringComparison.Ordinal);

        var rPr = runs
            .Select(r => r.RunProperties)
            .FirstOrDefault(p => p is not null)
            ?.CloneNode(deep: true) as RunProperties;

        foreach (var run in runs)
            run.Remove();

        var newRun = new Run();
        if (rPr is not null)
            newRun.RunProperties = rPr;

        var lines = newText.Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            if (i > 0)
                newRun.AppendChild(new Break());

            newRun.AppendChild(new Text(lines[i])
                { Space = SpaceProcessingModeValues.Preserve });
        }

        para.AppendChild(newRun);
    }

    private static void InsertBulletActions(WordprocessingDocument doc, string rawText)
    {
        var body = doc.MainDocumentPart!.Document.Body!;

        var placeholder = body.Descendants<Paragraph>()
            .FirstOrDefault(p =>
                string.Concat(p.Descendants<Text>().Select(t => t.Text))
                      .Contains("{{RECOMMENDED_ACTIONS}}", StringComparison.Ordinal));

        if (placeholder is null) return;

        var parent = placeholder.Parent!;

        var lines = ParseBulletLines(rawText).ToList();

        if (lines.Count == 0)
        {
            var fallback = new Paragraph(new Run(
                new RunProperties(SF(), new FontSize { Val = "22" }),
                new Text("N/A")));
            parent.InsertBefore(fallback, placeholder);
        }
        else
        {
            foreach (var (text, level) in lines)
                parent.InsertBefore(MakeBulletParagraph(text, level), placeholder);
        }

        placeholder.Remove();
    }

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

    private static Paragraph MakeBulletParagraph(string text, int level)
    {
        var indent  = level == 1 ? "360" : "720";
        var marker  = level == 1 ? "•" : "◦";

        return new Paragraph(
            new ParagraphProperties(
                new Indentation { Left = indent, Hanging = "240" },
                new SpacingBetweenLines { Before = "0", After = "60" }),
            new Run(
                new RunProperties(SF(), new FontSize { Val = "22" }),
                new Text($"{marker}  {text}") { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static void AppendEvidenceImages(
        WordprocessingDocument doc,
        IReadOnlyList<(string FileName, byte[] Data)> images)
    {
        var mainPart = doc.MainDocumentPart!;
        var body     = mainPart.Document.Body!;

        var sectPr = body.Elements<SectionProperties>().LastOrDefault();

        void InsertNode(OpenXmlElement node)
        {
            if (sectPr is not null)
                body.InsertBefore(node, sectPr);
            else
                body.AppendChild(node);
        }

        InsertNode(new Paragraph(new Run(new Break { Type = BreakValues.Page })));

        InsertNode(BuildSectionHeading("Evidências do Evento"));

        uint imageIndex = 1;

        foreach (var (fileName, data) in images)
        {
            var contentType = DetectImageContentType(data);
            var imgPart     = mainPart.AddImagePart(contentType);

            using var imgStream = new MemoryStream(data);
            imgPart.FeedData(imgStream);

            var relationshipId            = mainPart.GetIdOfPart(imgPart);
            var (widthEmu, heightEmu)     = CalculateImageEmu(data);

            InsertNode(BuildCaption(imageIndex, fileName));
            InsertNode(BuildImageParagraph(relationshipId, imageIndex, widthEmu, heightEmu));
            InsertNode(new Paragraph());

            imageIndex++;
        }
    }

    private static RunFonts SF() =>
        new() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri", EastAsia = "Calibri" };

    private static Paragraph BuildSectionHeading(string text)
    {
        var run = new Run(new Text(text));
        run.RunProperties = new RunProperties(SF(), new Bold(), new FontSize { Val = "28" });

        return new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "240", After = "120" }),
            run);
    }

    private static Paragraph BuildCaption(uint index, string fileName)
    {
        var run = new Run(
            new Text($"Figura {index} — {fileName}")
                { Space = SpaceProcessingModeValues.Preserve });

        run.RunProperties = new RunProperties(
            SF(),
            new Italic(),
            new Color { Val = "595959" },
            new FontSize { Val = "20" });

        return new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "120", After = "60" }),
            run);
    }

    private static Paragraph BuildImageParagraph(
        string relationshipId,
        uint imageId,
        long widthEmu,
        long heightEmu)
    {
        var drawing = new Drawing(
            new DW.Inline(
                new DW.Extent         { Cx = widthEmu, Cy = heightEmu },
                new DW.EffectExtent   { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties  { Id = imageId, Name = $"Evidencia_{imageId}" },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties
                                    { Id = 0U, Name = $"img{imageId}.png" },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip
                                {
                                    Embed            = relationshipId,
                                    CompressionState = A.BlipCompressionValues.Print
                                },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset  { X = 0L, Y = 0L },
                                    new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                                new A.PresetGeometry(new A.AdjustValueList())
                                    { Preset = A.ShapeTypeValues.Rectangle })))
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
            {
                DistanceFromTop    = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft   = 0U,
                DistanceFromRight  = 0U
            });

        return new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center }),
            new Run(drawing));
    }

    [SupportedOSPlatform("windows")]
    private static (long WidthEmu, long HeightEmu) CalculateImageEmuWindows(byte[] imageData)
    {
        const long maxWidthEmu  = 5_600_000L;
        const double emuPerInch = 914_400.0;

        using var ms  = new MemoryStream(imageData);
        using var img = SysImg.FromStream(ms);

        var dpiX = img.HorizontalResolution > 0 ? img.HorizontalResolution : 96.0;
        var dpiY = img.VerticalResolution   > 0 ? img.VerticalResolution   : 96.0;
        var rawW = (long)(img.Width  / dpiX * emuPerInch);
        var rawH = (long)(img.Height / dpiY * emuPerInch);

        if (rawW <= maxWidthEmu)
            return (rawW, rawH);

        var scale = (double)maxWidthEmu / rawW;
        return (maxWidthEmu, (long)(rawH * scale));
    }

    private static (long WidthEmu, long HeightEmu) CalculateImageEmu(byte[] imageData)
    {
        const long maxWidthEmu = 5_600_000L;
        try
        {
            if (OperatingSystem.IsWindows())
                return CalculateImageEmuWindows(imageData);
        }
        catch { }

        return (maxWidthEmu, 3_600_000L);
    }

    private static string DetectImageContentType(byte[] data)
    {
        if (data.Length >= 3
            && data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF)
            return "image/jpeg";

        if (data.Length >= 4
            && data[0] == 0x89 && data[1] == 0x50
            && data[2] == 0x4E && data[3] == 0x47)
            return "image/png";

        if (data.Length >= 3
            && data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46)
            return "image/gif";

        if (data.Length >= 2
            && data[0] == 0x42 && data[1] == 0x4D)
            return "image/bmp";

        return "image/png";
    }
}
