using Microsoft.AspNetCore.Mvc;
using ThinkReport.Models;
using ThinkReport.Services;

namespace ThinkReport.Controllers;

public sealed class ReportController : Controller
{
    private readonly IWordReportService _reportService;
    private readonly IPdfReportService  _pdfService;
    private readonly IAiExtractionService _aiService;
    private readonly ILogger<ReportController> _logger;

    private const long MaxImageSizeBytes   = 20 * 1024 * 1024;
    private const long MaxLogFileSizeBytes = 1 * 1024 * 1024;
    private const int  MaxLogFileCount     = 3;

    private static readonly HashSet<string> AllowedMimeTypes =
    [
        "image/jpeg", "image/png", "image/webp"
    ];

    private static readonly HashSet<string> AllowedLogExtensions =
    [
        ".txt", ".log", ".json"
    ];

    public ReportController(
        IWordReportService reportService,
        IPdfReportService  pdfService,
        IAiExtractionService aiService,
        ILogger<ReportController> logger)
    {
        _reportService = reportService;
        _pdfService    = pdfService;
        _aiService     = aiService;
        _logger        = logger;
    }

    [HttpGet]
    public IActionResult Index()
    {
        var model = new IncidentReportViewModel { IncidentDateTimeUtc = DateTime.UtcNow };
        return View(model);
    }

    [HttpPost]
    [RequestSizeLimit(200 * 1024 * 1024)]
    [RequestFormLimits(MultipartBodyLengthLimit = 200 * 1024 * 1024)]
    public async Task<IActionResult> Generate([FromForm] IncidentReportViewModel model)
    {
        if (!ModelState.IsValid)
            return View("Index", model);

        var images = new List<(string FileName, byte[] Data)>();
        var formFiles = Request.Form.Files.GetFiles("EvidenceImages");

        _logger.LogInformation("Arquivos recebidos no campo 'EvidenceImages': {Count}", formFiles.Count);

        foreach (var file in formFiles)
        {
            if (file.Length == 0)
            {
                _logger.LogWarning("Arquivo ignorado (tamanho zero): {Name}", file.FileName);
                continue;
            }

            if (file.Length > MaxImageSizeBytes)
            {
                ModelState.AddModelError(
                    nameof(model.EvidenceImages),
                    $"O arquivo '{file.FileName}' excede o limite de 20 MB.");
                return View("Index", model);
            }

            if (!AllowedMimeTypes.Contains(file.ContentType.ToLowerInvariant()))
            {
                ModelState.AddModelError(
                    nameof(model.EvidenceImages),
                    $"Tipo de arquivo não permitido: '{file.ContentType}'. Use JPEG, PNG ou WEBP.");
                return View("Index", model);
            }

            using var ms = new MemoryStream((int)file.Length);
            await file.CopyToAsync(ms);
            images.Add((file.FileName, ms.ToArray()));

            _logger.LogInformation("Imagem lida: {Name} ({Size} bytes)", file.FileName, file.Length);
        }

        try
        {
            var docBytes = _reportService.Generate(model, images);

            var safeId   = string.Concat(model.AlertId.Where(c => char.IsLetterOrDigit(c) || c == '-' || c == '_'));
            var fileName = $"ThinkIT_Incident_{safeId}_{DateTime.UtcNow:yyyyMMddHHmmss}.docx";

            _logger.LogInformation(
                "Relatório gerado. AlertId={AlertId}, Evidências={Count}",
                model.AlertId, images.Count);

            Response.Cookies.Append("downloadReady", "1",
                new CookieOptions { SameSite = SameSiteMode.Strict, MaxAge = TimeSpan.FromSeconds(10) });

            return File(
                docBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                fileName);
        }
        catch (FileNotFoundException ex)
        {
            _logger.LogError(ex, "Template .docx não encontrado.");
            ModelState.AddModelError(string.Empty, "Template do relatório não encontrado. Contate o administrador.");
            return View("Index", model);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Erro ao gerar o relatório.");
            ModelState.AddModelError(string.Empty, "Erro interno ao gerar o relatório. Verifique os logs.");
            return View("Index", model);
        }
    }

    [HttpPost]
    [RequestSizeLimit(200 * 1024 * 1024)]
    [RequestFormLimits(MultipartBodyLengthLimit = 200 * 1024 * 1024)]
    public async Task<IActionResult> GeneratePdf([FromForm] IncidentReportViewModel model)
    {
        if (!ModelState.IsValid)
            return View("Index", model);

        var images    = new List<(string FileName, byte[] Data)>();
        var formFiles = Request.Form.Files.GetFiles("EvidenceImages");

        _logger.LogInformation("PDF – Arquivos recebidos no campo 'EvidenceImages': {Count}", formFiles.Count);

        foreach (var file in formFiles)
        {
            if (file.Length == 0)
            {
                _logger.LogWarning("Arquivo ignorado (tamanho zero): {Name}", file.FileName);
                continue;
            }

            if (file.Length > MaxImageSizeBytes)
            {
                ModelState.AddModelError(
                    nameof(model.EvidenceImages),
                    $"O arquivo '{file.FileName}' excede o limite de 20 MB.");
                return View("Index", model);
            }

            if (!AllowedMimeTypes.Contains(file.ContentType.ToLowerInvariant()))
            {
                ModelState.AddModelError(
                    nameof(model.EvidenceImages),
                    $"Tipo de arquivo não permitido: '{file.ContentType}'. Use JPEG, PNG ou WEBP.");
                return View("Index", model);
            }

            using var ms = new MemoryStream((int)file.Length);
            await file.CopyToAsync(ms);
            images.Add((file.FileName, ms.ToArray()));

            _logger.LogInformation("Imagem lida: {Name} ({Size} bytes)", file.FileName, file.Length);
        }

        try
        {
            byte[] logoThink = System.IO.File.ReadAllBytes("wwwroot/img/image1.png");
            var pdfBytes = _pdfService.Generate(model, images, logoThink);

            var safeId   = string.Concat(model.AlertId.Where(c => char.IsLetterOrDigit(c) || c == '-' || c == '_'));
            var fileName = $"ThinkIT_Incident_{safeId}_{DateTime.UtcNow:yyyyMMddHHmmss}.pdf";

            _logger.LogInformation(
                "Relatório PDF gerado. AlertId={AlertId}, Evidências={Count}",
                model.AlertId, images.Count);

            Response.Cookies.Append("downloadReady", "1",
                new CookieOptions { SameSite = SameSiteMode.Strict, MaxAge = TimeSpan.FromSeconds(10) });

            return File(pdfBytes, "application/pdf", fileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Erro ao gerar o relatório PDF.");
            ModelState.AddModelError(string.Empty, "Erro interno ao gerar o relatório PDF. Verifique os logs.");
            return View("Index", model);
        }
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> FillWithAI(string rawText, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(rawText))
            return Json(new { error = "O campo de dados brutos não pode estar vazio." });

        var uploadedFiles = Request.Form.Files.GetFiles("LogFiles");

        if (uploadedFiles.Count > MaxLogFileCount)
            return Json(new { error = $"Máximo de {MaxLogFileCount} arquivos de log permitidos." });

        var logFiles = new List<(string FileName, string Content)>();

        foreach (var file in uploadedFiles)
        {
            if (file.Length == 0) continue;

            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            if (!AllowedLogExtensions.Contains(ext))
                return Json(new { error = $"Extensão não permitida: '{ext}'. Use .txt, .log ou .json." });

            if (file.Length > MaxLogFileSizeBytes)
                return Json(new { error = $"O arquivo '{file.FileName}' excede o limite de 1 MB." });

            using var reader = new StreamReader(file.OpenReadStream());
            var content = await reader.ReadToEndAsync(cancellationToken);
            logFiles.Add((file.FileName, content));
        }

        try
        {
            var result = await _aiService.ExtractAsync(rawText, logFiles, cancellationToken);
            _logger.LogInformation("AI extraction succeeded. AlertId={AlertId}", result.AlertId);
            return Json(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Erro na extração via IA.");
            return Json(new { error = "Erro ao chamar a IA. Verifique a API key e tente novamente." });
        }
    }
}
