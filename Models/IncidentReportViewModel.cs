using System.ComponentModel.DataAnnotations;

namespace ThinkReport.Models;
public sealed class IncidentReportViewModel : IValidatableObject
{
    [Required(ErrorMessage = "Sumário Executivo é obrigatório.")]
    [MaxLength(5000)]
    [Display(Name = "Sumário Executivo")]
    public string ExecutiveSummary { get; set; } = string.Empty;

    [Required(ErrorMessage = "ID do alerta é obrigatório.")]
    [MaxLength(100)]
    [Display(Name = "ID")]
    public string AlertId { get; set; } = string.Empty;

    [Required(ErrorMessage = "Título é obrigatório.")]
    [MaxLength(500)]
    [Display(Name = "Título")]
    public string Title { get; set; } = string.Empty;

    [Required(ErrorMessage = "Selecione a severidade.")]
    [Display(Name = "Severidade")]
    public SeverityLevel Severity { get; set; } = SeverityLevel.Medium;

    [Required(ErrorMessage = "Data/Hora UTC é obrigatória.")]
    [Display(Name = "Data / Hora (UTC)")]
    public DateTime IncidentDateTimeUtc { get; set; } = DateTime.UtcNow;

    [MaxLength(100)]
    [Display(Name = "Nº Chamado ITSM")]
    public string? ItsmTicketNumber { get; set; }

    [Display(Name = "Tática(s) MITRE ATT&CK")]
    public List<string> MitreTactic { get; set; } = [];

    [Required(ErrorMessage = "Resumo do evento é obrigatório.")]
    [MaxLength(10000)]
    [Display(Name = "Resumo do Evento")]
    public string EventSummary { get; set; } = string.Empty;

    [MaxLength(200)]
    [Display(Name = "Usuário")]
    public string? AffectedUser { get; set; }

    [MaxLength(45)]
    [Display(Name = "Endereço IP")]
    public string? IpAddress { get; set; }

    [MaxLength(253)]
    [Display(Name = "Hostname")]
    public string? Host { get; set; }

    [MaxLength(260)]
    [Display(Name = "Nome do Arquivo")]
    public string? FileName { get; set; }

    [MaxLength(40)]
    [Display(Name = "Hash (SHA1)")]
    [RegularExpression(@"^[0-9a-fA-F]{40}$", ErrorMessage = "Hash SHA1 inválido (deve ter 40 caracteres hex).")]
    public string? Sha1Hash { get; set; }

    [MaxLength(2000)]
    [Display(Name = "Caminho (Path)")]
    public string? FilePath { get; set; }

    [MaxLength(500)]
    [Display(Name = "Assinatura do Arquivo")]
    public string? FileSignature { get; set; }

    [Required(ErrorMessage = "Ações Realizadas pelo SOC é obrigatório.")]
    [Display(Name = "Tipo de Ação / Avaliação")]
    public SocAction SelectedSocAction { get; set; } = SocAction.SocAvaliation;

    [Required(ErrorMessage = "Avaliação do SOC é obrigatória.")]
    [MaxLength(10000)]
    public string SocAssessment { get; set; } = string.Empty;

    [MaxLength(10000)]
    [Display(Name = "Ações Tomadas pelo SOC")]
    public string? SocActionsTaken { get; set; }

    [MaxLength(10000)]
    [Display(Name = "Ações Recomendadas")]
    public string? RecommendedActions { get; set; }

    [MaxLength(5000)]
    [Display(Name = "Observação Final")]
    public string? FinalObservation { get; set; }

    [Display(Name = "Referências")]
    public List<string>? References { get; set; }

    [Display(Name = "Evidências (imagens)")]
    public List<IFormFile>? EvidenceImages { get; set; }

    public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
    {
        if (References is null) yield break;

        foreach (var url in References.Where(r => !string.IsNullOrWhiteSpace(r)))
        {
            if (url.Length > 2000)
            {
                yield return new ValidationResult(
                    "URL de referência muito longa (máx. 2000 caracteres).",
                    [nameof(References)]);
                yield break;
            }

            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
                (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                yield return new ValidationResult(
                    $"URL inválida ou protocolo não permitido (apenas http:// e https://): \"{url[..Math.Min(url.Length, 80)]}\"",
                    [nameof(References)]);
            }
        }
    }
}
