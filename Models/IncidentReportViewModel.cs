using System.ComponentModel.DataAnnotations;

namespace ThinkReport.Models;
public sealed class IncidentReportViewModel
{
    [Required(ErrorMessage = "Sumário Executivo é obrigatório.")]
    [Display(Name = "Sumário Executivo")]
    public string ExecutiveSummary { get; set; } = string.Empty;

    [Required(ErrorMessage = "ID do alerta é obrigatório.")]
    [Display(Name = "ID")]
    public string AlertId { get; set; } = string.Empty;

    [Required(ErrorMessage = "Título é obrigatório.")]
    [Display(Name = "Título")]
    public string Title { get; set; } = string.Empty;

    [Required(ErrorMessage = "Selecione a severidade.")]
    [Display(Name = "Severidade")]
    public SeverityLevel Severity { get; set; } = SeverityLevel.Medium;

    [Required(ErrorMessage = "Data/Hora UTC é obrigatória.")]
    [Display(Name = "Data / Hora (UTC)")]
    public DateTime IncidentDateTimeUtc { get; set; } = DateTime.UtcNow;

    [Display(Name = "Nº Chamado ITSM")]
    public string? ItsmTicketNumber { get; set; }

    [Display(Name = "Tática(s) MITRE ATT&CK")]
    public List<string> MitreTactic { get; set; } = [];

    [Required(ErrorMessage = "Resumo do evento é obrigatório.")]
    [Display(Name = "Resumo do Evento")]
    public string EventSummary { get; set; } = string.Empty;

    [Display(Name = "Usuário")]
    public string? AffectedUser { get; set; }

    [Display(Name = "Endereço IP")]
    public string? IpAddress { get; set; }

    [Display(Name = "Hostname")]
    public string? Host { get; set; }

    [Display(Name = "Nome do Arquivo")]
    public string? FileName { get; set; }

    [Display(Name = "Hash (SHA1)")]
    [RegularExpression(@"^[0-9a-fA-F]{40}$", ErrorMessage = "Hash SHA1 inválido (deve ter 40 caracteres hex).")]
    public string? Sha1Hash { get; set; }

    [Display(Name = "Caminho (Path)")]
    public string? FilePath { get; set; }

    [Display(Name = "Assinatura do Arquivo")]
    public string? FileSignature { get; set; }


    [Required(ErrorMessage = "Ações Realizadas pelo SOC é obrigatório.")]
    [Display(Name = "Tipo de Ação / Avaliação")]
    public SocAction SelectedSocAction {get; set; } =  SocAction.SocAvaliation;
    
    [Required(ErrorMessage = "Avaliação do SOC é obrigatória.")]
    public string SocAssessment { get; set; } = string.Empty;

    [Display(Name = "Ações Tomadas pelo SOC")]
    public string? SocActionsTaken { get; set; }

    [Display(Name = "Ações Recomendadas")]
    public string? RecommendedActions { get; set; }

    [Display(Name = "Observação Final")]
    public string? FinalObservation { get; set; }

    [Display(Name = "Evidências (imagens)")]
    public List<IFormFile>? EvidenceImages { get; set; }
}
