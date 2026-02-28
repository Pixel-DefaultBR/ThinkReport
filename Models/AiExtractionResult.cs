namespace ThinkReport.Models;

/// <summary>DTO retornado pelo serviço de extração de dados via IA.</summary>
public sealed class AiExtractionResult
{
    public string? ExecutiveSummary    { get; set; }
    public string? AlertId             { get; set; }
    public string? Title               { get; set; }
    /// <summary>ISO 8601 (UTC)</summary>
    public string? IncidentDateTimeUtc { get; set; }
    public string? ItsmTicketNumber    { get; set; }
    /// <summary>MITRE tactic strings — ex: ["TA0002 — Execution", "TA0003 — Persistence"]</summary>
    public string[]? MitreTactic       { get; set; }
    public string? EventSummary        { get; set; }
    public string? AffectedUser        { get; set; }
    public string? IpAddress           { get; set; }
    public string? Host                { get; set; }
    public string? FileName            { get; set; }
    public string? Sha1Hash            { get; set; }
    public string? FilePath            { get; set; }
    public string? FileSignature       { get; set; }
    public string? SocAssessment       { get; set; }
    public string? SocActionsTaken     { get; set; }
    public string? RecommendedActions  { get; set; }
    public string? FinalObservation    { get; set; }
    /// <summary>0=Informational … 4=Critical</summary>
    public int?    Severity            { get; set; }
    /// <summary>0=SocAvaliation, 1=SocTakenAction, 2=Both</summary>
    public int?    SelectedSocAction   { get; set; }
}
