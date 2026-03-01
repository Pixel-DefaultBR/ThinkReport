namespace ThinkReport.Models;

public sealed class AiExtractionResult
{
    public string? ExecutiveSummary    { get; set; }
    public string? AlertId             { get; set; }
    public string? Title               { get; set; }
    public string? IncidentDateTimeUtc { get; set; }
    public string? ItsmTicketNumber    { get; set; }
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
    public int?    Severity            { get; set; }
    public int?    SelectedSocAction   { get; set; }
}
