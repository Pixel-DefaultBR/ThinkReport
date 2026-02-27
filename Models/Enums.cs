using System.ComponentModel.DataAnnotations;
namespace ThinkReport.Models;

public enum SeverityLevel
{
    Informational,
    Low,
    Medium,
    High,
    Critical
}

public enum SocAction
{
    [Display(Name = "Avaliação do SOC")]
    SocAvaliation,
    [Display(Name = "Ação Tomada pelo SOC")]
    SocTakenAction
}

public static class MitreTactics
{
    public static readonly IReadOnlyList<(string Id, string Name)> All =
    [
        ("TA0001", "Initial Access"),
        ("TA0002", "Execution"),
        ("TA0003", "Persistence"),
        ("TA0004", "Privilege Escalation"),
        ("TA0005", "Defense Evasion"),
        ("TA0006", "Credential Access"),
        ("TA0007", "Discovery"),
        ("TA0008", "Lateral Movement"),
        ("TA0009", "Collection"),
        ("TA0011", "Command and Control"),
        ("TA0010", "Exfiltration"),
        ("TA0040", "Impact"),
    ];
}
