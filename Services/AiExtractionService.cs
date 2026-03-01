using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Configuration;
using OpenAI.Chat;
using ThinkReport.Models;

namespace ThinkReport.Services;

public sealed class AiExtractionService : IAiExtractionService
{
    private const int MaxLogContentChars = 40_000;

    private readonly ChatClient _chatClient;
    private readonly ILogger<AiExtractionService> _logger;

    private static readonly ChatCompletionOptions JsonOptions = new()
    {
        ResponseFormat = ChatResponseFormat.CreateJsonObjectFormat()
    };

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public AiExtractionService(IConfiguration config, ILogger<AiExtractionService> logger)
    {
        _logger = logger;

        var apiKey = config["OpenAI:ApiKey"]
            ?? throw new InvalidOperationException("OpenAI:ApiKey not configured.");
        var model = config["OpenAI:Model"] ?? "gpt-4.1-mini";

        var openAiClient = new OpenAI.OpenAIClient(apiKey);
        _chatClient = openAiClient.GetChatClient(model);
    }

    public async Task<AiExtractionResult> ExtractAsync(
        string rawText,
        IReadOnlyList<(string FileName, string Content)> logFiles,
        CancellationToken cancellationToken = default)
    {
        var userContent = BuildUserContent(rawText, logFiles);

        var messages = new List<ChatMessage>
        {
            new SystemChatMessage(SystemPrompt),
            new UserChatMessage(userContent)
        };

        _logger.LogInformation("Calling OpenAI for AI extraction. Raw text length: {Len}", rawText.Length);

        var response = await _chatClient.CompleteChatAsync(messages, JsonOptions, cancellationToken);
        var jsonText = response.Value.Content[0].Text;

        _logger.LogInformation("OpenAI response received ({Len} chars).", jsonText.Length);

        var result = JsonSerializer.Deserialize<AiExtractionResult>(jsonText, JsonOpts)
                     ?? new AiExtractionResult();

        PostProcess(result);
        return result;
    }

    private static string BuildUserContent(
        string rawText,
        IReadOnlyList<(string FileName, string Content)> logFiles)
    {
        var sb = new StringBuilder();
        sb.AppendLine("=== INCIDENT RAW DATA ===");
        sb.AppendLine(rawText);

        foreach (var (name, content) in logFiles)
        {
            sb.AppendLine();
            sb.AppendLine($"=== LOG FILE: {name} ===");
            var truncated = content.Length > MaxLogContentChars
                ? content[..MaxLogContentChars] + "\n[... truncated ...]"
                : content;
            sb.AppendLine(truncated);
        }

        return sb.ToString();
    }

    private static void PostProcess(AiExtractionResult r)
    {
        if (r.Sha1Hash is not null)
        {
            r.Sha1Hash = r.Sha1Hash.Trim().ToLowerInvariant();
            if (!Regex.IsMatch(r.Sha1Hash, @"^[0-9a-f]{40}$"))
                r.Sha1Hash = null;
        }

        if (r.Severity is < 0 or > 4)
            r.Severity = null;

        if (r.SelectedSocAction is < 0 or > 2)
            r.SelectedSocAction = null;
    }

    private const string SystemPrompt = """
        Você é um assistente especializado em cibersegurança para analistas Blue Team SOC (nível N1).
        Sua tarefa é extrair dados estruturados de relatório de incidente a partir do texto bruto fornecido pelo analista.
        Retorne SOMENTE um objeto JSON válido — sem markdown, sem code fences, sem texto adicional.
        TODOS os textos descritivos devem estar em português do Brasil (PT-BR).

        O JSON deve conter exatamente estes campos (todos anuláveis — use null se não encontrado):
        {
          "ExecutiveSummary":    "resumo executivo breve para o cliente em PT-BR (string ou null)",
          "AlertId":             "ID do alerta, ex: ALT-2025-00421 (string ou null)",
          "Title":               "título curto do incidente em PT-BR (string ou null)",
          "IncidentDateTimeUtc": "data/hora ISO 8601 UTC, ex: 2025-04-12T14:30:00Z (string ou null)",
          "ItsmTicketNumber":    "número do ticket ITSM, ex: INC0012345 (string ou null)",
          "MitreTactic":         ["array de táticas MITRE, cada uma no formato 'TAXXXX — Nome em inglês' usando em-dash (—), ex: ['TA0002 — Execution', 'TA0003 — Persistence'] — use array vazio [] se não encontrado"],
          "EventSummary":        "descrição técnica DETALHADA e completa do evento de segurança em PT-BR — inclua: o que aconteceu, quando foi detectado, qual ferramenta/regra disparou o alerta, processos envolvidos, comandos executados, conexões de rede observadas, arquivos criados/modificados, sequência cronológica dos eventos, contexto do ambiente e qualquer IOC identificado. Seja o mais completo possível com base nos dados fornecidos (string ou null)",
          "AffectedUser":        "usuário afetado, ex: DOMINIO\\usuario (string ou null)",
          "IpAddress":           "endereço IP do host afetado (string ou null)",
          "Host":                "hostname do sistema afetado (string ou null)",
          "FileName":            "nome do arquivo malicioso ou suspeito (string ou null)",
          "Sha1Hash":            "hash SHA1 — exatamente 40 caracteres hexadecimais minúsculos (string ou null)",
          "FilePath":            "caminho completo do arquivo (string ou null)",
          "FileSignature":       "status da assinatura digital em PT-BR, ex: 'Não assinado' ou 'Microsoft Corporation' (string ou null)",
          "SocAssessment":       "avaliação do analista SOC em PT-BR: contexto, confirmação/descarte de falso-positivo, IOCs identificados (string ou null)",
          "SocActionsTaken":     "ações tomadas pelo SOC em PT-BR (preencher somente se SelectedSocAction=2, caso contrário null) (string ou null)",
          "RecommendedActions":  "ações recomendadas em PT-BR como lista de bullets usando prefixo '- ' (string ou null)",
          "FinalObservation":    "observações adicionais em PT-BR, links de threat intel, tickets relacionados (string ou null)",
          "Severity":            "severidade como inteiro: 0=Informational, 1=Low, 2=Medium, 3=High, 4=Critical (inteiro ou null)",
          "SelectedSocAction":   "ação SOC como inteiro: 0=Avaliação do SOC, 1=Ação Tomada pelo SOC, 2=Ambos (inteiro ou null)"
        }

        Regras importantes:
        - Use null para qualquer campo que não puder ser determinado a partir dos dados.
        - Sha1Hash deve ter exatamente 40 caracteres hexadecimais minúsculos ou null.
        - IncidentDateTimeUtc deve ser ISO 8601 UTC ou null.
        - MitreTactic é um array JSON de strings; cada entrada deve usar o em-dash (—) entre o ID e o nome em inglês (padrão MITRE ATT&CK). Use [] se não encontrado.
        - Não invente dados que não estejam presentes na entrada.
        - RecommendedActions: use '- ' para bullets principais e '  - ' (2 espaços + hífen) para sub-bullets.
        - FileSignature: se "unsigned" ou "not signed" → traduzir para "Não assinado".
        """;
}
