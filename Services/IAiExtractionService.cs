using ThinkReport.Models;

namespace ThinkReport.Services;

public interface IAiExtractionService
{
    Task<AiExtractionResult> ExtractAsync(
        string rawText,
        IReadOnlyList<(string FileName, string Content)> logFiles,
        CancellationToken cancellationToken = default);
}
