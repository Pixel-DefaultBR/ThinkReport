using ThinkReport.Models;

namespace ThinkReport.Services;

public interface IWordReportService
{
    byte[] Generate(
        IncidentReportViewModel model,
        IReadOnlyList<(string FileName, byte[] Data)> images);
}
