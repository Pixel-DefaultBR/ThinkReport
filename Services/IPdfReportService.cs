using ThinkReport.Models;

namespace ThinkReport.Services;

public interface IPdfReportService
{
    byte[] Generate(
        IncidentReportViewModel model,
        IReadOnlyList<(string FileName, byte[] Data)> images, byte[] logoBytes);
}
