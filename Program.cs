using Microsoft.AspNetCore.Http.Features;
using QuestPDF.Infrastructure;
using ThinkReport.Services;
using ThinkReport.Tools;

var builder = WebApplication.CreateBuilder(args);

QuestPDF.Settings.License = LicenseType.Community;

builder.Services.AddControllersWithViews();
builder.Services.AddScoped<IWordReportService, WordReportService>();
builder.Services.AddScoped<IPdfReportService, PdfReportService>();

builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = 200 * 1024 * 1024;
});

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();
app.UseAuthorization();

app.MapControllerRoute(
    name:    "default",
    pattern: "{controller=Report}/{action=Index}/{id?}");

var templatePath = Path.Combine(app.Environment.WebRootPath, "templates", "incident_report_template.docx");
TemplateGenerator.EnsureTemplateExists(templatePath);

app.Run();
