// OutlookShredder.AddinHost — Program.cs
// Serves the Office.js add-in static files over HTTPS via IIS Express.
// No cert setup needed — IIS Express handles HTTPS automatically.

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
    {
        policy.WithOrigins(
                "https://localhost",
                "https://outlook.office.com",
                "https://outlook.office365.com",
                "https://*.office365.com",
                "https://*.microsoft.com")
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

var app = builder.Build();
app.UseCors();
app.UseDefaultFiles();
app.UseStaticFiles();
app.Run();
