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
                "https://*.microsoft.com",
                "https://outlook.cloud.microsoft",
                "https://*.cloud.microsoft")
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

var app = builder.Build();
app.UseCors();

// Chrome Private Network Access: allow public origins (OWA) to load iframes from localhost.
// Chrome 94+ blocks public → private network requests unless the server opts in.
// Edge is permissive by default; Chrome requires this header.
app.Use(async (context, next) =>
{
    context.Response.Headers["Access-Control-Allow-Private-Network"] = "true";
    await next();
});

app.UseDefaultFiles();
app.UseStaticFiles();
app.Run();
