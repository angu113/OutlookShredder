// OutlookShredder.Proxy — Program.cs
using OutlookShredder.Proxy.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
    {
        policy.WithOrigins("https://localhost:3000")
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

builder.Services.AddControllers();
builder.Services.AddSingleton<ClaudeService>();
builder.Services.AddSingleton<SharePointService>();

// Store the Anthropic API key in User Secrets (right-click project > Manage User Secrets)
// or appsettings.Development.json — never commit it to source control.
// In production, use Azure Key Vault or environment variables.

var app = builder.Build();
app.UseCors();
app.MapControllers();
app.Run();
