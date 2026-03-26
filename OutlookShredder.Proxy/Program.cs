// OutlookShredder.Proxy — Program.cs
using OutlookShredder.Proxy.Services;

var builder = WebApplication.CreateBuilder(args);

// Allow the AddinHost origin. Override Proxy:AllowedOrigin in appsettings or environment variables
// to match the deployed AddinHost URL in non-development environments.
var allowedOrigin = builder.Configuration["Proxy:AllowedOrigin"] ?? "https://localhost:3000";
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
    {
        policy.WithOrigins(allowedOrigin)
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

builder.Services.AddControllers();
builder.Services.AddSingleton<ClaudeService>();
builder.Services.AddSingleton<SharePointService>();
builder.Services.AddSingleton<MailService>();
builder.Services.AddHostedService<MailPollerService>();

// Store the Anthropic API key in User Secrets (right-click project > Manage User Secrets)
// or appsettings.Development.json — never commit it to source control.
// In production, use Azure Key Vault or environment variables.

var app = builder.Build();
app.UseCors();
app.MapControllers();
app.Run();
