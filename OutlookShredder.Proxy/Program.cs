// OutlookShredder.Proxy — Program.cs
using OutlookShredder.Proxy.Services;

var builder = WebApplication.CreateBuilder(args);

// Secrets file: gitignored, deployed alongside the exe.
// Overrides appsettings.json values — put real API keys and credentials here.
builder.Configuration.AddJsonFile("appsettings.secrets.json", optional: true, reloadOnChange: false);

// Allows the proxy to run as a Windows Service (start on boot, any user account).
// When run from the terminal for development it behaves as a normal console app.
builder.Host.UseWindowsService(options => options.ServiceName = "ShredderProxy");

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
builder.Services.AddSingleton<SupplierCacheService>();
builder.Services.AddHostedService(sp => sp.GetRequiredService<SupplierCacheService>());
builder.Services.AddSingleton<SharePointService>();
builder.Services.AddSingleton<MailService>();
builder.Services.AddSingleton<RfqNotificationService>();
builder.Services.AddSingleton<MailPollerService>();
builder.Services.AddHostedService(sp => sp.GetRequiredService<MailPollerService>());

// Store the Anthropic API key in User Secrets (right-click project > Manage User Secrets)
// or appsettings.Development.json — never commit it to source control.
// In production, use Azure Key Vault or environment variables.

var app = builder.Build();
app.UseCors();
app.MapControllers();
app.Run();
