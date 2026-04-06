// OutlookShredder.Proxy — Program.cs
using OutlookShredder.Proxy.Services;
using Serilog;
using Serilog.Events;

// ── Bootstrap logger (used until DI container is ready) ─────────────────────
var logPath = Path.Combine(AppContext.BaseDirectory, "Logs", "proxy-.log");
Log.Logger = new LoggerConfiguration()
    .MinimumLevel.Information()
    .MinimumLevel.Override("Microsoft.AspNetCore", LogEventLevel.Warning)
    .MinimumLevel.Override("Microsoft.Graph", LogEventLevel.Warning)
    .WriteTo.Console(outputTemplate: "[{Timestamp:HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}")
    .WriteTo.File(logPath,
        rollingInterval:       RollingInterval.Day,
        retainedFileCountLimit: 30,
        fileSizeLimitBytes:     50_000_000,
        rollOnFileSizeLimit:    true,
        outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
    .CreateBootstrapLogger();

try
{
    Log.Information("ShredderProxy starting — logs: {LogPath}", logPath);

    var builder = WebApplication.CreateBuilder(args);

    // Secrets file: gitignored, deployed alongside the exe.
    // Overrides appsettings.json values — put real API keys and credentials here.
    builder.Configuration.AddJsonFile("appsettings.secrets.json", optional: true, reloadOnChange: false);

    // Replace the default Microsoft.Extensions.Logging with Serilog so all
    // ILogger<T> calls (including from third-party libraries) go through Serilog.
    builder.Host.UseSerilog((ctx, _, config) => config
        .ReadFrom.Configuration(ctx.Configuration)
        .MinimumLevel.Information()
        .MinimumLevel.Override("Microsoft.AspNetCore", LogEventLevel.Warning)
        .MinimumLevel.Override("Microsoft.Graph", LogEventLevel.Warning)
        .Enrich.FromLogContext()
        .WriteTo.Console(outputTemplate: "[{Timestamp:HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}")
        .WriteTo.File(logPath,
            rollingInterval:       RollingInterval.Day,
            retainedFileCountLimit: 30,
            fileSizeLimitBytes:     50_000_000,
            rollOnFileSizeLimit:    true,
            outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}"));

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
    builder.Services.AddSingleton<ProductCatalogService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<ProductCatalogService>());
    builder.Services.AddSingleton<SharePointService>();
    builder.Services.AddSingleton<MailService>();
    builder.Services.AddSingleton<RfqNotificationService>();
    builder.Services.AddSingleton<MailPollerService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<MailPollerService>());
    builder.Services.AddHostedService<LqUpdateService>();

    var app = builder.Build();

    // ── Global unhandled-exception handler ───────────────────────────────────
    app.UseExceptionHandler(errorApp => errorApp.Run(async ctx =>
    {
        var ex = ctx.Features
            .Get<Microsoft.AspNetCore.Diagnostics.IExceptionHandlerFeature>()?.Error;
        var logger = ctx.RequestServices.GetRequiredService<ILogger<Program>>();
        if (ex is not null)
            logger.LogError(ex, "Unhandled exception on {Method} {Path}",
                ctx.Request.Method, ctx.Request.Path);
        ctx.Response.StatusCode  = 500;
        ctx.Response.ContentType = "application/json";
        await ctx.Response.WriteAsync("{\"success\":false,\"error\":\"Internal server error\"}");
    }));

    app.UseCors();
    app.MapControllers();
    app.Run();
}
catch (Exception ex) when (ex is not OperationCanceledException)
{
    Log.Fatal(ex, "ShredderProxy terminated unexpectedly");
}
finally
{
    Log.Information("ShredderProxy shut down");
    Log.CloseAndFlush();
}
