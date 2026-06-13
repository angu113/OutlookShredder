// OutlookShredder.Proxy — Program.cs
using System.Reflection;
using OutlookShredder.Proxy.Extensions;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;
using OutlookShredder.Proxy.Services.Ai;
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
    var version = typeof(Program).Assembly
        .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
        ?.InformationalVersion ?? "unknown";
    Log.Information("ShredderProxy {Version} starting — logs: {LogPath}", version, logPath);

    // Use the exe directory as the base for config files so that appsettings files are found
    // whether the process runs as a Windows service (working dir = System32) or from a terminal.
    var exeDir = Path.GetDirectoryName(Environment.ProcessPath)
              ?? AppContext.BaseDirectory;

    var builder = WebApplication.CreateBuilder(args);
    builder.Configuration.SetBasePath(exeDir);

    // Re-add appsettings.json from the exe directory so it overrides any file
    // that CreateBuilder auto-loaded from the working directory.
    builder.Configuration.AddJsonFile("appsettings.json", optional: true, reloadOnChange: false);
    builder.Configuration.AddJsonFile($"appsettings.{builder.Environment.EnvironmentName}.json", optional: true, reloadOnChange: false);

    // Secrets file: gitignored, deployed alongside the exe.
    // Overrides appsettings.json values — put real API keys and credentials here.
    builder.Configuration.AddJsonFile("appsettings.secrets.json", optional: true, reloadOnChange: false);

    // Persistent secrets file: survives reinstall; wins over the install-dir copy.
    // reinstall.ps1 merges new keys from the template on every install so new
    // secret fields are added with default values without overwriting existing ones.
    var persistentSecretsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ShredderData", "Secrets", "appsettings.secrets.json");
    builder.Configuration.AddJsonFile(persistentSecretsPath, optional: true, reloadOnChange: false);

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
        options.AddPolicy("PhoneSearch", policy =>
        {
            policy.AllowAnyOrigin().AllowAnyHeader().AllowAnyMethod();
        });
    });

    builder.Services.AddControllers();

    // Rate limit tracker — shared across both AI providers
    builder.Services.AddSingleton<AiRateLimitTracker>();
    builder.Services.AddHttpClient("gemini")
        .AddHttpMessageHandler(sp => new AiRateLimitHandler(
            "Gemini",
            sp.GetRequiredService<AiRateLimitTracker>(),
            sp.GetRequiredService<ILogger<AiRateLimitHandler>>()));

    // Register AI extraction services (pluggable)
    builder.Services.AddSingleton<ProductSynonymService>();
    builder.Services.AddSingleton<ClaudeExtractionService>();
    builder.Services.AddSingleton<GeminiExtractionService>();
    builder.Services.AddSingleton<AiServiceFactory>();
    builder.Services.AddSingleton<SupplierCacheService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<SupplierCacheService>());
    builder.Services.AddSingleton<CustomerCacheService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<CustomerCacheService>());
    builder.Services.AddSingleton<ProductCatalogService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<ProductCatalogService>());
    builder.Services.AddSingleton<SharePointService>();
    builder.Services.AddSingleton<SliCacheService>();
    builder.Services.AddSingleton<MailService>();
    builder.Services.AddSingleton<RfqNotificationService>();
    builder.Services.AddSingleton<ShrConvInRouter>();
    builder.Services.AddSingleton<MailPollerService>();
    builder.Services.AddSingleton<ErpAiService>();
    builder.Services.AddSingleton<CustomerImportService>();
    builder.Services.AddSingleton<WorkflowCardService>();
    builder.Services.AddSingleton<StockNeededService>();
    builder.Services.AddSingleton<FileWatcherService>();
    builder.Services.AddSingleton<TodoService>();
    builder.Services.AddSingleton<MailboxBridgeService>();
    builder.Services.AddSingleton<MailTaxonomyService>();
    builder.Services.AddSingleton<MailRuleService>();
    builder.Services.AddSingleton<MailCacheService>();
    builder.Services.AddSingleton<MailProjectService>();
    builder.Services.AddSingleton<MailClassifierService>();
    builder.Services.AddSingleton<BillExtractionService>();
    builder.Services.AddSingleton<ConfirmationExtractionService>();
    builder.Services.AddSingleton<MailWorkbenchService>();
    builder.Services.AddSingleton<MailAutoCaptureService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<MailCacheService>());
    builder.Services.AddHostedService(sp => sp.GetRequiredService<MailAutoCaptureService>());
    builder.Services.AddSingleton<PoConfirmationMatcherService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<PoConfirmationMatcherService>());
    builder.Services.AddSingleton<BillToPoMatcherService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<BillToPoMatcherService>());

    builder.Services.AddHostedService(sp => sp.GetRequiredService<MailPollerService>());
    builder.Services.AddHostedService(sp => sp.GetRequiredService<WorkflowCardService>());
    builder.Services.AddHostedService(sp => sp.GetRequiredService<StockNeededService>());
    builder.Services.AddHostedService(sp => sp.GetRequiredService<FileWatcherService>());
    builder.Services.AddHostedService(sp => sp.GetRequiredService<TodoService>());
    builder.Services.AddHostedService(sp => sp.GetRequiredService<MailboxBridgeService>());

    builder.Services.AddSingleton<OutlookComPollerService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<OutlookComPollerService>());

    builder.Services.AddSingleton<DelegatedTokenProvider>();
    builder.Services.AddSingleton<HackensackPollerService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<HackensackPollerService>());
    builder.Services.AddHostedService<LqUpdateService>();
    builder.Services.AddHostedService<RfqAutoCompleteService>();
    builder.Services.AddSingleton<ProxyLeaseService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<ProxyLeaseService>());
    builder.Services.AddHostedService<ZoomCallWatcherService>();
    builder.Services.AddSingleton<SignalWireService>();
    builder.Services.AddSingleton<MessagingService>();
    builder.Services.AddSingleton<PricingAnalysisService>();
    builder.Services.AddSingleton<RfqSummaryService>();
    builder.Services.AddSingleton<RfqStateOfPlayService>();
    builder.Services.AddSingleton<RfqSummaryQueue>();
    builder.Services.AddHostedService<RfqSummaryQueueProcessor>();
    builder.Services.AddSingleton<ForgeSchedulerQueue>();
    builder.Services.AddSingleton<ForgeTaskService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<ForgeTaskService>());
    builder.Services.AddSingleton(sp => new Lazy<ForgeTaskService>(() => sp.GetRequiredService<ForgeTaskService>()));
    builder.Services.AddSingleton<CatalogAnalysisService>();
    builder.Services.AddSingleton(sp => new Lazy<CatalogAnalysisService>(() => sp.GetRequiredService<CatalogAnalysisService>()));
    builder.Services.AddSingleton<SupplierProductMappingsCacheService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<SupplierProductMappingsCacheService>());

    // Persistent cache services — registered as ICacheStatusProvider for CacheController enumeration.
    builder.Services.AddSingleton<ICacheStatusProvider>(sp => sp.GetRequiredService<ArchiveCacheService>());
    builder.Services.AddSingleton<ICacheStatusProvider>(sp => sp.GetRequiredService<ProductCatalogService>());
    builder.Services.AddSingleton<ICacheStatusProvider>(sp => sp.GetRequiredService<SupplierCacheService>());
    builder.Services.AddSingleton<ICacheStatusProvider>(sp => sp.GetRequiredService<SliCacheService>());
    builder.Services.AddSingleton<ICacheStatusProvider>(sp => sp.GetRequiredService<MailCacheService>());
    builder.Services.AddSingleton<ArchiveCacheService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<ArchiveCacheService>());
    builder.Services.AddHostedService<ShutdownWatcherService>();

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

    // Per-request duration log for /api/* — helps diagnose UI lag. Logs only
    // method + path + status + ms, so noise stays low while still making slow
    // requests (>1s) and errors (>=500) immediately visible in the proxy log.
    app.Use(async (ctx, next) =>
    {
        if (!ctx.Request.Path.StartsWithSegments("/api"))
        {
            await next();
            return;
        }
        var logger = ctx.RequestServices.GetRequiredService<ILogger<Program>>();
        var sw = System.Diagnostics.Stopwatch.StartNew();
        await next();
        sw.Stop();
        var ms = sw.ElapsedMilliseconds;
        var status = ctx.Response.StatusCode;
        if (status >= 500 || ms >= 1000)
            logger.LogWarning("[HTTP] {Method} {Path} -> {Status} in {Ms}ms", ctx.Request.Method, ctx.Request.Path, status, ms);
        else
            logger.LogInformation("[HTTP] {Method} {Path} -> {Status} in {Ms}ms", ctx.Request.Method, ctx.Request.Path, status, ms);
    });

    app.MapControllers();

    // Fire-and-forget SharePoint pre-warm so the first user request skips ~500ms
    // of cold-path setup (auth token exchange, site ID + list ID resolution, HTTP/2 handshake).
    var lifetime = app.Services.GetRequiredService<IHostApplicationLifetime>();
    lifetime.ApplicationStarted.Register(() =>
    {
        _ = Task.Run(async () =>
        {
            try
            {
                var sp = app.Services.GetRequiredService<SharePointService>();
                await sp.PrewarmAsync();

                // Load synonym dictionary from SP into the AI extraction cache.
                // Seeds from product-synonyms.json on first run if SP list is empty.
                var synonyms = app.Services.GetRequiredService<ProductSynonymService>();
                await synonyms.LoadAsync();

                // Pre-populate the SLI cache so the first /api/items request
                // is served from memory rather than paginating SP live.
                // force=true ensures a fresh SP read even when a disk cache exists from
                // a previous session — disk data can be hours old and would otherwise be
                // served as "fresh" for 5 minutes before expiring.
                var sliCache = app.Services.GetRequiredService<SliCacheService>();
                await sliCache.PopulateAsync(force: true);

                // Subscribe this proxy to the rfq-updates topic so peer-proxy SR events
                // trigger targeted SliCache updates (MergeRfqRows / InvalidateRfq) rather
                // than waiting for the 5-minute TTL to expire.
                var notify = app.Services.GetRequiredService<RfqNotificationService>();
                await notify.StartBusListenerAsync();
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "SharePoint pre-warm failed (non-fatal)");
            }
        });
    });

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

