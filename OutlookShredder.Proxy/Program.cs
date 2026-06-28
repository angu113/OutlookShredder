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

    // Per-machine NON-secret overrides: gitignored, created per machine, and deliberately NOT registered
    // for the WS1 cleanup below — so config that is not a Key Vault secret (e.g. MailboxBridge:Mailboxes,
    // SignalWire:SpaceUrl/FromNumber) has a durable home that survives the secrets-file wipe. Loaded after
    // the secrets files so it can override them; genuine secrets still belong in Key Vault, never here.
    builder.Configuration.AddJsonFile("appsettings.local.json", optional: true, reloadOnChange: false);

    // Register the local cleartext copies so the post-startup WS1 cleanup can delete them once the vault
    // is proven healthy (see SecretsBootstrap.CleanupLocalSecretsIfVaultHealthy after PrewarmAsync).
    SecretsBootstrap.RegisterLocalSecretFile(Path.Combine(AppContext.BaseDirectory, "appsettings.secrets.json"));
    SecretsBootstrap.RegisterLocalSecretFile(persistentSecretsPath);

    // WS1 — overlay secrets from Azure Key Vault (signed-in user's WAM token, silent), added AFTER the
    // JSON files so it wins, and BEFORE the host is built so DI-time config readers see vault values.
    // Vault-first, file fallback: never throws — on any failure the file secrets above remain in effect.
    if (builder.Configuration.GetValue("KeyVault:Enabled", true))
        SecretsBootstrap.LoadInto(builder.Configuration, Log.Logger);

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

    // WS2 — local REST auth: mints + verifies the per-launch token / per-request HMAC.
    builder.Services.AddSingleton<ProxyAuthService>();

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
    // Sales-Order-history serving cache (indexed by customer) + resumable bulk-import runner, feeding the
    // Raptor incoming-call card. Built once at startup, write-through + 5-min safety re-scan thereafter.
    builder.Services.AddSingleton<SalesOrderHistoryService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<SalesOrderHistoryService>());
    builder.Services.AddSingleton<ProductCatalogService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<ProductCatalogService>());
    builder.Services.AddSingleton<SharePointService>();
    // Call-log CRM backfill (fills BpName/Contact/Popup on historic PhoneCallLog rows). Shared by the manual
    // /api/phone/call-log/backfill-crm endpoint and the post-import auto-backfill in ImportController.
    builder.Services.AddSingleton<CallLogCrmBackfillService>();
    builder.Services.AddSingleton<SliCacheService>();
    // In-memory unread-row index: builds the inbound-row set once + on write-through/bus, so the per-user
    // unread tally is an in-memory pass instead of a ~5-7s dual SharePoint scan on every call.
    builder.Services.AddSingleton<SupplierUnreadIndexService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<SupplierUnreadIndexService>());
    builder.Services.AddSingleton<MailService>();
    builder.Services.AddSingleton<RfqNotificationService>();
    builder.Services.AddSingleton<ShrConvInRouter>();
    builder.Services.AddSingleton<MailPollerService>();
    builder.Services.AddSingleton<ErpAiService>();
    builder.Services.AddSingleton<CustomerImportService>();
    builder.Services.AddSingleton<WorkflowCardService>();
    builder.Services.AddSingleton<PoSlipDependencyResolver>();
    builder.Services.AddSingleton<TransferLinkSuggestionService>();
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
    builder.Services.AddSingleton<MailEvalService>();
    builder.Services.AddSingleton<MailAutoCaptureService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<MailCacheService>());
    builder.Services.AddHostedService(sp => sp.GetRequiredService<MailAutoCaptureService>());
    builder.Services.AddSingleton<PoConfirmationMatcherService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<PoConfirmationMatcherService>());
    builder.Services.AddSingleton<BillToPoMatcherService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<BillToPoMatcherService>());

    builder.Services.AddHostedService(sp => sp.GetRequiredService<MailPollerService>());
    builder.Services.AddHostedService(sp => sp.GetRequiredService<WorkflowCardService>());
    builder.Services.AddHostedService(sp => sp.GetRequiredService<PoSlipDependencyResolver>());
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
    // SMS customer-inquiry pipeline. DAO seam (storage-agnostic) so it can port to Azure SQL by swapping
    // these two registrations — InquiryService + controllers depend only on the interfaces.
    builder.Services.AddSingleton<OutlookShredder.Proxy.Services.Storage.IInquiryStore,
                                  OutlookShredder.Proxy.Services.Storage.SharePointInquiryStore>();
    builder.Services.AddSingleton<OutlookShredder.Proxy.Services.Storage.IMessageStore,
                                  OutlookShredder.Proxy.Services.Storage.SharePointMessageStore>();
    builder.Services.AddSingleton<InquiryService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<InquiryService>());
    // Inbound SMS coverage: all-ingress webhook -> dedup queue (MessageSid) -> one competing-consumer.
    builder.Services.AddSingleton<SmsInboundQueue>();
    builder.Services.AddHostedService<SmsInboundQueueProcessor>();
    builder.Services.AddSingleton<ForgeSchedulerQueue>();
    builder.Services.AddSingleton<ForgeTaskService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<ForgeTaskService>());
    builder.Services.AddSingleton(sp => new Lazy<ForgeTaskService>(() => sp.GetRequiredService<ForgeTaskService>()));
    // ShadowCat — Payment Reconciliation (on-demand; no hosted service in v1).
    builder.Services.AddSingleton<PaymentReconciliationService>();
    builder.Services.AddSingleton<CatalogAnalysisService>();
    builder.Services.AddSingleton(sp => new Lazy<CatalogAnalysisService>(() => sp.GetRequiredService<CatalogAnalysisService>()));
    // Lazy wrapper so SharePointService can generate inbound-email summaries without a construction cycle.
    builder.Services.AddSingleton(sp => new Lazy<RfqSummaryService>(() => sp.GetRequiredService<RfqSummaryService>()));
    builder.Services.AddSingleton(sp => new Lazy<SalesOrderHistoryService>(() => sp.GetRequiredService<SalesOrderHistoryService>()));
    builder.Services.AddSingleton<SupplierProductMappingsCacheService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<SupplierProductMappingsCacheService>());

    // Persistent cache services — registered as ICacheStatusProvider for CacheController enumeration.
    builder.Services.AddSingleton<ICacheStatusProvider>(sp => sp.GetRequiredService<ArchiveCacheService>());
    builder.Services.AddSingleton<ICacheStatusProvider>(sp => sp.GetRequiredService<ProductCatalogService>());
    builder.Services.AddSingleton<ICacheStatusProvider>(sp => sp.GetRequiredService<SupplierCacheService>());
    builder.Services.AddSingleton<ICacheStatusProvider>(sp => sp.GetRequiredService<SliCacheService>());
    builder.Services.AddSingleton<ICacheStatusProvider>(sp => sp.GetRequiredService<SupplierUnreadIndexService>());
    builder.Services.AddSingleton<ICacheStatusProvider>(sp => sp.GetRequiredService<MailCacheService>());
    builder.Services.AddSingleton<ICacheStatusProvider>(sp => sp.GetRequiredService<CustomerCacheService>());
    builder.Services.AddSingleton<ArchiveCacheService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<ArchiveCacheService>());

    // Readiness: tracks per-cache warm state, serves GET /api/ready, pushes cache-ready / all-ready SSE.
    builder.Services.AddSingleton<ReadinessService>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<ReadinessService>());

    builder.Services.AddHostedService<ShutdownWatcherService>();

    var app = builder.Build();

    // WS2 — mint + write the launch token BEFORE Kestrel starts listening, so the file is on disk
    // before the first connection can be accepted (closes the startup token-less race).
    app.Services.GetRequiredService<ProxyAuthService>().EnsureTokenWritten();

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

    // WS2 — local REST auth (DENY-BY-DEFAULT for /api/* except the exempt allowlist).
    // Inserted before the HTTP-log middleware so the log line still records the final status
    // (including a 401 this produces). Warn-only this release: would-rejects are logged, not blocked.
    app.Use(async (ctx, next) =>
    {
        if (!ctx.Request.Path.StartsWithSegments("/api") || ProxyAuthService.IsExempt(ctx))
        {
            await next();
            return;
        }

        var auth   = ctx.RequestServices.GetRequiredService<ProxyAuthService>();
        var logger = ctx.RequestServices.GetRequiredService<ILogger<Program>>();
        var pid    = ctx.Request.Headers[ProxyAuthService.PidHeader].ToString();
        if (pid.Length == 0) pid = "-";

        // Per-second rate cap on non-exempt traffic (blunts a runaway/abusive loop).
        if (!auth.CheckRate())
        {
            if (auth.Enforce)
            {
                ctx.Response.StatusCode  = 429;
                ctx.Response.ContentType = "application/json";
                await ctx.Response.WriteAsync("{\"success\":false,\"error\":\"rate limited\"}");
                logger.LogWarning("[Auth] rate-limited {Method} {Path} pid={Pid}", ctx.Request.Method, ctx.Request.Path, pid);
                return;
            }
            logger.LogWarning("[Auth] rate-warn {Method} {Path} pid={Pid}", ctx.Request.Method, ctx.Request.Path, pid);
        }

        var result = auth.Verify(ctx);
        if (result != ProxyAuthService.AuthResult.Ok)
        {
            if (auth.Enforce)
            {
                ctx.Response.StatusCode  = 401;
                ctx.Response.ContentType = "application/json";
                await ctx.Response.WriteAsync("{\"success\":false,\"error\":\"unauthorized\"}");
                logger.LogWarning("[Auth] reject {Method} {Path} reason={Reason} pid={Pid}",
                    ctx.Request.Method, ctx.Request.Path, result, pid);
                return;
            }
            logger.LogWarning("[Auth] would-reject {Method} {Path} reason={Reason} pid={Pid}",
                ctx.Request.Method, ctx.Request.Path, result, pid);
        }

        // Audit every mutating call that passed (or warn-passed) — caller + intent recorded before dispatch.
        var method = ctx.Request.Method;
        if (HttpMethods.IsPost(method) || HttpMethods.IsPut(method)
            || HttpMethods.IsPatch(method) || HttpMethods.IsDelete(method))
        {
            logger.LogInformation("[Audit] {Method} {Path} caller-pid={Pid} result={Result}",
                method, ctx.Request.Path, pid, result == ProxyAuthService.AuthResult.Ok ? "allow" : "warn-pass");
        }

        await next();
    });

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

                // WS1 cleanup: SharePoint app-only auth (using the vault's ClientSecret) just succeeded, so
                // the vault credentials are proven healthy — delete the local cleartext secret copies. No-op
                // unless the vault supplied every secret. The central OneDrive copy is kept for now.
                SecretsBootstrap.CleanupLocalSecretsIfVaultHealthy(Log.Logger);

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

                // Read-state clean-slate is now PER USER and lazy: each user's SupplierReadProfiles row is
                // created with AdoptedAt=now on their first unread fetch. The old team-wide one-time backfill
                // (SupplierReadBackfillDone) is retired.
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

