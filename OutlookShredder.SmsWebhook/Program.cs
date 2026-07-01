using Azure.Messaging.ServiceBus;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureServices(services =>
    {
        // ServiceBusClient for inbound. Connection string carries a Send claim on sms-inbound-jobs ONLY
        // (least-privilege, entity-scoped SAS) — it cannot send to any other queue.
        services.AddSingleton(sp =>
        {
            var conn = sp.GetRequiredService<IConfiguration>()["ServiceBus:ConnectionString"]
                       ?? throw new InvalidOperationException("ServiceBus:ConnectionString is not configured");
            return new ServiceBusClient(conn);
        });
        // Separate client (distinct type, not keyed DI — avoids a package bump for FromKeyedServices) for
        // delivery-status callbacks: a distinct entity-scoped Send claim on sms-status-jobs. Kept as its own
        // client (not reusing the one above) so each queue's SAS policy stays scoped to exactly its entity.
        services.AddSingleton(sp =>
        {
            var conn = sp.GetRequiredService<IConfiguration>()["ServiceBus:StatusQueueConnectionString"]
                       ?? throw new InvalidOperationException("ServiceBus:StatusQueueConnectionString is not configured");
            return new StatusServiceBusClient(new ServiceBusClient(conn));
        });
    })
    .Build();

host.Run();

/// <summary>Distinct wrapper type so DI can hold two <see cref="ServiceBusClient"/> instances (inbound vs.
/// status) without keyed-services support — see the ConfigureServices comments above for why they must
/// stay separate (each carries a different entity-scoped SAS claim).</summary>
public sealed record StatusServiceBusClient(ServiceBusClient Client);
