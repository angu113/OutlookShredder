using Azure.Messaging.ServiceBus;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureServices(services =>
    {
        // One ServiceBusClient for the app. Connection string carries a Send claim on sms-inbound-jobs.
        services.AddSingleton(sp =>
        {
            var conn = sp.GetRequiredService<IConfiguration>()["ServiceBus:ConnectionString"]
                       ?? throw new InvalidOperationException("ServiceBus:ConnectionString is not configured");
            return new ServiceBusClient(conn);
        });
    })
    .Build();

host.Run();
