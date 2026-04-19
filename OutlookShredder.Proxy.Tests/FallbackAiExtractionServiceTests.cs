using Microsoft.Extensions.Logging.Abstractions;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Tests;

public class FallbackAiExtractionServiceTests
{
    private static FallbackAiExtractionService Build(StubAiService primary, StubAiService secondary) =>
        new(primary, secondary, NullLogger<FallbackAiExtractionService>.Instance);

    [Fact]
    public async Task PrimarySucceeds_SecondaryNotCalled()
    {
        var primary = new StubAiService("Primary") { RfqResult = new RfqExtraction { SupplierName = "ACME" } };
        var secondary = new StubAiService("Secondary");

        var sut = Build(primary, secondary);
        var result = await sut.ExtractRfqAsync(new ExtractRequest());

        Assert.Equal("ACME", result?.SupplierName);
        Assert.Equal(1, primary.RfqCallCount);
        Assert.Equal(0, secondary.RfqCallCount);
    }

    [Fact]
    public async Task PrimaryThrows_SecondaryReturnsResult()
    {
        var primary = new StubAiService("Primary") { RfqException = new HttpRequestException("simulated 503") };
        var secondary = new StubAiService("Secondary") { RfqResult = new RfqExtraction { SupplierName = "Backup Co" } };

        var sut = Build(primary, secondary);
        var result = await sut.ExtractRfqAsync(new ExtractRequest());

        Assert.Equal("Backup Co", result?.SupplierName);
        Assert.Equal(1, primary.RfqCallCount);
        Assert.Equal(1, secondary.RfqCallCount);
    }

    [Fact]
    public async Task PrimaryThrows_TimeoutTreatedAsFailure_SecondaryCalled()
    {
        var primary = new StubAiService("Primary") { RfqException = new TaskCanceledException("simulated http timeout") };
        var secondary = new StubAiService("Secondary") { RfqResult = new RfqExtraction { SupplierName = "Backup Co" } };

        var sut = Build(primary, secondary);
        // Caller did NOT cancel — TaskCanceledException originated inside HttpClient.
        var result = await sut.ExtractRfqAsync(new ExtractRequest(), CancellationToken.None);

        Assert.Equal("Backup Co", result?.SupplierName);
        Assert.Equal(1, primary.RfqCallCount);
        Assert.Equal(1, secondary.RfqCallCount);
    }

    [Fact]
    public async Task CallerCancelled_SecondaryNotCalled_ExceptionPropagates()
    {
        var primary = new StubAiService("Primary") { RfqException = new OperationCanceledException("caller cancelled") };
        var secondary = new StubAiService("Secondary") { RfqResult = new RfqExtraction { SupplierName = "Backup Co" } };

        var cts = new CancellationTokenSource();
        cts.Cancel();

        var sut = Build(primary, secondary);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => sut.ExtractRfqAsync(new ExtractRequest(), cts.Token));

        Assert.Equal(1, primary.RfqCallCount);
        Assert.Equal(0, secondary.RfqCallCount);
    }

    [Fact]
    public async Task PrimaryThrows_SecondaryAlsoThrows_SecondaryExceptionBubbles()
    {
        var primary = new StubAiService("Primary") { RfqException = new HttpRequestException("primary 503") };
        var secondary = new StubAiService("Secondary") { RfqException = new InvalidOperationException("secondary unavailable") };

        var sut = Build(primary, secondary);
        var ex = await Assert.ThrowsAsync<InvalidOperationException>(
            () => sut.ExtractRfqAsync(new ExtractRequest()));

        Assert.Equal("secondary unavailable", ex.Message);
        Assert.Equal(1, primary.RfqCallCount);
        Assert.Equal(1, secondary.RfqCallCount);
    }

    [Fact]
    public async Task PoExtraction_PrimaryThrows_SecondaryCalled()
    {
        var primary = new StubAiService("Primary") { PoException = new HttpRequestException("simulated 500") };
        var secondary = new StubAiService("Secondary") { PoResult = new PoExtraction { PoNumber = "HSK-PO-123" } };

        var sut = Build(primary, secondary);
        var result = await sut.ExtractPurchaseOrderAsync("base64==", "po.pdf", "context", "subject", []);

        Assert.Equal("HSK-PO-123", result?.PoNumber);
        Assert.Equal(1, primary.PoCallCount);
        Assert.Equal(1, secondary.PoCallCount);
    }

    private class StubAiService : IAiExtractionService
    {
        public StubAiService(string name) => ProviderName = name;
        public string ProviderName { get; }

        public RfqExtraction? RfqResult { get; set; }
        public Exception? RfqException { get; set; }
        public int RfqCallCount { get; private set; }

        public PoExtraction? PoResult { get; set; }
        public Exception? PoException { get; set; }
        public int PoCallCount { get; private set; }

        public Task<RfqExtraction?> ExtractRfqAsync(ExtractRequest request, CancellationToken ct = default)
        {
            RfqCallCount++;
            if (RfqException != null) throw RfqException;
            return Task.FromResult(RfqResult);
        }

        public Task<PoExtraction?> ExtractPurchaseOrderAsync(
            string base64Pdf, string fileName, string emailBodyContext,
            string emailSubject, List<string> jobRefs, CancellationToken ct = default)
        {
            PoCallCount++;
            if (PoException != null) throw PoException;
            return Task.FromResult(PoResult);
        }
    }
}
