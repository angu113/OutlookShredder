using Microsoft.Extensions.Logging.Abstractions;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Tests;

public class RoundRobinAiExtractionServiceTests
{
    private static RoundRobinAiExtractionService Build(StubAiService a, StubAiService b) =>
        new(a, b, NullLogger<RoundRobinAiExtractionService>.Instance);

    [Fact]
    public async Task AlternatesBetweenProvidersAcrossCalls()
    {
        var a = new StubAiService("A") { RfqResult = new RfqExtraction { SupplierName = "from-A" } };
        var b = new StubAiService("B") { RfqResult = new RfqExtraction { SupplierName = "from-B" } };

        var sut = Build(a, b);
        var r1 = await sut.ExtractRfqAsync(new ExtractRequest());
        var r2 = await sut.ExtractRfqAsync(new ExtractRequest());
        var r3 = await sut.ExtractRfqAsync(new ExtractRequest());
        var r4 = await sut.ExtractRfqAsync(new ExtractRequest());

        Assert.Equal("from-A", r1?.SupplierName);
        Assert.Equal("from-B", r2?.SupplierName);
        Assert.Equal("from-A", r3?.SupplierName);
        Assert.Equal("from-B", r4?.SupplierName);
        Assert.Equal(2, a.RfqCallCount);
        Assert.Equal(2, b.RfqCallCount);
    }

    [Fact]
    public async Task PrimaryA_Fails_FallsBackToB()
    {
        var a = new StubAiService("A") { RfqException = new HttpRequestException("A down") };
        var b = new StubAiService("B") { RfqResult = new RfqExtraction { SupplierName = "from-B" } };

        var sut = Build(a, b);
        var result = await sut.ExtractRfqAsync(new ExtractRequest());

        Assert.Equal("from-B", result?.SupplierName);
        Assert.Equal(1, a.RfqCallCount);
        Assert.Equal(1, b.RfqCallCount);
    }

    [Fact]
    public async Task PrimaryB_Fails_FallsBackToA()
    {
        var a = new StubAiService("A") { RfqResult = new RfqExtraction { SupplierName = "from-A" } };
        var b = new StubAiService("B") { RfqException = new HttpRequestException("B down") };

        var sut = Build(a, b);
        // First call picks A (counter 0 → A), so consume it to land the second call on B as primary.
        _ = await sut.ExtractRfqAsync(new ExtractRequest());
        var result = await sut.ExtractRfqAsync(new ExtractRequest());

        Assert.Equal("from-A", result?.SupplierName);
        Assert.Equal(1, b.RfqCallCount);
        Assert.Equal(2, a.RfqCallCount);
    }

    [Fact]
    public async Task TimeoutTreatedAsFailure_FallsBackToOther()
    {
        var a = new StubAiService("A") { RfqException = new TaskCanceledException("simulated http timeout") };
        var b = new StubAiService("B") { RfqResult = new RfqExtraction { SupplierName = "from-B" } };

        var sut = Build(a, b);
        var result = await sut.ExtractRfqAsync(new ExtractRequest(), CancellationToken.None);

        Assert.Equal("from-B", result?.SupplierName);
        Assert.Equal(1, a.RfqCallCount);
        Assert.Equal(1, b.RfqCallCount);
    }

    [Fact]
    public async Task CallerCancelled_NoFallback_ExceptionPropagates()
    {
        var a = new StubAiService("A") { RfqException = new OperationCanceledException("caller cancelled") };
        var b = new StubAiService("B") { RfqResult = new RfqExtraction { SupplierName = "from-B" } };

        using var cts = new CancellationTokenSource();
        cts.Cancel();

        var sut = Build(a, b);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => sut.ExtractRfqAsync(new ExtractRequest(), cts.Token));

        Assert.Equal(1, a.RfqCallCount);
        Assert.Equal(0, b.RfqCallCount);
    }

    [Fact]
    public async Task BothFail_SecondaryExceptionBubbles()
    {
        var a = new StubAiService("A") { RfqException = new HttpRequestException("A down") };
        var b = new StubAiService("B") { RfqException = new InvalidOperationException("B down") };

        var sut = Build(a, b);
        var ex = await Assert.ThrowsAsync<InvalidOperationException>(
            () => sut.ExtractRfqAsync(new ExtractRequest()));

        Assert.Equal("B down", ex.Message);
        Assert.Equal(1, a.RfqCallCount);
        Assert.Equal(1, b.RfqCallCount);
    }

    [Fact]
    public async Task PoExtraction_AlternatesAndFallsBack()
    {
        var a = new StubAiService("A") { PoResult = new PoExtraction { PoNumber = "PO-A" } };
        var b = new StubAiService("B") { PoException = new HttpRequestException("B down") };

        var sut = Build(a, b);

        var r1 = await sut.ExtractPurchaseOrderAsync("b64", "po.pdf", "ctx", "subj", []);
        Assert.Equal("PO-A", r1?.PoNumber);

        // Second call — primary is B (fails) → falls back to A.
        var r2 = await sut.ExtractPurchaseOrderAsync("b64", "po.pdf", "ctx", "subj", []);
        Assert.Equal("PO-A", r2?.PoNumber);

        Assert.Equal(2, a.PoCallCount);
        Assert.Equal(1, b.PoCallCount);
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
