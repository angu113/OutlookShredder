using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Behaviour tests for the MSPC-based fallback link suggestion: pair an active PO with an SO when they
// share a product code, excluding already-linked and rejected pairs. Pure (in/out).
public class TransferLinkSuggestionTests
{
    private static TransferLinkSuggestionService.PoView Po(string po, string[] mspcs,
        string[]? linked = null, string[]? rejected = null) =>
        new(po, "sp" + po, new(mspcs, StringComparer.OrdinalIgnoreCase),
            new(linked ?? [], StringComparer.OrdinalIgnoreCase),
            new(rejected ?? [], StringComparer.OrdinalIgnoreCase));

    [Fact]
    public void Mspcs_extracted_from_both_po_mspc_and_erp_code_fields()
    {
        Assert.Equal(new[] { "CTR1018D/5188" },
            TransferLinkSuggestionService.MspcsFromLineItems("""[{"product":"x","mspc":"CTR1018D/5188"}]"""));
        Assert.Equal(new[] { "HRF/50071" },
            TransferLinkSuggestionService.MspcsFromLineItems("""[{"Description":"y","Code":"HRF/50071"}]"""));
        // codes without '/' (free-text / non-MSPC) are ignored
        Assert.Empty(TransferLinkSuggestionService.MspcsFromLineItems("""[{"code":"WIDGET"}]"""));
        Assert.Empty(TransferLinkSuggestionService.MspcsFromLineItems(null));
    }

    [Fact]
    public void Suggests_a_pair_that_shares_an_mspc()
    {
        var pos = new[] { Po("HSK-PO1", ["A/1", "B/2"]) };
        var so  = new Dictionary<string, HashSet<string>> { ["HSK-SO9"] = new(["B/2"], System.StringComparer.OrdinalIgnoreCase) };
        var s = Assert.Single(TransferLinkSuggestionService.Compute(pos, so));
        Assert.Equal("HSK-PO1", s.PoNumber);
        Assert.Equal("HSK-SO9", s.SoNumber);
        Assert.Equal(1, s.SharedMspcs);
    }

    [Fact]
    public void No_shared_mspc_means_no_suggestion()
    {
        var pos = new[] { Po("HSK-PO1", ["A/1"]) };
        var so  = new Dictionary<string, HashSet<string>> { ["HSK-SO9"] = new(["Z/9"], System.StringComparer.OrdinalIgnoreCase) };
        Assert.Empty(TransferLinkSuggestionService.Compute(pos, so));
    }

    [Fact]
    public void Already_linked_or_rejected_pairs_are_excluded()
    {
        var so = new Dictionary<string, HashSet<string>> { ["HSK-SO9"] = new(["A/1"], System.StringComparer.OrdinalIgnoreCase) };
        Assert.Empty(TransferLinkSuggestionService.Compute(new[] { Po("HSK-PO1", ["A/1"], linked: ["HSK-SO9"]) }, so));
        Assert.Empty(TransferLinkSuggestionService.Compute(new[] { Po("HSK-PO1", ["A/1"], rejected: ["HSK-SO9"]) }, so));
    }

    [Fact]
    public void Lists_all_matching_candidates()
    {
        var pos = new[] { Po("HSK-PO1", ["A/1"]), Po("HSK-PO2", ["A/1", "C/3"]) };
        var so  = new Dictionary<string, HashSet<string>>
        {
            ["HSK-SO9"] = new(["A/1"], System.StringComparer.OrdinalIgnoreCase),
            ["HSK-SO8"] = new(["C/3"], System.StringComparer.OrdinalIgnoreCase),
        };
        var all = TransferLinkSuggestionService.Compute(pos, so);
        Assert.Equal(3, all.Count);   // PO1<->SO9, PO2<->SO9, PO2<->SO8
        Assert.Contains(all, s => s.PoNumber == "HSK-PO2" && s.SoNumber == "HSK-SO8");
    }
}
