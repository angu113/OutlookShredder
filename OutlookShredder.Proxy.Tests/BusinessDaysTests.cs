using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Guards the worklist read window: SubtractBusinessDays must skip weekends so the 5-business-day
// board window is a true working week (Mon cutoff = the prior Mon, not the prior Wed).
public class BusinessDaysTests
{
    [Fact]
    public void Zero_business_days_returns_same_date_stripped_of_time()
    {
        var from = new DateTime(2026, 6, 22, 14, 30, 0);   // Monday, with a time component
        Assert.Equal(new DateTime(2026, 6, 22), SharePointService.SubtractBusinessDays(from, 0));
    }

    [Fact]
    public void Five_business_days_back_from_monday_is_prior_monday()
    {
        var monday = new DateTime(2026, 6, 22);
        Assert.Equal(DayOfWeek.Monday, monday.DayOfWeek);            // sanity-anchor the fixture
        // Fri 19, Thu 18, Wed 17, Tue 16, Mon 15 — one weekend skipped.
        Assert.Equal(new DateTime(2026, 6, 15), SharePointService.SubtractBusinessDays(monday, 5));
    }

    [Fact]
    public void One_business_day_back_from_monday_is_prior_friday()
    {
        var monday = new DateTime(2026, 6, 22);
        Assert.Equal(new DateTime(2026, 6, 19), SharePointService.SubtractBusinessDays(monday, 1));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(5)]
    [InlineData(20)]
    public void Result_is_never_a_weekend(int days)
    {
        // Walk a full week of start dates; the cutoff must always land on a weekday.
        for (int i = 0; i < 7; i++)
        {
            var d = SharePointService.SubtractBusinessDays(new DateTime(2026, 6, 22).AddDays(i), days);
            Assert.True(d.DayOfWeek is not DayOfWeek.Saturday and not DayOfWeek.Sunday, $"start+{i}, {days}d -> {d:yyyy-MM-dd ddd}");
        }
    }
}
