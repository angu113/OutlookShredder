namespace OutlookShredder.Proxy.Models;

public record MailStatus(
    PollerStatus   Poller,
    ReprocessStatus Reprocess,
    RateLimitStatus RateLimit,
    List<InFlightItem> InFlight);

public record PollerStatus(
    bool            Running,
    DateTimeOffset? LastPollAt,
    int             MessagesFoundLastCycle);

public record ReprocessStatus(
    bool   Active,
    int    Total,
    int    Completed,
    int    Failed,
    double PercentComplete);

public record RateLimitStatus(
    int CallsInLastMinute,
    int MaxPerMinute,
    int SlotsAvailable);

public record InFlightItem(
    string Subject,
    string From,
    string StartedAt);
