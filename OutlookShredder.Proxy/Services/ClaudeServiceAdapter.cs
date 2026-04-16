using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services.Ai;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Adapter that makes the existing ClaudeService implement IAiProvider.
/// This allows the service to be used through the provider interface while maintaining
/// backward compatibility with existing code that depends on ClaudeService directly.
/// </summary>
public class ClaudeServiceAdapter : IAiProvider
{
    private readonly ClaudeService _claudeService;

    public string Name => "claude";

    public ClaudeServiceAdapter(ClaudeService claudeService)
    {
        _claudeService = claudeService;
    }

    public async Task<RfqExtraction?> ExtractAsync(ExtractRequest request, CancellationToken cancellationToken = default)
    {
        // The original ClaudeService.ExtractAsync doesn't support CancellationToken,
        // but we can wrap the call if needed in the future.
        // For now, we delegate directly.
        return await _claudeService.ExtractAsync(request);
    }
}
