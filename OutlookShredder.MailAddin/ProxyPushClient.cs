using System;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using OutlookShredder.MailAddin.Models;

namespace OutlookShredder.MailAddin;

internal class ProxyPushClient
{
    private readonly HttpClient _http;
    private readonly string     _proxyUrl;

    public ProxyPushClient(string proxyUrl)
    {
        _proxyUrl = proxyUrl.TrimEnd('/');
        _http     = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
    }

    public async Task SendAsync(MailMessagePayload payload)
    {
        try
        {
            var ser     = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var json    = ser.Serialize(payload);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            await _http.PostAsync($"{_proxyUrl}/api/addin/email-received", content).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            Log($"Push to proxy failed: {ex.Message}");
        }
    }

    internal static void Log(string msg)
    {
        try
        {
            var dir  = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? ".";
            var path = Path.Combine(dir, "addin.log");
            File.AppendAllText(path, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {msg}{Environment.NewLine}");
        }
        catch { }
    }
}
