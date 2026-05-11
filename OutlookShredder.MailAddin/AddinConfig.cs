using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Web.Script.Serialization;

namespace OutlookShredder.MailAddin;

public class AddinConfig
{
    public string       ProxyUrl           { get; set; } = "http://localhost:7000";
    public int          ListenPort         { get; set; } = 7002;
    public List<string> MonitoredMailboxes { get; set; } = new List<string>();
    public long         MaxAttachmentBytes { get; set; } = 52_428_800; // 50 MB

    public static AddinConfig Load()
    {
        var dir  = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? ".";
        var path = Path.Combine(dir, "addin-config.json");
        if (!File.Exists(path)) return new AddinConfig();
        try
        {
            var ser = new JavaScriptSerializer();
            return ser.Deserialize<AddinConfig>(File.ReadAllText(path)) ?? new AddinConfig();
        }
        catch
        {
            return new AddinConfig();
        }
    }
}
