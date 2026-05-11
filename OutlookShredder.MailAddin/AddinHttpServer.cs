using System;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using OutlookShredder.MailAddin.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookShredder.MailAddin;

internal class AddinHttpServer
{
    private readonly int                    _port;
    private readonly Outlook.Application    _app;
    private readonly SynchronizationContext _staCtx;
    private HttpListener?                   _listener;

    public AddinHttpServer(int port, Outlook.Application app, SynchronizationContext staCtx)
    {
        _port   = port;
        _app    = app;
        _staCtx = staCtx;
    }

    public void Start()
    {
        _listener = new HttpListener();
        _listener.Prefixes.Add($"http://localhost:{_port}/");
        _listener.Start();
        Task.Run(ListenLoop);
        ProxyPushClient.Log($"HTTP listener started on port {_port}");
    }

    public void Stop()
    {
        try { _listener?.Stop(); }
        catch { }
    }

    private async Task ListenLoop()
    {
        while (_listener?.IsListening == true)
        {
            try
            {
                var ctx = await _listener.GetContextAsync().ConfigureAwait(false);
                _ = Task.Run(() => HandleRequest(ctx));
            }
            catch (HttpListenerException) { break; }
            catch (ObjectDisposedException) { break; }
            catch { }
        }
    }

    private async Task HandleRequest(HttpListenerContext ctx)
    {
        var path   = ctx.Request.Url?.AbsolutePath ?? "/";
        var method = ctx.Request.HttpMethod;

        try
        {
            if (method == "GET" && path == "/health")
            {
                WriteJson(ctx.Response, 200, new { status = "ok", port = _port });
                return;
            }
            if (method == "POST" && path == "/fetch")
            {
                await HandleFetch(ctx).ConfigureAwait(false);
                return;
            }
            if (method == "POST" && path == "/send")
            {
                await HandleSend(ctx).ConfigureAwait(false);
                return;
            }
            WriteJson(ctx.Response, 404, new { error = "not found" });
        }
        catch (Exception ex)
        {
            ProxyPushClient.Log($"Request error {method} {path}: {ex.Message}");
            try { WriteJson(ctx.Response, 500, new { error = ex.Message }); }
            catch { }
        }
    }

    private async Task HandleFetch(HttpListenerContext ctx)
    {
        var req = await ReadJson<FetchRequest>(ctx.Request).ConfigureAwait(false);

        MailMessagePayload? payload = null;
        Exception?          error   = null;
        var                 done    = new ManualResetEventSlim(false);

        _staCtx.Post(_ =>
        {
            try
            {
                object raw = string.IsNullOrEmpty(req.StoreId)
                    ? _app.Session.GetItemFromID(req.EntryId)
                    : _app.Session.GetItemFromID(req.EntryId, req.StoreId);

                if (raw is Outlook.MailItem mail)
                    payload = OutlookReader.BuildPayload(mail, req.StoreId);
                else
                    error = new InvalidOperationException("Item is not a MailItem");
            }
            catch (Exception ex) { error = ex; }
            finally { done.Set(); }
        }, null);

        if (!done.Wait(TimeSpan.FromSeconds(30)))
            throw new TimeoutException("Outlook STA call timed out after 30 s");
        if (error != null) throw error;

        WriteJson(ctx.Response, 200, payload);
    }

    private async Task HandleSend(HttpListenerContext ctx)
    {
        var req = await ReadJson<SendRequest>(ctx.Request).ConfigureAwait(false);

        Exception? error = null;
        var        done  = new ManualResetEventSlim(false);

        _staCtx.Post(_ =>
        {
            Outlook.MailItem? mail = null;
            try
            {
                mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);
                mail.Subject = req.Subject ?? string.Empty;
                mail.To      = req.To      ?? string.Empty;
                if (!string.IsNullOrEmpty(req.Cc))  mail.CC  = req.Cc;
                if (!string.IsNullOrEmpty(req.Bcc)) mail.BCC = req.Bcc;

                if (!string.IsNullOrEmpty(req.BodyHtml))
                    mail.HTMLBody = req.BodyHtml;
                else
                    mail.Body = req.BodyText ?? string.Empty;

                if (!string.IsNullOrEmpty(req.FromAccount))
                {
                    foreach (Outlook.Account account in _app.Session.Accounts)
                    {
                        if (account.SmtpAddress.Equals(req.FromAccount, StringComparison.OrdinalIgnoreCase))
                        {
                            mail.SendUsingAccount = account;
                            break;
                        }
                    }
                }

                mail.Send();
            }
            catch (Exception ex) { error = ex; }
            finally
            {
                if (mail != null) Marshal.ReleaseComObject(mail);
                done.Set();
            }
        }, null);

        if (!done.Wait(TimeSpan.FromSeconds(30)))
            throw new TimeoutException("Outlook STA call timed out after 30 s");
        if (error != null) throw error;

        WriteJson(ctx.Response, 200, new { sent = true });
    }

    private static async Task<T> ReadJson<T>(HttpListenerRequest request) where T : new()
    {
        using var sr = new StreamReader(request.InputStream, Encoding.UTF8);
        var body = await sr.ReadToEndAsync().ConfigureAwait(false);
        if (string.IsNullOrWhiteSpace(body)) return new T();
        try
        {
            var ser = new JavaScriptSerializer();
            return ser.Deserialize<T>(body) ?? new T();
        }
        catch { return new T(); }
    }

    private static void WriteJson(HttpListenerResponse response, int status, object? obj)
    {
        var ser   = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        var bytes = Encoding.UTF8.GetBytes(ser.Serialize(obj));
        response.StatusCode      = status;
        response.ContentType     = "application/json; charset=utf-8";
        response.ContentLength64 = bytes.Length;
        response.OutputStream.Write(bytes, 0, bytes.Length);
        response.OutputStream.Close();
    }
}

internal class FetchRequest
{
    public string  EntryId { get; set; } = string.Empty;
    public string? StoreId { get; set; }
}

internal class SendRequest
{
    public string? FromAccount { get; set; }
    public string? To          { get; set; }
    public string? Cc          { get; set; }
    public string? Bcc         { get; set; }
    public string? Subject     { get; set; }
    public string? BodyHtml    { get; set; }
    public string? BodyText    { get; set; }
}
