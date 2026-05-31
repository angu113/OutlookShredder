using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Phase 1a of the mail workbench (wip/mail-classification.md): validate classification quality
/// on real forwarded mail BEFORE building SP persistence. Preview only — no writes.
/// </summary>
[ApiController]
[Route("api/mail-classify")]
public sealed class MailClassifyController : ControllerBase
{
    private readonly MailboxBridgeService _bridge;
    private readonly MailClassifierService _classifier;
    private readonly MailWorkbenchService _workbench;
    private readonly MailTaxonomyService _taxonomy;
    private readonly MailProjectService _projects;
    private readonly SharePointService _sp;

    public MailClassifyController(MailboxBridgeService bridge, MailClassifierService classifier,
        MailWorkbenchService workbench, MailTaxonomyService taxonomy, MailProjectService projects, SharePointService sp)
    {
        _bridge     = bridge;
        _classifier = classifier;
        _workbench  = workbench;
        _taxonomy   = taxonomy;
        _projects   = projects;
        _sp         = sp;
    }

    /// <summary>Idempotently provision the MailItems + MailClassifications SP lists (Phase 1b setup).</summary>
    [HttpPost("setup-lists")]
    public async Task<IActionResult> SetupLists() => Ok(await _sp.EnsureMailListsAsync());

    /// <summary>Capture a cached bridge message → MailItems (dedup) → classify → MailClassifications.</summary>
    [HttpPost("capture")]
    public async Task<IActionResult> Capture([FromBody] PreviewRequest req, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.Id)) return BadRequest(new { error = "id is required." });
        var upn = string.IsNullOrWhiteSpace(req.Upn) ? _bridge.GetStatuses().FirstOrDefault()?.WatchedUpn : req.Upn;
        if (string.IsNullOrWhiteSpace(upn)) return BadRequest(new { error = "No watched mailbox configured." });
        try { return Ok(await _workbench.CaptureAndClassifyAsync(upn, req.Id, ct: ct)); }
        catch (Exception ex) { return StatusCode(500, new { error = ex.Message }); }
    }

    /// <summary>Bulk capture+classify every message the bridge currently surfaces (background, idempotent).</summary>
    [HttpPost("capture-all")]
    public IActionResult CaptureAll([FromQuery] string? upn)
    {
        var u = string.IsNullOrWhiteSpace(upn) ? _bridge.GetStatuses().FirstOrDefault()?.WatchedUpn : upn;
        if (string.IsNullOrWhiteSpace(u)) return BadRequest(new { error = "No watched mailbox configured." });
        return Ok(_workbench.StartCaptureAll(u));
    }

    /// <summary>Backfill doc-library storage (attachments + raw .eml) for already-captured items missing it.</summary>
    [HttpPost("store-attachments")]
    public IActionResult StoreAttachments([FromQuery] string? upn)
    {
        var u = string.IsNullOrWhiteSpace(upn) ? _bridge.GetStatuses().FirstOrDefault()?.WatchedUpn : upn;
        if (string.IsNullOrWhiteSpace(u)) return BadRequest(new { error = "No watched mailbox configured." });
        return Ok(_workbench.StartStoreAttachments(u));
    }

    /// <summary>Full-folder backfill: capture+classify EVERY forward-as-attachment in the mirror folder (background).</summary>
    [HttpPost("backfill")]
    public IActionResult Backfill([FromQuery] string? upn)
    {
        var u = string.IsNullOrWhiteSpace(upn) ? _bridge.GetStatuses().FirstOrDefault()?.WatchedUpn : upn;
        if (string.IsNullOrWhiteSpace(u)) return BadRequest(new { error = "No watched mailbox configured." });
        return Ok(_workbench.StartBackfill(u));
    }

    /// <summary>Re-classify every stored MailItem with the current taxonomy (background; new version per item).</summary>
    [HttpPost("reclassify-all")]
    public IActionResult ReclassifyAll() => Ok(_workbench.StartReclassifyAll());

    /// <summary>Repair ReceivedAt on stored items from their archived .eml headers (original delivery time). Background.</summary>
    [HttpPost("repair-received")]
    public IActionResult RepairReceived() => Ok(_workbench.StartRepairReceived());

    /// <summary>Progress of the bulk capture / reclassify pass (shared tracker).</summary>
    [HttpGet("seed-status")]
    public IActionResult SeedStatus() => Ok(_workbench.GetSeedStatus());

    // ── Projects (Layer 2: cross-conversation grouping) ──────────────────────────────

    /// <summary>Auto-detected project suggestions: conversations sharing a business ref, not yet grouped.</summary>
    [HttpGet("project-suggestions")]
    public async Task<IActionResult> ProjectSuggestions(CancellationToken ct) => Ok(await _projects.SuggestAsync(ct));

    /// <summary>Active projects (cross-conversation clusters).</summary>
    [HttpGet("projects")]
    public async Task<IActionResult> Projects([FromQuery] bool includeArchived = false, CancellationToken ct = default)
        => Ok(await _projects.GetProjectsAsync(activeOnly: !includeArchived, ct));

    /// <summary>Create/confirm a project from a set of conversations (+ the refs that linked them).</summary>
    [HttpPost("projects")]
    public async Task<IActionResult> CreateProject([FromBody] CreateProjectRequest req, CancellationToken ct)
    {
        if (req is null || req.ConversationIds is not { Count: > 0 })
            return BadRequest(new { error = "conversationIds are required." });
        var p = await _projects.CreateAsync(req.Name ?? "Untitled project", req.ConversationIds, req.Refs ?? [], req.By, ct);
        return Ok(p);
    }

    /// <summary>Rename or change a project's conversation membership.</summary>
    [HttpPatch("projects/{projectId}")]
    public async Task<IActionResult> UpdateProject(string projectId, [FromBody] UpdateProjectRequest req, CancellationToken ct)
        => await _projects.UpdateAsync(projectId, req?.Name, req?.ConversationIds, ct)
            ? Ok(new { success = true }) : NotFound(new { error = "Project not found." });

    /// <summary>Archive (hide) a project.</summary>
    [HttpPost("projects/{projectId}/archive")]
    public async Task<IActionResult> ArchiveProject(string projectId, CancellationToken ct)
        => await _projects.ArchiveAsync(projectId, ct) ? Ok(new { success = true }) : NotFound(new { error = "Project not found." });

    public sealed class CreateProjectRequest
    {
        public string? Name { get; set; }
        public List<string>? ConversationIds { get; set; }
        public List<string>? Refs { get; set; }
        public string? By { get; set; }
    }
    public sealed class UpdateProjectRequest
    {
        public string? Name { get; set; }
        public List<string>? ConversationIds { get; set; }
    }

    /// <summary>The classification tree with per-leaf total/open/completed counts.</summary>
    [HttpGet("tree")]
    public async Task<IActionResult> Tree(CancellationToken ct) => Ok(await _workbench.GetTreeAsync(ct));

    /// <summary>Full detail for the viewer: HTML body (from the archived .eml), headers, attachments, classification, state.</summary>
    [HttpGet("item/{mailItemId}")]
    public async Task<IActionResult> Item(string mailItemId, CancellationToken ct)
    {
        var d = await _workbench.GetItemDetailAsync(mailItemId, ct);
        return d is null ? NotFound(new { error = "Item not found." }) : Ok(d);
    }

    /// <summary>Mark an item read/unread (global read-by-anyone). Updates cache + broadcasts a "Read" bus event.</summary>
    [HttpPost("read/{mailItemId}")]
    public async Task<IActionResult> Read(string mailItemId, [FromBody] ReadRequest req, CancellationToken ct)
    {
        var ok = await _workbench.MarkReadAsync(mailItemId, req?.Read ?? true, req?.By, ct);
        return ok ? Ok(new { success = true }) : NotFound(new { error = "Item not found." });
    }

    public sealed class ReadRequest { public bool Read { get; set; } = true; public string? By { get; set; } }

    /// <summary>Streams a stored attachment's bytes (re-rooted under this machine's OneDrive archive).</summary>
    [HttpGet("attachment/{mailItemId}")]
    public async Task<IActionResult> Attachment(string mailItemId, [FromQuery] string name, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(name)) return BadRequest(new { error = "name is required." });
        var r = await _workbench.GetItemAttachmentAsync(mailItemId, name, ct);
        return r is null ? NotFound() : File(r.Value.Bytes, r.Value.ContentType, r.Value.Name);
    }

    /// <summary>Items currently classified under a taxonomy path.</summary>
    [HttpGet("items")]
    public async Task<IActionResult> Items([FromQuery] string category, [FromQuery] bool includeCompleted = false, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(category)) return BadRequest(new { error = "category is required." });
        return Ok(await _workbench.GetItemsAsync(category, includeCompleted, ct));
    }

    /// <summary>Re-run classification on a stored item (writes a new version; never mutates the email).</summary>
    [HttpPost("reclassify/{mailItemId}")]
    public async Task<IActionResult> Reclassify(string mailItemId, CancellationToken ct)
    {
        try { return Ok(await _workbench.ReclassifyAsync(mailItemId, ct)); }
        catch (Exception ex) { return StatusCode(500, new { error = ex.Message }); }
    }

    /// <summary>Human classification correction (dev): writes a corrected version + logs AI-vs-human feedback locally.</summary>
    [HttpPost("amend/{mailItemId}")]
    public async Task<IActionResult> Amend(string mailItemId, [FromBody] AmendRequest req, CancellationToken ct)
    {
        if (req is null || string.IsNullOrWhiteSpace(req.CorrectedCategory))
            return BadRequest(new { error = "correctedCategory is required." });
        try { return Ok(await _workbench.AmendAsync(mailItemId, req.CorrectedCategory.Trim(), req.Reason, ct)); }
        catch (Exception ex) { return StatusCode(500, new { error = ex.Message }); }
    }

    /// <summary>Full reset (dev): delete all items+classifications, clear claim categories, remove stored files. Follow with backfill.</summary>
    [HttpPost("purge")]
    public async Task<IActionResult> Purge([FromQuery] string? upn, CancellationToken ct)
    {
        var u = string.IsNullOrWhiteSpace(upn) ? _bridge.GetStatuses().FirstOrDefault()?.WatchedUpn : upn;
        if (string.IsNullOrWhiteSpace(u)) return BadRequest(new { error = "No watched mailbox configured." });
        return Ok(await _workbench.PurgeAsync(u, ct));
    }

    /// <summary>
    /// Clean orphaned OneDrive archive folders (left by purge+backfill GUID churn or reclassify moves):
    /// deletes item folders no current item maps to + prunes empty dirs. dryRun=true reports only.
    /// </summary>
    [HttpPost("clean-archive")]
    public async Task<IActionResult> CleanArchive([FromQuery] bool dryRun = true, CancellationToken ct = default)
        => Ok(await _workbench.CleanArchiveAsync(dryRun, ct));

    /// <summary>Safety-net dedup sweep: remove MailItems sharing a WrapperGraphId (cross-proxy claim race).</summary>
    [HttpPost("dedup")]
    public async Task<IActionResult> Dedup(CancellationToken ct) => Ok(new { removed = await _workbench.DedupMailItemsAsync(ct) });

    /// <summary>Review the local classification-feedback log (dev analysis).</summary>
    [HttpGet("feedback")]
    public IActionResult Feedback() => Ok(_workbench.ReadFeedback());

    public sealed class AmendRequest
    {
        public string CorrectedCategory { get; set; } = "";
        public string? Reason { get; set; }
    }

    /// <summary>Mark a captured item complete/incomplete (configurable UI retention applies in the view).</summary>
    [HttpPost("complete/{mailItemId}")]
    public async Task<IActionResult> Complete(string mailItemId, [FromBody] CompleteRequest req, CancellationToken ct)
    {
        var ok = await _workbench.CompleteAsync(mailItemId, req?.Completed ?? true, req?.By, ct);
        return ok ? Ok(new { success = true }) : NotFound(new { error = "MailItem not found." });
    }

    public sealed class CompleteRequest
    {
        public bool Completed { get; set; } = true;
        public string? By { get; set; }
    }

    /// <summary>The effective taxonomy the classifier targets (static base + SP-confirmed leaves).</summary>
    [HttpGet("taxonomy")]
    public async Task<IActionResult> GetTaxonomy(CancellationToken ct) =>
        Ok((await _taxonomy.GetLeavesAsync(ct)).Select(l => new { l.Top, l.Sub, l.Path, l.Description }));

    /// <summary>
    /// Promote an "Other" suggestion to a real taxonomy leaf (dev): writes an SP hint so the
    /// classifier targets it on the next call (no deploy), then re-files the source item onto it.
    /// </summary>
    [HttpPost("confirm-leaf/{mailItemId}")]
    public async Task<IActionResult> ConfirmLeaf(string mailItemId, [FromBody] ConfirmLeafRequest req, CancellationToken ct)
    {
        if (req is null || string.IsNullOrWhiteSpace(req.CategoryPath))
            return BadRequest(new { error = "categoryPath is required." });
        try { return Ok(await _workbench.ConfirmLeafAsync(mailItemId, req.CategoryPath.Trim(), req.Hint, ct)); }
        catch (Exception ex) { return StatusCode(500, new { error = ex.Message }); }
    }

    public sealed class ConfirmLeafRequest
    {
        public string CategoryPath { get; set; } = "";
        public string? Hint { get; set; }
    }

    /// <summary>The operator-confirmed hints currently shaping the classifier prompt (dev inspection).</summary>
    [HttpGet("hints")]
    public async Task<IActionResult> Hints(CancellationToken ct) =>
        Ok(await _taxonomy.GetHintsAsync(ct));


    /// <summary>
    /// Classify a message without persisting. Provide either { upn?, id } to pull a cached mirror
    /// message, or raw { subject, from, body, attachmentNames } for ad-hoc testing.
    /// </summary>
    [HttpPost("preview")]
    public async Task<IActionResult> Preview([FromBody] PreviewRequest req, CancellationToken ct)
    {
        MailClassifyInput input;

        if (!string.IsNullOrWhiteSpace(req.Id))
        {
            var upn = string.IsNullOrWhiteSpace(req.Upn)
                ? _bridge.GetStatuses().FirstOrDefault()?.WatchedUpn
                : req.Upn;
            if (string.IsNullOrWhiteSpace(upn)) return BadRequest(new { error = "No watched mailbox configured." });

            var body = _bridge.GetMessage(upn, req.Id);
            if (body is null) return NotFound(new { error = "Message not in cache (may not have been polled yet)." });

            input = new MailClassifyInput
            {
                Subject         = body.Subject,
                FromAddress     = body.FromAddress,
                FromName        = body.FromName,
                ToLine          = body.ToLine,
                BodyText        = body.BodyText,
                AttachmentNames = body.Attachments.Select(a => a.Name).ToList(),
            };
        }
        else
        {
            input = new MailClassifyInput
            {
                Subject         = req.Subject ?? "",
                FromAddress     = req.From ?? "",
                FromName        = req.FromName ?? "",
                ToLine          = req.ToLine ?? "",
                BodyText        = req.Body ?? "",
                AttachmentNames = req.AttachmentNames ?? [],
            };
            if (string.IsNullOrWhiteSpace(input.Subject) && string.IsNullOrWhiteSpace(input.BodyText))
                return BadRequest(new { error = "Provide an id, or a subject/body to classify." });
        }

        var result = await _classifier.ClassifyAsync(input, ct);
        if (result is null) return StatusCode(503, new { error = "Both AI providers unavailable or failed." });

        return Ok(new
        {
            input = new { input.Subject, input.FromAddress, input.ToLine, attachments = input.AttachmentNames },
            classification = result,
        });
    }

    public sealed class PreviewRequest
    {
        public string? Upn { get; set; }
        public string? Id  { get; set; }
        public string? Subject  { get; set; }
        public string? From     { get; set; }
        public string? FromName { get; set; }
        public string? ToLine   { get; set; }
        public string? Body     { get; set; }
        public List<string>? AttachmentNames { get; set; }
    }
}
