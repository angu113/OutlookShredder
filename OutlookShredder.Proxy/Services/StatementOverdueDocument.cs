using System.Reflection;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Print-ready overdue-customer worklist PDF for the ShadowCat scheduled emails — mirrors the client
/// <c>OverdueReportDocument</c> (Customer / Terms / Due / Overdue / Notes). Returns the PDF bytes for an
/// email attachment (the client variant writes to a file path).
/// </summary>
public static class StatementOverdueDocument
{
    private const string Blue = "#1E5BAA", Red = "#C62828", Grey = "#7A7A7A",
                         Line = "#C9C9C9", LightGrey = "#F2F2F2", White = "#FFFFFF";

    private static readonly string ForgeVersion =
        typeof(StatementOverdueDocument).Assembly
            .GetCustomAttribute<System.Reflection.AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion ?? "unknown";

    public static byte[] Render(IReadOnlyList<OverdueRow> rows, string bucketLabel, DateTime asOf)
    {
        QuestPDF.Settings.License = LicenseType.Community;

        return Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.Letter);
                page.Margin(30);
                page.DefaultTextStyle(s => s.FontFamily("Arial").FontSize(11));

                page.Header().Element(h => ComposeHeader(h, bucketLabel, asOf, rows));
                page.Content().PaddingTop(6).Element(c => ComposeTable(c, rows));
                page.Footer().Row(row =>
                {
                    row.RelativeItem()
                        .Text(t => t.Span($"Copyright {DateTime.Today.Year} Silmaril Corp. Forge Version: {ForgeVersion}.")
                                    .FontColor(Grey).FontSize(7).Italic());
                    row.AutoItem().AlignRight().Text(t =>
                    {
                        t.Span("Page ").FontColor(Grey).FontSize(8);
                        t.CurrentPageNumber().FontColor(Grey).FontSize(8);
                        t.Span(" of ").FontColor(Grey).FontSize(8);
                        t.TotalPages().FontColor(Grey).FontSize(8);
                    });
                });
            });
        }).GeneratePdf();
    }

    private static void ComposeHeader(IContainer container, string bucketLabel, DateTime asOf, IReadOnlyList<OverdueRow> rows)
    {
        container.Column(col =>
        {
            col.Item().Text(t => t.Span($"ShadowCat: {bucketLabel} Customers Overdue as of {asOf:MMMM d, yyyy}")
                                  .Bold().FontSize(14).FontColor(Blue));
            col.Item().PaddingTop(1).Text(t => t.Span(
                $"{rows.Count} customer(s)  ·  Total overdue {rows.Sum(r => r.Overdue):C0}")
                .FontColor(Grey).FontSize(8));
            col.Item().PaddingTop(5).Height(2).Background(Blue);
        });
    }

    private static void ComposeTable(IContainer container, IReadOnlyList<OverdueRow> rows)
    {
        container.Table(table =>
        {
            table.ColumnsDefinition(cols =>
            {
                cols.RelativeColumn(5);   // Customer
                cols.RelativeColumn(2);   // Terms
                cols.RelativeColumn(2);   // Due
                cols.RelativeColumn(2);   // Overdue
                cols.RelativeColumn(4);   // Notes (blank)
            });

            table.Header(h =>
            {
                void HCell(string text, bool right = false)
                {
                    var c = h.Cell().Background(Blue).Padding(4);
                    (right ? c.AlignRight() : c).Text(t => t.Span(text).Bold().FontColor(White).FontSize(8));
                }
                HCell("Customer");
                HCell("Terms");
                HCell("Due", right: true);
                HCell("Overdue", right: true);
                HCell("Notes");
            });

            bool alt = false;
            foreach (var r in rows)
            {
                var bg = alt ? LightGrey : White;
                table.Cell().Background(bg).BorderBottom(0.5f).BorderColor(Line).Padding(3)
                    .Text(t => t.Span(r.Customer).SemiBold());
                table.Cell().Background(bg).BorderBottom(0.5f).BorderColor(Line).Padding(3)
                    .Text(t => t.Span(StatementOverdue.Label(r.Terms)).FontSize(9).FontColor(Grey));
                table.Cell().Background(bg).BorderBottom(0.5f).BorderColor(Line).Padding(3).AlignRight()
                    .Text(t => t.Span(r.Due.ToString("C0")));
                table.Cell().Background(bg).BorderBottom(0.5f).BorderColor(Line).Padding(3).AlignRight()
                    .Text(t => t.Span(r.Overdue.ToString("C0")).Bold().FontColor(Red));
                table.Cell().Background(White).BorderBottom(0.5f).BorderColor(Line).Padding(3);  // Notes
                alt = !alt;
            }

            table.Cell().ColumnSpan(2).BorderTop(1).BorderColor(Grey).Padding(3).AlignRight()
                .Text(t => t.Span("Totals").Bold());
            table.Cell().BorderTop(1).BorderColor(Grey).Padding(3).AlignRight()
                .Text(t => t.Span(rows.Sum(r => r.Due).ToString("C0")).Bold());
            table.Cell().BorderTop(1).BorderColor(Grey).Padding(3).AlignRight()
                .Text(t => t.Span(rows.Sum(r => r.Overdue).ToString("C0")).Bold().FontColor(Red));
            table.Cell().BorderTop(1).BorderColor(Grey).Padding(3);   // Notes (blank)
        });
    }
}
