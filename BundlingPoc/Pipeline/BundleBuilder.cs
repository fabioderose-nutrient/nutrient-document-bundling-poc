using System.Text.RegularExpressions;
using BundlingPoc.Models;
using GdPicture14;

namespace BundlingPoc.Pipeline;

/// <summary>
/// Orchestrates the whole bundling pipeline:
/// ingest (email expansion) → convert to PDF → cover sheets → merge →
/// selective OCR → table of contents → bookmarks → page numbering.
/// </summary>
public sealed class BundleBuilder(Manifest manifest, string workDir)
{
    private const float PageWidth = 595f;   // A4 portrait, points
    private const float PageHeight = 842f;
    private const int TocEntriesPerPage = 26;
    private const int MinCharsForTextPage = 20;

    // In evaluation mode GdPicture stamps "For Evaluation Purposes Only" on every page.
    // That stamp is text, so it is ignored when deciding whether a page needs OCR.
    private static readonly Regex IgnoredText =
        new(@"\b(?:For|Evaluation|Purposes|Only)\b|\s+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private sealed record BuiltEntry(string SectionTitle, string Title, List<string> Pdfs)
    {
        public int StartBodyPage { get; set; }
        public int PageCount { get; set; }
    }

    public void Build(string outputPath)
    {
        Directory.CreateDirectory(workDir);

        Console.WriteLine("[1/6] Ingesting documents (expanding emails)…");
        var entries = Flatten();

        Console.WriteLine("[2/6] Converting to PDF and generating cover sheets…");
        int order = 0, bodyPage = 1;
        foreach (var entry in entries)
        {
            for (var i = 0; i < entry.Pdfs.Count; i++)
                entry.Pdfs[i] = entry.Pdfs[i].EndsWith(".pdf|cover", StringComparison.Ordinal)
                    ? entry.Pdfs[i][..^6]
                    : Converter.ToPdf(entry.Pdfs[i], workDir, ++order);
            entry.StartBodyPage = bodyPage;
            entry.PageCount = entry.Pdfs.Sum(PdfPageCount);
            bodyPage += entry.PageCount;
            Console.WriteLine($"      {entry.Title} → body p.{entry.StartBodyPage} ({entry.PageCount} page(s))");
        }
        var totalBodyPages = bodyPage - 1;
        var tocPageCount = TocPageCount(entries);

        Console.WriteLine("[3/6] Merging…");
        var bodyPath = Path.Combine(workDir, "_body.pdf");
        Converter.Check(new GdPicturePDF().MergeDocuments(
            entries.SelectMany(e => e.Pdfs).ToArray(), bodyPath), "merge documents");

        Console.WriteLine("[4/6] OCR on pages without a text layer…");
        RunSelectiveOcr(bodyPath);

        Console.WriteLine("[5/6] Building table of contents…");
        var tocPath = Path.Combine(workDir, "_toc.pdf");
        BuildTocPdf(entries, tocPageCount, tocPath);

        Console.WriteLine("[6/6] Final assembly: bookmarks, links, page numbers…");
        var mergedPath = Path.Combine(workDir, "_merged.pdf");
        using (var merged = new GdPicturePDF())
        {
            Converter.Check(merged.MergeDocuments(new[] { tocPath, bodyPath }, mergedPath), "merge TOC + body");
        }
        using (var final = new GdPicturePDF())
        {
            // Load and save must target different files: saving onto the file the
            // document was opened from truncates the streams still being read.
            Converter.Check(final.LoadFromFile(mergedPath, false), $"reopen {mergedPath}");
            final.SetOrigin(PdfOrigin.PdfOriginTopLeft);
            final.SetMeasurementUnit(PdfMeasurementUnit.PdfMeasurementUnitPoint);

            NormaliseToA4(final);
            AddBookmarks(final, entries, tocPageCount);
            AddTocLinks(final, entries, tocPageCount);
            AddPageLabels(final, tocPageCount);
            StampPageNumbers(final, tocPageCount, totalBodyPages);

            Converter.Check(final.SaveToFile(outputPath), $"save {outputPath}");
            final.CloseDocument();
        }

        Console.WriteLine($"\nDone: {outputPath}");
        Console.WriteLine($"  {tocPageCount} TOC page(s) (roman numerals) + {totalBodyPages} body page(s), " +
                          $"{entries.Count} documents, {entries.Select(e => e.SectionTitle).Distinct().Count()} sections.");
    }

    /// <summary>Expands the manifest into the final linear document order.</summary>
    private List<BuiltEntry> Flatten()
    {
        var entries = new List<BuiltEntry>();
        var coverOrder = 500; // cover sheet temp files, numbered apart
        foreach (var section in manifest.Sections)
        foreach (var item in section.Items)
        {
            if (EmailIngestor.IsEmail(item.File))
            {
                foreach (var expanded in EmailIngestor.Expand(item, workDir))
                    entries.Add(new BuiltEntry(section.Title, expanded.DisplayTitle, new List<string> { expanded.File }));
                continue;
            }

            var pdfs = new List<string>();
            if (item.ExhibitRef is not null)
                pdfs.Add(CoverSheet.Generate(manifest, item, workDir, ++coverOrder) + "|cover");
            pdfs.Add(item.File);
            var title = item.ExhibitRef is null ? item.DisplayTitle : $"Exhibit {item.ExhibitRef} — {item.DisplayTitle}";
            entries.Add(new BuiltEntry(section.Title, title, pdfs));
        }
        return entries;
    }

    private static int PdfPageCount(string path)
    {
        using var pdf = new GdPicturePDF();
        Converter.Check(pdf.LoadFromFile(path, false), $"open {path}");
        var count = pdf.GetPageCount();
        pdf.CloseDocument();
        return count;
    }

    private void RunSelectiveOcr(string bodyPath)
    {
        var dictionaryPath = FindOcrResources();
        var ocrPath = Path.Combine(workDir, "_body-ocr.pdf");
        using var pdf = new GdPicturePDF();
        Converter.Check(pdf.LoadFromFile(bodyPath, false), $"open {bodyPath}");

        var pagesToOcr = new List<int>();
        for (var page = 1; page <= pdf.GetPageCount(); page++)
        {
            pdf.SelectPage(page);
            if (IgnoredText.Replace(pdf.GetPageText() ?? "", "").Length < MinCharsForTextPage)
                pagesToOcr.Add(page);
        }

        if (pagesToOcr.Count == 0)
        {
            Console.WriteLine("      All pages already carry a text layer, skipping OCR.");
        }
        else if (dictionaryPath is null)
        {
            Console.WriteLine($"      WARNING: {pagesToOcr.Count} page(s) need OCR but eng.traineddata was not found — skipped.");
        }
        else
        {
            Console.WriteLine($"      OCR on {pagesToOcr.Count}/{pdf.GetPageCount()} page(s): {string.Join(",", pagesToOcr)}");
            // Page by page, single-threaded and synchronous: a page whose embedded
            // image trips the native decoder is skipped with a warning instead of
            // taking the whole process down from an SDK worker thread.
            foreach (var page in pagesToOcr)
            {
                try
                {
                    var status = pdf.OcrPages(page.ToString(), 1, "eng", dictionaryPath, "", 200f,
                        OCRMode.FavorAccuracy, TimeoutMillisec: 120_000, Sync: true);
                    if (status != GdPictureStatus.OK)
                        Console.WriteLine($"      WARNING: OCR of page {page} returned {status} — skipped.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      WARNING: OCR of page {page} failed ({ex.GetType().Name}) — skipped.");
                }
            }
            Converter.Check(pdf.SaveToFile(ocrPath), "save OCR result");
            pdf.CloseDocument();
            File.Move(ocrPath, bodyPath, overwrite: true);
            return;
        }
        pdf.CloseDocument();
    }

    private static string? FindOcrResources()
    {
        string?[] candidates =
        {
            Environment.GetEnvironmentVariable("GDPICTURE_OCR_PATH"),
            AppContext.BaseDirectory,
            Path.Combine(AppContext.BaseDirectory, "runtimes", "any", "native"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "OCR"),
        };
        return candidates.FirstOrDefault(d => d is not null && File.Exists(Path.Combine(d, "eng.traineddata")));
    }

    private int TocPageCount(List<BuiltEntry> entries)
    {
        var rows = entries.Count + entries.Select(e => e.SectionTitle).Distinct().Count();
        return Math.Max(1, (int)Math.Ceiling(rows / (double)TocEntriesPerPage));
    }

    private void BuildTocPdf(List<BuiltEntry> entries, int tocPageCount, string tocPath)
    {
        using var toc = new GdPicturePDF();
        Converter.Check(toc.NewPDF(), "create TOC PDF");
        toc.SetOrigin(PdfOrigin.PdfOriginTopLeft);
        toc.SetMeasurementUnit(PdfMeasurementUnit.PdfMeasurementUnitPoint);

        var title = toc.AddStandardFont(PdfStandardFont.PdfStandardFontTimesBold);
        var heading = toc.AddStandardFont(PdfStandardFont.PdfStandardFontHelveticaBold);
        var body = toc.AddStandardFont(PdfStandardFont.PdfStandardFontHelvetica);

        float y = 0;
        var rowsOnPage = int.MaxValue;
        string? currentSection = null;

        void NewTocPage(bool first)
        {
            Converter.Check(toc.NewPage(PageWidth, PageHeight), "new TOC page");
            toc.SelectPage(toc.GetPageCount());
            if (first)
            {
                toc.SetTextSize(16);
                toc.DrawTextBox(title, 60, 50, PageWidth - 60, 74, TextAlignment.TextAlignmentCenter, TextAlignment.TextAlignmentNear, manifest.CaseTitle);
                toc.SetTextSize(12);
                toc.DrawTextBox(body, 60, 78, PageWidth - 60, 94, TextAlignment.TextAlignmentCenter, TextAlignment.TextAlignmentNear, $"Case ref: {manifest.CaseRef}");
                toc.SetTextSize(14);
                toc.DrawTextBox(heading, 60, 106, PageWidth - 60, 126, TextAlignment.TextAlignmentCenter, TextAlignment.TextAlignmentNear, manifest.BundleTitle.ToUpperInvariant() + " — INDEX");
                y = 150;
            }
            else
            {
                y = 60;
            }
            rowsOnPage = 0;
        }

        NewTocPage(first: true);
        foreach (var entry in entries)
        {
            if (rowsOnPage >= TocEntriesPerPage) NewTocPage(first: false);
            if (entry.SectionTitle != currentSection)
            {
                currentSection = entry.SectionTitle;
                toc.SetTextSize(12);
                toc.DrawTextBox(heading, 60, y + 6, PageWidth - 60, y + 24, TextAlignment.TextAlignmentNear, TextAlignment.TextAlignmentNear, currentSection);
                y += 28; rowsOnPage++;
                if (rowsOnPage >= TocEntriesPerPage) NewTocPage(first: false);
            }
            toc.SetTextSize(10);
            toc.DrawTextBox(body, 76, y, PageWidth - 110, y + 16, TextAlignment.TextAlignmentNear, TextAlignment.TextAlignmentNear, entry.Title);
            toc.DrawTextBox(body, PageWidth - 110, y, PageWidth - 60, y + 16, TextAlignment.TextAlignmentFar, TextAlignment.TextAlignmentNear, entry.StartBodyPage.ToString());
            y += 18; rowsOnPage++;
        }

        while (toc.GetPageCount() < tocPageCount)
            Converter.Check(toc.NewPage(PageWidth, PageHeight), "pad TOC page");

        Converter.Check(toc.SaveToFile(tocPath), "save TOC");
        toc.CloseDocument();
    }

    /// <summary>
    /// Source documents arrive in mixed page sizes (US Letter PDFs, image pages,
    /// rendered emails). A bundle has to print consistently, so every page
    /// is scaled onto A4 portrait, keeping its aspect ratio.
    /// </summary>
    private static void NormaliseToA4(GdPicturePDF pdf)
    {
        // ResizePage interprets the new height in the current coordinate system:
        // under a top-left origin it writes an inverted (negative-height) media box.
        pdf.SetOrigin(PdfOrigin.PdfOriginBottomLeft);
        var resized = 0;
        for (var page = 1; page <= pdf.GetPageCount(); page++)
        {
            pdf.SelectPage(page);
            var w = pdf.GetPageWidth();
            var h = pdf.GetPageHeight();
            if (Math.Abs(w - PageWidth) < 1 && Math.Abs(h - PageHeight) < 1) continue;

            var scale = Math.Min(PageWidth / w, PageHeight / h);
            if (pdf.ScalePage(scale, scale) != GdPictureStatus.OK) continue;
            if (pdf.ResizePage(PageWidth, PageHeight) != GdPictureStatus.OK) continue;
            // Rewrite the media box explicitly: the resize leaves it with its
            // vertical bounds inverted ([0 842 595 0]), which some viewers and
            // downstream tools report as a negative page height.
            pdf.SetPageBox(PdfPageBox.PdfPageBoxMediaBox, 0, 0, PageWidth, PageHeight);
            resized++;
        }
        pdf.SetOrigin(PdfOrigin.PdfOriginTopLeft);
        if (resized > 0) Console.WriteLine($"      Normalised {resized} page(s) to A4.");
    }

    private void AddBookmarks(GdPicturePDF pdf, List<BuiltEntry> entries, int tocPageCount)
    {
        // Drop whatever outline the source documents carried in through the merge,
        // so the bundle only exposes the index we build here.
        pdf.RemoveBookmarks();

        var root = pdf.GetBookmarkRootID();
        var indexBookmark = pdf.NewBookmark(root, "Index");
        pdf.SetBookmarkAction(indexBookmark, GoTo(pdf, 1));

        var sectionBookmarks = new Dictionary<string, int>();
        foreach (var entry in entries)
        {
            if (!sectionBookmarks.TryGetValue(entry.SectionTitle, out var sectionId))
            {
                sectionId = pdf.NewBookmark(root, entry.SectionTitle);
                pdf.SetBookmarkAction(sectionId, GoTo(pdf, tocPageCount + entry.StartBodyPage));
                sectionBookmarks[entry.SectionTitle] = sectionId;
            }
            var id = pdf.NewBookmark(sectionId, entry.Title);
            pdf.SetBookmarkAction(id, GoTo(pdf, tocPageCount + entry.StartBodyPage));
        }
    }

    private static int GoTo(GdPicturePDF pdf, int page) =>
        pdf.NewActionGoTo(PdfDestinationType.DestinationTypeFit, page, 0, 0, 0, 0, 0);

    /// <summary>Re-draws the TOC hit areas as clickable links (row geometry mirrors BuildTocPdf).</summary>
    private void AddTocLinks(GdPicturePDF pdf, List<BuiltEntry> entries, int tocPageCount)
    {
        var tocPage = 1;
        float y = 150;
        var rowsOnPage = 0;
        string? currentSection = null;

        void NextPage() { tocPage++; y = 60; rowsOnPage = 0; }

        pdf.SelectPage(tocPage);
        foreach (var entry in entries)
        {
            if (rowsOnPage >= TocEntriesPerPage) { NextPage(); pdf.SelectPage(tocPage); }
            if (entry.SectionTitle != currentSection)
            {
                currentSection = entry.SectionTitle;
                y += 28; rowsOnPage++;
                if (rowsOnPage >= TocEntriesPerPage) { NextPage(); pdf.SelectPage(tocPage); }
            }
            pdf.AddLinkToPageAnnotation(60, y, PageWidth - 120, 16, tocPageCount + entry.StartBodyPage, 0, 0, false, 0, 0, 200);
            y += 18; rowsOnPage++;
        }
    }

    private static void AddPageLabels(GdPicturePDF pdf, int tocPageCount)
    {
        pdf.DeletePageLabels();
        Converter.Check(pdf.AddPageLabelsRange(1, PdfPageLabelStyle.PdfPageLabelStyleLowercaseRomanNumerals, "", 1), "TOC page labels");
        Converter.Check(pdf.AddPageLabelsRange(tocPageCount + 1, PdfPageLabelStyle.PdfPageLabelStyleDecimalArabicNumerals, "", 1), "body page labels");
    }

    private static void StampPageNumbers(GdPicturePDF pdf, int tocPageCount, int totalBodyPages)
    {
        var font = pdf.AddStandardFont(PdfStandardFont.PdfStandardFontHelvetica);
        pdf.SetTextSize(9);
        for (var bodyPage = 1; bodyPage <= totalBodyPages; bodyPage++)
        {
            pdf.SelectPage(tocPageCount + bodyPage);
            var w = pdf.GetPageWidth();
            var h = pdf.GetPageHeight();
            pdf.DrawTextBox(font, 0, h - 30, w, h - 12,
                TextAlignment.TextAlignmentCenter, TextAlignment.TextAlignmentNear,
                $"{bodyPage} of {totalBodyPages}");
        }
    }
}
