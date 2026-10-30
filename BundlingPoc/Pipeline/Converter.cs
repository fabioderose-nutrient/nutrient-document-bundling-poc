using GdPicture14;
using GdDocumentFormat = GdPicture14.DocumentFormat;

namespace BundlingPoc.Pipeline;

/// <summary>Converts any supported input (PDF, DOCX/DOC, XLSX, images, MSG/EML…) to PDF.</summary>
public static class Converter
{
    public const float A4WidthPoints = 595f;
    public const float A4HeightPoints = 842f;

    private static readonly Dictionary<string, GdDocumentFormat> FormatsByExtension = new(StringComparer.OrdinalIgnoreCase)
    {
        [".pdf"] = GdDocumentFormat.DocumentFormatPDF,
        [".docx"] = GdDocumentFormat.DocumentFormatDOCX,
        [".doc"] = GdDocumentFormat.DocumentFormatDOC,
        [".rtf"] = GdDocumentFormat.DocumentFormatRTF,
        [".odt"] = GdDocumentFormat.DocumentFormatODT,
        [".xlsx"] = GdDocumentFormat.DocumentFormatXLSX,
        [".xls"] = GdDocumentFormat.DocumentFormatXLS,
        [".pptx"] = GdDocumentFormat.DocumentFormatPPTX,
        [".msg"] = GdDocumentFormat.DocumentFormatMSG,
        [".eml"] = GdDocumentFormat.DocumentFormatEML,
        [".txt"] = GdDocumentFormat.DocumentFormatTXT,
        [".html"] = GdDocumentFormat.DocumentFormatHTML,
        [".jpg"] = GdDocumentFormat.DocumentFormatJPEG,
        [".jpeg"] = GdDocumentFormat.DocumentFormatJPEG,
        [".png"] = GdDocumentFormat.DocumentFormatPNG,
        [".tif"] = GdDocumentFormat.DocumentFormatTIFF,
        [".tiff"] = GdDocumentFormat.DocumentFormatTIFF,
        [".bmp"] = GdDocumentFormat.DocumentFormatBMP,
    };

    /// <summary>Converts <paramref name="sourcePath"/> to PDF in <paramref name="workDir"/>. PDFs are copied as-is.</summary>
    public static string ToPdf(string sourcePath, string workDir, int order)
    {
        var ext = Path.GetExtension(sourcePath);
        var destPath = Path.Combine(workDir, $"{order:D3} {Path.GetFileNameWithoutExtension(sourcePath)}.pdf");

        if (ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            File.Copy(sourcePath, destPath, overwrite: true);
            return destPath;
        }

        if (!FormatsByExtension.TryGetValue(ext, out var format))
            throw new NotSupportedException($"Unsupported input format '{ext}': {sourcePath}");

        // Paginate HTML onto A4. Some Word-authored email bodies make the renderer
        // fail in that mode, so fall back to its single-page default and let the
        // A4 normalisation step in BundleBuilder rescale the result.
        if (!TryConvert(sourcePath, destPath, format, paginateHtml: true))
        {
            Console.WriteLine($"      NOTE: paginated conversion failed for " +
                              $"{Path.GetFileName(sourcePath)}, retrying single-page.");
            if (!TryConvert(sourcePath, destPath, format, paginateHtml: false))
                throw new InvalidOperationException($"GdPicture could not convert {sourcePath}");
        }
        return destPath;
    }

    private static bool TryConvert(string sourcePath, string destPath, GdDocumentFormat format, bool paginateHtml)
    {
        try
        {
            using var converter = new GdPictureDocumentConverter();
            if (paginateHtml)
            {
                // HTML defaults to a 1440pt-wide single page; bundles must stay A4.
                converter.HtmlPageWidth = A4WidthPoints;
                converter.HtmlPageHeight = A4HeightPoints;
                converter.HtmlPreferCSSPageSize = false;
                converter.HtmlPreferOnePage = false;
                converter.HtmlPageMarginTop = converter.HtmlPageMarginBottom = 36;
                converter.HtmlPageMarginLeft = converter.HtmlPageMarginRight = 36;
            }
            var status = converter.LoadFromFile(sourcePath, format);
            if (status == GdPictureStatus.OK)
                status = converter.SaveAsPDF(destPath, PdfConformance.PDF);
            if (status != GdPictureStatus.OK)
                Console.WriteLine($"      GdPicture status for {Path.GetFileName(sourcePath)}: {status}");
            return status == GdPictureStatus.OK;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"      Conversion of {Path.GetFileName(sourcePath)} failed: {ex.Message}");
            return false;
        }
    }

    public static void Check(GdPictureStatus status, string action)
    {
        if (status != GdPictureStatus.OK)
            throw new InvalidOperationException($"GdPicture failed to {action}: {status}");
    }
}
