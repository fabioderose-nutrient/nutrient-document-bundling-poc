using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using BundlingPoc.Models;

namespace BundlingPoc.Pipeline;

/// <summary>
/// Generates an exhibit cover sheet from a Word template (see
/// samples/documents/Exhibit cover sheet template.docx) by merging the bracketed
/// placeholders, then converting the result to PDF. The manifest must name the
/// template in "coverSheetTemplate" when it has exhibits.
/// </summary>
public static class CoverSheet
{
    public static string Generate(Manifest manifest, Item item, string workDir, int order)
    {
        var docxPath = Path.Combine(workDir, $"{order:D3} cover {SafeName(item.ExhibitRef!)}.docx");

        if (manifest.CoverSheetTemplate is { } template && File.Exists(template))
            MergeTemplate(template, docxPath, manifest, item);
        else
            throw new FileNotFoundException("Cover sheet template not found", manifest.CoverSheetTemplate);

        return Converter.ToPdf(docxPath, workDir, order);
    }

    private static void MergeTemplate(string templatePath, string outputPath, Manifest manifest, Item item)
    {
        File.Copy(templatePath, outputPath, overwrite: true);
        using var zip = ZipFile.Open(outputPath, ZipArchiveMode.Update);
        var entry = zip.GetEntry("word/document.xml")
                    ?? throw new InvalidDataException("Template has no word/document.xml");

        string xml;
        using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
            xml = reader.ReadToEnd();

        // Word may split a placeholder like "[ Defendant name ]" across several
        // <w:r> runs, so first collapse any run boundaries that occur inside
        // [ ... ] brackets, then substitute the placeholders.
        // Placeholders: [ Defendant name ], [ Case reference ], [ Exhibit ref ],
        // [ Document name ], [ Witness name ].
        xml = CollapseRunsInsideBrackets(xml);
        xml = ReplacePlaceholder(xml, "Defendant name", manifest.CaseTitle.Split(" v ").Last());
        xml = ReplacePlaceholder(xml, "Case reference", manifest.CaseRef);
        xml = ReplacePlaceholder(xml, "Exhibit ref", item.ExhibitRef ?? "");
        xml = ReplacePlaceholder(xml, "Document name", item.DisplayTitle);
        xml = ReplacePlaceholder(xml, "Witness name", (item.Witness ?? "").ToUpperInvariant());

        entry.Delete();
        var newEntry = zip.CreateEntry("word/document.xml");
        using var writer = new StreamWriter(newEntry.Open(), Encoding.UTF8);
        writer.Write(xml);
    }

    private static string CollapseRunsInsideBrackets(string xml) =>
        Regex.Replace(xml, @"\[[^\[\]]*?\]", m =>
            Regex.Replace(m.Value, @"</w:t></w:r>\s*<w:r(?:\s[^>]*)?>(?:<w:rPr>.*?</w:rPr>)?<w:t(?:\s[^>]*)?>", ""),
            RegexOptions.Singleline);

    private static string ReplacePlaceholder(string xml, string name, string value) =>
        Regex.Replace(xml, @"\[\s*" + Regex.Escape(name) + @"\s*\]", Escape(value));

    private static string Escape(string s) =>
        s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");

    private static string SafeName(string s) =>
        string.Concat(s.Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '-' : c));
}
