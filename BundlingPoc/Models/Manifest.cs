using System.Text.Json;
using System.Text.Json.Serialization;

namespace BundlingPoc.Models;

/// <summary>
/// Describes one bundle to build: case metadata plus an ordered list of sections
/// and documents. In production this would come from your case management
/// system; for the PoC it is a JSON file next to the source documents.
/// </summary>
public sealed class Manifest
{
    [JsonPropertyName("bundleTitle")] public string BundleTitle { get; set; } = "";
    [JsonPropertyName("caseTitle")] public string CaseTitle { get; set; } = "";
    [JsonPropertyName("caseRef")] public string CaseRef { get; set; } = "";
    [JsonPropertyName("coverSheetTemplate")] public string? CoverSheetTemplate { get; set; }
    [JsonPropertyName("sections")] public List<Section> Sections { get; set; } = new();

    public static Manifest Load(string path)
    {
        var manifest = JsonSerializer.Deserialize<Manifest>(File.ReadAllText(path))
                       ?? throw new InvalidDataException($"Cannot parse manifest: {path}");
        var baseDir = Path.GetDirectoryName(Path.GetFullPath(path))!;
        foreach (var item in manifest.Sections.SelectMany(s => s.Items))
        {
            if (!Path.IsPathRooted(item.File))
                item.File = Path.Combine(baseDir, item.File);
            if (!File.Exists(item.File))
                throw new FileNotFoundException($"Manifest references a missing file", item.File);
        }
        if (manifest.CoverSheetTemplate is { } tpl && !Path.IsPathRooted(tpl))
            manifest.CoverSheetTemplate = Path.Combine(baseDir, tpl);
        return manifest;
    }
}

public sealed class Section
{
    [JsonPropertyName("title")] public string Title { get; set; } = "";
    [JsonPropertyName("items")] public List<Item> Items { get; set; } = new();
}

public sealed class Item
{
    [JsonPropertyName("file")] public string File { get; set; } = "";
    [JsonPropertyName("title")] public string? Title { get; set; }

    /// <summary>Exhibit reference (e.g. "AM/01"); triggers a cover sheet when set.</summary>
    [JsonPropertyName("exhibitRef")] public string? ExhibitRef { get; set; }

    /// <summary>Witness the exhibit is referred to in the statement of.</summary>
    [JsonPropertyName("witness")] public string? Witness { get; set; }

    /// <summary>For email items (.eml or .msg): also pull the email's own attachments into the bundle.</summary>
    [JsonPropertyName("expandAttachments")] public bool ExpandAttachments { get; set; } = true;

    public string DisplayTitle => Title ?? Path.GetFileNameWithoutExtension(File);
}
