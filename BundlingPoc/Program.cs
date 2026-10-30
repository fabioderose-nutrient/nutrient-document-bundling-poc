using BundlingPoc.Models;
using BundlingPoc.Pipeline;
using GdPicture14;

if (args.Length < 1)
{
    Console.WriteLine("Usage: BundlingPoc <manifest.json> [output.pdf]");
    return 1;
}

// GDPICTURE_LICENSE_KEY empty or unset = evaluation mode. GdPicture enters evaluation
// mode when it is registered with an empty key: without any RegisterKEY call,
// Office conversion and OCR fail with "The GdPicture feature ... is required".
var licenseKey = Environment.GetEnvironmentVariable("GDPICTURE_LICENSE_KEY")?.Trim() ?? "";
if (licenseKey.Length == 0)
    Console.WriteLine("No GDPICTURE_LICENSE_KEY set: running in evaluation mode.");
if (!new LicenseManager().RegisterKEY(licenseKey))
    Console.WriteLine("WARNING: GdPicture rejected the license key.");

// GdPicture does not scan the macOS/Linux font locations by default, which makes
// Office and HTML conversion emit non-embedded font references (unreadable text).
foreach (var folder in new[]
         {
             "/System/Library/Fonts", "/System/Library/Fonts/Supplemental", "/Library/Fonts",
             Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Library/Fonts"),
             "/usr/share/fonts", "/usr/local/share/fonts",
         })
{
    if (Directory.Exists(folder)) GdPictureDocumentUtilities.AddFontFolder(folder);
}

var manifest = Manifest.Load(args[0]);
var output = args.Length > 1
    ? Path.GetFullPath(args[1])
    : Path.Combine(Path.GetDirectoryName(Path.GetFullPath(args[0]))!, manifest.BundleTitle + ".pdf");

var workDir = Path.Combine(Path.GetTempPath(), "bundling-poc", Guid.NewGuid().ToString("N")[..8]);
Console.WriteLine($"Bundle:  {manifest.BundleTitle} ({manifest.CaseTitle}, {manifest.CaseRef})");
Console.WriteLine($"Work dir: {workDir}\n");

new BundleBuilder(manifest, workDir).Build(output);
return 0;
