using System.Text;
using System.Text.RegularExpressions;
using BundlingPoc.Models;
using MsgReader.Outlook;
using MimeMessage = MsgReader.Mime.Message;

namespace BundlingPoc.Pipeline;

/// <summary>
/// Expands an email (.eml or Outlook .msg) into the email itself plus its
/// attachments, so that each attachment becomes its own bundle entry right after
/// the parent email. The email is rendered to a clean HTML page (headers + body)
/// so that dates are always displayed in one normalised format regardless of how
/// the sender's locale wrote them; GdPicture then converts that HTML to PDF.
/// </summary>
public static class EmailIngestor
{
    public const string DateFormat = "dd MMM yyyy HH:mm";

    private sealed record Attachment(string FileName, byte[] Data);

    private sealed record Email(
        string FromName, string FromAddress, IReadOnlyList<string> To, DateTimeOffset? Date,
        string? Subject, string? BodyHtml, string? BodyText, IReadOnlyList<Attachment> Attachments);

    public static bool IsEmail(string path) =>
        Path.GetExtension(path).ToLowerInvariant() is ".eml" or ".msg";

    public static List<Item> Expand(Item emailItem, string workDir)
    {
        var result = new List<Item>();
        var email = Path.GetExtension(emailItem.File).Equals(".eml", StringComparison.OrdinalIgnoreCase)
            ? ReadEml(emailItem.File)
            : ReadMsg(emailItem.File);

        var datePrefix = email.Date.HasValue
            ? email.Date.Value.ToString(DateFormat, System.Globalization.CultureInfo.InvariantCulture)
            : "Undated";
        var subject = string.IsNullOrWhiteSpace(email.Subject) ? emailItem.DisplayTitle : email.Subject;

        var extractDir = Path.Combine(workDir, "email-" + Path.GetFileNameWithoutExtension(emailItem.File));
        Directory.CreateDirectory(extractDir);

        result.Add(new Item
        {
            File = RenderToHtml(email, datePrefix, extractDir),
            Title = emailItem.Title ?? $"Email {datePrefix} — {subject}",
            ExpandAttachments = false,
        });

        if (!emailItem.ExpandAttachments) return result;

        foreach (var attachment in email.Attachments)
        {
            var path = Path.Combine(extractDir, attachment.FileName);
            File.WriteAllBytes(path, attachment.Data);
            result.Add(new Item
            {
                File = path,
                Title = $"Attachment to email {datePrefix} — {Path.GetFileNameWithoutExtension(attachment.FileName)}",
                ExpandAttachments = false,
            });
        }
        return result;
    }

    /// <summary>Reads an Outlook .msg file. Inline images (signature logos etc.) are skipped.</summary>
    private static Email ReadMsg(string path)
    {
        using var msg = new Storage.Message(path);
        var attachments = msg.Attachments
            .OfType<Storage.Attachment>()
            .Where(a => !a.IsInline)
            .Select(a => new Attachment(SafeFileName(a.FileName), a.Data))
            .ToList();
        return new Email(
            msg.Sender?.DisplayName ?? "", msg.Sender?.Email ?? "",
            msg.Recipients.Select(r => r.Email ?? r.DisplayName).ToList(),
            msg.SentOn ?? msg.ReceivedOn, msg.Subject, msg.BodyHtml, msg.BodyText, attachments);
    }

    /// <summary>Reads a MIME .eml file. Inline parts (images referenced by content-id) are skipped.</summary>
    private static Email ReadEml(string path)
    {
        var eml = MimeMessage.Load(new FileInfo(path));
        var headers = eml.Headers;
        var attachments = (eml.Attachments ?? [])
            .Where(a => !a.IsInline)
            .Select(a => new Attachment(SafeFileName(a.FileName), a.Body))
            .ToList();
        return new Email(
            headers.From?.DisplayName ?? "", headers.From?.Address ?? "",
            headers.To.Select(r => string.IsNullOrEmpty(r.Address) ? r.DisplayName : r.Address).ToList(),
            headers.DateSent == default ? null : headers.DateSent,
            headers.Subject, eml.HtmlBody?.GetBodyAsText(), eml.TextBody?.GetBodyAsText(), attachments);
    }

    private static string SafeFileName(string? name) =>
        string.IsNullOrWhiteSpace(name)
            ? "attachment.bin"
            : string.Concat(Path.GetFileName(name).Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '-' : c));

    /// <summary>Renders the message as a self-contained HTML page with a normalised header block.</summary>
    private static string RenderToHtml(Email email, string datePrefix, string extractDir)
    {
        var body = ExtractBodyFragment(email.BodyHtml);
        if (string.IsNullOrWhiteSpace(body))
            body = "<pre style=\"font-family:inherit;white-space:pre-wrap\">" +
                   Html(email.BodyText ?? "") + "</pre>";

        var attachmentNames = email.Attachments.Select(a => a.FileName).ToList();

        var html = new StringBuilder()
            .Append("<html><head><meta charset=\"utf-8\">")
            // Constrain to A4 content width, otherwise wide quoted mail tables make
            // the converter emit one very wide landscape page.
            // No @page rule here: the page box is imposed by the converter, and a
            // competing CSS declaration makes the renderer fail on Word-authored bodies.
            .Append("<style>table,img{max-width:100%!important}" +
                    "*{word-wrap:break-word;overflow-wrap:break-word}</style>")
            .Append("</head>")
            .Append("<body style=\"font-family:Helvetica,Arial,sans-serif;font-size:11pt\">")
            .Append("<div style=\"border:1px solid #999;background:#f4f4f4;padding:10px 14px;margin-bottom:16px\">")
            .Append($"<div><b>From:</b> {Html(email.FromName)} &lt;{Html(email.FromAddress)}&gt;</div>")
            .Append($"<div><b>To:</b> {Html(string.Join("; ", email.To))}</div>")
            .Append($"<div><b>Date:</b> {Html(datePrefix)}</div>")
            .Append($"<div><b>Subject:</b> {Html(email.Subject ?? "")}</div>")
            .Append(attachmentNames.Count > 0
                ? $"<div><b>Attachments:</b> {Html(string.Join("; ", attachmentNames))}</div>"
                : "")
            .Append("</div>")
            .Append(body)
            .Append("</body></html>")
            .ToString();

        var htmlPath = Path.Combine(extractDir, "email.html");
        File.WriteAllText(htmlPath, html, Encoding.UTF8);
        return htmlPath;
    }

    /// <summary>
    /// Mail clients store a complete HTML document (with its own head, Word styles and
    /// conditional comments) as the message body. Nesting that inside our wrapper
    /// produces malformed markup the renderer rejects, so keep only the body
    /// content and drop references the renderer cannot resolve.
    /// </summary>
    private static string ExtractBodyFragment(string? bodyHtml)
    {
        if (string.IsNullOrWhiteSpace(bodyHtml)) return "";
        var html = bodyHtml;

        var bodyOpen = Regex.Match(html, @"<body\b[^>]*>", RegexOptions.IgnoreCase);
        if (bodyOpen.Success)
        {
            var start = bodyOpen.Index + bodyOpen.Length;
            var bodyClose = Regex.Match(html[start..], @"</body\s*>", RegexOptions.IgnoreCase);
            html = bodyClose.Success ? html[start..(start + bodyClose.Index)] : html[start..];
        }
        else
        {
            html = Regex.Replace(html, @"<head\b.*?</head\s*>", "", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            html = Regex.Replace(html, @"</?html\b[^>]*>", "", RegexOptions.IgnoreCase);
        }

        // Inline images referenced by content-id are not resolvable from a plain file.
        html = Regex.Replace(html, @"<img\b[^>]*src\s*=\s*[""']?cid:[^>]*>", "", RegexOptions.IgnoreCase);
        // Word/Outlook conditional comments confuse the renderer.
        html = Regex.Replace(html, @"<!--\[if.*?<!\[endif\]-->", "", RegexOptions.IgnoreCase | RegexOptions.Singleline);
        return html;
    }

    private static string Html(string s) =>
        s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
}
