using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxReview;

/// <summary>
/// Enhanced document extraction that captures formatting, tables, images,
/// and other document layers beyond what DocumentReader provides.
/// Used by both DocumentDiffer and TextConv.
/// </summary>
public class DocumentExtraction
{
    public string FileName { get; set; } = "";
    public DocumentMetadata Metadata { get; set; } = new();
    public List<RichParagraphInfo> Paragraphs { get; set; } = new();
    public List<CommentInfo> Comments { get; set; } = new();
    public List<TableInfo> Tables { get; set; } = new();
    public List<ImageInfo> Images { get; set; } = new();
    public List<HeaderFooterInfo> HeadersFooters { get; set; } = new();
}

public class TableInfo
{
    public int Index { get; set; }
    public int ParagraphIndex { get; set; }  // position in document flow
    public int Rows { get; set; }
    public int Columns { get; set; }
    public List<List<string>> Cells { get; set; } = new();
}

public class ImageInfo
{
    public string RelationshipId { get; set; } = "";
    public string FileName { get; set; } = "";
    public string ContentType { get; set; } = "";
    public string Sha256 { get; set; } = "";
    public long SizeBytes { get; set; }
    public int ParagraphIndex { get; set; }
}

public class HeaderFooterInfo
{
    public string Type { get; set; } = "";  // "header" or "footer"
    public string Scope { get; set; } = "";  // "default", "first", "even"
    public string Text { get; set; } = "";
}

public static class DocumentExtractor
{
    /// <summary>
    /// Extract all document layers from a .docx file.
    /// </summary>
    public static DocumentExtraction Extract(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"File not found: {path}");

        var result = new DocumentExtraction
        {
            FileName = Path.GetFileName(path)
        };

        using var doc = WordprocessingDocument.Open(path, false);
        var mainPart = doc.MainDocumentPart!;
        var body = mainPart.Document.Body!;

        // Metadata
        result.Metadata = ExtractMetadata(doc);

        // Paragraphs with rich formatting
        int paraIndex = 0;
        int tableIndex = 0;
        foreach (var element in body.ChildElements)
        {
            if (element is Paragraph para)
            {
                result.Paragraphs.Add(ExtractRichParagraph(para, paraIndex));
                paraIndex++;
            }
            else if (element is Table table)
            {
                result.Tables.Add(ExtractTable(table, tableIndex, paraIndex));
                tableIndex++;
                paraIndex++;  // tables count as a position in the flow
            }
        }

        // Comments
        result.Comments = ExtractComments(doc, body.Elements<Paragraph>().ToList());

        // Images
        result.Images = ExtractImages(mainPart, body);

        // Headers and Footers
        result.HeadersFooters = ExtractHeadersFooters(mainPart);

        return result;
    }

    private static RichParagraphInfo ExtractRichParagraph(Paragraph para, int index)
    {
        var info = new RichParagraphInfo
        {
            Index = index,
            Style = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value
        };

        var textParts = new List<string>();
        var runs = new List<RichRunInfo>();
        var trackedChanges = new List<TrackedChangeInfo>();

        foreach (var child in para.ChildElements)
        {
            if (child is Run run)
            {
                string text = GetFullRunText(run);
                if (text.Length > 0)
                {
                    textParts.Add(text);
                    runs.Add(ExtractRunFormatting(run, text));
                }
            }
            else if (child is DeletedRun del)
            {
                string delText = string.Join("", del.Descendants<DeletedText>().Select(t => t.Text));
                if (delText.Length > 0)
                {
                    trackedChanges.Add(new TrackedChangeInfo
                    {
                        Type = "delete",
                        Text = delText,
                        Author = del.Author?.Value ?? "",
                        Date = del.Date?.HasValue == true ? del.Date.Value.ToString("yyyy-MM-ddTHH:mm:ssZ") : null,
                        Id = del.Id?.Value ?? ""
                    });
                }
            }
            else if (child is InsertedRun ins)
            {
                string insText = GetInsertedRunText(ins);
                if (insText.Length > 0)
                {
                    textParts.Add(insText);
                    trackedChanges.Add(new TrackedChangeInfo
                    {
                        Type = "insert",
                        Text = insText,
                        Author = ins.Author?.Value ?? "",
                        Date = ins.Date?.HasValue == true ? ins.Date.Value.ToString("yyyy-MM-ddTHH:mm:ssZ") : null,
                        Id = ins.Id?.Value ?? ""
                    });
                }
            }
        }

        info.Text = string.Join("", textParts);
        info.Runs = runs;
        info.TrackedChanges = trackedChanges;
        return info;
    }

    private static RichRunInfo ExtractRunFormatting(Run run, string text)
    {
        var info = new RichRunInfo { Text = text };
        var rPr = run.RunProperties;
        if (rPr == null) return info;

        info.Bold = rPr.Bold != null && (rPr.Bold.Val == null || rPr.Bold.Val.Value);
        info.Italic = rPr.Italic != null && (rPr.Italic.Val == null || rPr.Italic.Val.Value);
        info.Underline = rPr.Underline != null && rPr.Underline.Val != null
            && rPr.Underline.Val.Value != UnderlineValues.None;
        info.Strikethrough = rPr.Strike != null && (rPr.Strike.Val == null || rPr.Strike.Val.Value);

        // Font
        var fonts = rPr.RunFonts;
        if (fonts != null)
            info.FontName = fonts.Ascii?.Value ?? fonts.HighAnsi?.Value ?? fonts.ComplexScript?.Value;

        // Size (half-points in OOXML)
        if (rPr.FontSize?.Val?.Value != null)
            info.FontSize = rPr.FontSize.Val.Value;

        // Color
        if (rPr.Color?.Val?.Value != null)
            info.Color = rPr.Color.Val.Value;

        // Highlight
        if (rPr.Highlight?.Val != null)
            info.Highlight = rPr.Highlight.Val.Value.ToString();

        return info;
    }

    private static TableInfo ExtractTable(Table table, int tableIndex, int paraIndex)
    {
        var rows = table.Elements<TableRow>().ToList();
        var info = new TableInfo
        {
            Index = tableIndex,
            ParagraphIndex = paraIndex,
            Rows = rows.Count
        };

        foreach (var row in rows)
        {
            var cells = row.Elements<TableCell>().ToList();
            if (info.Columns == 0)
                info.Columns = cells.Count;

            info.Cells.Add(cells.Select(c => c.InnerText.Trim()).ToList());
        }

        return info;
    }

    private static List<ImageInfo> ExtractImages(MainDocumentPart mainPart, Body body)
    {
        var images = new List<ImageInfo>();

        foreach (var imagePart in mainPart.ImageParts)
        {
            var relId = mainPart.GetIdOfPart(imagePart);
            var info = new ImageInfo
            {
                RelationshipId = relId,
                ContentType = imagePart.ContentType,
                FileName = Path.GetFileName(imagePart.Uri.ToString()),
            };

            using var stream = imagePart.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            var bytes = ms.ToArray();
            info.SizeBytes = bytes.Length;
            info.Sha256 = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

            images.Add(info);
        }

        return images;
    }

    private static List<HeaderFooterInfo> ExtractHeadersFooters(MainDocumentPart mainPart)
    {
        var result = new List<HeaderFooterInfo>();

        foreach (var headerPart in mainPart.HeaderParts)
        {
            string relId = mainPart.GetIdOfPart(headerPart);
            string scope = DetectHeaderFooterScope(mainPart, relId, "header");
            string text = headerPart.Header?.InnerText?.Trim() ?? "";
            if (!string.IsNullOrEmpty(text))
            {
                result.Add(new HeaderFooterInfo { Type = "header", Scope = scope, Text = text });
            }
        }

        foreach (var footerPart in mainPart.FooterParts)
        {
            string relId = mainPart.GetIdOfPart(footerPart);
            string scope = DetectHeaderFooterScope(mainPart, relId, "footer");
            string text = footerPart.Footer?.InnerText?.Trim() ?? "";
            if (!string.IsNullOrEmpty(text))
            {
                result.Add(new HeaderFooterInfo { Type = "footer", Scope = scope, Text = text });
            }
        }

        return result;
    }

    private static string DetectHeaderFooterScope(MainDocumentPart mainPart, string relId, string type)
    {
        var body = mainPart.Document.Body!;
        var sectionProps = body.Descendants<SectionProperties>();

        foreach (var sp in sectionProps)
        {
            if (type == "header")
            {
                foreach (var hRef in sp.Elements<HeaderReference>())
                {
                    if (hRef.Id?.Value == relId)
                        return hRef.Type?.Value.ToString() ?? "default";
                }
            }
            else
            {
                foreach (var fRef in sp.Elements<FooterReference>())
                {
                    if (fRef.Id?.Value == relId)
                        return fRef.Type?.Value.ToString() ?? "default";
                }
            }
        }

        return "default";
    }

    private static DocumentMetadata ExtractMetadata(WordprocessingDocument doc)
    {
        var meta = new DocumentMetadata();
        var props = doc.PackageProperties;
        if (props != null)
        {
            meta.Title = props.Title;
            meta.Author = props.Creator;
            meta.LastModifiedBy = props.LastModifiedBy;
            if (props.Created.HasValue)
                meta.Created = props.Created.Value.ToString("yyyy-MM-ddTHH:mm:ssZ");
            if (props.Modified.HasValue)
                meta.Modified = props.Modified.Value.ToString("yyyy-MM-ddTHH:mm:ssZ");
            if (props.Revision != null && int.TryParse(props.Revision, out int rev))
                meta.Revision = rev;
        }

        // Count words and paragraphs
        var body = doc.MainDocumentPart!.Document.Body!;
        var paragraphs = body.Elements<Paragraph>().ToList();
        meta.ParagraphCount = paragraphs.Count;
        meta.WordCount = paragraphs
            .Select(p => p.InnerText)
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .Sum(t => t.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length);

        return meta;
    }

    private static List<CommentInfo> ExtractComments(WordprocessingDocument doc, List<Paragraph> paragraphs)
    {
        var comments = new List<CommentInfo>();
        var commentsPart = doc.MainDocumentPart!.WordprocessingCommentsPart;
        if (commentsPart?.Comments == null)
            return comments;

        // Build comment range map
        var rangeMap = new Dictionary<string, (string anchor, int paraIdx)>();
        for (int pi = 0; pi < paragraphs.Count; pi++)
        {
            var starts = new Dictionary<string, int>();
            var children = paragraphs[pi].ChildElements.ToList();
            for (int ci = 0; ci < children.Count; ci++)
            {
                if (children[ci] is CommentRangeStart crs)
                    starts[crs.Id?.Value ?? ""] = ci;
                else if (children[ci] is CommentRangeEnd cre)
                {
                    string id = cre.Id?.Value ?? "";
                    if (starts.TryGetValue(id, out int startCi))
                    {
                        // Extract anchor text between markers
                        var anchorParts = new List<string>();
                        for (int k = startCi + 1; k < ci; k++)
                        {
                            if (children[k] is Run r)
                                anchorParts.Add(GetFullRunText(r));
                        }
                        rangeMap[id] = (string.Join("", anchorParts), pi);
                    }
                }
            }
        }

        foreach (var comment in commentsPart.Comments.Elements<Comment>())
        {
            string id = comment.Id?.Value ?? "";
            string commentText = string.Join("\n",
                comment.Elements<Paragraph>()
                    .Select(p => string.Join("",
                        p.Elements<Run>()
                            .Where(r => !r.Elements<AnnotationReferenceMark>().Any())
                            .Select(r => GetFullRunText(r))))
                    .Where(s => !string.IsNullOrEmpty(s)));

            rangeMap.TryGetValue(id, out var range);
            comments.Add(new CommentInfo
            {
                Id = id,
                Author = comment.Author?.Value ?? "",
                Date = comment.Date?.HasValue == true
                    ? comment.Date.Value.ToString("yyyy-MM-ddTHH:mm:ssZ") : null,
                AnchorText = range.anchor ?? "",
                Text = commentText,
                ParagraphIndex = range.paraIdx
            });
        }

        return comments;
    }

    private static string GetFullRunText(Run run)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var child in run.ChildElements)
        {
            if (child is Text t)
                sb.Append(t.Text);
            else if (child is Break)
                sb.Append('\n');
        }
        return sb.ToString();
    }

    private static string GetInsertedRunText(OpenXmlElement container)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var run in container.Elements<Run>())
        {
            foreach (var child in run.ChildElements)
            {
                if (child is Text t)
                    sb.Append(t.Text);
                else if (child is Break)
                    sb.Append('\n');
            }
        }
        return sb.ToString();
    }
}
