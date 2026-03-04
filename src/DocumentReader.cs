using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxReview;

/// <summary>
/// Reads tracked changes, comments, and document structure from .docx files.
/// Produces a ReadResult that can be serialized to JSON or printed as human-readable text.
/// </summary>
public class DocumentReader
{
    /// <summary>
    /// Read and extract all review state from a document.
    /// </summary>
    public ReadResult Read(string inputPath)
    {
        if (!File.Exists(inputPath))
            throw new FileNotFoundException($"Input file not found: {inputPath}");

        var result = new ReadResult
        {
            File = Path.GetFileName(inputPath)
        };

        using var doc = WordprocessingDocument.Open(inputPath, false);
        var body = doc.MainDocumentPart!.Document.Body!;
        var paragraphs = body.Elements<Paragraph>().ToList();

        // Build comment range map: commentId → (startParaIndex, startCharOffset, endParaIndex, endCharOffset)
        var commentRanges = BuildCommentRangeMap(paragraphs);

        // Extract paragraphs with tracked changes
        int totalWordCount = 0;
        for (int i = 0; i < paragraphs.Count; i++)
        {
            var para = paragraphs[i];
            var paraInfo = ExtractParagraph(para, i);
            result.Paragraphs.Add(paraInfo);

            // Count words from visible text
            if (!string.IsNullOrWhiteSpace(paraInfo.Text))
            {
                totalWordCount += paraInfo.Text.Split(
                    (char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
            }
        }

        // Extract comments
        result.Comments = ExtractComments(doc, paragraphs, commentRanges);

        // Extract metadata
        result.Metadata = ExtractMetadata(doc, totalWordCount, paragraphs.Count);

        // Build summary
        result.Summary = BuildSummary(result);

        return result;
    }

    /// <summary>
    /// Extract paragraph text, style, and tracked changes.
    /// The "text" field is the CURRENT visible text (includes insertions, excludes deletions).
    /// </summary>
    private ParagraphInfo ExtractParagraph(Paragraph para, int index)
    {
        var info = new ParagraphInfo { Index = index };

        // Get paragraph style
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        info.Style = styleId;

        var textParts = new List<string>();
        var trackedChanges = new List<TrackedChangeInfo>();

        foreach (var child in para.ChildElements)
        {
            if (child is Run run)
            {
                // Regular run — part of the visible text
                string runText = GetRunText(run);
                if (runText.Length > 0)
                    textParts.Add(runText);
            }
            else if (child is DeletedRun del)
            {
                // w:del — tracked deletion (NOT visible in current text)
                string delText = GetDeletedRunText(del);
                if (delText.Length > 0)
                {
                    trackedChanges.Add(new TrackedChangeInfo
                    {
                        Type = "delete",
                        Text = delText,
                        Author = del.Author?.Value ?? "",
                        Date = FormatDate(del.Date),
                        Id = del.Id?.Value ?? ""
                    });
                }
            }
            else if (child is InsertedRun ins)
            {
                // w:ins — tracked insertion (IS visible in current text)
                string insText = GetInsertedRunText(ins);
                if (insText.Length > 0)
                {
                    textParts.Add(insText);
                    trackedChanges.Add(new TrackedChangeInfo
                    {
                        Type = "insert",
                        Text = insText,
                        Author = ins.Author?.Value ?? "",
                        Date = FormatDate(ins.Date),
                        Id = ins.Id?.Value ?? ""
                    });
                }
            }
            else if (child is DeletedMathControl)
            {
                // Skip deleted math controls
            }
            else if (child is MoveFromRun moveFrom)
            {
                // Move-from is like a deletion — not visible
                string moveText = GetInsertedRunText(moveFrom);
                if (moveText.Length > 0)
                {
                    trackedChanges.Add(new TrackedChangeInfo
                    {
                        Type = "delete",
                        Text = moveText,
                        Author = moveFrom.Author?.Value ?? "",
                        Date = FormatDate(moveFrom.Date),
                        Id = moveFrom.Id?.Value ?? ""
                    });
                }
            }
            else if (child is MoveToRun moveTo)
            {
                // Move-to is like an insertion — visible
                string moveText = GetInsertedRunText(moveTo);
                if (moveText.Length > 0)
                {
                    textParts.Add(moveText);
                    trackedChanges.Add(new TrackedChangeInfo
                    {
                        Type = "insert",
                        Text = moveText,
                        Author = moveTo.Author?.Value ?? "",
                        Date = FormatDate(moveTo.Date),
                        Id = moveTo.Id?.Value ?? ""
                    });
                }
            }
        }

        info.Text = string.Join("", textParts);
        info.TrackedChanges = trackedChanges;

        return info;
    }

    /// <summary>
    /// Format a DateTimeValue to ISO 8601 string, or null if not present.
    /// </summary>
    private static string? FormatDate(DateTimeValue? dtv)
    {
        if (dtv == null || !dtv.HasValue) return null;
        return dtv.Value.ToString("yyyy-MM-ddTHH:mm:ssZ");
    }

    /// <summary>
    /// Get text from a regular Run element, emitting \n for &lt;w:br/&gt; line breaks.
    /// </summary>
    private static string GetRunText(Run run)
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

    /// <summary>
    /// Get text from a DeletedRun (w:del) element.
    /// The inner runs contain DeletedText elements instead of Text.
    /// </summary>
    private static string GetDeletedRunText(DeletedRun del)
    {
        var texts = new List<string>();
        foreach (var run in del.Elements<Run>())
        {
            foreach (var dt in run.Elements<DeletedText>())
            {
                texts.Add(dt.Text);
            }
        }
        return string.Join("", texts);
    }

    /// <summary>
    /// Get text from an InsertedRun (w:ins) or MoveFromRun/MoveToRun element.
    /// The inner runs contain regular Text elements.
    /// </summary>
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

    /// <summary>
    /// Build a map of comment ID → anchor text and paragraph index
    /// by scanning for CommentRangeStart/CommentRangeEnd markers.
    /// </summary>
    private Dictionary<string, (string anchorText, int paragraphIndex)> BuildCommentRangeMap(
        List<Paragraph> paragraphs)
    {
        var result = new Dictionary<string, (string anchorText, int paragraphIndex)>();

        // Track where each comment range starts
        var rangeStarts = new Dictionary<string, (int paraIndex, int childIndex)>();

        for (int pi = 0; pi < paragraphs.Count; pi++)
        {
            var para = paragraphs[pi];
            var children = para.ChildElements.ToList();

            for (int ci = 0; ci < children.Count; ci++)
            {
                if (children[ci] is CommentRangeStart crs)
                {
                    string id = crs.Id?.Value ?? "";
                    if (!string.IsNullOrEmpty(id))
                        rangeStarts[id] = (pi, ci);
                }
                else if (children[ci] is CommentRangeEnd cre)
                {
                    string id = cre.Id?.Value ?? "";
                    if (!string.IsNullOrEmpty(id) && rangeStarts.TryGetValue(id, out var start))
                    {
                        // Extract anchor text between start and end markers
                        string anchorText = ExtractAnchorText(paragraphs, start.paraIndex, start.childIndex, pi, ci);
                        result[id] = (anchorText, start.paraIndex);
                        rangeStarts.Remove(id);
                    }
                }
            }
        }

        return result;
    }

    /// <summary>
    /// Extract the visible text between CommentRangeStart and CommentRangeEnd markers,
    /// potentially spanning multiple paragraphs.
    /// </summary>
    private static string ExtractAnchorText(List<Paragraph> paragraphs,
        int startPara, int startChild, int endPara, int endChild)
    {
        var texts = new List<string>();

        for (int pi = startPara; pi <= endPara; pi++)
        {
            var para = paragraphs[pi];
            var children = para.ChildElements.ToList();

            int fromChild = (pi == startPara) ? startChild + 1 : 0;
            int toChild = (pi == endPara) ? endChild : children.Count;

            for (int ci = fromChild; ci < toChild; ci++)
            {
                var child = children[ci];
                if (child is Run run)
                {
                    texts.Add(GetRunText(run));
                }
                else if (child is InsertedRun ins)
                {
                    texts.Add(GetInsertedRunText(ins));
                }
                // Deleted text is not visible, so skip DeletedRun
            }

            // Add space between paragraphs
            if (pi < endPara)
                texts.Add(" ");
        }

        return string.Join("", texts);
    }

    /// <summary>
    /// Extract comments from the WordprocessingCommentsPart.
    /// </summary>
    private List<CommentInfo> ExtractComments(WordprocessingDocument doc,
        List<Paragraph> paragraphs,
        Dictionary<string, (string anchorText, int paragraphIndex)> commentRanges)
    {
        var comments = new List<CommentInfo>();

        var commentsPart = doc.MainDocumentPart!.WordprocessingCommentsPart;
        if (commentsPart?.Comments == null)
            return comments;

        foreach (var comment in commentsPart.Comments.Elements<Comment>())
        {
            string id = comment.Id?.Value ?? "";
            string author = comment.Author?.Value ?? "";
            string? date = FormatDate(comment.Date);

            // Get comment text (all paragraph text within the comment)
            string commentText = string.Join("\n",
                comment.Elements<Paragraph>()
                    .Select(p =>
                    {
                        // Skip the AnnotationReferenceMark run
                        var runs = p.Elements<Run>()
                            .Where(r => !r.Elements<AnnotationReferenceMark>().Any());
                        return string.Join("", runs.Select(r => GetRunText(r)));
                    })
                    .Where(s => !string.IsNullOrEmpty(s)));

            string anchorText = "";
            int paraIndex = -1;
            if (commentRanges.TryGetValue(id, out var range))
            {
                anchorText = range.anchorText;
                paraIndex = range.paragraphIndex;
            }

            comments.Add(new CommentInfo
            {
                Id = id,
                Author = author,
                Date = date,
                AnchorText = anchorText,
                Text = commentText,
                ParagraphIndex = paraIndex
            });
        }

        return comments;
    }

    /// <summary>
    /// Extract document metadata from core properties.
    /// </summary>
    private DocumentMetadata ExtractMetadata(WordprocessingDocument doc, int wordCount, int paragraphCount)
    {
        var meta = new DocumentMetadata
        {
            WordCount = wordCount,
            ParagraphCount = paragraphCount
        };

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

        return meta;
    }

    /// <summary>
    /// Build aggregated summary statistics.
    /// </summary>
    private ReadSummary BuildSummary(ReadResult result)
    {
        var allChanges = result.Paragraphs.SelectMany(p => p.TrackedChanges).ToList();
        var insertions = allChanges.Count(c => c.Type == "insert");
        var deletions = allChanges.Count(c => c.Type == "delete");

        return new ReadSummary
        {
            TotalTrackedChanges = allChanges.Count,
            Insertions = insertions,
            Deletions = deletions,
            TotalComments = result.Comments.Count,
            ChangeAuthors = allChanges.Select(c => c.Author).Distinct().OrderBy(a => a).ToList(),
            CommentAuthors = result.Comments.Select(c => c.Author).Distinct().OrderBy(a => a).ToList()
        };
    }

    /// <summary>
    /// Print a human-readable summary of the read result.
    /// </summary>
    public static void PrintHumanReadable(ReadResult result)
    {
        Console.WriteLine();
        Console.WriteLine($"docx-review — {result.File}");
        Console.WriteLine(new string('═', 60));

        // Metadata
        Console.WriteLine();
        Console.WriteLine("Document Metadata");
        Console.WriteLine(new string('─', 40));
        if (result.Metadata.Title != null)
            Console.WriteLine($"  Title:          {result.Metadata.Title}");
        if (result.Metadata.Author != null)
            Console.WriteLine($"  Author:         {result.Metadata.Author}");
        if (result.Metadata.LastModifiedBy != null)
            Console.WriteLine($"  Last Modified:  {result.Metadata.LastModifiedBy}");
        if (result.Metadata.Created != null)
            Console.WriteLine($"  Created:        {result.Metadata.Created}");
        if (result.Metadata.Modified != null)
            Console.WriteLine($"  Modified:       {result.Metadata.Modified}");
        if (result.Metadata.Revision != null)
            Console.WriteLine($"  Revision:       {result.Metadata.Revision}");
        Console.WriteLine($"  Words:          {result.Metadata.WordCount}");
        Console.WriteLine($"  Paragraphs:     {result.Metadata.ParagraphCount}");

        // Summary
        Console.WriteLine();
        Console.WriteLine("Review Summary");
        Console.WriteLine(new string('─', 40));
        Console.WriteLine($"  Tracked Changes: {result.Summary.TotalTrackedChanges} ({result.Summary.Insertions} insertions, {result.Summary.Deletions} deletions)");
        Console.WriteLine($"  Comments:        {result.Summary.TotalComments}");
        if (result.Summary.ChangeAuthors.Count > 0)
            Console.WriteLine($"  Change Authors:  {string.Join(", ", result.Summary.ChangeAuthors)}");
        if (result.Summary.CommentAuthors.Count > 0)
            Console.WriteLine($"  Comment Authors: {string.Join(", ", result.Summary.CommentAuthors)}");

        // Paragraphs with tracked changes
        var parasWithChanges = result.Paragraphs
            .Where(p => p.TrackedChanges.Count > 0)
            .ToList();

        if (parasWithChanges.Count > 0)
        {
            Console.WriteLine();
            Console.WriteLine("Paragraphs with Tracked Changes");
            Console.WriteLine(new string('─', 40));

            foreach (var para in parasWithChanges)
            {
                string styleLabel = para.Style != null ? $" ({para.Style})" : "";
                Console.WriteLine($"\n  ¶{para.Index}{styleLabel}:");

                // Build inline view of the paragraph with change markers
                string inlineText = BuildInlineText(para);
                // Wrap to ~76 chars with indent
                foreach (var line in WordWrap(inlineText, 72))
                    Console.WriteLine($"    {line}");

                foreach (var tc in para.TrackedChanges)
                {
                    string icon = tc.Type == "insert" ? "+" : "-";
                    Console.WriteLine($"      [{icon}{tc.Type}] \"{Truncate(tc.Text, 50)}\" by {tc.Author}");
                }
            }
        }

        // Comments
        if (result.Comments.Count > 0)
        {
            Console.WriteLine();
            Console.WriteLine("Comments");
            Console.WriteLine(new string('─', 40));

            foreach (var comment in result.Comments)
            {
                Console.WriteLine($"\n  Comment #{comment.Id} by {comment.Author}");
                if (comment.Date != null)
                    Console.WriteLine($"    Date: {comment.Date}");
                if (!string.IsNullOrEmpty(comment.AnchorText))
                    Console.WriteLine($"    On: \"{Truncate(comment.AnchorText, 60)}\" (¶{comment.ParagraphIndex})");
                Console.WriteLine($"    {comment.Text}");
            }
        }

        Console.WriteLine();
    }

    /// <summary>
    /// Build a text representation of a paragraph with inline change markers.
    /// Deletions shown as [-deleted-], insertions as [+inserted+].
    /// </summary>
    private static string BuildInlineText(ParagraphInfo para)
    {
        // Simple approach: show the text with markers for changes
        // We can't perfectly reconstruct the inline view without the original XML,
        // so we show the visible text and list changes separately.
        // For a better inline view, we'd need to walk the XML again.
        // For now, show the plain text with a note about changes.
        return para.Text;
    }

    private static string Truncate(string s, int max) =>
        s.Length <= max ? s : s[..max] + "…";

    private static List<string> WordWrap(string text, int maxWidth)
    {
        if (string.IsNullOrEmpty(text))
            return new List<string> { "" };

        var lines = new List<string>();
        while (text.Length > maxWidth)
        {
            int breakAt = text.LastIndexOf(' ', maxWidth);
            if (breakAt <= 0) breakAt = maxWidth;
            lines.Add(text[..breakAt]);
            text = text[breakAt..].TrimStart();
        }
        if (text.Length > 0)
            lines.Add(text);
        return lines;
    }
}
