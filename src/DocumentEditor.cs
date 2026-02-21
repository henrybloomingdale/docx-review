using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxReview;

/// <summary>
/// Core editing engine for adding tracked changes and comments to .docx files
/// using the Open XML SDK. Produces proper w:del/w:ins/comment markup that
/// renders natively in Microsoft Word.
/// </summary>
public class DocumentEditor
{
    private readonly string _author;
    private readonly string _dateStr;
    private int _revId = 100;

    public DocumentEditor(string author, DateTime? date = null)
    {
        _author = author;
        _dateStr = (date ?? DateTime.UtcNow).ToString("yyyy-MM-ddTHH:mm:ssZ");
    }

    /// <summary>
    /// Process a complete edit manifest against a document.
    /// </summary>
    public ProcessingResult Process(string inputPath, string outputPath, EditManifest manifest, bool dryRun = false)
    {
        var result = new ProcessingResult
        {
            Input = inputPath,
            Output = dryRun ? null : outputPath,
            Author = _author,
            ChangesAttempted = manifest.Changes?.Count ?? 0,
            CommentsAttempted = manifest.Comments?.Count ?? 0
        };

        if (!dryRun)
            File.Copy(inputPath, outputPath, true);

        // For dry-run, open a temp copy read-only-ish to check matches
        string workPath = dryRun ? CreateTempCopy(inputPath) : outputPath;

        try
        {
            using var doc = WordprocessingDocument.Open(workPath, !dryRun);
            var body = doc.MainDocumentPart!.Document.Body!;
            var paragraphs = body.Elements<Paragraph>().ToList();

            // --- Comments first (before tracked changes modify XML) ---
            if (manifest.Comments != null)
            {
                bool hasAddCommentOps = manifest.Comments.Any(c =>
                {
                    string op = (c.Op ?? "add").Trim();
                    return op.Length == 0 || op.Equals("add", StringComparison.OrdinalIgnoreCase);
                });

                if (!dryRun && hasAddCommentOps)
                    EnsureCommentsPart(doc);

                int commentId = (dryRun || !hasAddCommentOps) ? 0 : GetNextCommentId(doc);

                for (int i = 0; i < manifest.Comments.Count; i++)
                {
                    var cdef = manifest.Comments[i];
                    string op = (cdef.Op ?? "add").Trim().ToLowerInvariant();
                    if (op.Length == 0) op = "add";

                    var er = new EditResult
                    {
                        Index = i,
                        Type = op == "update" ? "comment_update" : "comment"
                    };

                    if (op == "update")
                    {
                        if (cdef.Id == null)
                        {
                            er.Success = false;
                            er.Message = "Missing 'id' for comment update";
                        }
                        else if (cdef.Text == null)
                        {
                            er.Success = false;
                            er.Message = "Missing 'text' for comment update";
                        }
                        else
                        {
                            string idStr = cdef.Id.Value.ToString();
                            bool exists = CommentExistsById(doc, idStr);
                            if (!exists)
                            {
                                er.Success = false;
                                er.Message = $"Comment id not found: {idStr}";
                            }
                            else if (dryRun)
                            {
                                er.Success = true;
                                er.Message = "Comment found for update";
                            }
                            else
                            {
                                bool ok = UpdateCommentById(doc, idStr, cdef.Text);
                                er.Success = ok;
                                er.Message = ok ? "Comment updated" : $"Comment id not found: {idStr}";
                            }
                        }
                    }
                    else if (op == "add")
                    {
                        if (string.IsNullOrEmpty(cdef.Anchor))
                        {
                            er.Success = false;
                            er.Message = "Empty anchor text";
                        }
                        else if (dryRun)
                        {
                            bool found = FindAnchorInParagraphs(paragraphs, cdef.Anchor!);
                            er.Success = found;
                            er.Message = found ? "Anchor found" : $"Anchor not found: \"{Truncate(cdef.Anchor!, 60)}\"";
                        }
                        else
                        {
                            bool ok = AddComment(doc, paragraphs, cdef.Anchor!, cdef.Text ?? "", commentId);
                            er.Success = ok;
                            er.Message = ok ? "Comment added" : $"Anchor not found: \"{Truncate(cdef.Anchor!, 60)}\"";
                            if (ok) commentId++;
                        }
                    }
                    else
                    {
                        er.Success = false;
                        er.Message = $"Unknown comment op: {op}";
                    }

                    result.Results.Add(er);
                    if (er.Success) result.CommentsSucceeded++;
                }

                // Refresh paragraphs after comment markers inserted
                if (!dryRun)
                    paragraphs = body.Elements<Paragraph>().ToList();
            }

            // --- Tracked Changes ---
            if (manifest.Changes != null)
            {
                for (int i = 0; i < manifest.Changes.Count; i++)
                {
                    var change = manifest.Changes[i];
                    var er = new EditResult { Index = i, Type = change.Type };

                    try
                    {
                        switch (change.Type.ToLowerInvariant())
                        {
                            case "replace":
                                er = ProcessReplace(paragraphs, change, i, dryRun);
                                break;
                            case "delete":
                                er = ProcessDelete(paragraphs, change, i, dryRun);
                                break;
                            case "insert_after":
                                er = ProcessInsert(paragraphs, change, i, dryRun, after: true);
                                break;
                            case "insert_before":
                                er = ProcessInsert(paragraphs, change, i, dryRun, after: false);
                                break;
                            default:
                                er.Success = false;
                                er.Message = $"Unknown change type: {change.Type}";
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        er.Success = false;
                        er.Message = $"Error: {ex.Message}";
                    }

                    result.Results.Add(er);
                    if (er.Success) result.ChangesSucceeded++;
                }
            }

            if (!dryRun)
                doc.MainDocumentPart!.Document.Save();
        }
        finally
        {
            if (dryRun && File.Exists(workPath))
                File.Delete(workPath);
        }

        result.Success = result.ChangesSucceeded == result.ChangesAttempted
                      && result.CommentsSucceeded == result.CommentsAttempted;

        return result;
    }

    #region Change Processors

    private EditResult ProcessReplace(List<Paragraph> paragraphs, Change change, int index, bool dryRun)
    {
        var er = new EditResult { Index = index, Type = "replace" };

        if (string.IsNullOrEmpty(change.Find))
        {
            er.Success = false;
            er.Message = "Missing 'find' field";
            return er;
        }

        if (change.Replace == null)
        {
            er.Success = false;
            er.Message = "Missing 'replace' field";
            return er;
        }

        if (dryRun)
        {
            bool found = FindTextInParagraphs(paragraphs, change.Find);
            er.Success = found;
            er.Message = found
                ? $"Match found for: \"{Truncate(change.Find, 60)}\""
                : $"No match for: \"{Truncate(change.Find, 60)}\"";
            return er;
        }

        int n = ReplaceWithTracking(paragraphs, change.Find, change.Replace);
        er.Success = n > 0;
        er.Message = n > 0 ? "Replaced" : $"No match for: \"{Truncate(change.Find, 60)}\"";
        return er;
    }

    private EditResult ProcessDelete(List<Paragraph> paragraphs, Change change, int index, bool dryRun)
    {
        var er = new EditResult { Index = index, Type = "delete" };

        if (string.IsNullOrEmpty(change.Find))
        {
            er.Success = false;
            er.Message = "Missing 'find' field";
            return er;
        }

        if (dryRun)
        {
            bool found = FindTextInParagraphs(paragraphs, change.Find);
            er.Success = found;
            er.Message = found
                ? $"Match found for deletion: \"{Truncate(change.Find, 60)}\""
                : $"No match for: \"{Truncate(change.Find, 60)}\"";
            return er;
        }

        int n = DeleteWithTracking(paragraphs, change.Find);
        er.Success = n > 0;
        er.Message = n > 0 ? "Deleted" : $"No match for: \"{Truncate(change.Find, 60)}\"";
        return er;
    }

    private EditResult ProcessInsert(List<Paragraph> paragraphs, Change change, int index, bool dryRun, bool after)
    {
        string direction = after ? "insert_after" : "insert_before";
        var er = new EditResult { Index = index, Type = direction };

        if (string.IsNullOrEmpty(change.Anchor))
        {
            er.Success = false;
            er.Message = "Missing 'anchor' field";
            return er;
        }

        if (string.IsNullOrEmpty(change.Text))
        {
            er.Success = false;
            er.Message = "Missing 'text' field";
            return er;
        }

        if (dryRun)
        {
            bool found = FindTextInParagraphs(paragraphs, change.Anchor);
            er.Success = found;
            er.Message = found
                ? $"Anchor found for {direction}: \"{Truncate(change.Anchor, 60)}\""
                : $"No match for: \"{Truncate(change.Anchor, 60)}\"";
            return er;
        }

        int n = InsertWithTracking(paragraphs, change.Anchor, change.Text, after);
        er.Success = n > 0;
        er.Message = n > 0
            ? $"Inserted {(after ? "after" : "before")} anchor"
            : $"No match for: \"{Truncate(change.Anchor, 60)}\"";
        return er;
    }

    #endregion

    #region Core XML Operations

    /// <summary>
    /// Replace first occurrence of text with proper tracked changes markup.
    /// Creates w:del (DeletedRun) and w:ins (InsertedRun) elements.
    /// Handles text spanning multiple XML runs and preserves formatting.
    /// </summary>
    private int ReplaceWithTracking(List<Paragraph> paragraphs, string find, string replace)
    {
        foreach (var para in paragraphs)
        {
            var runs = para.Elements<Run>().ToList();
            string paraText = string.Join("", runs.SelectMany(r => r.Elements<Text>()).Select(t => t.Text));
            int idx = paraText.IndexOf(find, StringComparison.Ordinal);
            if (idx < 0) continue;

            // Map runs to character positions
            int charPos = 0;
            var runMap = new List<(Run run, int start, int end, RunProperties? rPr)>();
            foreach (var run in runs)
            {
                var textEl = run.Elements<Text>().FirstOrDefault();
                if (textEl != null)
                {
                    int len = textEl.Text.Length;
                    runMap.Add((run, charPos, charPos + len,
                        run.RunProperties?.CloneNode(true) as RunProperties));
                    charPos += len;
                }
            }

            int matchEnd = idx + find.Length;
            var affected = runMap.Where(r => r.start < matchEnd && r.end > idx).ToList();
            if (affected.Count == 0) continue;

            var rPr = affected[0].rPr;
            var firstAffected = affected.First();
            string revId = (_revId++).ToString();

            string prefix = "";
            if (idx > firstAffected.start)
            {
                var t = firstAffected.run.Elements<Text>().First();
                prefix = t.Text.Substring(0, idx - firstAffected.start);
            }

            var lastAffected = affected.Last();
            string suffix = "";
            if (matchEnd < lastAffected.end)
            {
                var t = lastAffected.run.Elements<Text>().First();
                suffix = t.Text.Substring(matchEnd - lastAffected.start);
            }

            var insertPoint = firstAffected.run;

            if (prefix.Length > 0)
            {
                var prefixRun = new Run();
                if (rPr != null) prefixRun.Append(rPr.CloneNode(true));
                prefixRun.Append(new Text(prefix) { Space = SpaceProcessingModeValues.Preserve });
                para.InsertBefore(prefixRun, insertPoint);
            }

            // w:del
            var del = new DeletedRun()
            {
                Author = new StringValue(_author),
                Date = new DateTimeValue(DateTime.Parse(_dateStr)),
                Id = revId
            };
            var delRun = new Run();
            if (rPr != null) delRun.Append(rPr.CloneNode(true));
            delRun.Append(new DeletedText(find) { Space = SpaceProcessingModeValues.Preserve });
            del.Append(delRun);
            para.InsertBefore(del, insertPoint);

            // w:ins
            var ins = new InsertedRun()
            {
                Author = new StringValue(_author),
                Date = new DateTimeValue(DateTime.Parse(_dateStr)),
                Id = (_revId++).ToString()
            };
            var insRun = new Run();
            if (rPr != null) insRun.Append(rPr.CloneNode(true));
            insRun.Append(new Text(replace) { Space = SpaceProcessingModeValues.Preserve });
            ins.Append(insRun);
            para.InsertBefore(ins, insertPoint);

            if (suffix.Length > 0)
            {
                var suffixRun = new Run();
                if (rPr != null) suffixRun.Append(rPr.CloneNode(true));
                suffixRun.Append(new Text(suffix) { Space = SpaceProcessingModeValues.Preserve });
                para.InsertBefore(suffixRun, insertPoint);
            }

            foreach (var (run, _, _, _) in affected)
                run.Remove();

            return 1;
        }
        return 0;
    }

    /// <summary>
    /// Delete text with tracked changes (w:del only, no w:ins).
    /// </summary>
    private int DeleteWithTracking(List<Paragraph> paragraphs, string find)
    {
        foreach (var para in paragraphs)
        {
            var runs = para.Elements<Run>().ToList();
            string paraText = string.Join("", runs.SelectMany(r => r.Elements<Text>()).Select(t => t.Text));
            int idx = paraText.IndexOf(find, StringComparison.Ordinal);
            if (idx < 0) continue;

            int charPos = 0;
            var runMap = new List<(Run run, int start, int end, RunProperties? rPr)>();
            foreach (var run in runs)
            {
                var textEl = run.Elements<Text>().FirstOrDefault();
                if (textEl != null)
                {
                    int len = textEl.Text.Length;
                    runMap.Add((run, charPos, charPos + len,
                        run.RunProperties?.CloneNode(true) as RunProperties));
                    charPos += len;
                }
            }

            int matchEnd = idx + find.Length;
            var affected = runMap.Where(r => r.start < matchEnd && r.end > idx).ToList();
            if (affected.Count == 0) continue;

            var rPr = affected[0].rPr;
            var firstAffected = affected.First();
            string revId = (_revId++).ToString();

            string prefix = "";
            if (idx > firstAffected.start)
            {
                var t = firstAffected.run.Elements<Text>().First();
                prefix = t.Text.Substring(0, idx - firstAffected.start);
            }

            var lastAffected = affected.Last();
            string suffix = "";
            if (matchEnd < lastAffected.end)
            {
                var t = lastAffected.run.Elements<Text>().First();
                suffix = t.Text.Substring(matchEnd - lastAffected.start);
            }

            var insertPoint = firstAffected.run;

            if (prefix.Length > 0)
            {
                var prefixRun = new Run();
                if (rPr != null) prefixRun.Append(rPr.CloneNode(true));
                prefixRun.Append(new Text(prefix) { Space = SpaceProcessingModeValues.Preserve });
                para.InsertBefore(prefixRun, insertPoint);
            }

            // w:del only — no w:ins
            var del = new DeletedRun()
            {
                Author = new StringValue(_author),
                Date = new DateTimeValue(DateTime.Parse(_dateStr)),
                Id = revId
            };
            var delRun = new Run();
            if (rPr != null) delRun.Append(rPr.CloneNode(true));
            delRun.Append(new DeletedText(find) { Space = SpaceProcessingModeValues.Preserve });
            del.Append(delRun);
            para.InsertBefore(del, insertPoint);

            if (suffix.Length > 0)
            {
                var suffixRun = new Run();
                if (rPr != null) suffixRun.Append(rPr.CloneNode(true));
                suffixRun.Append(new Text(suffix) { Space = SpaceProcessingModeValues.Preserve });
                para.InsertBefore(suffixRun, insertPoint);
            }

            foreach (var (run, _, _, _) in affected)
                run.Remove();

            return 1;
        }
        return 0;
    }

    /// <summary>
    /// Insert text before or after an anchor with tracked insertion (w:ins).
    /// </summary>
    private int InsertWithTracking(List<Paragraph> paragraphs, string anchor, string text, bool after)
    {
        foreach (var para in paragraphs)
        {
            var runs = para.Elements<Run>().ToList();
            string paraText = string.Join("", runs.SelectMany(r => r.Elements<Text>()).Select(t => t.Text));
            int idx = paraText.IndexOf(anchor, StringComparison.Ordinal);
            if (idx < 0) continue;

            int charPos = 0;
            var runMap = new List<(Run run, int start, int end, RunProperties? rPr)>();
            foreach (var run in runs)
            {
                var textEl = run.Elements<Text>().FirstOrDefault();
                if (textEl != null)
                {
                    int len = textEl.Text.Length;
                    runMap.Add((run, charPos, charPos + len,
                        run.RunProperties?.CloneNode(true) as RunProperties));
                    charPos += len;
                }
            }

            int anchorEnd = idx + anchor.Length;

            // Find the run that contains the insertion point
            (Run run, int start, int end, RunProperties? rPr) targetEntry;
            int splitPos;

            if (after)
            {
                targetEntry = runMap.First(r => r.end >= anchorEnd && r.start < anchorEnd);
                splitPos = anchorEnd;
            }
            else
            {
                targetEntry = runMap.First(r => r.end > idx && r.start <= idx);
                splitPos = idx;
            }

            var rPr = targetEntry.rPr;
            var targetRun = targetEntry.run;

            var textEl2 = targetRun.Elements<Text>().First();
            string fullText = textEl2.Text;
            int localSplit = splitPos - targetEntry.start;

            string beforeSplit = fullText.Substring(0, localSplit);
            string afterSplit = fullText.Substring(localSplit);

            // Build: beforeRun + w:ins + afterRun, then remove original
            if (beforeSplit.Length > 0)
            {
                var beforeRun = new Run();
                if (rPr != null) beforeRun.Append(rPr.CloneNode(true));
                beforeRun.Append(new Text(beforeSplit) { Space = SpaceProcessingModeValues.Preserve });
                para.InsertBefore(beforeRun, targetRun);
            }

            var ins = new InsertedRun()
            {
                Author = new StringValue(_author),
                Date = new DateTimeValue(DateTime.Parse(_dateStr)),
                Id = (_revId++).ToString()
            };
            var insRun = new Run();
            if (rPr != null) insRun.Append(rPr.CloneNode(true));
            insRun.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            ins.Append(insRun);
            para.InsertBefore(ins, targetRun);

            if (afterSplit.Length > 0)
            {
                var afterRun = new Run();
                if (rPr != null) afterRun.Append(rPr.CloneNode(true));
                afterRun.Append(new Text(afterSplit) { Space = SpaceProcessingModeValues.Preserve });
                para.InsertBefore(afterRun, targetRun);
            }

            targetRun.Remove();
            return 1;
        }
        return 0;
    }

    /// <summary>
    /// Add a comment anchored to specific text. Places CommentRangeStart/End
    /// markers around the anchor text and adds the comment to comments.xml.
    /// </summary>
    private bool AddComment(WordprocessingDocument doc, List<Paragraph> paragraphs,
        string anchorText, string commentText, int id)
    {
        foreach (var para in paragraphs)
        {
            string pText = para.InnerText;
            int idx = pText.IndexOf(anchorText, StringComparison.Ordinal);
            if (idx < 0) continue;

            var runs = para.Descendants<Run>().ToList();
            int charPos = 0;
            Run? startRun = null;
            Run? endRun = null;
            int anchorEnd = idx + anchorText.Length;

            foreach (var run in runs)
            {
                string runText = run.InnerText;
                int runEnd = charPos + runText.Length;

                if (startRun == null && runEnd > idx)
                    startRun = run;
                if (runEnd >= anchorEnd)
                {
                    endRun = run;
                    break;
                }
                charPos += runText.Length;
            }

            if (startRun == null) continue;
            endRun ??= startRun;

            string idStr = id.ToString();

            // Add comment to comments.xml
            var commentsPart = doc.MainDocumentPart!.WordprocessingCommentsPart!;
            var comment = new Comment()
            {
                Id = idStr,
                Author = new StringValue(_author),
                Date = new DateTimeValue(DateTime.Parse(_dateStr)),
                Initials = new StringValue(_author.Length > 0 ? _author[0].ToString() : "R")
            };
            comment.Append(new Paragraph(
                new Run(
                    new RunProperties(new RunStyle() { Val = "CommentReference" }),
                    new AnnotationReferenceMark()
                ),
                new Run(
                    new Text(commentText) { Space = SpaceProcessingModeValues.Preserve }
                )
            ));
            commentsPart.Comments.Append(comment);
            commentsPart.Comments.Save();

            // Insert CommentRangeStart before the startRun
            startRun.InsertBeforeSelf(new CommentRangeStart() { Id = idStr });

            // Insert CommentRangeEnd and reference after the endRun
            endRun.InsertAfterSelf(new CommentRangeEnd() { Id = idStr });
            var rangeEnd = endRun.NextSibling<CommentRangeEnd>();
            if (rangeEnd != null)
            {
                rangeEnd.InsertAfterSelf(new Run(
                    new RunProperties(new RunStyle() { Val = "CommentReference" }),
                    new CommentReference() { Id = idStr }
                ));
            }

            return true;
        }
        return false;
    }

    private static bool CommentExistsById(WordprocessingDocument doc, string id)
    {
        var commentsPart = doc.MainDocumentPart!.WordprocessingCommentsPart;
        if (commentsPart?.Comments == null) return false;

        return commentsPart.Comments.Elements<Comment>()
            .Any(c => string.Equals(c.Id?.Value, id, StringComparison.Ordinal));
    }

    /// <summary>
    /// Replace the text payload for an existing comment ID while preserving
    /// comment metadata (author/date/initials) and anchor references in the body.
    /// </summary>
    private static bool UpdateCommentById(WordprocessingDocument doc, string id, string newText)
    {
        var commentsPart = doc.MainDocumentPart!.WordprocessingCommentsPart;
        if (commentsPart?.Comments == null) return false;

        var comment = commentsPart.Comments.Elements<Comment>()
            .FirstOrDefault(c => string.Equals(c.Id?.Value, id, StringComparison.Ordinal));
        if (comment == null) return false;

        RewriteCommentText(comment, newText);
        commentsPart.Comments.Save();
        return true;
    }

    private static void RewriteCommentText(Comment comment, string text)
    {
        var firstPara = comment.Elements<Paragraph>().FirstOrDefault();
        var paraAttrs = firstPara?.GetAttributes().ToList() ?? new List<OpenXmlAttribute>();
        var paraProps = firstPara?.ParagraphProperties?.CloneNode(true) as ParagraphProperties;

        var newPara = new Paragraph();
        if (paraProps != null)
            newPara.Append(paraProps);
        foreach (var attr in paraAttrs)
            newPara.SetAttribute(attr);

        newPara.Append(new Run(
            new RunProperties(new RunStyle() { Val = "CommentReference" }),
            new AnnotationReferenceMark()
        ));
        newPara.Append(new Run(
            new Text(text) { Space = SpaceProcessingModeValues.Preserve }
        ));

        foreach (var para in comment.Elements<Paragraph>().ToList())
            para.Remove();
        comment.Append(newPara);
    }

    #endregion

    #region Helpers

    private static void EnsureCommentsPart(WordprocessingDocument doc)
    {
        if (doc.MainDocumentPart!.WordprocessingCommentsPart == null)
        {
            var part = doc.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
            part.Comments = new Comments();
            part.Comments.Save();
        }
    }

    private static int GetNextCommentId(WordprocessingDocument doc)
    {
        var part = doc.MainDocumentPart!.WordprocessingCommentsPart;
        if (part?.Comments == null) return 0;

        var ids = part.Comments.Elements<Comment>()
            .Select(c => { int.TryParse(c.Id?.Value, out int id); return id; })
            .ToList();

        return ids.Count > 0 ? ids.Max() + 1 : 0;
    }

    private static bool FindTextInParagraphs(List<Paragraph> paragraphs, string text)
    {
        foreach (var para in paragraphs)
        {
            var runs = para.Elements<Run>().ToList();
            string paraText = string.Join("", runs.SelectMany(r => r.Elements<Text>()).Select(t => t.Text));
            if (paraText.Contains(text, StringComparison.Ordinal))
                return true;
        }
        return false;
    }

    private static bool FindAnchorInParagraphs(List<Paragraph> paragraphs, string anchor)
    {
        foreach (var para in paragraphs)
        {
            if (para.InnerText.Contains(anchor, StringComparison.Ordinal))
                return true;
        }
        return false;
    }

    private static string CreateTempCopy(string path)
    {
        string tmp = Path.Combine(Path.GetTempPath(), $"docx-review-{Guid.NewGuid()}.docx");
        File.Copy(path, tmp);
        return tmp;
    }

    private static string Truncate(string s, int max) =>
        s.Length <= max ? s : s.Substring(0, max) + "…";

    #endregion
}
