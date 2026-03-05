using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using DocxReview;

class Program
{
    static int Main(string[] args)
    {
        // Parse arguments
        string? inputPath = null;
        string? manifestPath = null;
        string? outputPath = null;
        string? author = null;
        bool jsonOutput = false;
        bool dryRun = false;
        bool readMode = false;
        bool diffMode = false;
        bool textConvMode = false;
        bool gitSetup = false;
        bool createMode = false;
        string? templatePath = null;
        bool showHelp = false;
        bool showVersion = false;
        bool acceptExisting = true;
        var positionalArgs = new List<string>();

        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "-v":
                case "--version":
                    showVersion = true;
                    break;
                case "-o":
                case "--output":
                    if (i + 1 < args.Length) outputPath = args[++i];
                    break;
                case "--author":
                    if (i + 1 < args.Length) author = args[++i];
                    break;
                case "--json":
                    jsonOutput = true;
                    break;
                case "--dry-run":
                    dryRun = true;
                    break;
                case "--read":
                    readMode = true;
                    break;
                case "--diff":
                    diffMode = true;
                    break;
                case "--textconv":
                    textConvMode = true;
                    break;
                case "--git-setup":
                    gitSetup = true;
                    break;
                case "--create":
                    createMode = true;
                    break;
                case "--template":
                    if (i + 1 < args.Length) templatePath = args[++i];
                    break;
                case "--accept-existing":
                    acceptExisting = true;
                    break;
                case "--no-accept-existing":
                    acceptExisting = false;
                    break;
                case "-h":
                case "--help":
                    showHelp = true;
                    break;
                default:
                    if (!args[i].StartsWith("-"))
                        positionalArgs.Add(args[i]);
                    break;
            }
        }

        // Map positional args
        if (positionalArgs.Count >= 1) inputPath = positionalArgs[0];
        if (positionalArgs.Count >= 2) manifestPath = positionalArgs[1];

        if (showVersion)
        {
            Console.WriteLine($"docx-review {GetVersion()}");
            return 0;
        }

        // ── Git setup ──────────────────────────────────────────────
        if (gitSetup)
        {
            PrintGitSetup();
            return 0;
        }

        // ── Create mode ──────────────────────────────────────────
        if (createMode)
        {
            if (outputPath == null && !dryRun)
            {
                Error("--create requires -o/--output path: docx-review --create -o manuscript.docx");
                return 1;
            }

            // In create mode, positionalArgs[0] is the manifest (not an input docx)
            string? createManifestPath = positionalArgs.Count >= 1 ? positionalArgs[0] : null;
            EditManifest? createManifest = null;

            if (createManifestPath != null)
            {
                if (!File.Exists(createManifestPath))
                {
                    Error($"Manifest file not found: {createManifestPath}");
                    return 1;
                }
                string mJson = File.ReadAllText(createManifestPath);
                try
                {
                    createManifest = JsonSerializer.Deserialize(mJson, DocxReviewJsonContext.Default.EditManifest)
                        ?? throw new Exception("Manifest deserialized to null");
                }
                catch (Exception ex)
                {
                    Error($"Failed to parse manifest JSON: {ex.Message}");
                    return 1;
                }
            }
            else if (Console.IsInputRedirected)
            {
                string mJson = Console.In.ReadToEnd();
                if (!string.IsNullOrWhiteSpace(mJson))
                {
                    try
                    {
                        createManifest = JsonSerializer.Deserialize(mJson, DocxReviewJsonContext.Default.EditManifest)
                            ?? throw new Exception("Manifest deserialized to null");
                    }
                    catch (Exception ex)
                    {
                        Error($"Failed to parse manifest JSON: {ex.Message}");
                        return 1;
                    }
                }
            }

            string createAuthor = author ?? createManifest?.Author ?? "Author";

            try
            {
                var creator = new DocumentCreator();
                var createResult = creator.Create(outputPath ?? "", createManifest, createAuthor, templatePath, dryRun);

                if (jsonOutput)
                {
                    Console.WriteLine(JsonSerializer.Serialize(createResult, DocxReviewJsonContext.Default.CreateResult));
                }
                else
                {
                    PrintCreateResult(createResult, dryRun);
                }
                return createResult.Success ? 0 : 1;
            }
            catch (Exception ex)
            {
                Error($"Create failed: {ex.Message}");
                return 1;
            }
        }

        if (showHelp || (inputPath == null && !gitSetup))
        {
            PrintUsage();
            return showHelp ? 0 : 1;
        }

        // ── Diff mode ─────────────────────────────────────────────
        if (diffMode)
        {
            if (manifestPath == null)
            {
                Error("--diff requires two files: docx-review --diff old.docx new.docx");
                return 1;
            }

            if (!File.Exists(inputPath!))
            {
                Error($"Old file not found: {inputPath}");
                return 1;
            }
            if (!File.Exists(manifestPath))
            {
                Error($"New file not found: {manifestPath}");
                return 1;
            }

            try
            {
                var oldDoc = DocumentExtractor.Extract(inputPath!);
                var newDoc = DocumentExtractor.Extract(manifestPath);
                var diffResult = DocumentDiffer.Diff(oldDoc, newDoc);

                if (jsonOutput)
                {
                    Console.WriteLine(JsonSerializer.Serialize(diffResult, DocxReviewJsonContext.Default.DiffResult));
                }
                else
                {
                    DocumentDiffer.PrintHumanReadable(diffResult);
                }
                return 0;
            }
            catch (Exception ex)
            {
                Error($"Diff failed: {ex.Message}");
                return 1;
            }
        }

        // ── TextConv mode ─────────────────────────────────────────
        if (textConvMode)
        {
            if (!File.Exists(inputPath!))
            {
                Error($"File not found: {inputPath}");
                return 1;
            }

            try
            {
                var extraction = DocumentExtractor.Extract(inputPath!);
                Console.Write(TextConv.Convert(extraction));
                return 0;
            }
            catch (Exception ex)
            {
                Error($"TextConv failed: {ex.Message}");
                return 1;
            }
        }

        // Validate input file
        if (!File.Exists(inputPath))
        {
            Error($"Input file not found: {inputPath}");
            return 1;
        }

        // ── Read mode ─────────────────────────────────────────────
        if (readMode)
        {
            try
            {
                var reader = new DocumentReader();
                var readResult = reader.Read(inputPath);

                if (jsonOutput)
                {
                    Console.WriteLine(JsonSerializer.Serialize(readResult, DocxReviewJsonContext.Default.ReadResult));
                }
                else
                {
                    DocumentReader.PrintHumanReadable(readResult);
                }
                return 0;
            }
            catch (Exception ex)
            {
                Error($"Read failed: {ex.Message}");
                return 1;
            }
        }

        // ── Edit mode (original behavior) ─────────────────────────
        // Read manifest from file or stdin
        string manifestJson;
        if (manifestPath != null)
        {
            if (!File.Exists(manifestPath))
            {
                Error($"Manifest file not found: {manifestPath}");
                return 1;
            }
            manifestJson = File.ReadAllText(manifestPath);
        }
        else if (!Console.IsInputRedirected)
        {
            Error("No manifest file specified and no stdin input.\nUsage: docx-review <input.docx> <edits.json> -o <output.docx>");
            return 1;
        }
        else
        {
            manifestJson = Console.In.ReadToEnd();
        }

        // Default output path
        if (outputPath == null && !dryRun)
        {
            string dir = Path.GetDirectoryName(inputPath) ?? ".";
            string name = Path.GetFileNameWithoutExtension(inputPath);
            outputPath = Path.Combine(dir, $"{name}_reviewed.docx");
        }

        // Deserialize manifest (using source-generated context for trim/AOT safety)
        EditManifest manifest;
        try
        {
            manifest = JsonSerializer.Deserialize(manifestJson, DocxReviewJsonContext.Default.EditManifest)
                ?? throw new Exception("Manifest deserialized to null");
        }
        catch (Exception ex)
        {
            Error($"Failed to parse manifest JSON: {ex.Message}");
            return 1;
        }

        // Resolve author (CLI flag > manifest > default)
        string effectiveAuthor = author ?? manifest.Author ?? "Reviewer";

        // Process
        var editor = new DocumentEditor(effectiveAuthor);
        ProcessingResult result;

        try
        {
            result = editor.Process(inputPath, outputPath ?? "", manifest, dryRun, acceptExisting);
        }
        catch (Exception ex)
        {
            Error($"Processing failed: {ex.Message}");
            return 1;
        }

        // Output
        if (jsonOutput)
        {
            Console.WriteLine(JsonSerializer.Serialize(result, DocxReviewJsonContext.Default.ProcessingResult));
        }
        else
        {
            PrintHumanResult(result, dryRun);
        }

        return result.Success ? 0 : 1;
    }

    static void PrintUsage()
    {
        Console.Error.WriteLine(@"docx-review — Read, write, create, and diff Word documents with full revision awareness

Usage:
  docx-review <input.docx> --read [--json]              Read review state
  docx-review <input.docx> <edits.json> [options]       Write tracked changes/comments
  docx-review --create -o <output.docx> [manifest.json]  Create from NIH template
  docx-review --diff <old.docx> <new.docx> [--json]     Semantic document diff
  docx-review --textconv <file.docx>                     Git textconv (normalized text)
  docx-review --git-setup                                Print git configuration
  cat edits.json | docx-review <input.docx> [options]

Create Options:
  --create               Create new document from bundled NIH template
  --template <path>      Use custom template instead of built-in NIH template
  -o, --output <path>    Output file path (required for create)

Diff & Git Integration:
  --diff                 Compare two documents semantically (text, comments,
                         tracked changes, formatting, styles, metadata)
  --textconv             Output normalized text for use as git diff textconv driver
  --git-setup            Print .gitattributes and .gitconfig setup instructions

Read/Write Options:
  --read                 Read mode: extract tracked changes, comments, metadata
  -o, --output <path>    Output file path (default: <input>_reviewed.docx)
  --author <name>        Reviewer name (overrides manifest author)
  --json                 Output results as JSON
  --dry-run              Validate manifest without modifying
  --no-accept-existing   Preserve existing tracked changes (default: accept them)
  -h, --help             Show this help

JSON Manifest Format:
  {
    ""author"": ""Reviewer Name"",
    ""changes"": [
      { ""type"": ""replace"", ""find"": ""old"", ""replace"": ""new"" },
      { ""type"": ""delete"", ""find"": ""text to remove"" },
      { ""type"": ""insert_after"", ""anchor"": ""after this"", ""text"": ""new text"" },
      { ""type"": ""insert_before"", ""anchor"": ""before this"", ""text"": ""new text"" }
    ],
    ""comments"": [
      { ""anchor"": ""text to comment on"", ""text"": ""Comment content"" },
      { ""op"": ""update"", ""id"": 12, ""text"": ""Updated comment text"" }
    ]
  }");
    }

    static void PrintGitSetup()
    {
        Console.WriteLine(@"Git Integration for Word Documents
══════════════════════════════════

Add to your repository's .gitattributes:

  *.docx diff=docx

Add to your .gitconfig (global or per-repo):

  [diff ""docx""]
      textconv = docx-review --textconv

Now `git diff` will show meaningful content changes for .docx files,
including text, comments, tracked changes, formatting, and metadata.

For two-file comparison outside git:

  docx-review --diff old.docx new.docx
  docx-review --diff old.docx new.docx --json
");
    }

    static void PrintHumanResult(ProcessingResult result, bool dryRun)
    {
        string mode = dryRun ? "[DRY RUN] " : "";
        Console.WriteLine($"\n{mode}docx-review results");
        Console.WriteLine(new string('─', 50));
        Console.WriteLine($"  Input:    {result.Input}");
        if (!dryRun && result.Output != null)
            Console.WriteLine($"  Output:   {result.Output}");
        Console.WriteLine($"  Author:   {result.Author}");
        Console.WriteLine($"  Changes:  {result.ChangesSucceeded}/{result.ChangesAttempted}");
        Console.WriteLine($"  Comments: {result.CommentsSucceeded}/{result.CommentsAttempted}");
        Console.WriteLine();

        foreach (var r in result.Results)
        {
            string icon = r.Success ? "✓" : "✗";
            Console.WriteLine($"  {icon} [{r.Type}] {r.Message}");
        }

        Console.WriteLine();
        if (result.Success)
            Console.WriteLine(dryRun ? "✅ All edits would succeed" : "✅ All edits applied successfully");
        else
            Console.WriteLine("⚠️  Some edits failed (see above)");
    }

    static void PrintCreateResult(CreateResult result, bool dryRun)
    {
        string mode = dryRun ? "[DRY RUN] " : "";
        Console.WriteLine($"\n{mode}docx-review create");
        Console.WriteLine(new string('─', 50));
        Console.WriteLine($"  Template: {result.Template}");
        if (!dryRun && result.Output != null)
            Console.WriteLine($"  Output:   {result.Output}");

        if (result.Populated)
        {
            Console.WriteLine($"  Changes:  {result.ChangesSucceeded}/{result.ChangesAttempted}");
            Console.WriteLine($"  Comments: {result.CommentsSucceeded}/{result.CommentsAttempted}");
            Console.WriteLine();

            foreach (var r in result.Results)
            {
                string icon = r.Success ? "✓" : "✗";
                Console.WriteLine($"  {icon} [{r.Type}] {r.Message}");
            }
        }

        Console.WriteLine();
        if (!result.Populated)
            Console.WriteLine(dryRun ? "✅ Template would be created successfully" : "✅ Template copied — ready for editing");
        else if (result.Success)
            Console.WriteLine(dryRun ? "✅ All populate edits would succeed" : "✅ Template created and populated successfully");
        else
            Console.WriteLine("⚠️  Some populate edits failed (see above)");
    }

    static string GetVersion()
    {
        var asm = System.Reflection.Assembly.GetExecutingAssembly();
        var ver = asm.GetName().Version;
        return ver != null ? $"{ver.Major}.{ver.Minor}.{ver.Build}" : "1.0.0";
    }

    static void Error(string msg) => Console.Error.WriteLine($"Error: {msg}");
}
