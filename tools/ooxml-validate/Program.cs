// Validates .pptx files against the Office Open XML schemas using the Open XML SDK.
// PowerPoint is stricter than LibreOffice; this catches files that would trigger a "repair" prompt.
// Usage: dotnet run --project tools/ooxml-validate -- <file.pptx>... [--ignore-template-errors <template.pptx>...]
// Exit code 1 if any file has errors that its template did not already have.
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var files = new List<string>();
var baselines = new List<string>();
var target = files;
foreach (var arg in args)
{
    if (arg == "--baseline") { target = baselines; continue; }
    target.Add(arg);
}

static List<string> Errors(string path)
{
    using var doc = PresentationDocument.Open(path, false);
    var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
    return validator.Validate(doc)
        .Select(e => $"{e.Part?.Uri} {e.Path?.XPath} :: {e.Description}")
        .ToList();
}

// Errors already present in the source templates are not the engine's fault; report them once.
var known = new HashSet<string>();
foreach (var b in baselines)
{
    var errs = Errors(b);
    foreach (var e in errs) known.Add(Normalize(e));
    if (errs.Count > 0) Console.WriteLine($"note: template {Path.GetFileName(b)} already has {errs.Count} schema error(s); ignored in outputs.");
}

var failed = 0;
foreach (var f in files)
{
    List<string> errs;
    try { errs = Errors(f); }
    catch (Exception ex) { Console.WriteLine($"FAIL {f}: cannot open: {ex.Message}"); failed++; continue; }
    var fresh = errs.Where(e => !known.Contains(Normalize(e))).ToList();
    if (fresh.Count == 0) { Console.WriteLine($"ok   {f}"); continue; }
    failed++;
    Console.WriteLine($"FAIL {f}: {fresh.Count} schema error(s)");
    foreach (var e in fresh.Take(20)) Console.WriteLine($"     {e}");
}
return failed == 0 ? 0 : 1;

// Slide/chart numbers differ between template and output; compare errors by kind, not location index.
static string Normalize(string e) => System.Text.RegularExpressions.Regex.Replace(e, @"\[\d+\]|\d+\.xml", "#");
