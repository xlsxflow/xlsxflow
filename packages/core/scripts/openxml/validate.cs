// Checks xlsx files against the Office Open XML schema with Microsoft's Open XML SDK, the same rules
// Excel is strict about. Run by check.py; on its own: dotnet run validate.cs -- <file or folder>...
#:package DocumentFormat.OpenXml@3.5.1
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var files = args.SelectMany(a => Directory.Exists(a) ? Directory.GetFiles(a, "*.xlsx") : [a]).Order().ToList();
var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
var failed = 0;
foreach (var file in files)
{
    List<string> errors;
    try
    {
        using var doc = SpreadsheetDocument.Open(file, false);
        errors = validator.Validate(doc).Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}").ToList();
    }
    catch (Exception e)
    {
        Console.WriteLine($"FAIL {Path.GetFileName(file)}: cannot open: {e.Message}");
        failed++;
        continue;
    }
    Console.WriteLine((errors.Count == 0 ? "ok   " : "FAIL ") + Path.GetFileName(file));
    foreach (var e in errors.Take(20))
        Console.WriteLine("     " + e);
    if (errors.Count > 20) Console.WriteLine($"     ...and {errors.Count - 20} more");
    if (errors.Count > 0) failed++;
}
Console.WriteLine($"{files.Count - failed} of {files.Count} files passed the schema check");
return failed == 0 ? 0 : 1;
