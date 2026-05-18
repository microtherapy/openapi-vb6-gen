using System.Globalization;
using System.Text;

namespace OpenApiVb6Gen;

internal static class Vb6Naming
{
    private static readonly HashSet<string> ReservedWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "And", "As", "Boolean", "ByRef", "Byte", "ByVal", "Call", "Case", "Class", "Close", "Const",
        "Currency", "Date", "Declare", "Dim", "Do", "Double", "Each", "Else", "ElseIf", "Empty",
        "End", "EndIf", "Enum", "Eqv", "Erase", "Error", "Event", "Exit", "False", "For", "Friend",
        "Function", "Get", "Global", "GoSub", "GoTo", "If", "Imp", "Implements", "In", "Input",
        "Integer", "Is", "Let", "Like", "LongLong", "Long", "LongPtr", "Loop", "Me", "Mod", "Name",
        "New", "Next", "Not", "Nothing", "Null", "Object", "On", "Open", "Option", "Optional", "Or",
        "ParamArray", "Preserve", "Print", "Private", "Property", "Public", "Put", "RaiseEvent",
        "Read", "ReadOnly", "ReDim", "Resume", "Return", "Select", "Set", "Single", "Static", "Step",
        "Stop", "String", "Sub", "Then", "Time", "To", "True", "Type", "TypeOf", "Until", "Variant",
        "Wend", "While", "With", "WithEvents", "Write", "Xor"
    };

    public static string PascalCase(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "X";
        var cleaned = new StringBuilder(raw.Length);
        bool nextUpper = true;
        foreach (var ch in raw)
        {
            if (char.IsLetterOrDigit(ch))
            {
                cleaned.Append(nextUpper ? char.ToUpper(ch, CultureInfo.InvariantCulture) : ch);
                nextUpper = false;
            }
            else
            {
                nextUpper = true;
            }
        }
        if (cleaned.Length == 0) return "X";
        if (char.IsDigit(cleaned[0])) cleaned.Insert(0, '_');
        return cleaned.ToString();
    }

    public static string CamelCase(string raw)
    {
        var pc = PascalCase(raw);
        if (pc.Length == 0) return pc;
        return char.ToLower(pc[0], CultureInfo.InvariantCulture) + pc[1..];
    }

    public static string SafeIdentifier(string raw)
    {
        var id = PascalCase(raw);
        if (ReservedWords.Contains(id)) id += "_";
        return id;
    }

    public static string SafeParameter(string raw)
    {
        var id = CamelCase(raw);
        if (ReservedWords.Contains(id)) id += "_";
        return id;
    }

    public static string PropertyName(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "X";
        var sb = new StringBuilder(raw.Length);
        bool upper = true;
        foreach (var ch in raw)
        {
            if (char.IsLetterOrDigit(ch))
            {
                sb.Append(upper ? char.ToUpper(ch, CultureInfo.InvariantCulture) : ch);
                upper = false;
            }
            else if (ch == '_')
            {
                sb.Append('_');
                upper = true;
            }
            else
            {
                upper = true;
            }
        }
        if (sb.Length == 0) return "X";
        if (char.IsDigit(sb[0])) sb.Insert(0, '_');
        var id = sb.ToString();
        if (ReservedWords.Contains(id)) id += "_";
        return id;
    }

    public static string ModuleName(string tag) => "mod" + PascalCase(tag) + "Client";

    public static string ClassName(string schemaName) => "c" + PascalCase(schemaName);

    public static string EnumName(string schemaName) => "e" + PascalCase(schemaName);

    private static readonly System.Text.RegularExpressions.Regex VersionSegment =
        new(@"^v[\d.]+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

    public static string OperationName(string operationId, string fallbackTag, string method, string path)
    {
        if (!string.IsNullOrWhiteSpace(operationId))
        {
            var idx = operationId.IndexOf('_');
            var trimmed = idx > 0 ? operationId[(idx + 1)..] : operationId;
            return PascalCase(trimmed);
        }

        var verb = method.ToUpperInvariant() switch
        {
            "GET" => "Get",
            "POST" => "Create",
            "PUT" => "Update",
            "DELETE" => "Delete",
            "PATCH" => "Patch",
            _ => PascalCase(method)
        };

        var segments = path.Split('/', StringSplitOptions.RemoveEmptyEntries)
            .Where(s => !VersionSegment.IsMatch(s))
            .ToList();

        var trailingParam = segments.Count > 0 && segments[^1].StartsWith('{') ? segments[^1] : null;
        var staticSegs = segments
            .Where(s => !s.StartsWith('{'))
            .Select(PascalCase)
            .ToList();
        var body = string.Concat(staticSegs);
        if (trailingParam is not null)
            body += "By" + PascalCase(trailingParam.Trim('{', '}'));

        if (body.Length == 0) body = PascalCase(fallbackTag);
        return verb + body;
    }
}
