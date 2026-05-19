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

    private static IReadOnlyDictionary<string, string>? _schemaAliases;

    public static void SetSchemaAliases(IReadOnlyDictionary<string, string>? map) => _schemaAliases = map;

    private static string Alias(string schemaName)
        => _schemaAliases is not null && _schemaAliases.TryGetValue(schemaName, out var v) ? v : schemaName;

    public static string ClassName(string schemaName) => "c" + PascalCase(Alias(schemaName));

    public static string EnumName(string schemaName) => "e" + PascalCase(Alias(schemaName));

    public static IReadOnlyDictionary<string, string> BuildSchemaAliasMap(IEnumerable<string> rawNames, int maxAliasLen = int.MaxValue)
    {
        var names = rawNames.Distinct().ToList();
        if (names.Count == 0) return new Dictionary<string, string>();

        var parts = names.ToDictionary(n => n, SplitNamespace);
        var depth = names.ToDictionary(n => n, _ => 1);

        Dictionary<string, string> aliases = names.ToDictionary(n => n, _ => "");
        for (int iter = 0; iter < 32; iter++)
        {
            aliases = names.ToDictionary(n => n, n => TakeLastSegments(parts[n], depth[n]));
            var collisions = aliases
                .GroupBy(kv => PascalCase(kv.Value), StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1)
                .ToList();
            if (collisions.Count == 0) break;

            bool advanced = false;
            foreach (var g in collisions)
                foreach (var kv in g)
                    if (depth[kv.Key] < parts[kv.Key].Count) { depth[kv.Key]++; advanced = true; }

            if (!advanced) break;
        }

        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in aliases.Keys.OrderBy(k => k, StringComparer.Ordinal).ToList())
        {
            var pc = PascalCase(aliases[key]);
            if (pc.Length > maxAliasLen) pc = pc[..maxAliasLen];
            if (used.Contains(pc))
            {
                for (int n = 2; n < 10000; n++)
                {
                    var suf = n.ToString();
                    var trim = Math.Min(pc.Length, Math.Max(1, maxAliasLen - suf.Length));
                    var cand = pc[..trim] + suf;
                    if (!used.Contains(cand)) { pc = cand; break; }
                }
            }
            aliases[key] = pc;
            used.Add(pc);
        }
        return aliases;
    }

    private static List<string> SplitNamespace(string raw)
        => raw.Split(new[] { '.', '+', '/' }, StringSplitOptions.RemoveEmptyEntries).ToList();

    private static string TakeLastSegments(List<string> parts, int count)
    {
        if (parts.Count == 0) return "X";
        var n = Math.Min(count, parts.Count);
        return string.Join("_", parts.Skip(parts.Count - n));
    }

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
