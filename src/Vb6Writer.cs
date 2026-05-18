using System.Text;

namespace OpenApiVb6Gen;

internal sealed class Vb6Writer
{
    private readonly StringBuilder _sb = new();
    private int _indent;

    public int IndentSize { get; init; } = 4;

    public Vb6Writer Indent() { _indent++; return this; }
    public Vb6Writer Outdent() { if (_indent > 0) _indent--; return this; }

    public Vb6Writer Line(string text = "")
    {
        if (text.Length == 0) _sb.Append("\r\n");
        else _sb.Append(new string(' ', _indent * IndentSize)).Append(text).Append("\r\n");
        return this;
    }

    public Vb6Writer Comment(string text)
    {
        foreach (var line in text.ReplaceLineEndings("\n").Split('\n'))
            Line("' " + line);
        return this;
    }

    public Vb6Writer Block(string open, Action<Vb6Writer> body, string close)
    {
        Line(open);
        Indent();
        body(this);
        Outdent();
        Line(close);
        return this;
    }

    public override string ToString() => _sb.ToString();

    public void WriteTo(string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, _sb.ToString(), new UTF8Encoding(false));
    }

    public static string EscapeStringLiteral(string s) => "\"" + s.Replace("\"", "\"\"") + "\"";
}
