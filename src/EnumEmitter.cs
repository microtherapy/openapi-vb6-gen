namespace OpenApiVb6Gen;

internal sealed class EnumEmitter
{
    public string Emit(EnumModel e)
    {
        var w = new Vb6Writer();
        if (!string.IsNullOrWhiteSpace(e.Description))
            w.Comment(e.Description!);
        w.Line($"Public Enum {e.EnumName}");
        w.Indent();
        foreach (var m in e.Members)
            w.Line($"{e.EnumName}_{m.VbName} = {m.IntValue}");
        if (e.Members.Count == 0)
            w.Line($"{e.EnumName}_None = 0");
        w.Outdent();
        w.Line("End Enum");
        return w.ToString();
    }
}
