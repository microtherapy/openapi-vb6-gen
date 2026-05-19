namespace OpenApiVb6Gen;

internal sealed class EnumEmitter
{
    public string Emit(IEnumerable<EnumModel> enums)
    {
        var w = new Vb6Writer();
        w.Line("VERSION 1.0 CLASS");
        w.Line("BEGIN");
        w.Line("  MultiUse = -1  'True");
        w.Line("  Persistable = 0  'NotPersistable");
        w.Line("  DataBindingBehavior = 0  'vbNone");
        w.Line("  DataSourceBehavior  = 0  'vbNone");
        w.Line("  MTSTransactionMode  = 0  'NotAnMTSObject");
        w.Line("END");
        w.Line("Attribute VB_Name = \"cEnums\"");
        w.Line("Attribute VB_GlobalNameSpace = False");
        w.Line("Attribute VB_Creatable = False");
        w.Line("Attribute VB_PredeclaredId = True");
        w.Line("Attribute VB_Exposed = True");
        w.Line("Option Explicit");
        w.Line();
        w.Comment("Public enums exposed as the DLL's typelib surface.");
        w.Comment("Hosted in a .cls (not .bas) because BAS modules in an OleDll are not exposed to clients.");
        w.Line();

        foreach (var e in enums)
        {
            if (!string.IsNullOrWhiteSpace(e.Description))
                w.Comment(e.Description!);
            w.Line($"Public Enum {e.EnumName}");
            w.Indent();
            if (e.Members.Count == 0)
                w.Line($"{e.EnumName}_None = 0");
            else
                foreach (var m in e.Members)
                    w.Line($"{e.EnumName}_{m.VbName} = {m.IntValue}");
            w.Outdent();
            w.Line("End Enum");
            w.Line();
        }
        return w.ToString();
    }
}
