namespace OpenApiVb6Gen;

internal sealed class FacadeEmitter
{
    public string Emit(IReadOnlyList<ControllerModel> controllers)
    {
        var w = new Vb6Writer();
        WriteClsHeader(w, "cApi");
        w.Line("Option Explicit");
        w.Line();
        w.Comment("Facade: the only class host code New's directly. Init must be called before any controller property.");
        w.Line();
        w.Line("Public BaseURL As String");
        w.Line("Public BearerToken As String");
        w.Line();
        foreach (var c in controllers)
            w.Line($"Private m{c.PropertyName} As {c.ClassName}");
        w.Line();
        w.Line("Public Sub Init(ByVal baseUrl As String, ByVal bearerToken As String)");
        w.Indent();
        w.Line("Me.BaseURL = modGenApi.NormalizeBaseUrl(baseUrl)");
        w.Line("Me.BearerToken = bearerToken");
        foreach (var c in controllers)
        {
            w.Line($"Set m{c.PropertyName} = New {c.ClassName}");
            w.Line($"m{c.PropertyName}.Init Me");
        }
        w.Outdent();
        w.Line("End Sub");
        w.Line();
        foreach (var c in controllers)
        {
            w.Line($"Public Property Get {c.PropertyName}() As {c.ClassName}");
            w.Indent().Line($"Set {c.PropertyName} = m{c.PropertyName}").Outdent();
            w.Line("End Property");
            w.Line();
        }
        return w.ToString();
    }

    private static void WriteClsHeader(Vb6Writer w, string className)
    {
        w.Line("VERSION 1.0 CLASS");
        w.Line("BEGIN");
        w.Line("  MultiUse = -1  'True");
        w.Line("  Persistable = 0  'NotPersistable");
        w.Line("  DataBindingBehavior = 0  'vbNone");
        w.Line("  DataSourceBehavior  = 0  'vbNone");
        w.Line("  MTSTransactionMode  = 0  'NotAnMTSObject");
        w.Line("END");
        w.Line($"Attribute VB_Name = \"{className}\"");
        w.Line("Attribute VB_GlobalNameSpace = False");
        w.Line("Attribute VB_Creatable = True");
        w.Line("Attribute VB_PredeclaredId = False");
        w.Line("Attribute VB_Exposed = True");
    }
}
