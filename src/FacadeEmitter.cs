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

        WriteUnlockChilkat(w);
        w.Line();
        WriteSaveBytesToFile(w);
        return w.ToString();
    }

    private static void WriteUnlockChilkat(Vb6Writer w)
    {
        w.Comment("Unlocks Chilkat 11 process-wide. Call once at app startup before any API call.");
        w.Comment("Returns True on success. On failure, the Chilkat global's LastErrorText is in the raised error.");
        w.Line("Public Function UnlockChilkat(ByVal licenseKey As String) As Boolean");
        w.Indent();
        w.Line("Dim g As ChilkatGlobal: Set g = New ChilkatGlobal");
        w.Line("If g.UnlockBundle(licenseKey) <> 1 Then");
        w.Indent().Line("Err.Raise vbObjectError + 513, \"cApi.UnlockChilkat\", g.LastErrorText").Outdent();
        w.Line("End If");
        w.Line("UnlockChilkat = True");
        w.Outdent();
        w.Line("End Function");
    }

    private static void WriteSaveBytesToFile(Vb6Writer w)
    {
        w.Comment("Writes a Variant byte array (as returned by binary endpoints) to disk.");
        w.Line("Public Sub SaveBytesToFile(ByVal data As Variant, ByVal path As String)");
        w.Indent();
        w.Line("Dim b() As Byte: b = data");
        w.Line("Dim fnum As Integer: fnum = FreeFile");
        w.Line("Open path For Binary Access Write As #fnum");
        w.Line("Put #fnum, , b");
        w.Line("Close #fnum");
        w.Outdent();
        w.Line("End Sub");
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
