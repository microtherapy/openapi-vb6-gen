namespace OpenApiVb6Gen;

internal sealed class DtoEmitter
{
    public string Emit(DtoModel dto)
    {
        var w = new Vb6Writer();
        WriteClsHeader(w, dto.ClassName);
        w.Line("Option Explicit");
        w.Line();
        if (!string.IsNullOrWhiteSpace(dto.Description))
        {
            w.Comment(dto.Description!);
            w.Line();
        }

        foreach (var p in dto.Properties)
            w.Line($"Private m{p.VbName} As {p.Type.Declaration}");
        w.Line();

        foreach (var p in dto.Properties)
        {
            WriteAccessor(w, p);
            w.Line();
        }

        WriteFromJson(w, dto);
        w.Line();
        WriteToJson(w, dto);
        w.Line();

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

    private static void WriteAccessor(Vb6Writer w, DtoPropertyModel p)
    {
        if (p.Type.IsCollection || p.Type.IsDtoRef || p.Type.Kind == Vb6Kind.ChilkatJsonObject)
        {
            w.Line($"Public Property Get {p.VbName}() As {p.Type.Declaration}");
            w.Indent().Line($"Set {p.VbName} = m{p.VbName}").Outdent();
            w.Line("End Property");
            w.Line($"Public Property Set {p.VbName}(ByVal v As {p.Type.Declaration})");
            w.Indent().Line($"Set m{p.VbName} = v").Outdent();
            w.Line("End Property");
        }
        else
        {
            w.Line($"Public Property Get {p.VbName}() As {p.Type.Declaration}");
            w.Indent().Line($"{p.VbName} = m{p.VbName}").Outdent();
            w.Line("End Property");
            w.Line($"Public Property Let {p.VbName}(ByVal v As {p.Type.Declaration})");
            w.Indent().Line($"m{p.VbName} = v").Outdent();
            w.Line("End Property");
        }
    }

    private static void WriteFromJson(Vb6Writer w, DtoModel dto)
    {
        w.Line("Public Sub FromJson(ByVal obj As ChilkatJsonObject)");
        w.Indent();
        w.Line("If obj Is Nothing Then Exit Sub");
        foreach (var p in dto.Properties)
            WriteFromJsonProp(w, p);
        w.Outdent();
        w.Line("End Sub");
    }

    private static void WriteFromJsonProp(Vb6Writer w, DtoPropertyModel p)
    {
        var jn = Vb6Writer.EscapeStringLiteral(p.JsonName);
        switch (p.Type.Kind)
        {
            case Vb6Kind.String:
                w.Line($"m{p.VbName} = obj.StringOf({jn})");
                break;
            case Vb6Kind.Long:
                w.Line($"m{p.VbName} = obj.IntOf({jn})");
                break;
            case Vb6Kind.Currency:
                w.Line($"m{p.VbName} = CCur(obj.IntOf({jn}))");
                break;
            case Vb6Kind.Double:
                w.Line($"m{p.VbName} = CDbl(Val(obj.StringOf({jn})))");
                break;
            case Vb6Kind.Boolean:
                w.Line($"m{p.VbName} = obj.BoolOf({jn})");
                break;
            case Vb6Kind.Date:
                w.Line($"m{p.VbName} = modGenApi.ParseIso(obj.StringOf({jn}))");
                break;
            case Vb6Kind.Enum:
                w.Line($"m{p.VbName} = obj.IntOf({jn})");
                break;
            case Vb6Kind.Variant:
                w.Line($"If obj.IsNullOf({jn}) Then");
                w.Indent().Line($"m{p.VbName} = Empty").Outdent();
                w.Line("Else");
                w.Indent().Line($"m{p.VbName} = obj.StringOf({jn})").Outdent();
                w.Line("End If");
                break;
            case Vb6Kind.DtoRef:
                w.Line($"Set m{p.VbName} = modGenApi.LoadDto_{p.Type.DtoClassName}(obj, {jn})");
                break;
            case Vb6Kind.Collection:
                EmitCollectionLoad(w, p, jn);
                break;
            case Vb6Kind.ChilkatJsonObject:
                w.Line($"Set m{p.VbName} = obj.ObjectOf({jn})");
                break;
        }
    }

    private static void EmitCollectionLoad(Vb6Writer w, DtoPropertyModel p, string jn)
    {
        var item = p.Type.ItemType!;
        var helper = item.Kind switch
        {
            Vb6Kind.DtoRef => $"LoadList_{item.DtoClassName}",
            Vb6Kind.String => "LoadList_String",
            Vb6Kind.Long => "LoadList_Long",
            Vb6Kind.Currency => "LoadList_Currency",
            Vb6Kind.Double => "LoadList_Double",
            Vb6Kind.Boolean => "LoadList_Boolean",
            Vb6Kind.Date => "LoadList_Date",
            Vb6Kind.Enum => "LoadList_Long",
            _ => "LoadList_Variant"
        };
        w.Line($"Set m{p.VbName} = modGenApi.{helper}(obj, {jn})");
    }

    private static void WriteToJson(Vb6Writer w, DtoModel dto)
    {
        w.Line("Public Function ToJson() As ChilkatJsonObject");
        w.Indent();
        w.Line("Set ToJson = New ChilkatJsonObject");
        foreach (var p in dto.Properties)
            WriteToJsonProp(w, p);
        w.Outdent();
        w.Line("End Function");
    }

    private static void WriteToJsonProp(Vb6Writer w, DtoPropertyModel p)
    {
        var jn = Vb6Writer.EscapeStringLiteral(p.JsonName);
        switch (p.Type.Kind)
        {
            case Vb6Kind.String:
                w.Line($"ToJson.UpdateString {jn}, m{p.VbName}");
                break;
            case Vb6Kind.Long:
            case Vb6Kind.Enum:
                w.Line($"ToJson.UpdateInt {jn}, m{p.VbName}");
                break;
            case Vb6Kind.Currency:
                w.Line($"ToJson.UpdateInt {jn}, CLng(m{p.VbName})");
                break;
            case Vb6Kind.Double:
                w.Line($"ToJson.UpdateNumber {jn}, CStr(m{p.VbName})");
                break;
            case Vb6Kind.Boolean:
                w.Line($"ToJson.UpdateBool {jn}, m{p.VbName}");
                break;
            case Vb6Kind.Date:
                w.Line($"ToJson.UpdateString {jn}, modGenApi.Iso(m{p.VbName})");
                break;
            case Vb6Kind.Variant:
                w.Line($"If IsEmpty(m{p.VbName}) Then");
                w.Indent().Line($"ToJson.UpdateNull {jn}").Outdent();
                w.Line("Else");
                w.Indent().Line($"ToJson.UpdateString {jn}, CStr(m{p.VbName})").Outdent();
                w.Line("End If");
                break;
            case Vb6Kind.DtoRef:
                w.Line($"modGenApi.AppendDto_{p.Type.DtoClassName} ToJson, {jn}, m{p.VbName}");
                break;
            case Vb6Kind.Collection:
                EmitCollectionSave(w, p, jn);
                break;
            case Vb6Kind.ChilkatJsonObject:
                w.Line($"If Not m{p.VbName} Is Nothing Then");
                w.Indent();
                w.Line($"ToJson.AddObjectAt -1, {jn}");
                w.Line($"ToJson.ObjectOf({jn}).Load m{p.VbName}.Emit()");
                w.Outdent();
                w.Line("End If");
                break;
        }
    }

    private static void EmitCollectionSave(Vb6Writer w, DtoPropertyModel p, string jn)
    {
        var item = p.Type.ItemType!;
        var helper = item.Kind switch
        {
            Vb6Kind.DtoRef => $"AppendList_{item.DtoClassName}",
            Vb6Kind.String => "AppendList_String",
            Vb6Kind.Long => "AppendList_Long",
            Vb6Kind.Currency => "AppendList_Currency",
            Vb6Kind.Double => "AppendList_Double",
            Vb6Kind.Boolean => "AppendList_Boolean",
            Vb6Kind.Date => "AppendList_Date",
            Vb6Kind.Enum => "AppendList_Long",
            _ => "AppendList_Variant"
        };
        w.Line($"modGenApi.{helper} ToJson, {jn}, m{p.VbName}");
    }
}
