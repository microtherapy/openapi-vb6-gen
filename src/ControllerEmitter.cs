using System.Text;

namespace OpenApiVb6Gen;

internal sealed class ControllerEmitter
{
    public string Emit(ControllerModel controller)
    {
        var w = new Vb6Writer();
        WriteClsHeader(w, controller.ClassName);
        w.Line("Option Explicit");
        w.Line();
        w.Line("Private mApi As cApi");
        w.Line();
        w.Line("Friend Sub Init(ByVal parent As cApi)");
        w.Indent().Line("Set mApi = parent").Outdent();
        w.Line("End Sub");
        w.Line();

        foreach (var op in controller.Operations)
        {
            WriteOperation(w, op);
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

    private static void WriteOperation(Vb6Writer w, OperationModel op)
    {
        w.Comment($"{op.HttpMethod.ToUpperInvariant()} {op.PathTemplate}");
        if (!string.IsNullOrWhiteSpace(op.Description))
            w.Comment(op.Description!);

        if (!string.IsNullOrWhiteSpace(op.SkipReason))
        {
            w.Comment("SKIPPED: " + op.SkipReason!);
            w.Line($"' Public ... ' {op.VbMethodName}");
            return;
        }

        var paramList = BuildParamList(op);
        var (isFunction, signature) = BuildSignature(op, paramList);

        w.Line(signature);
        w.Indent();
        w.Line("Dim url As String");
        EmitUrlBuild(w, op);
        EmitCall(w, op, isFunction);
        w.Outdent();
        w.Line(isFunction ? "End Function" : "End Sub");
    }

    private static string BuildParamList(OperationModel op)
    {
        var sb = new StringBuilder();
        bool first = true;
        foreach (var p in op.PathParameters.Concat(op.QueryParameters))
        {
            if (!first) sb.Append(", ");
            first = false;
            var byRef = p.Type.IsCollection || p.Type.IsDtoRef || p.Type.Kind == Vb6Kind.ChilkatJsonObject
                ? "ByVal"
                : "ByVal";
            sb.Append($"{byRef} {p.VbName} As {p.Type.Declaration}");
        }
        if (op.Body is not null)
        {
            if (!first) sb.Append(", ");
            sb.Append($"ByVal body As {op.Body.Type.Declaration}");
        }
        return sb.ToString();
    }

    private static (bool isFunction, string signature) BuildSignature(OperationModel op, string paramList)
    {
        if (op.Response is null)
            return (false, $"Public Sub {op.VbMethodName}({paramList})");
        return (true, $"Public Function {op.VbMethodName}({paramList}) As {op.Response.Declaration}");
    }

    private static void EmitUrlBuild(Vb6Writer w, OperationModel op)
    {
        var template = op.PathTemplate.TrimStart('/');
        var sb = new StringBuilder();
        int i = 0;
        sb.Append("url = mApi.BaseURL & ");
        bool firstPart = true;
        while (i < template.Length)
        {
            var open = template.IndexOf('{', i);
            if (open < 0)
            {
                var rest = template[i..];
                if (rest.Length > 0)
                {
                    if (!firstPart) sb.Append(" & ");
                    sb.Append(Vb6Writer.EscapeStringLiteral(rest));
                    firstPart = false;
                }
                break;
            }
            if (open > i)
            {
                if (!firstPart) sb.Append(" & ");
                sb.Append(Vb6Writer.EscapeStringLiteral(template[i..open]));
                firstPart = false;
            }
            var close = template.IndexOf('}', open);
            var pname = template[(open + 1)..close];
            var match = op.PathParameters.FirstOrDefault(p => p.ApiName.Equals(pname, StringComparison.OrdinalIgnoreCase));
            var vbName = match?.VbName ?? Vb6Naming.SafeParameter(pname);
            if (!firstPart) sb.Append(" & ");
            sb.Append($"modGenApi.PathSeg({vbName})");
            firstPart = false;
            i = close + 1;
        }
        if (firstPart) sb.Append(Vb6Writer.EscapeStringLiteral(""));
        w.Line(sb.ToString());

        foreach (var q in op.QueryParameters)
            EmitQueryParam(w, q);
    }

    private static void EmitQueryParam(Vb6Writer w, ParameterModel q)
    {
        var jn = Vb6Writer.EscapeStringLiteral(q.ApiName);
        if (q.Type.IsCollection)
        {
            var loopVar = "qsItem_" + q.VbName;
            w.Line($"If Not {q.VbName} Is Nothing Then");
            w.Indent();
            w.Line($"Dim {loopVar} As Variant");
            w.Line($"For Each {loopVar} In {q.VbName}");
            w.Indent();
            w.Line($"modGenApi.AppendQS url, {jn}, {loopVar}");
            w.Outdent();
            w.Line($"Next {loopVar}");
            w.Outdent();
            w.Line("End If");
        }
        else
        {
            w.Line($"modGenApi.AppendQS url, {jn}, {q.VbName}");
        }
    }

    private static void EmitCall(Vb6Writer w, OperationModel op, bool isFunction)
    {
        var verb = op.HttpMethod.ToUpperInvariant();
        var resp = op.Response;

        switch (verb)
        {
            case "GET":
                EmitGet(w, op, resp, isFunction);
                break;
            case "POST":
                EmitPost(w, op, resp, isFunction);
                break;
            case "PUT":
                EmitPut(w, op, resp, isFunction);
                break;
            case "DELETE":
                EmitDelete(w, op, resp, isFunction);
                break;
            case "PATCH":
                EmitPatch(w, op, resp, isFunction);
                break;
            default:
                w.Comment($"Unhandled HTTP method: {verb}");
                if (isFunction) w.Line($"' return default of {resp!.Declaration}");
                break;
        }
    }

    private static string ResponseAssignPrefix(Vb6Type resp)
        => resp.Kind is Vb6Kind.DtoRef or Vb6Kind.Collection or Vb6Kind.ChilkatJsonObject
            ? "Set "
            : "";

    private static void EmitGet(Vb6Writer w, OperationModel op, Vb6Type? resp, bool isFunction)
    {
        if (resp is null) { w.Line("modGenApi.GetVoid mApi, url"); return; }
        w.Line($"{ResponseAssignPrefix(resp)}{op.VbMethodName} = {GetCall(resp)}");
    }

    private static string GetCall(Vb6Type resp) => resp.Kind switch
    {
        Vb6Kind.DtoRef => $"modGenApi.GetJsonAs_{resp.DtoClassName}(mApi, url)",
        Vb6Kind.Collection when resp.ItemType!.IsDtoRef =>
            $"modGenApi.GetJsonArrayAs_{resp.ItemType.DtoClassName}(mApi, url)",
        Vb6Kind.Collection => $"modGenApi.GetJsonArrayAs_{ItemSuffix(resp.ItemType!)}(mApi, url)",
        Vb6Kind.String => "modGenApi.GetString(mApi, url)",
        Vb6Kind.Long or Vb6Kind.Enum => "modGenApi.GetLong(mApi, url)",
        Vb6Kind.Currency => "CCur(modGenApi.GetDouble(mApi, url))",
        Vb6Kind.Double => "modGenApi.GetDouble(mApi, url)",
        Vb6Kind.Boolean => "modGenApi.GetBool(mApi, url)",
        Vb6Kind.Date => "modGenApi.ParseIso(modGenApi.GetString(mApi, url))",
        Vb6Kind.ChilkatJsonObject => "modGenApi.GetJsonObject(mApi, url)",
        Vb6Kind.Binary => "modGenApi.GetBytes(mApi, url)",
        _ => "modGenApi.GetString(mApi, url)"
    };

    private static string ItemSuffix(Vb6Type item) => item.Kind switch
    {
        Vb6Kind.String => "String",
        Vb6Kind.Long or Vb6Kind.Enum => "Long",
        Vb6Kind.Currency => "Currency",
        Vb6Kind.Double => "Double",
        Vb6Kind.Boolean => "Boolean",
        Vb6Kind.Date => "Date",
        _ => "Variant"
    };

    private static void EmitPost(Vb6Writer w, OperationModel op, Vb6Type? resp, bool isFunction)
    {
        var bodyArg = op.Body is null ? "Nothing" : BodyArg(op.Body);
        if (resp is null) { w.Line($"modGenApi.PostJsonVoid mApi, url, {bodyArg}"); return; }
        w.Line($"{ResponseAssignPrefix(resp)}{op.VbMethodName} = {PostCall(resp, bodyArg)}");
    }

    private static string PostCall(Vb6Type resp, string bodyArg) => resp.Kind switch
    {
        Vb6Kind.DtoRef => $"modGenApi.PostJsonAs_{resp.DtoClassName}(mApi, url, {bodyArg})",
        Vb6Kind.Collection when resp.ItemType!.IsDtoRef =>
            $"modGenApi.PostJsonArrayAs_{resp.ItemType.DtoClassName}(mApi, url, {bodyArg})",
        Vb6Kind.String => $"modGenApi.PostJsonReturnString(mApi, url, {bodyArg})",
        Vb6Kind.Long or Vb6Kind.Enum => $"modGenApi.PostJsonReturnLong(mApi, url, {bodyArg})",
        Vb6Kind.ChilkatJsonObject => $"modGenApi.PostJsonReturnObject(mApi, url, {bodyArg})",
        Vb6Kind.Binary => $"modGenApi.PostJsonReturnBytes(mApi, url, {bodyArg})",
        _ => $"modGenApi.PostJsonReturnString(mApi, url, {bodyArg})"
    };

    private static void EmitPut(Vb6Writer w, OperationModel op, Vb6Type? resp, bool isFunction)
    {
        var bodyArg = op.Body is null ? "Nothing" : BodyArg(op.Body);
        if (resp is null) { w.Line($"modGenApi.PutJsonVoid mApi, url, {bodyArg}"); return; }
        if (resp.Kind == Vb6Kind.Binary)
        {
            w.Line($"' PUT with binary response body ({resp.Declaration}) not supported by generated helpers");
            return;
        }
        w.Line($"{ResponseAssignPrefix(resp)}{op.VbMethodName} = {PutCall(resp, bodyArg)}");
    }

    private static string PutCall(Vb6Type resp, string bodyArg) => resp.Kind switch
    {
        Vb6Kind.DtoRef => $"modGenApi.PutJsonAs_{resp.DtoClassName}(mApi, url, {bodyArg})",
        Vb6Kind.Collection when resp.ItemType!.IsDtoRef =>
            $"modGenApi.PutJsonArrayAs_{resp.ItemType.DtoClassName}(mApi, url, {bodyArg})",
        Vb6Kind.String => $"modGenApi.PutJsonReturnString(mApi, url, {bodyArg})",
        Vb6Kind.Long or Vb6Kind.Enum => $"modGenApi.PutJsonReturnLong(mApi, url, {bodyArg})",
        Vb6Kind.ChilkatJsonObject => $"modGenApi.PutJsonReturnObject(mApi, url, {bodyArg})",
        _ => $"modGenApi.PutJsonReturnString(mApi, url, {bodyArg})"
    };

    private static void EmitDelete(Vb6Writer w, OperationModel op, Vb6Type? resp, bool isFunction)
    {
        if (resp is null) { w.Line("modGenApi.DeleteResource mApi, url"); return; }
        if (resp.Kind is Vb6Kind.String or Vb6Kind.Long or Vb6Kind.Enum or Vb6Kind.Currency or Vb6Kind.Double or Vb6Kind.Boolean or Vb6Kind.Date)
            w.Line($"{op.VbMethodName} = modGenApi.DeleteReturnString(mApi, url)");
        else
            w.Line($"' DELETE with typed response body ({resp.Declaration}) not supported by generated helpers");
    }

    private static void EmitPatch(Vb6Writer w, OperationModel op, Vb6Type? resp, bool isFunction)
    {
        var bodyArg = op.Body is null ? "Nothing" : BodyArg(op.Body);
        w.Comment("PATCH support is minimal; treated like PUT.");
        if (resp is null) { w.Line($"modGenApi.PutJsonVoid mApi, url, {bodyArg}"); return; }
        w.Line($"{ResponseAssignPrefix(resp)}{op.VbMethodName} = {PutCall(resp, bodyArg)}");
    }

    private static string BodyArg(ParameterModel body) => body.Type.Kind switch
    {
        Vb6Kind.DtoRef => "body.ToJson()",
        Vb6Kind.ChilkatJsonObject => "body",
        _ => "modGenApi.WrapScalarJson(body)"
    };
}
