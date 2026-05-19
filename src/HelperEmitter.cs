using System.Text;

namespace OpenApiVb6Gen;

internal sealed class HelperEmitter
{
    public string Emit(IEnumerable<DtoModel> dtos)
    {
        var w = new Vb6Writer();
        w.Line("Attribute VB_Name = \"modGenApi\"");
        w.Line("Option Explicit");
        w.Line();
        w.Comment("Generated HTTP / JSON helpers. Not for external use.");
        w.Comment("Uses Chilkat ActiveX for HTTP and JSON.");
        w.Line();

        WriteConstants(w);
        w.Line();
        WriteUrlHelpers(w);
        w.Line();
        WriteHttpCore(w);
        w.Line();
        WriteScalarGetters(w);
        w.Line();
        WritePostPutDelete(w);
        w.Line();
        WriteBinaryHelpers(w);
        w.Line();
        WriteDtoLoaders(w, dtos);
        w.Line();
        WritePrimitiveListHelpers(w);
        w.Line();
        WriteScalarWrap(w);
        w.Line();
        WriteDateHelpers(w);
        return w.ToString();
    }

    private static void WriteConstants(Vb6Writer w)
    {
        w.Line("Private Const gcContentType As String = \"application/json\"");
        w.Line("Private Const gcCharSet As String = \"utf-8\"");
    }

    private static void WriteUrlHelpers(Vb6Writer w)
    {
        w.Line("Public Function NormalizeBaseUrl(ByVal s As String) As String");
        w.Indent();
        w.Line("If Len(s) = 0 Then NormalizeBaseUrl = \"\": Exit Function");
        w.Line("If Right$(s, 1) = \"/\" Then NormalizeBaseUrl = s Else NormalizeBaseUrl = s & \"/\"");
        w.Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Public Function PathSeg(ByVal v As Variant) As String");
        w.Indent();
        w.Line("PathSeg = UrlEncode(CStr(v))");
        w.Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Public Sub AppendQS(ByRef url As String, ByVal name As String, ByVal v As Variant)");
        w.Indent();
        w.Line("Dim s As String");
        w.Line("If IsMissing(v) Or IsEmpty(v) Then Exit Sub");
        w.Line("If VarType(v) = vbDate Then s = Iso(CDate(v)) Else s = CStr(v)");
        w.Line("If Len(s) = 0 Then Exit Sub");
        w.Line("Dim sep As String: If InStr(url, \"?\") > 0 Then sep = \"&\" Else sep = \"?\"");
        w.Line("url = url & sep & UrlEncode(name) & \"=\" & UrlEncode(s)");
        w.Outdent();
        w.Line("End Sub");
        w.Line();

        w.Line("Private Function UrlEncode(ByVal s As String) As String");
        w.Indent();
        w.Line("Dim i As Long, ch As String, code As Integer, out As String");
        w.Line("For i = 1 To Len(s)");
        w.Indent();
        w.Line("ch = Mid$(s, i, 1)");
        w.Line("code = AscW(ch)");
        w.Line("If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) _");
        w.Indent();
        w.Line("Or ch = \"-\" Or ch = \"_\" Or ch = \".\" Or ch = \"~\" Then");
        w.Outdent();
        w.Indent().Line("out = out & ch").Outdent();
        w.Line("Else");
        w.Indent().Line("out = out & \"%\" & Right$(\"00\" & Hex$(code), 2)").Outdent();
        w.Line("End If");
        w.Outdent();
        w.Line("Next i");
        w.Line("UrlEncode = out");
        w.Outdent();
        w.Line("End Function");
    }

    private static void WriteHttpCore(Vb6Writer w)
    {
        w.Line("Private Function NewHttp(ByVal api As cApi) As ChilkatHttp");
        w.Indent();
        w.Line("Set NewHttp = New ChilkatHttp");
        w.Line("NewHttp.ConnectTimeout = 10");
        w.Line("If Len(api.BearerToken) > 0 Then NewHttp.SetRequestHeader \"Authorization\", \"bearer \" & api.BearerToken");
        w.Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Private Sub EnsureOk(ByVal http As ChilkatHttp, ByVal context As String)");
        w.Indent();
        w.Line("If http.LastStatus >= 400 Or http.LastStatus = 0 Then");
        w.Indent();
        w.Line("Err.Raise vbObjectError + http.LastStatus, context, _");
        w.Indent().Line("\"HTTP \" & http.LastStatus & \": \" & http.LastStatusText").Outdent();
        w.Outdent();
        w.Line("End If");
        w.Outdent();
        w.Line("End Sub");
        w.Line();

        w.Line("Private Sub EnsureRespOk(ByVal resp As ChilkatHttpResponse, ByVal context As String)");
        w.Indent();
        w.Line("If resp Is Nothing Then Err.Raise vbObjectError + 1, context, \"null response\"");
        w.Line("If resp.StatusCode >= 400 Then _");
        w.Indent().Line("Err.Raise vbObjectError + resp.StatusCode, context, \"HTTP \" & resp.StatusCode & \": \" & resp.StatusText").Outdent();
        w.Outdent();
        w.Line("End Sub");
    }

    private static void WriteScalarGetters(Vb6Writer w)
    {
        w.Line("Public Function GetString(ByVal api As cApi, ByVal url As String) As String");
        w.Indent();
        w.Line("Dim http As ChilkatHttp: Set http = NewHttp(api)");
        w.Line("GetString = http.QuickGetStr(url)");
        w.Line("EnsureOk http, \"GET \" & url");
        w.Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Public Function GetLong(ByVal api As cApi, ByVal url As String) As Long");
        w.Indent().Line("GetLong = CLng(Val(GetString(api, url)))").Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Public Function GetDouble(ByVal api As cApi, ByVal url As String) As Double");
        w.Indent().Line("GetDouble = CDbl(Val(GetString(api, url)))").Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Public Function GetBool(ByVal api As cApi, ByVal url As String) As Boolean");
        w.Indent().Line("GetBool = (LCase$(GetString(api, url)) = \"true\")").Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Public Sub GetVoid(ByVal api As cApi, ByVal url As String)");
        w.Indent().Line("Dim s As String: s = GetString(api, url)").Outdent();
        w.Line("End Sub");
        w.Line();

        w.Line("Public Function GetJsonObject(ByVal api As cApi, ByVal url As String) As ChilkatJsonObject");
        w.Indent();
        w.Line("Dim s As String: s = GetString(api, url)");
        w.Line("Set GetJsonObject = New ChilkatJsonObject");
        w.Line("If Len(s) > 0 Then GetJsonObject.Load s");
        w.Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Public Function GetJsonArray(ByVal api As cApi, ByVal url As String) As ChilkatJsonArray");
        w.Indent();
        w.Line("Dim s As String: s = GetString(api, url)");
        w.Line("Set GetJsonArray = New ChilkatJsonArray");
        w.Line("If Len(s) > 0 Then GetJsonArray.Load s");
        w.Outdent();
        w.Line("End Function");
    }

    private static void WritePostPutDelete(Vb6Writer w)
    {
        w.Line("Public Function PostJson(ByVal api As cApi, ByVal url As String, ByVal body As ChilkatJsonObject) As ChilkatJsonObject");
        w.Indent();
        w.Line("Dim http As ChilkatHttp: Set http = NewHttp(api)");
        w.Line("Dim jsonStr As String: If Not body Is Nothing Then jsonStr = body.Emit() Else jsonStr = \"\"");
        w.Line("Dim resp As ChilkatHttpResponse");
        w.Line("Set resp = http.PostJson2(url, gcContentType, jsonStr)");
        w.Line("EnsureRespOk resp, \"POST \" & url");
        w.Line("Set PostJson = New ChilkatJsonObject");
        w.Line("If Len(resp.BodyStr) > 0 Then PostJson.Load resp.BodyStr");
        w.Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Public Function PostJsonReturnString(ByVal api As cApi, ByVal url As String, ByVal body As ChilkatJsonObject) As String");
        w.Indent();
        w.Line("Dim http As ChilkatHttp: Set http = NewHttp(api)");
        w.Line("Dim jsonStr As String: If Not body Is Nothing Then jsonStr = body.Emit() Else jsonStr = \"\"");
        w.Line("Dim resp As ChilkatHttpResponse");
        w.Line("Set resp = http.PostJson2(url, gcContentType, jsonStr)");
        w.Line("EnsureRespOk resp, \"POST \" & url");
        w.Line("PostJsonReturnString = resp.BodyStr");
        w.Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Public Function PostJsonReturnLong(ByVal api As cApi, ByVal url As String, ByVal body As ChilkatJsonObject) As Long");
        w.Indent().Line("PostJsonReturnLong = CLng(Val(PostJsonReturnString(api, url, body)))").Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Public Function PostJsonReturnObject(ByVal api As cApi, ByVal url As String, ByVal body As ChilkatJsonObject) As ChilkatJsonObject");
        w.Indent().Line("Set PostJsonReturnObject = PostJson(api, url, body)").Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Public Sub PostJsonVoid(ByVal api As cApi, ByVal url As String, ByVal body As ChilkatJsonObject)");
        w.Indent().Line("Dim s As String: s = PostJsonReturnString(api, url, body)").Outdent();
        w.Line("End Sub");
        w.Line();

        w.Line("Public Function PutJsonReturnString(ByVal api As cApi, ByVal url As String, ByVal body As ChilkatJsonObject) As String");
        w.Indent();
        w.Line("Dim http As ChilkatHttp: Set http = NewHttp(api)");
        w.Line("Dim jsonStr As String: If Not body Is Nothing Then jsonStr = body.Emit() Else jsonStr = \"\"");
        w.Line("PutJsonReturnString = http.PutText(url, jsonStr, gcCharSet, gcContentType, 0, 0)");
        w.Line("EnsureOk http, \"PUT \" & url");
        w.Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Public Function PutJson(ByVal api As cApi, ByVal url As String, ByVal body As ChilkatJsonObject) As ChilkatJsonObject");
        w.Indent();
        w.Line("Dim s As String: s = PutJsonReturnString(api, url, body)");
        w.Line("Set PutJson = New ChilkatJsonObject");
        w.Line("If Len(s) > 0 Then PutJson.Load s");
        w.Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Public Function PutJsonReturnLong(ByVal api As cApi, ByVal url As String, ByVal body As ChilkatJsonObject) As Long");
        w.Indent().Line("PutJsonReturnLong = CLng(Val(PutJsonReturnString(api, url, body)))").Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Public Function PutJsonReturnObject(ByVal api As cApi, ByVal url As String, ByVal body As ChilkatJsonObject) As ChilkatJsonObject");
        w.Indent().Line("Set PutJsonReturnObject = PutJson(api, url, body)").Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Public Sub PutJsonVoid(ByVal api As cApi, ByVal url As String, ByVal body As ChilkatJsonObject)");
        w.Indent().Line("Dim s As String: s = PutJsonReturnString(api, url, body)").Outdent();
        w.Line("End Sub");
        w.Line();

        w.Line("Public Sub DeleteResource(ByVal api As cApi, ByVal url As String)");
        w.Indent();
        w.Line("Dim http As ChilkatHttp: Set http = NewHttp(api)");
        w.Line("Dim resp As ChilkatHttpResponse");
        w.Line("Set resp = http.QuickRequest(\"DELETE\", url)");
        w.Line("EnsureRespOk resp, \"DELETE \" & url");
        w.Outdent();
        w.Line("End Sub");
        w.Line();

        w.Line("Public Function DeleteReturnString(ByVal api As cApi, ByVal url As String) As String");
        w.Indent();
        w.Line("Dim http As ChilkatHttp: Set http = NewHttp(api)");
        w.Line("Dim resp As ChilkatHttpResponse");
        w.Line("Set resp = http.QuickRequest(\"DELETE\", url)");
        w.Line("EnsureRespOk resp, \"DELETE \" & url");
        w.Line("DeleteReturnString = resp.BodyStr");
        w.Outdent();
        w.Line("End Function");
    }

    private static void WriteDtoLoaders(Vb6Writer w, IEnumerable<DtoModel> dtos)
    {
        foreach (var dto in dtos)
        {
            var cls = dto.ClassName;
            w.Line($"Public Function GetJsonAs_{cls}(ByVal api As cApi, ByVal url As String) As {cls}");
            w.Indent();
            w.Line("Dim obj As ChilkatJsonObject: Set obj = GetJsonObject(api, url)");
            w.Line($"Set GetJsonAs_{cls} = New {cls}");
            w.Line($"GetJsonAs_{cls}.FromJson obj");
            w.Outdent();
            w.Line("End Function");
            w.Line();

            w.Line($"Public Function GetJsonArrayAs_{cls}(ByVal api As cApi, ByVal url As String) As Collection");
            w.Indent();
            w.Line("Dim arr As ChilkatJsonArray: Set arr = GetJsonArray(api, url)");
            w.Line($"Set GetJsonArrayAs_{cls} = New Collection");
            w.Line("Dim i As Long, n As Long: n = arr.Size");
            w.Line("For i = 0 To n - 1");
            w.Indent();
            w.Line($"Dim item As {cls}: Set item = New {cls}");
            w.Line("item.FromJson arr.ObjectAt(i)");
            w.Line($"GetJsonArrayAs_{cls}.Add item");
            w.Outdent();
            w.Line("Next i");
            w.Outdent();
            w.Line("End Function");
            w.Line();

            w.Line($"Public Function PostJsonAs_{cls}(ByVal api As cApi, ByVal url As String, ByVal body As ChilkatJsonObject) As {cls}");
            w.Indent();
            w.Line("Dim obj As ChilkatJsonObject: Set obj = PostJson(api, url, body)");
            w.Line($"Set PostJsonAs_{cls} = New {cls}");
            w.Line($"PostJsonAs_{cls}.FromJson obj");
            w.Outdent();
            w.Line("End Function");
            w.Line();

            w.Line($"Public Function PostJsonArrayAs_{cls}(ByVal api As cApi, ByVal url As String, ByVal body As ChilkatJsonObject) As Collection");
            w.Indent();
            w.Line("Dim http As ChilkatHttp: Set http = NewHttp(api)");
            w.Line("Dim jsonStr As String: If Not body Is Nothing Then jsonStr = body.Emit() Else jsonStr = \"\"");
            w.Line("Dim resp As ChilkatHttpResponse");
            w.Line("Set resp = http.PostJson2(url, gcContentType, jsonStr)");
            w.Line("EnsureRespOk resp, \"POST \" & url");
            w.Line("Dim arr As New ChilkatJsonArray: arr.Load resp.BodyStr");
            w.Line($"Set PostJsonArrayAs_{cls} = New Collection");
            w.Line("Dim i As Long, n As Long: n = arr.Size");
            w.Line("For i = 0 To n - 1");
            w.Indent();
            w.Line($"Dim item As {cls}: Set item = New {cls}");
            w.Line("item.FromJson arr.ObjectAt(i)");
            w.Line($"PostJsonArrayAs_{cls}.Add item");
            w.Outdent();
            w.Line("Next i");
            w.Outdent();
            w.Line("End Function");
            w.Line();

            w.Line($"Public Function PutJsonAs_{cls}(ByVal api As cApi, ByVal url As String, ByVal body As ChilkatJsonObject) As {cls}");
            w.Indent();
            w.Line("Dim obj As ChilkatJsonObject: Set obj = PutJson(api, url, body)");
            w.Line($"Set PutJsonAs_{cls} = New {cls}");
            w.Line($"PutJsonAs_{cls}.FromJson obj");
            w.Outdent();
            w.Line("End Function");
            w.Line();

            w.Line($"Public Function PutJsonArrayAs_{cls}(ByVal api As cApi, ByVal url As String, ByVal body As ChilkatJsonObject) As Collection");
            w.Indent();
            w.Line("Dim s As String: s = PutJsonReturnString(api, url, body)");
            w.Line("Dim arr As New ChilkatJsonArray: If Len(s) > 0 Then arr.Load s");
            w.Line($"Set PutJsonArrayAs_{cls} = New Collection");
            w.Line("Dim i As Long, n As Long: n = arr.Size");
            w.Line("For i = 0 To n - 1");
            w.Indent();
            w.Line($"Dim item As {cls}: Set item = New {cls}");
            w.Line("item.FromJson arr.ObjectAt(i)");
            w.Line($"PutJsonArrayAs_{cls}.Add item");
            w.Outdent();
            w.Line("Next i");
            w.Outdent();
            w.Line("End Function");
            w.Line();

            w.Line($"Public Function LoadDto_{cls}(ByVal parent As ChilkatJsonObject, ByVal propName As String) As {cls}");
            w.Indent();
            w.Line("If parent Is Nothing Then Exit Function");
            w.Line("If parent.IsNullOf(propName) Then Exit Function");
            w.Line("Dim child As ChilkatJsonObject: Set child = parent.ObjectOf(propName)");
            w.Line("If child Is Nothing Then Exit Function");
            w.Line($"Set LoadDto_{cls} = New {cls}");
            w.Line($"LoadDto_{cls}.FromJson child");
            w.Outdent();
            w.Line("End Function");
            w.Line();

            w.Line($"Public Sub AppendDto_{cls}(ByVal parent As ChilkatJsonObject, ByVal propName As String, ByVal dto As {cls})");
            w.Indent();
            w.Line("If parent Is Nothing Then Exit Sub");
            w.Line("If dto Is Nothing Then parent.UpdateNull propName: Exit Sub");
            w.Line("parent.AddObjectAt -1, propName");
            w.Line("Dim child As ChilkatJsonObject");
            w.Line("Set child = parent.ObjectOf(propName)");
            w.Line("child.Load dto.ToJson().Emit()");
            w.Outdent();
            w.Line("End Sub");
            w.Line();

            w.Line($"Public Function LoadList_{cls}(ByVal parent As ChilkatJsonObject, ByVal propName As String) As Collection");
            w.Indent();
            w.Line($"Set LoadList_{cls} = New Collection");
            w.Line("If parent Is Nothing Then Exit Function");
            w.Line("Dim arr As ChilkatJsonArray: Set arr = parent.ArrayOf(propName)");
            w.Line("If arr Is Nothing Then Exit Function");
            w.Line("Dim i As Long, n As Long: n = arr.Size");
            w.Line("For i = 0 To n - 1");
            w.Indent();
            w.Line($"Dim item As {cls}: Set item = New {cls}");
            w.Line("item.FromJson arr.ObjectAt(i)");
            w.Line($"LoadList_{cls}.Add item");
            w.Outdent();
            w.Line("Next i");
            w.Outdent();
            w.Line("End Function");
            w.Line();

            w.Line($"Public Sub AppendList_{cls}(ByVal parent As ChilkatJsonObject, ByVal propName As String, ByVal items As Collection)");
            w.Indent();
            w.Line("If parent Is Nothing Then Exit Sub");
            w.Line("parent.AddArrayAt -1, propName");
            w.Line("If items Is Nothing Then Exit Sub");
            w.Line("Dim arr As ChilkatJsonArray");
            w.Line("Set arr = parent.ArrayOf(propName)");
            w.Line($"Dim item As {cls}");
            w.Line("Dim child As ChilkatJsonObject");
            w.Line("For Each item In items");
            w.Indent();
            w.Line("arr.AddObjectAt -1");
            w.Line("Set child = arr.ObjectAt(arr.Size - 1)");
            w.Line("child.Load item.ToJson().Emit()");
            w.Outdent();
            w.Line("Next item");
            w.Outdent();
            w.Line("End Sub");
            w.Line();
        }
    }

    private static void WritePrimitiveListHelpers(Vb6Writer w)
    {
        EmitListHelper(w, "String", "obj.StringOf", "UpdateString", "Variant");
        EmitListHelper(w, "Long", "obj.IntOf", "UpdateInt", "Variant");
        EmitListHelper(w, "Currency", "obj.IntOf", "UpdateInt", "Variant");
        EmitListHelper(w, "Double", "obj.NumberOf", "UpdateNumber", "Variant");
        EmitListHelper(w, "Boolean", "obj.BoolOf", "UpdateBool", "Variant");
        EmitListHelper(w, "Date", "obj.StringOf", "UpdateString", "Variant");
        EmitVariantList(w);
    }

    private static void EmitListHelper(Vb6Writer w, string suffix, string reader, string writer, string itemDecl)
    {
        w.Line($"Public Function LoadList_{suffix}(ByVal parent As ChilkatJsonObject, ByVal propName As String) As Collection");
        w.Indent();
        w.Line($"Set LoadList_{suffix} = New Collection");
        w.Line("If parent Is Nothing Then Exit Function");
        w.Line("Dim arr As ChilkatJsonArray: Set arr = parent.ArrayOf(propName)");
        w.Line("If arr Is Nothing Then Exit Function");
        w.Line("Dim i As Long, n As Long: n = arr.Size");
        w.Line("For i = 0 To n - 1");
        w.Indent();
        w.Line($"LoadList_{suffix}.Add {ReaderForPrimitive(suffix)}");
        w.Outdent();
        w.Line("Next i");
        w.Outdent();
        w.Line("End Function");
        w.Line();

        w.Line($"Public Function GetJsonArrayAs_{suffix}(ByVal api As cApi, ByVal url As String) As Collection");
        w.Indent();
        w.Line("Dim arr As ChilkatJsonArray: Set arr = GetJsonArray(api, url)");
        w.Line($"Set GetJsonArrayAs_{suffix} = New Collection");
        w.Line("If arr Is Nothing Then Exit Function");
        w.Line("Dim i As Long, n As Long: n = arr.Size");
        w.Line("For i = 0 To n - 1");
        w.Indent();
        w.Line($"GetJsonArrayAs_{suffix}.Add {ReaderForPrimitive(suffix)}");
        w.Outdent();
        w.Line("Next i");
        w.Outdent();
        w.Line("End Function");
        w.Line();

        w.Line($"Public Sub AppendList_{suffix}(ByVal parent As ChilkatJsonObject, ByVal propName As String, ByVal items As Collection)");
        w.Indent();
        w.Line("If parent Is Nothing Then Exit Sub");
        w.Line("parent.AddArrayAt -1, propName");
        w.Line("If items Is Nothing Then Exit Sub");
        w.Line("Dim arr As ChilkatJsonArray");
        w.Line("Set arr = parent.ArrayOf(propName)");
        w.Line($"Dim v As {itemDecl}");
        w.Line("For Each v In items");
        w.Indent();
        w.Line(AppendForPrimitive(suffix));
        w.Outdent();
        w.Line("Next v");
        w.Outdent();
        w.Line("End Sub");
        w.Line();
    }

    private static string ReaderForPrimitive(string suffix) => suffix switch
    {
        "String" => "arr.StringAt(i)",
        "Long" => "arr.IntAt(i)",
        "Currency" => "arr.IntAt(i)",
        "Double" => "CDbl(Val(arr.StringAt(i)))",
        "Boolean" => "arr.BoolAt(i)",
        "Date" => "arr.StringAt(i)",
        _ => "arr.StringAt(i)"
    };

    private static string AppendForPrimitive(string suffix) => suffix switch
    {
        "String" => "arr.AddStringAt -1, CStr(v)",
        "Long" => "arr.AddIntAt -1, CLng(v)",
        "Currency" => "arr.AddIntAt -1, CLng(v)",
        "Double" => "arr.AddNumberAt -1, CStr(v)",
        "Boolean" => "arr.AddBoolAt -1, CBool(v)",
        "Date" => "arr.AddStringAt -1, Iso(CDate(v))",
        _ => "arr.AddStringAt -1, CStr(v)"
    };

    private static void EmitVariantList(Vb6Writer w)
    {
        EmitListHelper(w, "Variant", "obj.StringOf", "UpdateString", "Variant");
    }

    private static void WriteScalarWrap(Vb6Writer w)
    {
        w.Line("Public Function WrapScalarJson(ByVal v As Variant) As ChilkatJsonObject");
        w.Indent();
        w.Line("Set WrapScalarJson = New ChilkatJsonObject");
        w.Line("WrapScalarJson.UpdateString \"value\", CStr(v)");
        w.Outdent();
        w.Line("End Function");
    }

    private static void WriteBinaryHelpers(Vb6Writer w)
    {
        w.Line("Public Function GetBytes(ByVal api As cApi, ByVal url As String) As Variant");
        w.Indent();
        w.Line("Dim http As ChilkatHttp: Set http = NewHttp(api)");
        w.Line("Dim resp As ChilkatHttpResponse");
        w.Line("Set resp = http.QuickRequest(\"GET\", url)");
        w.Line("EnsureRespOk resp, \"GET \" & url");
        w.Line("GetBytes = resp.Body");
        w.Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Public Function PostJsonReturnBytes(ByVal api As cApi, ByVal url As String, ByVal body As ChilkatJsonObject) As Variant");
        w.Indent();
        w.Line("Dim http As ChilkatHttp: Set http = NewHttp(api)");
        w.Line("Dim jsonStr As String: If Not body Is Nothing Then jsonStr = body.Emit() Else jsonStr = \"\"");
        w.Line("Dim resp As ChilkatHttpResponse");
        w.Line("Set resp = http.PostJson2(url, gcContentType, jsonStr)");
        w.Line("EnsureRespOk resp, \"POST \" & url");
        w.Line("PostJsonReturnBytes = resp.Body");
        w.Outdent();
        w.Line("End Function");
        w.Line();
    }

    private static void WriteDateHelpers(Vb6Writer w)
    {
        w.Line("Public Function Iso(ByVal d As Date) As String");
        w.Indent();
        w.Line("If d = 0 Then Iso = \"\": Exit Function");
        w.Line("Iso = Format$(d, \"yyyy-mm-dd\\Thh:nn:ss\")");
        w.Outdent();
        w.Line("End Function");
        w.Line();

        w.Line("Public Function ParseIso(ByVal s As String) As Date");
        w.Indent();
        w.Line("On Error Resume Next");
        w.Line("If Len(s) = 0 Then ParseIso = 0: Exit Function");
        w.Line("Dim s2 As String: s2 = Replace$(s, \"T\", \" \")");
        w.Line("Dim p As Long: p = InStr(s2, \".\"): If p > 0 Then s2 = Left$(s2, p - 1)");
        w.Line("p = InStr(s2, \"+\"): If p > 0 Then s2 = Left$(s2, p - 1)");
        w.Line("p = InStrRev(s2, \"-\"): If p > 11 Then s2 = Left$(s2, p - 1)");
        w.Line("If Right$(s2, 1) = \"Z\" Then s2 = Left$(s2, Len(s2) - 1)");
        w.Line("ParseIso = CDate(s2)");
        w.Outdent();
        w.Line("End Function");
    }
}
