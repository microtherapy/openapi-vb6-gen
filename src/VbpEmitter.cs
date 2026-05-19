using System.Text;

namespace OpenApiVb6Gen;

internal sealed class VbpEmitterInputs
{
    public required string ProjectName { get; init; }
    public required IReadOnlyList<string> ClassFiles { get; init; }
    public required IReadOnlyList<string> ModuleFiles { get; init; }
    public string? MainVbpPath { get; init; }
}

internal sealed class VbpEmitter
{
    public string EmitVbp(VbpEmitterInputs i)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Type=OleDll");
        var refs = TryCollectChilkat11References(i.MainVbpPath);
        if (refs.Count == 0)
            refs.Add(@"Reference=*\G{06FB4061-5E43-42E0-8A6E-4A1C869E59AF}#1.0#0#C:\Windows\SysWow64\ChilkatAx-win32.dll#Chilkat ActiveX v11.0.0");
        foreach (var r in refs)
            sb.AppendLine(r);
        foreach (var m in i.ModuleFiles)
            sb.AppendLine($"Module={Path.GetFileNameWithoutExtension(m)}; {Path.GetFileName(m)}");
        foreach (var c in i.ClassFiles)
            sb.AppendLine($"Class={Path.GetFileNameWithoutExtension(c)}; {Path.GetFileName(c)}");
        sb.AppendLine("Startup=\"(None)\"");
        sb.AppendLine($"ExeName32=\"{i.ProjectName}.dll\"");
        sb.AppendLine("Command32=\"\"");
        sb.AppendLine($"Name=\"{i.ProjectName}\"");
        sb.AppendLine("HelpContextID=\"0\"");
        sb.AppendLine($"Description=\"Generated API client for {i.ProjectName}\"");
        sb.AppendLine("CompatibleMode=\"0\"");
        sb.AppendLine("MajorVer=1");
        sb.AppendLine("MinorVer=0");
        sb.AppendLine("RevisionVer=0");
        sb.AppendLine("AutoIncrementVer=1");
        sb.AppendLine("ServerSupportFiles=0");
        sb.AppendLine("VersionCompanyName=\"Generated\"");
        sb.AppendLine("CompilationType=0");
        sb.AppendLine("OptimizationType=0");
        sb.AppendLine("FavorPentiumPro(tm)=0");
        sb.AppendLine("CodeViewDebugInfo=0");
        sb.AppendLine("NoAliasing=0");
        sb.AppendLine("BoundsCheck=0");
        sb.AppendLine("OverflowCheck=0");
        sb.AppendLine("FlPointCheck=0");
        sb.AppendLine("FDIVCheck=0");
        sb.AppendLine("UnroundedFP=0");
        sb.AppendLine("StartMode=0");
        sb.AppendLine("Unattended=0");
        sb.AppendLine("Retained=0");
        sb.AppendLine("ThreadPerObject=0");
        sb.AppendLine("MaxNumberOfThreads=1");
        return sb.ToString();
    }

    public string EmitVbw(VbpEmitterInputs i)
    {
        var sb = new StringBuilder();
        foreach (var m in i.ModuleFiles)
            sb.AppendLine($"{Path.GetFileNameWithoutExtension(m)} = 0, 0, 0, 0, C");
        foreach (var c in i.ClassFiles)
            sb.AppendLine($"{Path.GetFileNameWithoutExtension(c)} = 0, 0, 0, 0, C");
        return sb.ToString();
    }

    public string EmitVbg(string projectName, string? hostVbpAbsolutePath, string clientVbpRelativeFromVbg)
    {
        var sb = new StringBuilder();
        sb.AppendLine("VBGROUP 5.0");
        sb.AppendLine($"StartupProject={clientVbpRelativeFromVbg}");
        sb.AppendLine($"Project={clientVbpRelativeFromVbg}");
        if (!string.IsNullOrWhiteSpace(hostVbpAbsolutePath) && File.Exists(hostVbpAbsolutePath))
            sb.AppendLine($"Project={hostVbpAbsolutePath}");
        return sb.ToString();
    }

    private static List<string> TryCollectChilkat11References(string? mainVbpPath)
    {
        var result = new List<string>();
        if (string.IsNullOrWhiteSpace(mainVbpPath) || !File.Exists(mainVbpPath))
            return result;
        const string chilkat11TypelibGuid = "{06FB4061-5E43-42E0-8A6E-4A1C869E59AF}";
        foreach (var line in File.ReadAllLines(mainVbpPath))
        {
            if (!line.StartsWith("Reference=", StringComparison.OrdinalIgnoreCase) &&
                !line.StartsWith("Object=", StringComparison.OrdinalIgnoreCase))
                continue;
            if (!line.Contains("Chilkat", StringComparison.OrdinalIgnoreCase)) continue;
            if (line.Contains(chilkat11TypelibGuid, StringComparison.OrdinalIgnoreCase))
                result.Add(line);
        }
        return result;
    }
}
