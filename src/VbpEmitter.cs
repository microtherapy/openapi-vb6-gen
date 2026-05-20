using System.Text;

namespace OpenApiVb6Gen;

internal enum ChilkatVersion
{
    V9_5,
    V11
}

internal sealed record ChilkatTarget(string TypelibGuid, string DllFileName, string ReferenceDescription, string TypelibVersion)
{
    public static readonly ChilkatTarget V9_5 = new(
        "{004CB902-F437-4D01-BD85-9E18836DA5C2}",
        "ChilkatAx-9.5.0-win32.dll",
        "Chilkat ActiveX v9.5.0",
        "1.0");

    public static readonly ChilkatTarget V11 = new(
        "{06FB4061-5E43-42E0-8A6E-4A1C869E59AF}",
        "ChilkatAx-win32.dll",
        "Chilkat ActiveX v11.0.0",
        "1.0");

    public static ChilkatTarget For(ChilkatVersion v) => v switch
    {
        ChilkatVersion.V9_5 => V9_5,
        ChilkatVersion.V11 => V11,
        _ => V11
    };
}

internal sealed record VbpEmitterInputs
{
    public required string ProjectName { get; init; }
    public required string OutputDir { get; init; }
    public required IReadOnlyList<string> ClassFiles { get; init; }
    public required IReadOnlyList<string> ModuleFiles { get; init; }
    public string? MainVbpPath { get; init; }
    public string? CompatibleExePath { get; init; }
    public ChilkatVersion Chilkat { get; init; } = ChilkatVersion.V11;
}

internal sealed class VbpEmitter
{
    public string EmitVbp(VbpEmitterInputs i)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Type=OleDll");
        var target = ChilkatTarget.For(i.Chilkat);
        var refs = TryCollectChilkatReferences(i.MainVbpPath, i.OutputDir, target.TypelibGuid);
        if (refs.Count == 0)
            refs.Add(BuildChilkatReferenceFromRegistry(i.OutputDir, target));
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
        if (!string.IsNullOrWhiteSpace(i.CompatibleExePath) && File.Exists(i.CompatibleExePath))
        {
            sb.AppendLine("CompatibleMode=\"2\"");
            sb.AppendLine($"CompatibleEXE32=\"{MakeRelativeIfPossible(i.OutputDir, i.CompatibleExePath)}\"");
        }
        else
        {
            sb.AppendLine("CompatibleMode=\"0\"");
        }
        sb.AppendLine("MajorVer=1");
        sb.AppendLine("MinorVer=0");
        sb.AppendLine("RevisionVer=0");
        sb.AppendLine("AutoIncrementVer=0");
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

    public string EmitVbg(string projectName, string? hostVbpAbsolutePath, string clientVbpRelativeFromVbg, string outputDir)
    {
        var sb = new StringBuilder();
        sb.AppendLine("VBGROUP 5.0");
        sb.AppendLine($"StartupProject={clientVbpRelativeFromVbg}");
        sb.AppendLine($"Project={clientVbpRelativeFromVbg}");
        if (!string.IsNullOrWhiteSpace(hostVbpAbsolutePath) && File.Exists(hostVbpAbsolutePath))
            sb.AppendLine($"Project={MakeRelativeIfPossible(outputDir, hostVbpAbsolutePath)}");
        return sb.ToString();
    }

    private static List<string> TryCollectChilkatReferences(string? mainVbpPath, string outputDir, string typelibGuid)
    {
        var result = new List<string>();
        if (string.IsNullOrWhiteSpace(mainVbpPath) || !File.Exists(mainVbpPath))
            return result;
        var mainVbpDir = Path.GetDirectoryName(Path.GetFullPath(mainVbpPath));
        foreach (var line in File.ReadAllLines(mainVbpPath))
        {
            if (!line.StartsWith("Reference=", StringComparison.OrdinalIgnoreCase) &&
                !line.StartsWith("Object=", StringComparison.OrdinalIgnoreCase))
                continue;
            if (!line.Contains("Chilkat", StringComparison.OrdinalIgnoreCase)) continue;
            if (!line.Contains(typelibGuid, StringComparison.OrdinalIgnoreCase)) continue;
            // Reference line in main vbp may have a path relative to main vbp's folder.
            // Re-anchor it to outputDir so it remains valid when written to the generated vbp.
            result.Add(RebaseVbpReferenceLine(line, mainVbpDir, outputDir));
        }
        return result;
    }

    /// <summary>
    /// Builds a synthetic Chilkat Reference= line for the requested version, looking
    /// up the typelib's registered file path and making it relative to
    /// <paramref name="outputDir"/>. Falls back to a sensible default path if the
    /// typelib isn't registered on this machine.
    /// </summary>
    private static string BuildChilkatReferenceFromRegistry(string outputDir, ChilkatTarget target)
    {
        var registeredPath = Vb6Bootstrap.LookupTypelibPath(target.TypelibGuid, target.TypelibVersion);
        string path;
        if (registeredPath is not null)
        {
            path = MakeRelativeIfPossible(outputDir, registeredPath);
        }
        else if (target == ChilkatTarget.V11)
        {
            path = @"..\..\..\Program Files (x86)\Chilkat Software, Inc\Chilkat 32-bit ActiveX\ChilkatAx-win32.dll";
        }
        else
        {
            // Chilkat 9.5 doesn't have a canonical install location; just use the bare
            // filename and let VB6 fall back to its DLL search order.
            path = target.DllFileName;
        }
        return $@"Reference=*\G{target.TypelibGuid}#{target.TypelibVersion}#0#{path}#{target.ReferenceDescription}";
    }

    /// <summary>
    /// VB6 Reference= lines have a path token (the 4th field, after the #-separated
    /// GUID/version/lcid) that may be absolute or relative. If relative, it's
    /// resolved against the .vbp's folder. When copying a Reference= line from one
    /// vbp into another that lives in a different folder, the path token must be
    /// re-anchored or the new vbp won't find the DLL.
    /// </summary>
    private static string RebaseVbpReferenceLine(string line, string? oldVbpDir, string newVbpDir)
    {
        if (oldVbpDir is null) return line;
        // Format: Reference=*\G{guid}#major.minor#lcid#PATH#description
        var prefix = "Reference=";
        if (!line.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) return line;
        var body = line[prefix.Length..];
        var parts = body.Split('#');
        if (parts.Length < 5) return line;
        var pathToken = parts[3];
        var absolute = Path.IsPathRooted(pathToken)
            ? pathToken
            : Path.GetFullPath(Path.Combine(oldVbpDir, pathToken));
        parts[3] = MakeRelativeIfPossible(newVbpDir, absolute);
        return prefix + string.Join('#', parts);
    }

    private static string MakeRelativeIfPossible(string anchorDir, string targetPath)
    {
        if (string.IsNullOrEmpty(anchorDir) || string.IsNullOrEmpty(targetPath))
            return targetPath;
        try
        {
            var rel = Path.GetRelativePath(Path.GetFullPath(anchorDir), Path.GetFullPath(targetPath));
            // If GetRelativePath couldn't bridge (e.g. different drive letters) it
            // returns the absolute path unchanged.
            return rel;
        }
        catch
        {
            return targetPath;
        }
    }
}
