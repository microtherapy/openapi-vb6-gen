using System.Text;

namespace OpenApiVb6Gen;

internal sealed record VbpEmitterInputs
{
    public required string ProjectName { get; init; }
    public required string OutputDir { get; init; }
    public required IReadOnlyList<string> ClassFiles { get; init; }
    public required IReadOnlyList<string> ModuleFiles { get; init; }
    public string? MainVbpPath { get; init; }
    public string? CompatibleExePath { get; init; }
}

internal sealed class VbpEmitter
{
    public string EmitVbp(VbpEmitterInputs i)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Type=OleDll");
        var refs = TryCollectChilkat11References(i.MainVbpPath, i.OutputDir);
        if (refs.Count == 0)
            refs.Add(BuildChilkat11ReferenceFromRegistry(i.OutputDir));
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

    private const string Chilkat11TypelibGuid = "{06FB4061-5E43-42E0-8A6E-4A1C869E59AF}";

    private static List<string> TryCollectChilkat11References(string? mainVbpPath, string outputDir)
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
            if (!line.Contains(Chilkat11TypelibGuid, StringComparison.OrdinalIgnoreCase)) continue;
            // Reference line in main vbp may have a path relative to main vbp's folder.
            // Re-anchor it to outputDir so it remains valid when written to the generated vbp.
            result.Add(RebaseVbpReferenceLine(line, mainVbpDir, outputDir));
        }
        return result;
    }

    /// <summary>
    /// Builds a synthetic Chilkat 11 Reference= line, looking up the typelib's
    /// registered file path and making it relative to <paramref name="outputDir"/>.
    /// Falls back to the absolute path if not registered.
    /// </summary>
    private static string BuildChilkat11ReferenceFromRegistry(string outputDir)
    {
        var registeredPath = Vb6Bootstrap.LookupTypelibPath(Chilkat11TypelibGuid, "1.0");
        var path = registeredPath is not null
            ? MakeRelativeIfPossible(outputDir, registeredPath)
            : @"..\..\..\Program Files (x86)\Chilkat Software, Inc\Chilkat 32-bit ActiveX\ChilkatAx-win32.dll";
        return $@"Reference=*\G{Chilkat11TypelibGuid}#1.0#0#{path}#Chilkat ActiveX v11.0.0";
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
