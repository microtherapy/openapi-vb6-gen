using System.Diagnostics;

namespace OpenApiVb6Gen;

internal static class Vb6Bootstrap
{
    public static string FindVb6Exe(string? explicitPath)
    {
        if (!string.IsNullOrWhiteSpace(explicitPath))
        {
            if (!File.Exists(explicitPath))
                throw new FileNotFoundException($"VB6.EXE not found at {explicitPath}");
            return explicitPath;
        }
        string[] candidates =
        {
            @"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE",
            @"C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
        };
        foreach (var c in candidates)
            if (File.Exists(c)) return c;
        throw new FileNotFoundException("VB6.EXE not found in default locations; pass --vb6-exe to override");
    }

    public static int RunMake(string vb6Exe, string vbpPath, string? logPath)
    {
        var psi = new ProcessStartInfo(vb6Exe)
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            WorkingDirectory = Path.GetDirectoryName(vbpPath) ?? Environment.CurrentDirectory
        };
        psi.ArgumentList.Add("/make");
        psi.ArgumentList.Add(vbpPath);
        if (!string.IsNullOrWhiteSpace(logPath))
        {
            psi.ArgumentList.Add("/out");
            psi.ArgumentList.Add(logPath);
        }
        using var p = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start VB6.EXE");
        if (!p.WaitForExit((int)TimeSpan.FromMinutes(5).TotalMilliseconds))
        {
            try { p.Kill(true); } catch { }
            throw new TimeoutException("VB6 /make exceeded 5 minutes — likely stuck on a dialog or missing reference.");
        }
        return p.ExitCode;
    }
}
