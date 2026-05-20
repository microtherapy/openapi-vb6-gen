using System.Diagnostics;
using System.Text;

namespace OpenApiVb6Gen;

internal static class Vb6Bootstrap
{
    /// <summary>
    /// Looks up the file path registered for a TypeLib in HKCR. Returns null if the
    /// typelib is not registered. Uses reg.exe to stay net10.0 cross-platform —
    /// the generator only runs on Windows, but this avoids pulling in the
    /// Microsoft.Win32.Registry dependency.
    /// </summary>
    public static string? LookupTypelibPath(string libId, string version = "1.0")
    {
        var key = $@"HKCR\TypeLib\{libId}\{version}\0\win32";
        var psi = new ProcessStartInfo("reg.exe")
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        psi.ArgumentList.Add("query");
        psi.ArgumentList.Add(key);
        psi.ArgumentList.Add("/ve");
        using var p = Process.Start(psi);
        if (p is null) return null;
        var sb = new StringBuilder();
        p.OutputDataReceived += (_, e) => { if (e.Data != null) sb.AppendLine(e.Data); };
        p.BeginOutputReadLine();
        if (!p.WaitForExit(5_000))
        {
            try { p.Kill(true); } catch { }
            return null;
        }
        if (p.ExitCode != 0) return null;
        // reg query output:
        //   HKEY_CLASSES_ROOT\TypeLib\{...}\1.0\0\win32
        //       (Default)    REG_SZ    C:\path\to\foo.dll
        foreach (var line in sb.ToString().Split('\n'))
        {
            var idx = line.IndexOf("REG_SZ", StringComparison.Ordinal);
            if (idx < 0) continue;
            var path = line[(idx + "REG_SZ".Length)..].Trim();
            if (path.Length > 0) return path;
        }
        return null;
    }

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
