using Microsoft.OpenApi.Models;
using Microsoft.OpenApi.Readers;

namespace OpenApiVb6Gen;

internal static class SpecLoader
{
    public static OpenApiDocument Load(string input)
    {
        using var stream = OpenStream(input);
        var reader = new OpenApiStreamReader();
        var doc = reader.Read(stream, out var diagnostic);
        if (diagnostic.Errors.Count > 0)
        {
            Console.Error.WriteLine("OpenAPI parse errors:");
            foreach (var err in diagnostic.Errors)
                Console.Error.WriteLine($"  {err.Pointer}: {err.Message}");
        }
        return doc;
    }

    private static Stream OpenStream(string input)
    {
        if (input.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            || input.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            using var client = new HttpClient();
            var bytes = client.GetByteArrayAsync(input).GetAwaiter().GetResult();
            return new MemoryStream(bytes);
        }
        return File.OpenRead(input);
    }
}
