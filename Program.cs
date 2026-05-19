using System.Text.RegularExpressions;
using Microsoft.OpenApi.Models;
using OpenApiVb6Gen;

return App.Run(args);

internal static class App
{
    public static int Run(string[] args)
    {
        var opts = ParseArgs(args);
        if (opts is null) return 64;

        Console.WriteLine($"Loading OpenAPI spec from: {opts.Input}");
        OpenApiDocument doc;
        try { doc = SpecLoader.Load(opts.Input); }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to load spec: {ex.Message}");
            return 1;
        }

        var compatDllPath = Path.Combine(opts.Output, opts.ProjectName + ".compat.dll");
        if (opts.Clean && Directory.Exists(opts.Output))
        {
            var preserveName = Path.GetFileName(compatDllPath);
            foreach (var f in Directory.EnumerateFiles(opts.Output))
                if (!string.Equals(Path.GetFileName(f), preserveName, StringComparison.OrdinalIgnoreCase))
                    File.Delete(f);
        }
        Directory.CreateDirectory(opts.Output);

        try
        {
            var maxAliasLen = Math.Max(4, 39 - opts.ProjectName.Length - 1 - 1);
            Vb6Naming.SetSchemaAliases(doc.Components?.Schemas is { } s
                ? Vb6Naming.BuildSchemaAliasMap(s.Keys, maxAliasLen)
                : null);

            var typeMapper = new Vb6TypeMapper();
            var (dtos, enums) = BuildSchemaModels(doc, typeMapper, opts.SchemaFilter);
            var controllers = BuildControllerModels(doc, typeMapper, opts.TagFilter, opts.ProjectName);

            EmitAll(opts, dtos, enums, controllers, compatDllPath);

            Console.WriteLine($"Generated {dtos.Count} DTO classes, {enums.Count} enums, {controllers.Count} controllers");
            Console.WriteLine($"Output: {opts.Output}");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Emission failed: {ex}");
            return 2;
        }
    }

    private static Options? ParseArgs(string[] args)
    {
        string? input = null, output = null, projectName = "OpenApiClient", mainVbp = null;
        string? tagFilter = null, schemaFilter = null, vb6Exe = null;
        bool clean = false, noSeed = false;

        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "--input": input = args[++i]; break;
                case "--output": output = args[++i]; break;
                case "--project-name": projectName = args[++i]; break;
                case "--main-vbp": mainVbp = args[++i]; break;
                case "--tag-filter": tagFilter = args[++i]; break;
                case "--schema-filter": schemaFilter = args[++i]; break;
                case "--clean": clean = true; break;
                case "--vb6-exe": vb6Exe = args[++i]; break;
                case "--no-seed": noSeed = true; break;
                case "-h" or "--help": PrintHelp(); return null;
                default:
                    Console.Error.WriteLine($"Unknown argument: {args[i]}");
                    PrintHelp();
                    return null;
            }
        }

        if (string.IsNullOrWhiteSpace(input) || string.IsNullOrWhiteSpace(output))
        {
            Console.Error.WriteLine("--input and --output are required.");
            PrintHelp();
            return null;
        }

        return new Options
        {
            Input = input!,
            Output = output!,
            ProjectName = projectName!,
            MainVbp = mainVbp,
            TagFilter = string.IsNullOrWhiteSpace(tagFilter) ? null : new Regex(tagFilter!),
            SchemaFilter = string.IsNullOrWhiteSpace(schemaFilter) ? null : new Regex(schemaFilter!),
            Clean = clean,
            Vb6Exe = vb6Exe,
            NoSeed = noSeed
        };
    }

    private static void PrintHelp()
    {
        Console.Error.WriteLine("""
            openapi-vb6-gen
              --input <path-or-url>     swagger.json file or http(s) URL
              --output <dir>            target directory for generated .vbp/.cls/.bas
              --project-name <name>     default: OpenApiClient
              --main-vbp <path>         optional: host .vbp to copy Chilkat References from
              --tag-filter <regex>      optional: include only matching tags
              --schema-filter <regex>   optional: include only matching schemas
              --clean                   wipe output dir (preserves <project>.compat.dll)
              --vb6-exe <path>          override default VB6.EXE location for the compat-DLL build
              --no-seed                 skip the VB6 build (CI mode); .vbp falls back to CompatibleMode=0 if no .compat.dll
            """);
    }

    private static (List<DtoModel>, List<EnumModel>) BuildSchemaModels(OpenApiDocument doc, Vb6TypeMapper mapper, Regex? filter)
    {
        var dtos = new List<DtoModel>();
        var enums = new List<EnumModel>();
        if (doc.Components?.Schemas is null) return (dtos, enums);

        foreach (var kv in doc.Components.Schemas)
        {
            var name = kv.Key;
            if (filter is not null && !filter.IsMatch(name)) continue;
            var schema = kv.Value;
            if (schema.Enum is { Count: > 0 })
            {
                enums.Add(BuildEnum(name, schema));
                continue;
            }
            if (schema.Type == "object" || schema.Properties is { Count: > 0 })
                dtos.Add(BuildDto(name, schema, mapper));
        }
        return (dtos, enums);
    }

    private static EnumModel BuildEnum(string name, OpenApiSchema schema)
    {
        var en = new EnumModel
        {
            EnumName = Vb6Naming.EnumName(name),
            IsString = schema.Type == "string",
            Description = schema.Description
        };
        long fallback = 0;
        foreach (var v in schema.Enum)
        {
            if (v is Microsoft.OpenApi.Any.OpenApiInteger oi)
                en.Members.Add(new EnumMember
                {
                    VbName = Vb6Naming.SafeIdentifier(oi.Value < 0 ? $"VNeg{-oi.Value}" : $"V{oi.Value}"),
                    IntValue = oi.Value
                });
            else if (v is Microsoft.OpenApi.Any.OpenApiLong ol)
                en.Members.Add(new EnumMember
                {
                    VbName = Vb6Naming.SafeIdentifier(ol.Value < 0 ? $"VNeg{-ol.Value}" : $"V{ol.Value}"),
                    IntValue = ol.Value
                });
            else if (v is Microsoft.OpenApi.Any.OpenApiString os)
                en.Members.Add(new EnumMember
                {
                    VbName = Vb6Naming.SafeIdentifier(os.Value),
                    IntValue = fallback++,
                    StringValue = os.Value
                });
        }
        return en;
    }

    private static DtoModel BuildDto(string name, OpenApiSchema schema, Vb6TypeMapper mapper)
    {
        var dto = new DtoModel
        {
            SchemaName = name,
            ClassName = Vb6Naming.ClassName(name),
            Description = schema.Description
        };
        if (schema.Properties is null) return dto;
        foreach (var pkv in schema.Properties)
        {
            var t = mapper.Map(pkv.Value);
            dto.Properties.Add(new DtoPropertyModel
            {
                JsonName = pkv.Key,
                VbName = Vb6Naming.PropertyName(pkv.Key),
                Type = t,
                Description = pkv.Value.Description
            });
        }
        return dto;
    }

    private static List<ControllerModel> BuildControllerModels(OpenApiDocument doc, Vb6TypeMapper mapper, Regex? filter, string projectName)
    {
        var byTag = new Dictionary<string, ControllerModel>(StringComparer.OrdinalIgnoreCase);
        if (doc.Paths is null) return new List<ControllerModel>();

        var maxClassLen = Math.Max(8, 39 - projectName.Length - 1);
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var (path, item) in doc.Paths)
        {
            foreach (var (method, op) in item.Operations)
            {
                var tag = op.Tags?.FirstOrDefault()?.Name ?? "Default";
                if (filter is not null && !filter.IsMatch(tag)) continue;
                if (!byTag.TryGetValue(tag, out var c))
                {
                    var full = "c" + Vb6Naming.PascalCase(tag) + "Api";
                    var cls = TruncateUnique(full, maxClassLen, used);
                    used.Add(cls);
                    c = new ControllerModel
                    {
                        Tag = tag,
                        ClassName = cls,
                        PropertyName = Vb6Naming.SafeIdentifier(tag)
                    };
                    byTag[tag] = c;
                }
                c.Operations.Add(BuildOperation(method.ToString(), path, op, mapper));
            }
        }
        return byTag.Values.OrderBy(c => c.ClassName).ToList();
    }

    private static string TruncateUnique(string name, int maxLen, HashSet<string> used)
    {
        if (name.Length <= maxLen && !used.Contains(name)) return name;
        var baseName = name.Length <= maxLen ? name : name[..maxLen];
        if (!used.Contains(baseName)) return baseName;
        for (int n = 2; n < 1000; n++)
        {
            var suffix = n.ToString();
            var candidate = baseName[..Math.Min(baseName.Length, maxLen - suffix.Length)] + suffix;
            if (!used.Contains(candidate)) return candidate;
        }
        return baseName;
    }

    private static OperationModel BuildOperation(string method, string path, OpenApiOperation op, Vb6TypeMapper mapper)
    {
        var model = new OperationModel
        {
            OperationId = op.OperationId ?? "",
            VbMethodName = Vb6Naming.OperationName(op.OperationId ?? "", op.Tags?.FirstOrDefault()?.Name ?? "", method, path),
            HttpMethod = method,
            PathTemplate = path,
            Description = op.Description ?? op.Summary
        };

        foreach (var p in op.Parameters)
        {
            var t = mapper.Map(p.Schema);
            var pm = new ParameterModel
            {
                ApiName = p.Name,
                VbName = Vb6Naming.SafeParameter(p.Name),
                Type = t,
                Required = p.Required,
                Description = p.Description
            };
            if (p.In == ParameterLocation.Path) model.PathParameters.Add(pm);
            else if (p.In == ParameterLocation.Query) model.QueryParameters.Add(pm);
        }

        if (op.RequestBody is not null)
        {
            var json = op.RequestBody.Content.FirstOrDefault(kv => kv.Key.Contains("json", StringComparison.OrdinalIgnoreCase));
            if (json.Value?.Schema is not null)
            {
                model.Body = new ParameterModel
                {
                    ApiName = "body",
                    VbName = "body",
                    Type = mapper.Map(json.Value.Schema),
                    Required = op.RequestBody.Required
                };
            }
            else
            {
                model.SkipReason = "non-JSON request body (multipart/form-data or octet-stream) is not supported";
            }
        }

        var success = op.Responses
            .Where(kv => kv.Key.StartsWith("2", StringComparison.Ordinal))
            .OrderBy(kv => kv.Key)
            .FirstOrDefault();
        if (success.Value is not null)
        {
            if (success.Key == "204" || success.Value.Content.Count == 0)
                model.Response = null;
            else
            {
                var json = success.Value.Content.FirstOrDefault(kv => kv.Key.Contains("json", StringComparison.OrdinalIgnoreCase));
                if (json.Value?.Schema is not null)
                    model.Response = mapper.Map(json.Value.Schema);
                else if (LooksLikeBinaryResponse(success.Value.Content))
                    model.Response = new Vb6Type { Kind = Vb6Kind.Binary, Declaration = "Variant" };
                else
                    model.Response = new Vb6Type { Kind = Vb6Kind.String, Declaration = "String" };
            }
        }
        return model;
    }

    private static bool LooksLikeBinaryResponse(IDictionary<string, OpenApiMediaType> content)
    {
        foreach (var (mt, media) in content)
        {
            if (mt.StartsWith("application/pdf", StringComparison.OrdinalIgnoreCase)) return true;
            if (mt.StartsWith("application/octet-stream", StringComparison.OrdinalIgnoreCase)) return true;
            if (mt.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) return true;
            if (media.Schema is { Type: "string", Format: "binary" }) return true;
        }
        return false;
    }

    private static void EmitAll(Options opts, List<DtoModel> dtos, List<EnumModel> enums, List<ControllerModel> controllers, string compatDllPath)
    {
        var dtoEmitter = new DtoEmitter();
        var controllerEmitter = new ControllerEmitter();
        var facadeEmitter = new FacadeEmitter();
        var enumEmitter = new EnumEmitter();
        var helperEmitter = new HelperEmitter();
        var vbpEmitter = new VbpEmitter();

        var classFiles = new List<string>();
        var moduleFiles = new List<string>();

        foreach (var dto in dtos)
        {
            var p = Path.Combine(opts.Output, dto.ClassName + ".cls");
            File.WriteAllText(p, dtoEmitter.Emit(dto));
            classFiles.Add(p);
        }
        foreach (var c in controllers)
        {
            var p = Path.Combine(opts.Output, c.ClassName + ".cls");
            File.WriteAllText(p, controllerEmitter.Emit(c));
            classFiles.Add(p);
        }
        var facadePath = Path.Combine(opts.Output, "cApi.cls");
        File.WriteAllText(facadePath, facadeEmitter.Emit(controllers));
        classFiles.Add(facadePath);

        if (enums.Count > 0)
        {
            var enumPath = Path.Combine(opts.Output, "cEnums.cls");
            File.WriteAllText(enumPath, enumEmitter.Emit(enums));
            classFiles.Add(enumPath);
        }

        var helperPath = Path.Combine(opts.Output, "modGenApi.bas");
        File.WriteAllText(helperPath, helperEmitter.Emit(dtos));
        moduleFiles.Add(helperPath);

        var baseInputs = new VbpEmitterInputs
        {
            ProjectName = opts.ProjectName,
            ClassFiles = classFiles,
            ModuleFiles = moduleFiles,
            MainVbpPath = opts.MainVbp,
            CompatibleExePath = null
        };
        var vbpPath = Path.Combine(opts.Output, opts.ProjectName + ".vbp");

        if (!opts.NoSeed)
            BuildCompatDll(opts, baseInputs, vbpEmitter, vbpPath, compatDllPath);

        var finalInputs = baseInputs with { CompatibleExePath = File.Exists(compatDllPath) ? compatDllPath : null };
        File.WriteAllText(vbpPath, vbpEmitter.EmitVbp(finalInputs));
        File.WriteAllText(Path.Combine(opts.Output, opts.ProjectName + ".vbw"), vbpEmitter.EmitVbw(finalInputs));

        var clientVbpRel = opts.ProjectName + ".vbp";
        File.WriteAllText(Path.Combine(opts.Output, opts.ProjectName + ".vbg"),
            vbpEmitter.EmitVbg(opts.ProjectName, opts.MainVbp, clientVbpRel));
    }

    private static void BuildCompatDll(Options opts, VbpEmitterInputs baseInputs, VbpEmitter vbpEmitter,
        string vbpPath, string compatDllPath)
    {
        var anchorExists = File.Exists(compatDllPath);
        var buildInputs = baseInputs with { CompatibleExePath = anchorExists ? compatDllPath : null };
        File.WriteAllText(vbpPath, vbpEmitter.EmitVbp(buildInputs));
        File.WriteAllText(Path.Combine(opts.Output, opts.ProjectName + ".vbw"), vbpEmitter.EmitVbw(buildInputs));

        var builtDll = Path.Combine(opts.Output, opts.ProjectName + ".dll");
        if (File.Exists(builtDll)) File.Delete(builtDll);
        var vb6 = Vb6Bootstrap.FindVb6Exe(opts.Vb6Exe);
        var buildLog = Path.Combine(opts.Output, "_build.log");
        var verb = anchorExists ? "Rebuilding compat DLL against existing anchor" : "Seeding compat DLL (no existing anchor)";
        Console.WriteLine($"{verb}: {vb6} /make {Path.GetFileName(vbpPath)}");
        var exit = Vb6Bootstrap.RunMake(vb6, vbpPath, buildLog);
        if (!File.Exists(builtDll))
        {
            var log = File.Exists(buildLog) ? File.ReadAllText(buildLog) : "(no log)";
            throw new InvalidOperationException($"VB6 build failed (exit={exit}):\n{log}");
        }

        if (File.Exists(compatDllPath)) File.Delete(compatDllPath);
        File.Move(builtDll, compatDllPath);

        if (File.Exists(buildLog)) File.Delete(buildLog);

        Console.WriteLine($"Wrote compat DLL: {Path.GetFileName(compatDllPath)}");
    }

    private sealed class Options
    {
        public required string Input { get; init; }
        public required string Output { get; init; }
        public required string ProjectName { get; init; }
        public string? MainVbp { get; init; }
        public Regex? TagFilter { get; init; }
        public Regex? SchemaFilter { get; init; }
        public bool Clean { get; init; }
        public string? Vb6Exe { get; init; }
        public bool NoSeed { get; init; }
    }
}
