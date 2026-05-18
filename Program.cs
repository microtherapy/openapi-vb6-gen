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

        if (opts.Clean && Directory.Exists(opts.Output))
            foreach (var f in Directory.EnumerateFiles(opts.Output))
                File.Delete(f);
        Directory.CreateDirectory(opts.Output);

        try
        {
            var typeMapper = new Vb6TypeMapper();
            var (dtos, enums) = BuildSchemaModels(doc, typeMapper, opts.SchemaFilter);
            var controllers = BuildControllerModels(doc, typeMapper, opts.TagFilter);

            EmitAll(opts, dtos, enums, controllers);

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
        string? tagFilter = null, schemaFilter = null;
        bool clean = false;

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
            Clean = clean
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
              --clean                   wipe output dir first
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
                    VbName = Vb6Naming.SafeIdentifier($"V{oi.Value}"),
                    IntValue = oi.Value
                });
            else if (v is Microsoft.OpenApi.Any.OpenApiLong ol)
                en.Members.Add(new EnumMember
                {
                    VbName = Vb6Naming.SafeIdentifier($"V{ol.Value}"),
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

    private static List<ControllerModel> BuildControllerModels(OpenApiDocument doc, Vb6TypeMapper mapper, Regex? filter)
    {
        var byTag = new Dictionary<string, ControllerModel>(StringComparer.OrdinalIgnoreCase);
        if (doc.Paths is null) return new List<ControllerModel>();

        foreach (var (path, item) in doc.Paths)
        {
            foreach (var (method, op) in item.Operations)
            {
                var tag = op.Tags?.FirstOrDefault()?.Name ?? "Default";
                if (filter is not null && !filter.IsMatch(tag)) continue;
                if (!byTag.TryGetValue(tag, out var c))
                {
                    var cls = Vb6Naming.PascalCase(tag) + "Api";
                    c = new ControllerModel
                    {
                        Tag = tag,
                        ClassName = "c" + cls,
                        PropertyName = Vb6Naming.PascalCase(tag)
                    };
                    byTag[tag] = c;
                }
                c.Operations.Add(BuildOperation(method.ToString(), path, op, mapper));
            }
        }
        return byTag.Values.OrderBy(c => c.ClassName).ToList();
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
                else
                    model.Response = new Vb6Type { Kind = Vb6Kind.String, Declaration = "String" };
            }
        }
        return model;
    }

    private static void EmitAll(Options opts, List<DtoModel> dtos, List<EnumModel> enums, List<ControllerModel> controllers)
    {
        var dtoEmitter = new DtoEmitter();
        var controllerEmitter = new ControllerEmitter();
        var facadeEmitter = new FacadeEmitter();
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

        var helperPath = Path.Combine(opts.Output, "modGenApi.bas");
        File.WriteAllText(helperPath, helperEmitter.Emit(dtos, enums));
        moduleFiles.Add(helperPath);

        var inputs = new VbpEmitterInputs
        {
            ProjectName = opts.ProjectName,
            ClassFiles = classFiles,
            ModuleFiles = moduleFiles,
            MainVbpPath = opts.MainVbp
        };
        File.WriteAllText(Path.Combine(opts.Output, opts.ProjectName + ".vbp"), vbpEmitter.EmitVbp(inputs));
        File.WriteAllText(Path.Combine(opts.Output, opts.ProjectName + ".vbw"), vbpEmitter.EmitVbw(inputs));

        var clientVbpRel = opts.ProjectName + ".vbp";
        File.WriteAllText(Path.Combine(opts.Output, opts.ProjectName + ".vbg"),
            vbpEmitter.EmitVbg(opts.ProjectName, opts.MainVbp, clientVbpRel));
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
    }
}
