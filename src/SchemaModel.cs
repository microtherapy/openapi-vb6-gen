namespace OpenApiVb6Gen;

internal sealed class DtoModel
{
    public required string SchemaName { get; init; }
    public required string ClassName { get; init; }
    public List<DtoPropertyModel> Properties { get; } = new();
    public string? Description { get; init; }
}

internal sealed class DtoPropertyModel
{
    public required string JsonName { get; init; }
    public required string VbName { get; init; }
    public required Vb6Type Type { get; init; }
    public string? Description { get; init; }
}

internal sealed class EnumModel
{
    public required string EnumName { get; init; }
    public required bool IsString { get; init; }
    public List<EnumMember> Members { get; } = new();
    public string? Description { get; init; }
}

internal sealed class EnumMember
{
    public required string VbName { get; init; }
    public required long IntValue { get; init; }
    public string? StringValue { get; init; }
}
