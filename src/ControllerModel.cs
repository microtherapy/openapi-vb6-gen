namespace OpenApiVb6Gen;

internal sealed class ControllerModel
{
    public required string Tag { get; init; }
    public required string ClassName { get; init; }
    public required string PropertyName { get; init; }
    public List<OperationModel> Operations { get; } = new();
}

internal sealed class OperationModel
{
    public required string OperationId { get; init; }
    public required string VbMethodName { get; init; }
    public required string HttpMethod { get; init; }
    public required string PathTemplate { get; init; }
    public List<ParameterModel> PathParameters { get; } = new();
    public List<ParameterModel> QueryParameters { get; } = new();
    public ParameterModel? Body { get; set; }
    public Vb6Type? Response { get; set; }
    public string? Description { get; init; }
    public string? SkipReason { get; set; }
}

internal sealed class ParameterModel
{
    public required string ApiName { get; init; }
    public required string VbName { get; init; }
    public required Vb6Type Type { get; init; }
    public bool Required { get; init; }
    public string? Description { get; init; }
}
