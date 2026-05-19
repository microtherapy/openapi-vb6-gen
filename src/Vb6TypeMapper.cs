using Microsoft.OpenApi.Models;

namespace OpenApiVb6Gen;

internal enum Vb6Kind
{
    String,
    Long,
    Currency,
    Double,
    Boolean,
    Date,
    Variant,
    Enum,
    DtoRef,
    ChilkatJsonObject,
    Collection,
    Binary
}

internal sealed class Vb6Type
{
    public required Vb6Kind Kind { get; init; }
    public required string Declaration { get; init; }

    public bool IsNullable { get; init; }

    public string? DtoClassName { get; init; }
    public string? EnumName { get; init; }

    public Vb6Type? ItemType { get; init; }

    public string? ItemSchemaName { get; init; }

    public bool IsScalar =>
        Kind is Vb6Kind.String or Vb6Kind.Long or Vb6Kind.Currency or Vb6Kind.Double
              or Vb6Kind.Boolean or Vb6Kind.Date or Vb6Kind.Variant or Vb6Kind.Enum;

    public bool IsCollection => Kind == Vb6Kind.Collection;
    public bool IsDtoRef => Kind == Vb6Kind.DtoRef;
}

internal sealed class Vb6TypeMapper
{
    public Vb6Type Map(OpenApiSchema? schema)
    {
        if (schema is null)
            return new Vb6Type { Kind = Vb6Kind.Variant, Declaration = "Variant" };

        if (schema.Reference is { Id: { } refId })
        {
            if (LooksLikeEnumSchema(schema))
            {
                var en = Vb6Naming.EnumName(refId);
                return new Vb6Type { Kind = Vb6Kind.Enum, Declaration = en, EnumName = en };
            }
            var cls = Vb6Naming.ClassName(refId);
            return new Vb6Type
            {
                Kind = Vb6Kind.DtoRef,
                Declaration = cls,
                DtoClassName = cls,
                ItemSchemaName = refId,
                IsNullable = true
            };
        }

        if (schema.Enum is { Count: > 0 } && !string.IsNullOrEmpty(schema.Title))
        {
            var en = Vb6Naming.EnumName(schema.Title);
            return new Vb6Type { Kind = Vb6Kind.Enum, Declaration = en, EnumName = en };
        }

        return schema.Type switch
        {
            "integer" => MapInteger(schema),
            "number" => Scalar(Vb6Kind.Double, "Double", schema.Nullable),
            "boolean" => Scalar(Vb6Kind.Boolean, "Boolean", schema.Nullable),
            "string" => MapString(schema),
            "array" => MapArray(schema),
            "object" => MapObject(schema),
            _ => new Vb6Type { Kind = Vb6Kind.Variant, Declaration = "Variant" }
        };
    }

    private static bool LooksLikeEnumSchema(OpenApiSchema schema)
        => schema.Enum is { Count: > 0 };

    private static Vb6Type MapInteger(OpenApiSchema schema)
    {
        var kind = schema.Format == "int64" ? Vb6Kind.Currency : Vb6Kind.Long;
        var decl = kind == Vb6Kind.Currency ? "Currency" : "Long";
        return Scalar(kind, decl, schema.Nullable);
    }

    private static Vb6Type MapString(OpenApiSchema schema)
    {
        if (schema.Format is "date-time" or "date")
            return Scalar(Vb6Kind.Date, "Date", schema.Nullable);
        return Scalar(Vb6Kind.String, "String", schema.Nullable);
    }

    private Vb6Type MapArray(OpenApiSchema schema)
    {
        var itemType = Map(schema.Items);
        return new Vb6Type
        {
            Kind = Vb6Kind.Collection,
            Declaration = "Collection",
            ItemType = itemType,
            ItemSchemaName = schema.Items?.Reference?.Id
        };
    }

    private static Vb6Type MapObject(OpenApiSchema _)
        => new() { Kind = Vb6Kind.ChilkatJsonObject, Declaration = "ChilkatJsonObject" };

    private static Vb6Type Scalar(Vb6Kind kind, string decl, bool nullable)
    {
        if (nullable)
            return new Vb6Type { Kind = Vb6Kind.Variant, Declaration = "Variant", IsNullable = true };
        return new Vb6Type { Kind = kind, Declaration = decl };
    }
}
