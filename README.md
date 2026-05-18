# openapi-vb6-gen

Generates a strongly-typed VB6 API client from an OpenAPI 3 spec.

Output is a complete VB6 ActiveX DLL sub-project (`.vbp` + `.cls` + `.bas`) that the host VB6 application references once. Regenerating the client does not modify the host `.vbp`.

The runtime depends on the [Chilkat ActiveX](https://www.chilkatsoft.com/) component for HTTP and JSON.

## Why

VB6 has no first-class HTTP/JSON tooling and no equivalent of NSwag or OpenAPI Generator. Hand-written API plumbing in VB6 rots whenever the upstream API changes. This generator emits a typed VB6 client whose surface is regenerated from the OpenAPI spec, so renames and new endpoints surface as compile errors instead of runtime "method not found".

## Design choices

| | |
|---|---|
| **Output shape** | ActiveX DLL sub-project. Host references the compiled DLL once; regenerating the client never touches the host `.vbp`. |
| **DTO style** | Strongly-typed `.cls` per OpenAPI schema with `FromJson` / `ToJson`. Typos like `c.CleintName` fail at compile time, not at runtime. |
| **Debugging** | Generated `.vbg` (project group) opens host and client in the same VB6 IDE session for native step-through. |
| **HTTP / JSON** | Chilkat ActiveX. Other libraries (WinHTTP / MSXML / VBA-JSON) could be plugged in by replacing `modGenApi.bas` template. |

## Build

```powershell
dotnet build OpenApiVb6Gen.csproj
```

## Run

```powershell
dotnet run --project OpenApiVb6Gen.csproj -- `
    --input http://localhost:5000/swagger/v1/swagger.json `
    --output C:\code\my-api-client `
    --main-vbp "C:\code\my-app\MyApp.vbp"
```

### Arguments

| Flag | Required | Notes |
|---|---|---|
| `--input` | yes | Path or http(s) URL to `swagger.json` |
| `--output` | yes | Target directory for generated files |
| `--project-name` | no | DLL/typelib name. Default `OpenApiClient` |
| `--main-vbp` | no | Host `.vbp` to copy Chilkat `Reference=` / `Object=` lines from |
| `--tag-filter` | no | Regex; only matching tags emitted |
| `--schema-filter` | no | Regex; only matching schemas emitted |
| `--clean` | no | Wipe output dir before emitting |

## Generated layout

```
{output}/
  {ProjectName}.vbp     ActiveX DLL project
  {ProjectName}.vbw     workspace state
  {ProjectName}.vbg     project group (host + DLL) for IDE step-through
  cApi.cls              facade — only class host New's directly
  c{Schema}.cls         one per OpenAPI schema (typed DTO)
  c{Tag}Api.cls         one per OpenAPI tag (operations)
  modGenApi.bas         internal helpers + Public Enum types
```

## One-time host setup

1. Open `{output}\{ProjectName}.vbp` in VB6 IDE.
2. If the Chilkat reference wasn't copied from `--main-vbp`: Project → References → Chilkat ActiveX.
3. File → Make `{ProjectName}.dll`.
4. In your host `.vbp`: Project → References → Browse → select the compiled DLL. Save.

## Host usage

```vb
Dim api As New OpenApiClient.cApi
api.Init "https://my-api.example.com/", myBearerToken

' Single object
Dim c As OpenApiClient.cClient
Set c = api.Clients.GetById(123)
Debug.Print c.ClientName    ' compile-checked

' Collection of typed objects
Dim list As Collection
Set list = api.Clients.GetAll(agencyFirmId)
Dim item As OpenApiClient.cClient
For Each item In list
    Debug.Print item.PK_Client, item.ClientName
Next
```

## Regeneration workflow

1. Re-run `dotnet run …` with the same arguments.
2. Open `{ProjectName}.vbp` in VB6 IDE, File → Make DLL. If signatures changed, click through the binary-compat dialog.
3. If signatures changed: in host VB6, Project → References → uncheck `{ProjectName}` → recheck. Save.
4. The host `.vbp` is never edited by hand.

Open `{ProjectName}.vbg` instead of either `.vbp` to load host + client together for step-through debugging.

## Type mapping

| OpenAPI | VB6 |
|---|---|
| `string` | `String` |
| `string` date / date-time | `Date` (helpers `modGenApi.Iso` / `ParseIso`) |
| `integer` int32 | `Long` |
| `integer` int64 | `Currency` (closest 64-bit; not full int64 range) |
| `number` | `Double` |
| `boolean` | `Boolean` |
| `nullable` primitive | `Variant` (Empty = null) |
| enum | `Public Enum e{Name}` in `modGenApi.bas` |
| `$ref` to object | `c{Schema}` typed class |
| array of `$ref` | `Collection` of `c{Schema}` |
| array of primitive | `Collection` of native type |
| inline object (no name) | `ChilkatJsonObject` |

`Collection` items are Variant at the language level — assign to a typed Dim'd variable to recover compile-time checks:

```vb
Dim c As cClient: Set c = list(1)   ' typed
```

## Naming

- Schemas: `c{PascalCase(schemaName)}` (e.g. `cClient`, `cPlanMarketSub`)
- Controllers: `c{PascalCase(tagName)}Api` (e.g. `cClientsApi`)
- Enums: `e{PascalCase(name)}`, members `e{Name}_{Value}`
- Property names: preserve underscores from JSON keys (`PK_Client` stays `PK_Client`)
- Method names: prefer `operationId`; fallback is `{Verb}{Path}` with version segments (`v0.17`, `v2`, etc.) stripped, trailing path-param converted to `By{Param}`
- Verbs map: GET→Get, POST→Create, PUT→Update, DELETE→Delete, PATCH→Patch
- Reserved VB6 words (`Open`, `ReadOnly`, `Name`, `Time`, …) get a trailing `_` on parameters

## Non-goals

- multipart/form-data, octet-stream, file upload/download
- OAuth/OIDC flows (Bearer token only)
- streaming responses
- retry / rate-limit
- automatic DLL compilation (requires VB6 IDE)
- automatic editing of the host `.vbp` (intentional — host file untouched)

Operations whose request body is non-JSON are emitted as comments only, with a `SKIPPED:` reason.

## Status

Early. Tested against multi-controller, multi-schema specs; not yet round-trip tested through a compiled DLL build in VB6 IDE. PRs and issue reports for spec-shape edge cases welcome.

## License

MIT
