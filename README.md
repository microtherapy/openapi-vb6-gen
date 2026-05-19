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

Requires a .NET 10 SDK (the project targets `net10.0`).

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
| `--tag-filter` | no | Regex; only matching tags emitted. Triggers a reachability prune (drops DTOs/enums no kept op references). |
| `--operation-id-filter` | no | Regex; only matching `operationId`s emitted. Triggers a reachability prune. |
| `--path-filter` | no | Regex; only matching paths emitted. Useful for specs without `operationId`s. Triggers a reachability prune. |
| `--schema-filter` | no | Regex; only matching schemas emitted. **Applied before** the reachability prune — combining with the filters above can produce build failures if a kept op references a DTO this filter dropped. |
| `--clean` | no | Wipe output dir before emitting (preserves `<project>.compat.dll`) |
| `--vb6-exe` | no | Override VB6.EXE path used for the seed bootstrap (default: `C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE`) |
| `--no-seed` | no | Skip the seed bootstrap (use when VB6 isn't installed, e.g. CI). Generated `.vbp` falls back to `CompatibleMode=0` |

## Generated layout

```
{output}/
  {ProjectName}.vbp         ActiveX DLL project
  {ProjectName}.vbw         workspace state
  {ProjectName}.vbg         project group (host + DLL) for IDE step-through
  {ProjectName}.compat.dll  cached compat DLL — VB6 binary-compatibility target (commit to source control)
  cApi.cls                  facade — only class host New's directly
  cEnums.cls                public enums (ActiveX-exposed)
  c{Schema}.cls             one per OpenAPI schema (typed DTO)
  c{Tag}Api.cls             one per OpenAPI tag (operations)
  modGenApi.bas             internal helpers (not part of the ActiveX surface)
```

## One-time host setup

1. Run the generator. It invokes VB6 to build the client DLL, then renames the output to `<project>.compat.dll`. The first run starts from scratch (CompatibleMode=0); every subsequent run anchors against the existing compat DLL so CLSIDs for existing types are preserved, and **new types added between runs are captured in the refreshed compat DLL** — no risk of a type silently getting a fresh CLSID on every build because the anchor was stale.
2. In your host `.vbp`: Project → References → Browse → select `<project>.compat.dll`. Save.

**Commit `<project>.compat.dll` to source control** alongside the generated `.cls` files. It is the on-disk identity of the COM contract; lose it and CLSIDs reshuffle, breaking host references.

If VB6 isn't on the box (e.g. CI codegen, Docker), pass `--no-seed` to skip the VB6 build. The generated `.vbp` then emits `CompatibleMode=0` (if no compat DLL exists) or `CompatibleMode=2` referencing the existing one — but no fresh DLL is produced; the user has to open VB6 to build it.

## Host usage

```vb
Dim api As New OpenApiClient.cApi
api.UnlockChilkat "MY-CHILKAT-LICENSE-KEY"   ' once at startup
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

1. Re-run `dotnet run …` with the same arguments. The generator emits source, invokes VB6 `/make` anchored against the existing `<project>.compat.dll`, and replaces the compat DLL with the freshly built one. Existing types' CLSIDs are preserved; new types' CLSIDs are minted once and locked in by the refresh.
2. The host `.vbp` is **not** edited. As long as `<project>.compat.dll` is present and committed, the host reference stays valid.
3. If a method signature changed incompatibly, VB6's binary-compatibility logic warns and the build may fail with a vtable-mismatch error in `_build.log`. Resolve by reverting the breaking change, or — for an intentional break — delete `<project>.compat.dll` and re-run (fresh CLSIDs, host must be re-referenced).

Open `{ProjectName}.vbg` instead of `.vbp` to load host + client together for step-through debugging in the IDE.

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
| enum | `Public Enum e{Name}` in `cEnums.cls` |
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
