# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A .NET 8 console tool that reads an OpenAPI 3 spec and emits a VB6 ActiveX DLL sub-project (`.vbp` + `.cls` files + `modGenApi.bas`) implementing a typed HTTP client. Output is VB6 source. Verification of generated code requires either the VB6 IDE or the VB6 command line — see "Compiling the generated client" below.

`README.md` covers user-facing usage (flags, generated layout, host integration). This file covers internals and the VB6/Chilkat 11 gotchas.

## Build and run

```powershell
dotnet build OpenApiVb6Gen.csproj
dotnet run --project OpenApiVb6Gen.csproj -- --input <swagger.json|url> --output <dir> [--main-vbp <host.vbp>] [--clean] [--tag-filter <regex>] [--schema-filter <regex>] [--project-name <name>] [--vb6-exe <path>] [--no-seed]
```

`TreatWarningsAsErrors=true` and `Nullable=enable` are set — any new C# warning fails the build. There is **no test project**.

If the dev box has a newer .NET runtime than `net8.0` (e.g. only `net10`), set `DOTNET_ROLL_FORWARD=Major` for that shell.

## Compiling the generated client

Two paths from `.vbp` to `.dll`:

1. **VB6 IDE** — File → Make `<ProjectName>.dll`. Surfaces errors with line numbers in a dialog.
2. **Command line** — what the workspace's `wixproj`s use via the `Microtherapy.MsBuild.VB6` task. Or just call VB6 directly:
   ```
   "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make <file>.vbp /out build.log
   ```
   Exit code 0 = success, non-zero = compile error. The log file lists `Compile Error in File 'X.cls', Line N : <message>`. **VB6 line numbers in compile errors are often off by several lines — search the file for the cited token, not the line.** A common pattern is a parse failure on a later line being attributed to an earlier `If` block opening.
   
   `"Error accessing the system registry"` in the log is **not** a build failure. VB6's post-make step tries to self-register the DLL into HKLM; without elevation it prints that and continues. The `.dll` is still written. Register separately with `regsvr32 <dll>`.

## Architecture

Two-stage pipeline. Parse → model → emit. The boundary is deliberate: emitters never look at `OpenApiSchema` directly.

1. **Parse** (`SpecLoader.cs`) — Microsoft.OpenApi reads JSON/YAML into `OpenApiDocument`. HTTP and local file inputs are both supported. Parse errors are logged to stderr but don't abort.
2. **Map to language-neutral models** (`Program.cs` → `BuildSchemaModels` / `BuildControllerModels`):
   - `DtoModel` / `DtoPropertyModel` / `EnumModel` / `EnumMember` (`src/SchemaModel.cs`)
   - `ControllerModel` / `OperationModel` / `ParameterModel` (`src/ControllerModel.cs`)
   - Every OpenAPI type is funneled through `Vb6TypeMapper.Map(OpenApiSchema)` → `Vb6Type { Kind, Declaration, … }` (`src/Vb6TypeMapper.cs`). `Vb6Kind` is the discriminator emitters branch on.
3. **Emit** one file per concern, all going through `Vb6Writer` (CRLF line endings, indent stack):
   - `DtoEmitter` — one `.cls` per schema, with `FromJson` / `ToJson`
   - `ControllerEmitter` — one `c{Tag}Api.cls` per OpenAPI tag
   - `FacadeEmitter` — single `cApi.cls` that host code instantiates; holds `BaseURL` / `BearerToken`, lazy-inits controllers, and exposes the ActiveX-visible utilities `UnlockChilkat(licenseKey)` and `SaveBytesToFile(data, path)`. These are on `cApi` (not `modGenApi.bas`) precisely because `.bas` symbols are not reachable across the ActiveX boundary.
   - `EnumEmitter` — single `cEnums.cls` holding every `Public Enum`. **Must be a `.cls`, not the `.bas`** — see "ActiveX export rules" below.
   - `HelperEmitter` — `modGenApi.bas`: HTTP core, JSON scalar getters, ISO date helpers, per-DTO `GetJsonAs_/GetJsonArrayAs_/PostJsonAs_/PostJsonArrayAs_/PutJsonAs_/PutJsonArrayAs_/LoadDto_/LoadList_/AppendDto_/AppendList_`, and primitive variants of `LoadList_/GetJsonArrayAs_/AppendList_`.
   - `VbpEmitter` — emits `.vbp`, `.vbw`, `.vbg`. Reads the host `.vbp` (if `--main-vbp` given) and copies any `Reference=` / `Object=` lines mentioning `Chilkat` into the generated `.vbp`. Without `--main-vbp` a TODO comment is left and the user must add the Chilkat reference in the IDE. If `CompatibleExePath` is set on `VbpEmitterInputs` and the file exists, emits `CompatibleMode=2` + `CompatibleEXE32=<path>`; otherwise `CompatibleMode=0`.

To add a new output file kind: write a new emitter and wire it from `App.EmitAll`. To change the C# → VB6 type mapping for *all* outputs: edit `Vb6TypeMapper.Map`.

## Type system invariants

`Vb6Type` is canonical; emitters use its `Kind` and `Declaration` (never the raw OpenAPI type). When extending:

- `Kind = Collection` always has `ItemType` set. Collections of `$ref` schemas also carry `ItemSchemaName` (raw OpenAPI name) so `HelperEmitter` can synthesize a typed loader.
- `Kind = DtoRef` carries `DtoClassName` (`c{PascalCase}`) and `ItemSchemaName` (raw OpenAPI name). `IsNullable = true` for DTO refs because VB6 object refs are always nullable.
- `Kind = Enum` carries `EnumName` (`e{PascalCase}`). Enums are emitted into `cEnums.cls` (see below), **not** `modGenApi.bas`.
- Nullable primitives collapse to `Variant` (VB6 has no `int?`). `IsNullable` on the resulting type is `true`.
- `int64` → `Currency`. This is a known precision compromise — Currency holds 64 bits but with implicit ×10000 scaling, so the usable integer range is ±922,337,203,685,477. Do not "fix" by switching to `LongLong` (VB6 has no LongLong).
- `type: object` with no `$ref` becomes `Kind = ChilkatJsonObject` — an escape hatch passing the raw Chilkat JSON node through. Inline schemas are intentionally **not** lifted into anonymous DTO classes.
- Treating reference-typed properties (DTO ref, Collection, ChilkatJsonObject) requires `Set` assignment and `Property Get` + `Property Set` — value-typed properties use plain `Property Get` + `Property Let`. Several emitters branch on `p.Type.IsCollection || p.Type.IsDtoRef || p.Type.Kind == Vb6Kind.ChilkatJsonObject` — keep that triple in sync.

## ActiveX export rules

The output is an `Type=OleDll` (ActiveX DLL). The DLL's public surface is **only** what's reachable from `.cls` files with `Attribute VB_Exposed = True`. This has two important consequences:

- **`Public Enum` in a `.bas` module is NOT exposed to clients of the DLL.** A host that adds the compiled DLL as a Reference cannot see those enums; worse, a public class property whose return type is such an enum fails to compile with `Private Enum and user defined types cannot be used as parameters or return types for public procedures...`. So all enums live in `cEnums.cls` (emitted by `EnumEmitter`) with `VB_Exposed = True`, `VB_Creatable = False`, `VB_PredeclaredId = True`, `VB_GlobalNameSpace = False`.
- **Same applies to public UDTs** (not used today, but would need the same treatment).

`modGenApi.bas` remains a `.bas` deliberately — its functions are DLL-internal helpers, not part of the public ActiveX surface. Anything host code needs to call lives on a `.cls` with `VB_Exposed = True`.

## VB6 emission constraints

- All generated text must be CRLF. `Vb6Writer.Line` enforces this; do not bypass it with raw `_sb.AppendLine`.
- `.cls` files need the exact `VERSION 1.0 CLASS … END` + five `Attribute VB_…` lines header that `DtoEmitter`, `ControllerEmitter`, `FacadeEmitter`, and `EnumEmitter` all duplicate. If you change one, change the others (no shared helper today — by design, to keep emitters independent).
- `.bas` files start with `Attribute VB_Name = "modName"` then `Option Explicit`. See `HelperEmitter`.
- `.vbp` lines are order-sensitive. `Module=` lines must come before `Class=` lines; `Startup=` and the `ExeName32=` block are required.
- VB6 reserved words (see `Vb6Naming.ReservedWords`) become parameter/identifier collisions. `SafeIdentifier` / `SafeParameter` / `PropertyName` append `_` to deconflict. When adding a new name slot, route it through one of these helpers — including `ControllerModel.PropertyName` (otherwise tags like `Time` collide with VB6 intrinsics).

## ProgID 39-character limit

VB6 ActiveX DLLs are subject to a hard limit of **39 characters total for `<LibraryName>.<ClassName>`** (the COM ProgID). Exceeding this fails the link step with `Programmatic ID string too long '...'`.

The generator enforces this on two paths:

- **Controller classes** (`BuildControllerModels` in `Program.cs`): `maxClassLen = 39 - projectName.Length - 1` (minus 1 for the ProgID dot), truncated + de-duplicated via `TruncateUnique`. The whole `"c"+Tag+"Api"` string is what gets truncated.
- **DTO / enum classes** (`Vb6Naming.BuildSchemaAliasMap`, seeded once from `App.Run`): `maxAliasLen = 39 - projectName.Length - 2` (minus 1 for the dot, minus 1 for the `c`/`e` prefix). The map is keyed by the raw OpenAPI schema name and stored as static state on `Vb6Naming`; `ClassName`/`EnumName` look up the alias on every call, so `Vb6TypeMapper.Map` automatically picks up truncated names when resolving `$ref`s. For namespaced names like `Foo.Bar.Baz` the alias starts as the last segment (`Baz`) and walks back through the namespace on collision (`Bar_Baz`, `Foo_Bar_Baz`).

Callers should still pick short `--project-name` (≤8 chars is comfortable; the default `OpenApiClient` is 13 which leaves only 24 chars for DTO names and may force more truncation than ideal).

Only schemas declared in `Components.Schemas` get an alias entry — an inline schema referenced by `$ref` from outside components would bypass the truncation. The specs we've tested don't hit this case.

## Chilkat 11 API gotchas

Generated code calls Chilkat 11 ActiveX (typelib `{06FB4061-5E43-42E0-8A6E-4A1C869E59AF}`, ProgID prefix `Chilkat_11_0_0.*`). Several methods that *seem* like they should exist do not. When emitting helpers, only call methods listed in the authoritative reference at https://www.chilkatsoft.com/refdoc/xChilkatJsonObjectRef.html and https://www.chilkatsoft.com/refdoc/xChilkatJsonArrayRef.html.

Notable absences and workarounds:

| What you'd expect | Reality in Chilkat 11 | Workaround |
|---|---|---|
| `JsonArray.NumberAt(i)` | Doesn't exist | `CDbl(Val(arr.StringAt(i)))` |
| `JsonObject.NumberOf(name)` | Doesn't exist | `CDbl(Val(obj.StringOf(name)))` |
| `JsonObject.AppendObject(name, child)` (2-arg) | Only `AppendObject(name)` (1-arg), returns a new empty child | `parent.AddObjectAt -1, name` then `parent.ObjectOf(name).Load child.Emit()` |
| `JsonObject.AddArrayCopyAt(name, idx, arr)` | Signature differs | `parent.AddArrayAt -1, name` then populate via `parent.ArrayOf(name)` |
| `JsonArray.AddObjectAt(idx, name)` | Arrays have no named slots; the `(idx, name)` overload doesn't exist on arrays | `arr.AddObjectAt -1` then `arr.ObjectAt(arr.Size - 1).Load child.Emit()` |

Confirmed-present methods we *do* rely on: `UpdateString`, `UpdateInt`, `UpdateBool`, `UpdateNull`, `UpdateNumber` (takes a string), `IsNullOf`, `StringOf`, `IntOf`, `BoolOf`, `Load`, `Emit`, `ObjectOf`, `ArrayOf`, `AddObjectAt`, `AddArrayAt`; on arrays: `Size`, `StringAt`, `IntAt`, `BoolAt`, `ObjectAt`, `AddStringAt`, `AddIntAt`, `AddBoolAt`, `AddNumberAt`, `AddObjectAt(idx)`, `Load`.

When introducing a new Chilkat call, verify it against the Chilkat 11 ActiveX reference docs (linked above) before emitting it. Methods that exist in Chilkat 9.5 sometimes don't in 11 (and vice versa), and the .NET / Python / C++ docs aren't reliable proxies for the ActiveX surface — the COM typelib is the source of truth.

## Per-verb response handling

Each HTTP verb's response is generated through a `*Call` switch in `ControllerEmitter` that branches on `Vb6Kind`. **All four (GET / POST / PUT / DELETE / PATCH) must stay in sync** when adding a new response shape:

- `GetCall` → `Vb6Kind.Collection` of DTO → `GetJsonArrayAs_<cls>`; primitive → `GetJsonArrayAs_<suffix>`; etc.
- `PostCall` and `PutCall` are structurally identical (same DTO helpers per kind, just with the `Post`/`Put` prefix). When the spec declares a 2xx body schema, the controller must pick the typed helper; otherwise it falls through to `*ReturnString`.
- `EmitDelete` only handles primitive return shapes; a DELETE with a typed JSON body comes through as a comment.

Per-DTO `Put*` helpers were added retroactively — early versions emitted `Set foo = PutJsonReturnString(...)` for Collection-returning PUTs, which fails to compile (String → Collection mismatch). If you add a new verb, also add the corresponding per-DTO helpers in `HelperEmitter.WriteDtoLoaders`.

## Spec quirks that aren't generator bugs

- **Empty 2xx response schemas.** ASP.NET endpoints without `[ProducesResponseType(typeof(T), 200)]` emit `200: {description: "OK"}` with no content schema. The generator treats those as void (`Sub`, not `Function`). To get typed returns, fix the server attribute — the generator deliberately doesn't guess `ChilkatJsonObject` for missing schemas because that would mask real spec bugs.
- **One-tag-per-endpoint specs.** Some swaggers produce 50+ tiny "controllers" each with a single operation. The generator emits one `c{Tag}Api.cls` per tag as designed — the resulting `cApi` facade just becomes very wide. Not a generator bug.

## Naming rules (single source of truth: `src/Vb6Naming.cs`)

- Schemas → `c{PascalCase(schemaName)}` (e.g. `cClient`). Enums → `e{PascalCase}`, members `{EnumName}_{Value}`.
- Property names preserve underscores from JSON keys (`PK_Client` stays `PK_Client`). This is deliberate — host code references DB-derived names that already carry underscores.
- Operation names: if `operationId` is `Controller_Method` (NSwag convention), the prefix before `_` is stripped. Otherwise `{Verb}{StaticPathSegments}By{TrailingParam}` is synthesized, with version-like segments (`v1`, `v0.17`) dropped. See `Vb6Naming.OperationName`.
- `ControllerModel.PropertyName` goes through `SafeIdentifier` (the reserved-word collision check) — *not* raw `PascalCase`. Tags like `Time` would otherwise shadow the VB6 intrinsic.

## Non-goals (codified in `Program.BuildOperation`)

- Multipart, form-data, octet-stream request bodies → operation is emitted as a `' SKIPPED: …` comment, not executable code.
- OAuth flows beyond Bearer token, streaming responses, retries, rate-limits, file upload/download — out of scope.
- The host `.vbp` is **never** edited. Users wire the compiled DLL via the IDE (or via a reg-free manifest for production).

## Editing checklist for codegen changes

When changing emitter output, regenerate against a real spec and:

1. Confirm `dotnet build` passes (warnings-as-errors).
2. Run `VB6.EXE /make` (see "Compiling the generated client") against the generated `.vbp`. The IDE alone is not enough for a CI loop — command-line build is what catches regressions fastest.
3. If the change touches `Vb6Type` shape or `Vb6Kind` cases, grep all emitters — `DtoEmitter`, `ControllerEmitter`, `FacadeEmitter`, `EnumEmitter`, `HelperEmitter`, `VbpEmitter` — for switch/branch on `Kind` to ensure the new case is handled.
4. If you add a new Chilkat method call, confirm it exists in the Chilkat 11 ActiveX reference docs at <https://www.chilkatsoft.com/refdoc/>.

## Binary-compatibility seeding

VB6 mints fresh CLSIDs on every `/make` unless the project file declares a `CompatibleEXE32=<path>` to compare against. Without that, every regen breaks the host's references. The generator anchors against `<project>.compat.dll` and **rebuilds it on every run**, so newly-added types get their CLSIDs captured immediately rather than drifting on each rebuild.

Flow (`Program.BuildCompatDll`, invoked from `EmitAll` unless `--no-seed`):

1. Write `.vbp` / `.vbw` with `CompatibleEXE32=<...compat.dll>` if one already exists, else `CompatibleMode=0` for the very first run.
2. Invoke VB6 via `Vb6Bootstrap.RunMake` — `VB6.EXE /make <project>.vbp /out _build.log`, 5-minute timeout, kill on hang.
3. If `<project>.dll` doesn't exist after the run, surface the build log and abort. Pre-existing `<project>.compat.dll` is left untouched.
4. Replace `<project>.compat.dll` with the freshly built `<project>.dll` — this snapshots the current type set so the *next* regen anchors against types added in *this* regen.
5. `EmitAll` re-emits `.vbp` / `.vbw` with `CompatibleEXE32=<...compat.dll>` (now guaranteed to exist) so the IDE workflow uses the same anchor.

The .vbp's `ExeName32` is `<project>.dll` while `CompatibleEXE32` is `<project>.compat.dll`. These are deliberately different files: VB6 can build `<project>.dll` without overwriting the anchor it's reading from, and the generator does the swap after the build succeeds.

Why rebuild every time, not just on first run? Without the refresh, types added between regens get a fresh CLSID on every subsequent build (the stale compat DLL doesn't list them, so VB6 has nothing to preserve). The refresh locks new CLSIDs in on the regen that introduces them, so they're stable from that point forward.

Contracts:

- **`<project>.compat.dll` must be source-controlled** by the user. It *is* the on-disk identity of the COM contract — there is no separate metadata file. Lose the DLL and CLSIDs reshuffle on the next build.
- **`--clean` preserves it** (`Program.cs` filename allowlist) — wiping the rest of the output dir is safe.
- **`--no-seed`** skips the VB6 build entirely. Use in CI without VB6 installed. The `.vbp` is still emitted with `CompatibleEXE32` pointing at the existing compat DLL (if any), but no fresh DLL is produced.
- **Deleting `.compat.dll`** forces the next run to seed from scratch with fresh CLSIDs. Host needs to be re-referenced.

There used to be a `TlbPatcher` / `GuidStore` step that rewrote the seed DLL's CLSIDs to deterministic values cached in a `.gen.json` file. It was removed: once the binary is committed to source control, the JSON adds no information not already encoded in the DLL, and the patching path doesn't capture CLSIDs that VB6 mints incrementally between seeds (for new types added later). The simpler "VB6 picks, we commit, regen refreshes" model is what's left.

## Deployment notes

The generated DLL targets **Chilkat 11.x ActiveX** (typelib `{06FB4061-5E43-42E0-8A6E-4A1C869E59AF}`). Consumers must have it registered (`regsvr32 ChilkatAx-win32.dll`) or declared in a side-by-side manifest.

For production reg-free COM (host `.exe` + `.exe.manifest`), the generated DLL's typelib entries must be added to the host manifest as a `<file name="…\GeneratedClient.dll"><typelib …/><comClass …/>…</file>` block. The CLSIDs and TLBID can be read from the registry after `regsvr32` (under `HKLM\SOFTWARE\WOW6432Node\Classes\CLSID\…` on 64-bit Windows for the 32-bit DLL), or extracted directly from `<project>.compat.dll`'s typelib resource. With binary-compat seeding active and the compat DLL committed, those CLSIDs **do not change across regens** — the manifest is authored once and survives every codegen. If `<project>.compat.dll` is deleted (or the user runs with `--no-seed`), CLSIDs become volatile again and the manifest must be regenerated whenever the DLL is rebuilt.
