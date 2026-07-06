# Excel-DNA NativeAOT add-ins

NativeAOT support (the `ExcelDna.AddIn.NativeAOT` package) compiles an Excel-DNA add-in to a native `.xll`
that runs without the .NET Desktop Runtime installed. It is a preview / specialized path, distinct from the
ordinary managed `ExcelDna.AddIn` package.

This document covers the areas that differ from a managed add-in and the points that have caused confusion
(see https://groups.google.com/g/exceldna/c/qwQLiufM5d4).

## Ribbon

For a NativeAOT add-in a ribbon class implements `ExcelDna.Integration.CustomUI.IExcelRibbon` (it does **not**
derive from `ExcelRibbon`):

```csharp
using ExcelDna.Integration.CustomUI;

public class RibbonController : IExcelRibbon
{
    public string GetCustomUI(string ribbonId) => /* ribbon XML */;
}
```

Ribbon callbacks are dispatched by name. The callback signatures use Excel-DNA wrapper types instead of the
COM interfaces used by managed add-ins:

| Ribbon XML callback | NativeAOT callback signature |
|---|---|
| `onAction` (button) | `void OnButton(RibbonControl control)` |
| `onAction` (toggleButton/checkBox) | `void OnToggle(RibbonControl control, bool pressed)` |
| `onChange` (editBox/comboBox) | `void OnChange(RibbonControl control, string text)` |
| `onAction` (dropDown) | `void OnSelect(RibbonControl control, string selectedId, int selectedIndex)` |
| `getLabel` / `getScreentip` / `getText` | `string GetLabel(RibbonControl control)` |
| `getVisible` / `getEnabled` / `getPressed` | `bool GetEnabled(RibbonControl control)` |
| `getItemCount` | `int GetItemCount(RibbonControl control)` |
| `getItemLabel` | `string GetItemLabel(RibbonControl control, int index)` |
| `getImage` | `byte[] GetImage(RibbonControl control)` (raw image bytes) or an `imageMso` string |
| `onLoad` | `void OnLoad(RibbonUI ribbon)` |

`RibbonControl` exposes the control's `Id` and `Tag`. Callbacks that return a value (the `get…` callbacks)
have their result marshalled back to Excel.

### Dynamic ribbons and `onLoad`

To update the ribbon at runtime, add `onLoad='OnLoad'` to the `customUI` element and keep the `RibbonUI`
that is passed to the callback. `RibbonUI` exposes `Invalidate()`, `InvalidateControl(id)`,
`InvalidateControlMso(id)`, `ActivateTab(id)`, `ActivateTabMso(id)` and `ActivateTabQ(id, ns)`:

```csharp
public class RibbonController : IExcelRibbon
{
    private RibbonUI? _ribbon;
    private bool _enabled;

    public string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='OnLoad'>
  <ribbon><tabs><tab id='t' label='Demo'><group id='g' label='Demo'>
    <button id='toggle' label='Toggle' onAction='OnToggle'/>
    <button id='action' getEnabled='GetEnabled' label='Run' onAction='OnRun'/>
  </group></tab></tabs></ribbon>
</customUI>";

    public void OnLoad(RibbonUI ribbon) => _ribbon = ribbon;

    public void OnToggle(RibbonControl control)
    {
        _enabled = !_enabled;
        _ribbon?.InvalidateControl("action");   // re-query GetEnabled
    }

    public bool GetEnabled(RibbonControl control) => _enabled;

    public void OnRun(RibbonControl control) { /* ... */ }
}
```

### Ribbon images (`loadImage`)

As with managed Excel-DNA add-ins, image loading is opt-in: you must add `loadImage='LoadImage'` to the
`customUI` element. `LoadImage` is a built-in Excel-DNA callback - you do **not** implement it yourself.
With `loadImage='LoadImage'` set, the `image` attribute of a control names an **embedded manifest resource**
in the add-in assembly:

```xml
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' loadImage='LoadImage'>
  ...
  <button id='b' label='Go' image='MyAddIn.app.ico' onAction='OnGo'/>
  ...
</customUI>
```

```xml
<!-- in the .csproj -->
<ItemGroup>
  <EmbeddedResource Include="app.ico" />
</ItemGroup>
```

The resource name is the default manifest resource name (usually `<RootNamespace>.<FileName>`).

## IntelliSense

IntelliSense **works** for a NativeAOT add-in - but it is deployed as a separate add-in, not referenced into
the project.

### Do this: deploy the standalone IntelliSense loader alongside the add-in

Ship `ExcelDna.IntelliSense.xll` (32-bit) / `ExcelDna.IntelliSense64.xll` (64-bit) - from the
[ExcelDna.IntelliSense releases](https://github.com/Excel-DNA/IntelliSense/releases) - next to your add-in and
load it (open it as an add-in, or auto-load it). It runs as its own ordinary (JIT) add-in and overlays the
function help.

This works because IntelliSense discovery does not depend on the add-in being managed, and needs no registration
step in your add-in. Every Excel-DNA add-in, including a NativeAOT one, registers a hidden `RegistrationInfo_<guid>`
function during `AutoOpen` that returns its function/argument metadata (driven by your `[ExcelFunction]` /
`[ExcelArgument]` descriptions). The standalone IntelliSense server monitors add-in loads and probes each one
through that channel, so it picks up a NativeAOT add-in's functions automatically. The overlay itself
(WinForms / UI Automation) runs entirely inside the separate IntelliSense add-in; and because the NativeAOT add-in
carries no .NET runtime, there is no in-process runtime conflict.

So: write good `[ExcelFunction(Description=...)]` and `[ExcelArgument(Description=...)]` metadata, and the
standalone IntelliSense loader will surface it.

### Do not: reference the ExcelDna.IntelliSense package into the AOT project

Adding the `ExcelDna.IntelliSense` **NuGet package** to a NativeAOT project does not work:

* It is not AOT-compatible (WinForms / UI Automation), so it cannot be compiled into the native image.
* It brings in the standard (non-NativeAOT) `ExcelDna.Integration` assembly, which replaces the NativeAOT build -
  types such as `IExcelRibbon`, `RibbonControl` and `RibbonUI` then disappear and the project no longer compiles.

The build emits warning `EXCELDNA001` if the `ExcelDna.IntelliSense` package is referenced from a NativeAOT add-in.

## 32-bit (win-x86) and 64-bit (win-x64)

NativeAOT output is runtime-identifier specific, so a single add-in build targets one architecture. Both
`win-x64` (64-bit Excel) and `win-x86` (32-bit Excel) are supported.

* `-r win-x64` produces `<name>-AddIn64.xll` (load this in 64-bit Excel).
* `-r win-x86` produces `<name>-AddIn.xll` (load this in 32-bit Excel).

```
dotnet publish -c Release -r win-x64
dotnet publish -c Release -r win-x86
```

The correct loader stub for the target architecture is selected automatically from the `RuntimeIdentifier`.
Other runtime identifiers (e.g. `win-arm64`) are not supported and produce a build error.
