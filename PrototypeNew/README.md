# PrototypeNew

Clean-slate sandbox for rebuilding UI loading with object controls.

## Goal
No hardcoded control names in VBA. UI page is described in XML, and runtime builds controls via interface objects.

## Current scope
- Read Dev layout from `PrototypeNew/ui/DevUI.xml`.
- For each declared control, use `type` as root and auto-resolve:
	- control UI: `PrototypeNew/vba/classes/obj_<Type>ControlUI.xml`
	- ViewModel: `obj_<Type>ControlViewModel`
- Build controls through `obj_IControl` + `ex_ControlFactory`.
- Render controls through object ViewModels (currently `obj_HelloWorldControlViewModel`).

## Entry macro
Run:
- `ex_Actions.m_LoadPrototypeNewUi`

This calls `ex_ControlRuntime.m_RenderDevLayout`.

## Runtime API
- `ex_ControlRuntime.m_RenderWorksheet(ws, wsUiPath)`
	- `ws` required: target worksheet.
	- `wsUiPath` optional: relative UI path for this worksheet page.
	- if `wsUiPath` is empty, runtime auto-loads `PrototypeNew/ui/<SheetName>UI.xml`.
