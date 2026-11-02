# UI palettes

A page can keep its layout, named styles and style pipeline in one XAML file while importing color-only palettes:

```xml
<page xmlns="urn:excelprototype:profiles"
      xmlns:controls="urn:excelprototype:controls"
      version="1"
      paletteSources="Dark.xaml;Light.xaml"
      palette="{Binding Path=Settings.Theme}">
```

Each imported file has a `palette` root in the profiles namespace, `name` and `version="1"`. It contains only named `color` elements. All imported palettes must expose the same nonempty set of resource names; duplicate palette names and duplicate colors are errors. Palette files are resolved relative to the page folder. Missing files, missing resources and unknown active palettes are errors; there is no implicit fallback.

```xml
<palette xmlns="urn:excelprototype:profiles"
         name="Dark"
         version="1">
  <color name="FieldBackground"
         value="#2D2D2D"/>
</palette>
```

Use `{ThemeResource FieldBackground}` in `backColor`, `textColor`, `fontColor`, `borderColor` and `shadowColor`. Pipeline declarations use the same reference, for example `styles="{backColor:{ThemeResource FieldBackground};}"`. The declaration parser removes only the outer braces, splits declarations at semicolons and splits a property at its first colon; resource references contain neither separator. Literal colors remain supported for pages that do not need themes. Palette values use `#RRGGBB`; `#AARRGGBB` is also supported for backgrounds, subject to the existing cell/shape transparency restrictions.

`obj_UiPalette` retains the latest color applications and their order. A repeated application to the same target/property replaces its earlier entry and moves to the end. This preserves the effective cascade, including pipeline rules, read-only field colors and selection overlays, without rerunning layout or changing cell values. Deleted popup shapes are discarded on the next switch. Full rendering resets the retained applications. Pages without palettes do not retain color applications.

Call `obj_PageBase.ApplyTheme(name, diagnostic)` from a page command. It recolors the active page without rebuilding it. After success, update the binding used by the page's `palette` attribute so later full rendering uses the same theme. Changing that binding alone does not automatically apply a theme. Screen updating and events are restored after success or failure; a repaint failure attempts to restore the previous palette and reports the original error. Caption caches are updated so subsequent text refreshes retain the active theme.

Inline caption colors accept palette references, for example `[[color={ThemeResource ThemeLightText}]]...[[/color]]`. The style catalog retains the raw caption and resolves its resources each time it applies the caption, including full/partial rendering and theme changes. Existing literal color markers remain supported. TDB's theme button fills its allocated cell without margins or padding, uses the shared Border resource, and highlights the current theme name while muting the other name and divider.

Modes that opt into palettes must provide the `text.StylePalette*` diagnostic templates used by `obj_UiPalette`, plus `text.StyleThemeContext`, in wsConfig. TDB includes these templates; the generic engine does not supply hardcoded user-facing fallback messages.

TDB exposes `Commands.ToggleTheme` and a button replacing the actions heading at the top of the right panel. Its initial theme comes from `TravelDocumentsBuilder::ui.theme` in wsConfig; switching lasts for the current page session and does not save a preference into the workbook. Button text and the failure template also come from wsConfig.

Run `python PROJECTS_EXCEL/WorkbookUpdater/Test-UiPaletteSources.py` and `pwsh -File Scripts/PowerShell/Test-UiStyleSources.ps1` for static palette and style-source checks. These checks do not execute or compile VBA and do not establish Excel runtime correctness.