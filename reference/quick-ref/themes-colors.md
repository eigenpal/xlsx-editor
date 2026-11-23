# Themes and colors quick reference

A summary of the workbook theme and of the two color systems in an XLSX file, written for this project in the DOCX Editor format. Cells, fonts, fills, and borders use SpreadsheetML colors (`CT_Color` in `sml.xsd`). Drawings and charts use DrawingML colors (`a:srgbClr`, `a:schemeClr`, and the others), which `packages/ooxml-core/src/drawingml-color.ts` resolves.

"Schema" means the Transitional schemas in `reference/ecma-376/part4/transitional-schemas/`. "Measure" means that the standard does not settle the behavior: confirm it on a fixture and its reference PDF before code depends on it, and record the result here.

## The theme part

A workbook has one theme part, usually `/xl/theme/theme1.xml`, linked from the workbook part. Its root is `a:theme` (`CT_OfficeStyleSheet`, DrawingML, `dml-main.xsd`). `a:themeElements` (`CT_BaseStyles`) holds the three schemes:

```xml
<a:theme name="...">
  <a:themeElements>
    <a:clrScheme name="...">   <!-- 12 colors -->
    <a:fontScheme name="...">  <!-- major and minor fonts -->
    <a:fmtScheme name="...">   <!-- fills, lines, effects for drawings -->
  </a:themeElements>
</a:theme>
```

- `a:clrScheme` (`CT_ColorScheme`) holds, in schema order: `a:dk1`, `a:lt1`, `a:dk2`, `a:lt2`, `a:accent1` to `a:accent6`, `a:hlink`, `a:folHlink`. Each holds one color of `EG_ColorChoice`: `a:scrgbClr`, `a:srgbClr`, `a:hslClr`, `a:sysClr`, `a:schemeClr`, or `a:prstClr` (schema). Files usually hold `a:srgbClr/@val` or `a:sysClr` (`@val` names a system color, `@lastClr` holds its last sRGB value).
- `a:fontScheme` (`CT_FontScheme`) holds `a:majorFont` (headings) and `a:minorFont` (body). Each (`CT_FontCollection`) has `a:latin`, `a:ea`, and `a:cs` typefaces, and optional `a:font` entries for each script.
- A file without a theme part is valid. Then theme colors and theme fonts have no value: fall back to the explicit values and record the fallback.

## Theme fonts in cells

A SpreadsheetML font (`x:font` in `styles.xml`, or `x:rPr` in a rich text run) can carry `x:scheme/@val` (`ST_FontScheme`):

- `minor`: the font follows the theme's minor font.
- `major`: the font follows the theme's major font.
- `none`, or no `x:scheme`: the font is the name in `x:name/@val`.

`x:name/@val` is also present on theme fonts, as the name that the writer resolved. Measured on fixture `colors-fonts` (`fixtures/measured.json`): a font with `x:scheme val="minor"` or `val="major"` and `x:name val="Courier New"` shows the theme's minor or major typeface. The theme wins.

## SpreadsheetML colors (`CT_Color`)

`x:color`, `x:fgColor`, `x:bgColor`, and the other color elements of `styles.xml` have these attributes (schema, `CT_Color`, Part 1, 18.8.19). No attribute is required:

| Attribute | Meaning |
| --- | --- |
| `rgb` | An ARGB hex value, for example `FFFF0000`. The first byte is alpha |
| `theme` | A 0-based index of a theme color |
| `indexed` | An index into the indexed color palette |
| `auto` | The automatic color |
| `tint` | A lightness change from -1.0 to 1.0, `0.0` by default. It applies to the color that the other attribute gives |

Resolve the base color from `rgb`, `theme`, `indexed`, or `auto`, then apply `tint`. Every attribute value is attacker-controlled: an index out of range is invalid input. Fall back to the default color and record it, never read past a list.

### Theme color index

Files in common use store black body text as `theme="1"` and white backgrounds as `theme="0"`. So the index reads the first four scheme colors in the order light 1, dark 1, light 2, dark 2, not in schema order:

| `theme` | Scheme color |
| --- | --- |
| 0 | `a:lt1` |
| 1 | `a:dk1` |
| 2 | `a:lt2` |
| 3 | `a:dk2` |
| 4 to 9 | `a:accent1` to `a:accent6` |
| 10 | `a:hlink` |
| 11 | `a:folHlink` |

The ECMA-376 text describes the index in `a:clrScheme` schema order. Files in common use do not follow it. Measured on fixture `colors-fonts` (`fixtures/measured.json`): indexes 0 to 9 show `lt1`, `dk1`, `lt2`, `dk2`, and `accent1` to `accent6`, as this table states.

### Tint

`tint` changes the lightness of the base color in the HLS color space, with lightness from 0 to `HLSMAX`:

- If `tint` is less than 0: `L' = L * (1 + tint)`.
- If `tint` is greater than 0: `L' = L * (1 - tint) + (HLSMAX - HLSMAX * (1 - tint))`.

Hue and saturation do not change. This is not the DrawingML `a:tint` or `a:shade` transform, so do not use `packages/ooxml-core/src/drawingml-color.ts` for it. The `tint` clause of Part 1 (`CT_Color`) gives `HLSMAX` as 255. Measured on fixture `colors-fonts` (`fixtures/measured.json`): `accent1` `4472C4` with tints -0.5, -0.25, 0.25, 0.4, 0.5, 0.6, 0.8, and 0.9 shows `203764`, `305496`, `7395D2`, `8EA9DB`, `A2B8E1`, `B4C6E7`, `D9E1F2`, and `ECF0F9`. These formulas in floating point give each channel within 2 of 255, and no integer HLS variant tried matched all eight. The difference is the known approximation "Tint rounding".

### Indexed colors

`indexed` selects a color from the legacy palette. `x:styleSheet/x:colors/x:indexedColors` in `styles.xml` replaces the whole palette when it is present (schema, `CT_Colors` and `CT_IndexedColors`, Part 1, 18.8.27). Each entry is an `x:rgbColor`. Indexes 64 and 65 are the system foreground and background colors.

The default palette:

| Index | Color | Index | Color | Index | Color | Index | Color |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 0 | `000000` | 16 | `800000` | 32 | `000080` | 48 | `3366FF` |
| 1 | `FFFFFF` | 17 | `008000` | 33 | `FF00FF` | 49 | `33CCCC` |
| 2 | `FF0000` | 18 | `000080` | 34 | `FFFF00` | 50 | `99CC00` |
| 3 | `00FF00` | 19 | `808000` | 35 | `00FFFF` | 51 | `FFCC00` |
| 4 | `0000FF` | 20 | `800080` | 36 | `800080` | 52 | `FF9900` |
| 5 | `FFFF00` | 21 | `008080` | 37 | `800000` | 53 | `FF6600` |
| 6 | `FF00FF` | 22 | `C0C0C0` | 38 | `008080` | 54 | `666699` |
| 7 | `00FFFF` | 23 | `808080` | 39 | `0000FF` | 55 | `969696` |
| 8 | `000000` | 24 | `9999FF` | 40 | `00CCFF` | 56 | `003366` |
| 9 | `FFFFFF` | 25 | `993366` | 41 | `CCFFFF` | 57 | `339966` |
| 10 | `FF0000` | 26 | `FFFFCC` | 42 | `CCFFCC` | 58 | `003300` |
| 11 | `00FF00` | 27 | `CCFFFF` | 43 | `FFFF99` | 59 | `333300` |
| 12 | `0000FF` | 28 | `660066` | 44 | `99CCFF` | 60 | `993300` |
| 13 | `FFFF00` | 29 | `FF8080` | 45 | `FF99CC` | 61 | `993366` |
| 14 | `FF00FF` | 30 | `0066CC` | 46 | `CC99FF` | 62 | `333399` |
| 15 | `00FFFF` | 31 | `CCCCFF` | 47 | `FFCC99` | 63 | `333333` |

The table follows Part 1, 18.8.27 (`indexedColors`).

### Alpha and automatic colors

- The alpha byte of `rgb` is ignored. Measured on fixture `colors-fonts` (`fixtures/measured.json`): a fill with `rgb="00FF0000"` shows `FF0000`.
- `auto="1"` selects the automatic color. Measured on fixture `colors-fonts` (`fixtures/measured.json`): a font with `auto` is black, and a solid fill with `auto` prints white (the reference PDF paints a white rectangle). Indexed color 64 on a border is black.

## DrawingML colors in drawings and charts

Shapes, pictures, and charts use DrawingML colors with the same transforms as DOCX Editor and PPTX Editor. `resolveDrawingColor` in `packages/ooxml-core/src/drawingml-color.ts` applies the transforms. The caller resolves scheme color names:

- `dk1`, `lt1`, `dk2`, `lt2`, the accents, `hlink`, and `folHlink` read the theme's `a:clrScheme` directly.
- `tx1`, `bg1`, `tx2`, and `bg2` are logical names. A workbook has no color map element, so map them to `dk1`, `lt1`, `dk2`, and `lt2`. Measure this on a drawing fixture.
- `phClr` is the color of the shape style that holds the reference.
