# XLSX Internal Structure — Developer Reference

## Overview

An `.xlsx` file is a standard ZIP archive containing XML files and folders.
Excel reads and writes all data through these XML files — understanding their
structure is the prerequisite for any surgical, code-based repair.

This document is grounded in a direct analysis of `BadPipeIn.xlsx` and reflects
the specific layout of the PipeUpdUV workbook. It serves as the foundation for
developing `FixPipe.py` fixes step by step.

---

## 1. ZIP Container Layout

```
BadPipeIn.xlsx (ZIP archive)
│
├── [Content_Types].xml              ← Registry of all parts and their MIME types
├── _rels/
│   └── .rels                        ← Root: points to xl/workbook.xml
│
├── docProps/
│   ├── app.xml                      ← Application metadata (app version, etc.)
│   ├── core.xml                     ← Core properties (author, created, modified)
│   └── custom.xml                   ← Custom document properties
│
└── xl/
    ├── workbook.xml                 ← Sheet list, names, active sheet
    ├── _rels/
    │   └── workbook.xml.rels        ← Maps rId→ file for each sheet, cache, style, etc.
    │
    ├── worksheets/
    │   ├── sheet1.xml  … sheet12.xml   ← Cell data, formulas, column/row definitions
    │   └── _rels/
    │       ├── sheet4.xml.rels      ← Refs to pivot tables for "By Sales"
    │       ├── sheet5.xml.rels      ← (and other sheets that have rels)
    │       └── …
    │
    ├── pivotTables/
    │   ├── pivotTable1.xml … pivotTable10.xml   ← Layout, field config, filter state
    │   └── _rels/
    │       └── pivotTableN.xml.rels ← Each PT → its pivotCacheDefinition
    │
    ├── pivotCache/
    │   ├── pivotCacheDefinition1.xml … 4.xml    ← Field metadata + shared-item lists
    │   ├── pivotCacheRecords1.xml    … 4.xml    ← Actual cached data rows
    │   └── _rels/
    │       └── pivotCacheDefinitionN.xml.rels   ← Def → its Records file
    │
    ├── charts/
    │   └── chart1.xml, chart2.xml, chart3.xml  ← Chart definitions (type, series, style)
    │
    ├── drawings/
    │   ├── drawing1.xml                         ← Anchors charts to worksheet grid
    │   └── _rels/
    │       └── drawing1.xml.rels                ← drawing → chart file(s)
    │
    ├── sharedStrings.xml            ← Deduplicated string table (cell text)
    ├── styles.xml                   ← Number formats, fonts, fills, borders, cell styles
    ├── calcChain.xml                ← Formula recalculation order
    └── theme/
        └── theme1.xml               ← Color/font theme
```

---

## 2. Content_Types.xml

Every file in the ZIP must be registered here so Excel knows how to handle it.

```xml
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <!-- Default: applies to any file with this extension -->
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>

  <!-- Override: applies to a specific path -->
  <Override PartName="/xl/workbook.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet4.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/charts/chart1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheRecords1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
</Types>
```

**Rules:**
- Adding a new chart/drawing/cache without a `<Override>` entry will cause Excel to
  ignore or corrupt the file on open.
- Use **string manipulation** (not ElementTree) to edit this file. ET re-serializes
  the namespace as `ns0:` which breaks the file.

---

## 3. Relationship Files (.rels)

Every `.rels` file lives in a `_rels/` subdirectory alongside its parent file.

```xml
<!-- xl/worksheets/_rels/sheet4.xml.rels -->
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"  Type="…/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
  <Relationship Id="rId9"  Type="…/drawing"    Target="../drawings/drawing1.xml"/>
</Relationships>
```

**Key rules:**
- `Id` values are unique **only within that .rels file** — `rId1` in `sheet4.xml.rels`
  is unrelated to `rId1` in `drawing1.xml.rels`.
- The `Type` URI identifies what kind of link this is (worksheet, chart, drawing,
  pivotTable, pivotCacheDefinition, sharedStrings, styles, theme, …).
- `Target` is a **relative path** from the parent file's directory.
- Relationship files are optional: only sheets that have charts, pivot tables, or
  other linked objects have a `.rels` file.

---

## 4. Workbook (xl/workbook.xml)

Lists all sheets with their display names, internal sheet IDs, and relationship IDs:

```xml
<workbook>
  <sheets>
    <sheet name="Pipeline Sell Out" sheetId="1" r:id="rId1"/>
    <sheet name="By Sales"          sheetId="3" r:id="rId4"/>
    <sheet name="Pipe Analysis"     sheetId="4" r:id="rId5"/>
    <sheet name="Pipe Log"          sheetId="7" r:id="rId8"/>
    <!-- hidden sheets: -->
    <sheet name="Pipeline Run Rate" sheetId="2" state="hidden" r:id="rId3"/>
  </sheets>
</workbook>
```

`xl/_rels/workbook.xml.rels` maps each `r:id` to a file path, and also maps the
four pivot cache definitions to their workbook-level relationship IDs.

---

## 5. Worksheets (xl/worksheets/sheetN.xml)

Contains cell data, column widths, row heights, and references to embedded objects.

```xml
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>…</sheetViews>
  <sheetFormatPr …/>
  <cols>…</cols>
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>42</v></c>   <!-- string: index into sharedStrings -->
      <c r="B1" t="n"><v>1234.5</v></c>
      <c r="C1"><f>A1+B1</f><v>…</v></c>
    </row>
  </sheetData>
  <!-- Reference to the drawing layer (charts): -->
  <drawing r:id="rId10"/>
</worksheet>
```

The `<drawing r:id="…"/>` element at the **end** of the worksheet XML links the
sheet to its `drawingN.xml` via the sheet's `.rels` file.
Removing or failing to add this element means charts are invisible even if all
other files are present.

---

## 6. Pivot Cache Architecture (Three-File System)

Each pivot cache is three files: definition, records, and the definition's `.rels`.

### 6a. pivotCacheDefinition.xml

Describes the data source and the field catalogue (schema + shared value lists).

```xml
<pivotCacheDefinition r:id="rId1" recordCount="936">
  <cacheSource type="worksheet">
    <!-- Option 1: direct range reference -->
    <worksheetSource ref="A2:U1048576" sheet="Pipeline Sell Out"/>
    <!-- Option 2: named range (DATASELLOUT in BadPipeIn cache4) -->
    <worksheetSource name="DATASELLOUT"/>
  </cacheSource>

  <cacheFields count="21">
    <!-- Field with shared-item list (used by <x> refs in records) -->
    <cacheField name="Propriétaire de l'opportunité" numFmtId="0">
      <sharedItems containsBlank="1" count="33">
        <s v="Kajanan SHAN"/>
        <s v="Charles TEZENAS"/>
        <!-- … one <s> per unique owner value … -->
      </sharedItems>
    </cacheField>

    <!-- Field whose values are stored inline in records (no shared items) -->
    <cacheField name="Opportunity Number" numFmtId="0">
      <sharedItems containsBlank="1" containsSemiMixedTypes="1" containsNonDate="1" containsString="1"/>
    </cacheField>

    <!-- Numeric field -->
    <cacheField name="Quantité" numFmtId="0">
      <sharedItems containsNumber="1" containsInteger="1" minValue="1" maxValue="10000000"/>
    </cacheField>
  </cacheFields>
</pivotCacheDefinition>
```

**Shared items** are the unique values for a field, stored as an indexed list.
Fields **without** shared items store their values directly in each record.

### 6b. pivotCacheRecords.xml

One `<r>` element per data row. Each child element corresponds to one field
in the **same order** as `cacheFields`:

```xml
<pivotCacheRecords count="936">
  <r>
    <x v="0"/>                          <!-- Owner: index 0 → "Kajanan SHAN" -->
    <d v="2024-12-19T00:00:00"/>        <!-- Created Date: inline date -->
    <d v="2025-06-30T00:00:00"/>        <!-- Close Date: inline date -->
    <s v="Closed Won"/>                 <!-- Stage: inline string -->
    <s v="OP0000218588"/>               <!-- Opportunity Number: inline string -->
    <s v="CFI"/>                        <!-- Revendeur: inline string -->
    <s v="France Travail"/>             <!-- Client Final: inline string -->
    <n v="110000"/>                     <!-- Quantité: inline number -->
    <n v="367"/>                        <!-- Prix de vente: inline number -->
    <n v="40370000"/>                   <!-- Prix total: inline number -->
    <!-- … more fields … -->
    <x v="0"/>                          <!-- Product Family: index into its sharedItems -->
    <x v="3"/>                          <!-- Quarter Invoice: index -->
    <x v="2"/>                          <!-- Forecast: index -->
    <m/>                                <!-- Blank / missing value -->
  </r>
</pivotCacheRecords>
```

**Element types:**
| Tag | Meaning |
|-----|---------|
| `<x v="N"/>` | Shared-item reference — index N into that field's `<sharedItems>` list |
| `<s v="…"/>` | Inline string |
| `<n v="…"/>` | Inline number |
| `<d v="…"/>` | Inline date (ISO 8601) |
| `<m/>` | Missing / null value |
| `<b v="0\|1"/>` | Boolean |
| `<e v="…"/>` | Error value |

### 6c. Cache Inventory in BadPipeIn.xlsx

| Cache | Source | Records | Used by Pivot Tables | Notes |
|-------|--------|---------|---------------------|-------|
| cache1 | Owner Opty Tracking sheet | — | PT5 (pivotTable5) | Owner tracking pivot; **do not modify** |
| cache2 | `Pipeline Sell Out` A2:U∞ | 4799 | PT1,2,4,6,7,9 | Main pipeline data |
| cache3 | — | — | PT10 | Used by Pipeline Close Lost sheet |
| cache4 | Named range `DATASELLOUT` | 936 | PT3, PT8 | "By Sales" main data; refreshes on load |

---

## 7. Pivot Table (xl/pivotTables/pivotTableN.xml)

Defines how the cache data is arranged and displayed.

```xml
<pivotTableDefinition name="Tableau croisé dynamique1" cacheId="2">
  <location ref="S6:V11" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>

  <pivotFields count="24">
    <!-- One pivotField per cache field, in the same order -->

    <!-- A row field (axis="axisRow") -->
    <pivotField axis="axisRow" showAll="0">
      <items count="40">
        <item sd="0" x="0"/>     <!-- cache index 0 → visible (no h attr) -->
        <item h="1" sd="0" x="17"/>  <!-- h="1" → hidden in pivot -->
        <item t="default" sd="0"/>   <!-- the Grand Total row marker -->
      </items>
    </pivotField>

    <!-- A page/filter field (axis="axisPage") -->
    <pivotField axis="axisPage" showAll="0" multipleItemSelectionAllowed="1">
      <items count="25">
        <item x="0"/>            <!-- selected (no h attr) -->
        <item h="1" x="1"/>     <!-- hidden / unchecked in filter dropdown -->
      </items>
    </pivotField>

    <!-- A data field (no axis) -->
    <pivotField dataField="1" showAll="0"/>
  </pivotFields>

  <!-- Which fields go where -->
  <rowFields>  <field x="0"/>  </rowFields>
  <colFields>  <field x="3"/>  </colFields>
  <pageFields> <pageField fld="18" item="4294967295"/> </pageFields>
  <dataFields> <dataField name="Sum of Qty" fld="7" subtotal="sum"/> </dataFields>
</pivotTableDefinition>
```

**Key concepts:**
- `<item x="N"/>` — the `x` attribute is a **cache sharedItems index** for that field.
- `h="1"` — the item is hidden (unchecked) in the pivot/filter.
- `item="4294967295"` on `<pageField>` — sentinel value meaning "multiple items selected".
- A pivot table references its cache via `cacheId` which maps to the workbook's
  `pivotCacheDefinition` relationship.

---

## 8. Charts and Drawings

### 8a. Chart (xl/charts/chartN.xml)

Defines chart type, title, series, and visual styling.

```xml
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:title><c:tx>…<a:t>My Chart Title</a:t>…</c:tx></c:title>
    <c:plotArea>
      <c:barChart>   <!-- or lineChart, line3DChart, pieChart, … -->
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <!-- Series label from a cell -->
          <c:tx><c:strRef><c:f>'Pipe Analysis'!$B$2</c:f></c:strRef></c:tx>
          <!-- Category axis values -->
          <c:cat><c:numRef><c:f>'Pipe Analysis'!$A$3:$A$34</c:f></c:numRef></c:cat>
          <!-- Data values -->
          <c:val><c:numRef><c:f>'Pipe Analysis'!$C$3:$C$34</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
```

**In BadPipeIn.xlsx:**
- `chart1.xml` — `line3DChart`, title "Evolution OPTY 30 Derniers Jours", references
  `Pipe Analysis` sheet. Belongs to `drawing1.xml` via `rId1`.
- `chart2.xml` — `barChart`, references `Pipe Analysis` sheet data. `rId2`.
- `chart3.xml` — `barChart`, references `Pipe Analysis` sheet data. `rId3`.
- All three charts live in `drawing1.xml` which is anchored to a **different sheet**
  (not "By Sales" — `drawing1.xml` is already present in another sheet's `.rels`).

### 8b. Drawing (xl/drawings/drawingN.xml)

Acts as the **anchor layer** that positions charts (or images/shapes) on the worksheet grid.

```xml
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">

  <xdr:twoCellAnchor>   <!-- anchor defined by two cell corners -->
    <xdr:from>
      <xdr:col>4</xdr:col>   <xdr:colOff>260350</xdr:colOff>
      <xdr:row>2</xdr:row>   <xdr:rowOff>63500</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>14</xdr:col>  <xdr:colOff>438150</xdr:colOff>
      <xdr:row>22</xdr:row>  <xdr:rowOff>107950</xdr:rowOff>
    </xdr:to>

    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <!-- This r:id links to chart1.xml via drawing1.xml.rels -->
          <c:chart xmlns:c="…" xmlns:r="…" r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>

  <!-- Second chart anchor in the same drawing file -->
  <xdr:twoCellAnchor>…r:id="rId2"…</xdr:twoCellAnchor>

</xdr:wsDr>
```

`colOff` and `rowOff` are in **EMUs** (English Metric Units): 914400 EMU = 1 inch.

---

## 9. The Complete Link Chain: Sheet → Drawing → Chart

```
xl/workbook.xml
  sheet name="By Sales" r:id="rId4"
        ↓
xl/_rels/workbook.xml.rels
  rId4 → worksheets/sheet4.xml
        ↓
xl/worksheets/sheet4.xml
  <drawing r:id="rId9"/>    ← must be present inside the worksheet XML
        ↓
xl/worksheets/_rels/sheet4.xml.rels
  rId9 → ../drawings/drawing2.xml   (new drawing for By Sales)
        ↓
xl/drawings/drawing2.xml
  <c:chart r:id="rId1"/>            ← first chart
  <c:chart r:id="rId2"/>            ← second chart
        ↓
xl/drawings/_rels/drawing2.xml.rels
  rId1 → ../charts/chart4.xml       (new chart files)
  rId2 → ../charts/chart5.xml
        ↓
xl/charts/chart4.xml  (bar chart — By Sales pivot chart 1)
xl/charts/chart5.xml  (bar chart — By Sales pivot chart 2)
```

And in `[Content_Types].xml`, each new file needs an `<Override>` entry.

---

## 10. "By Sales" Tab — Current State of BadPipeIn.xlsx

**Sheet file:** `xl/worksheets/sheet4.xml` (mapped via `rId4` in workbook.xml.rels)

**Pivot tables on this sheet (9 total):**

| File | Name | Cache | Location | Description |
|------|------|-------|----------|-------------|
| pivotTable1 | PivotTable5 | cache2 | J60:L64 | Small summary |
| pivotTable2 | TCD2 | cache2 | AT4:AV25 | Column pivot |
| pivotTable3 | PivotTable7 | cache4 | AA5:AE12 | By Sales main (DATASELLOUT) |
| pivotTable4 | TCD1 | cache2 | S6:V11 | Filter-heavy pivot |
| pivotTable5 | TCD4 | cache1 | AH7:AI9 | Owner Opty (skip) |
| pivotTable6 | TCD6 | cache2 | O6:P32 | Long list pivot |
| pivotTable7 | PivotTable6 | cache2 | O60:R69 | Summary |
| pivotTable8 | TCD3 | cache4 | B5:D13 | By Sales top-left (DATASELLOUT) |
| pivotTable9 | PivotTable1 | cache2 | B60:F78 | Large summary |

**Charts:** Currently `drawing1.xml` (which has `chart1`, `chart2`, `chart3`) is
**NOT linked to sheet4**. Sheet4 has no `<drawing>` element and no drawing `.rels` entry.
This is the corruption `fix_by_sales_charts` addresses.

**Target pivot table for chart fix:** `pivotTable8` (TCD3, cache4, location B5:D13)
— this is the "By Sales" summary table that the new charts should visualize.

---

## 11. XML Manipulation Rules

1. **Use string operations** (not `ElementTree.tostring`) for `.rels` and
   `[Content_Types].xml` — ET adds `ns0:` namespace prefixes that corrupt the file.

2. **Use `ElementTree.fromstring()`** freely for *reading* (parsing) XML — it is
   safe to parse without modifying the serialized form.

3. **Preserve ZIP metadata** when rewriting: use `zf_out.writestr(item, new_bytes)`
   where `item` is the original `ZipInfo` object to keep compression method and
   other metadata consistent.

4. **XML-escape attribute values** when building XML strings:
   `&` → `&amp;`, `"` → `&quot;`, `<` → `&lt;`, `>` → `&gt;`

5. **Orphaned file cleanup** (planned): before injecting new chart/drawing files,
   scan all `.rels` files to find which `chartN.xml` / `drawingN.xml` files are
   actually referenced. Remove unreferenced files from the ZIP and their
   `[Content_Types].xml` entries. This prevents accumulation of dead files
   (`chart4`, `chart6`, `chart8`…) from repeated fix cycles.

---

## 12. Namespace Reference

| Prefix | URI | Usage |
|--------|-----|-------|
| (default) | `http://schemas.openxmlformats.org/spreadsheetml/2006/main` | Worksheets, pivot tables, cache |
| `r:` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | Relationship IDs in XML |
| `c:` | `http://schemas.openxmlformats.org/drawingml/2006/chart` | Chart XML |
| `a:` | `http://schemas.openxmlformats.org/drawingml/2006/main` | DrawingML (shapes, text) |
| `xdr:` | `http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing` | Drawing anchors |
| `mc:` | `http://schemas.openxmlformats.org/markup-compatibility/2006` | Markup compatibility |
| `xr:` | `http://schemas.microsoft.com/office/spreadsheetml/2014/revision` | Revision tracking |
| (rels) | `http://schemas.openxmlformats.org/package/2006/relationships` | .rels files |
| (ct) | `http://schemas.openxmlformats.org/package/2006/content-types` | Content_Types.xml |
