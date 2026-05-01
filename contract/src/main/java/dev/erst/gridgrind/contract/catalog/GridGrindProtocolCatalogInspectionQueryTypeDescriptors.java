package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.query.InspectionQuery;
import java.util.List;

/** Inspection-query descriptors for the public protocol catalog. */
final class GridGrindProtocolCatalogInspectionQueryTypeDescriptors {
  static final List<CatalogTypeDescriptor> INSPECTION_QUERY_TYPES =
      List.of(
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetWorkbookSummary.class,
              "GET_WORKBOOK_SUMMARY",
              "Return workbook-level summary facts including sheet count, sheet names,"
                  + " named range count, and formula recalculation flag."
                  + " Empty workbooks return workbook.kind=EMPTY;"
                  + " non-empty workbooks return workbook.kind=WITH_SHEETS with activeSheetName"
                  + " and selectedSheetNames."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetPackageSecurity.class,
              "GET_PACKAGE_SECURITY",
              "Return OOXML package-encryption facts and package-signature validation results."
                  + " This reports the currently loaded workbook package state;"
                  + " after in-memory mutations, previously valid source signatures read back as"
                  + " INVALIDATED_BY_MUTATION until the saved output is re-signed."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetWorkbookProtection.class,
              "GET_WORKBOOK_PROTECTION",
              "Return workbook-level protection facts including structure, windows, and"
                  + " revisions lock state plus whether password hashes are present."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetCustomXmlMappings.class,
              "GET_CUSTOM_XML_MAPPINGS",
              "Return workbook custom-XML mapping metadata, including map identifiers,"
                  + " schema metadata, linked single cells, linked tables, and optional data"
                  + " binding facts."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.ExportCustomXmlMapping.class,
              "EXPORT_CUSTOM_XML_MAPPING",
              "Export one existing workbook custom-XML mapping as serialized XML."
                  + " mapping locates one existing map by mapId and/or name;"
                  + " validateSchema defaults to false and encoding defaults to UTF-8 when"
                  + " omitted.",
              "mapping",
              "validateSchema",
              "encoding"),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetNamedRanges.class,
              "GET_NAMED_RANGES",
              "Return named ranges matched by the supplied selection."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetSheetSummary.class,
              "GET_SHEET_SUMMARY",
              "Return structural summary facts for one sheet."
                  + " Includes visibility and sheet protection state."
                  + " physicalRowCount is the number of physically materialized rows (sparse)."
                  + " lastRowIndex is the 0-based index of the last materialized row"
                  + " (-1 when empty), including metadata-only rows."
                  + " lastColumnIndex is the 0-based index of the last materialized column"
                  + " in any row (-1 when empty)."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetCells.class,
              "GET_CELLS",
              "Return exact cell snapshots for explicit addresses."
                  + " An invalid or out-of-range address returns INVALID_CELL_ADDRESS, not a blank."
                  + " Each snapshot includes address, declaredType, effectiveType, displayValue,"
                  + " style, and metadata fields. Type-specific fields: stringValue (STRING),"
                  + " numberValue (NUMBER), booleanValue (BOOLEAN), errorValue (ERROR),"
                  + " formula and evaluation (FORMULA). For FORMULA cells, effectiveType is FORMULA"
                  + " and the evaluated result type is in evaluation.effectiveType."
                  + " style.font.fontHeight is a plain object with both twips and points fields,"
                  + " nested under the style.font group."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetWindow.class,
              "GET_WINDOW",
              "Return a rectangular window of cell snapshots."
                  + " rowCount * columnCount must not exceed "
                  + InspectionQuery.MAX_WINDOW_CELLS
                  + ". Each cell snapshot has the same shape as GET_CELLS: address,"
                  + " declaredType, effectiveType, displayValue, style, metadata,"
                  + " and type-specific value fields."
                  + " For FORMULA cells, effectiveType is FORMULA and the evaluated result type"
                  + " is in evaluation.effectiveType."
                  + " Response shape: { \"window\": { \"rows\": [ { \"cells\": [...] } ] } }."
                  + " Note: the top-level key is \"window\" and cells are nested under"
                  + " window.rows[N].cells, unlike GET_CELLS which places cells directly"
                  + " under the top-level \"cells\" key."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetMergedRegions.class,
              "GET_MERGED_REGIONS",
              "Return the merged regions defined on one sheet."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetHyperlinks.class,
              "GET_HYPERLINKS",
              "Return hyperlink metadata for selected cells in the same discriminated shape used"
                  + " by SET_HYPERLINK targets. FILE targets are returned as normalized path"
                  + " strings, not file: URIs."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetComments.class,
              "GET_COMMENTS",
              "Return comment metadata for selected cells."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetDrawingObjects.class,
              "GET_DRAWING_OBJECTS",
              "Return factual drawing-object metadata for one sheet."
                  + " Read families include pictures, signature lines, simple shapes,"
                  + " connectors, groups, charts, graphic frames, and embedded objects with"
                  + " truthful anchor and package facts."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetCharts.class,
              "GET_CHARTS",
              "Return factual chart metadata for one sheet."
                  + " Supported authored chart families are modeled authoritatively;"
                  + " unsupported plot families or unsupported loaded detail are surfaced as"
                  + " explicit UNSUPPORTED entries with preserved plot-type tokens."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetArrayFormulas.class,
              "GET_ARRAY_FORMULAS",
              "Return factual array-formula group metadata for the selected sheets,"
                  + " including the stored range, top-left anchor cell, normalized formula text,"
                  + " and whether the group is single-cell."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetPivotTables.class,
              "GET_PIVOT_TABLES",
              "Return factual pivot-table metadata selected by workbook-global pivot-table name"
                  + " or ALL."
                  + " Supported pivots surface source, anchor, row or column labels,"
                  + " report filters, data fields, and values-axis placement."
                  + " Unsupported or malformed pivots are returned explicitly with preserved"
                  + " detail instead of causing read failure."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetDrawingObjectPayload.class,
              "GET_DRAWING_OBJECT_PAYLOAD",
              "Return the extracted binary payload for one existing named picture or embedded"
                  + " object."
                  + " Non-binary drawing shapes such as connectors and simple shapes are"
                  + " rejected."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetSheetLayout.class,
              "GET_SHEET_LAYOUT",
              GridGrindContractText.sheetLayoutReadSummary()),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetPrintLayout.class,
              "GET_PRINT_LAYOUT",
              "Return supported print-layout metadata for one sheet, including print area,"
                  + " orientation, scaling, repeating rows or columns, and plain header or"
                  + " footer text."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetDataValidations.class,
              "GET_DATA_VALIDATIONS",
              "Return factual data-validation structures for the selected sheet ranges."
                  + " Supported rules include explicit lists, formula lists, comparison rules,"
                  + " and custom formulas; unsupported rules are surfaced explicitly with typed"
                  + " detail."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetConditionalFormatting.class,
              "GET_CONDITIONAL_FORMATTING",
              "Return factual conditional-formatting blocks for the selected sheet ranges."
                  + " Read families include authored formula and cell-value rules plus loaded"
                  + " color scales, data bars, icon sets, and explicitly unsupported rules."
                  + " Each block preserves its stored ordered ranges and rule priority data."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetAutofilters.class,
              "GET_AUTOFILTERS",
              "Return sheet- and table-owned autofilter metadata for one sheet."
                  + " Each entry is typed as SHEET or TABLE so ownership is explicit."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetTables.class,
              "GET_TABLES",
              "Return factual table metadata selected by workbook-global table name or ALL."
                  + " Each table includes range, header and totals row counts, column names,"
                  + " style metadata, and whether a table-owned autofilter is present."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetFormulaSurface.class,
              "GET_FORMULA_SURFACE",
              GridGrindContractText.formulaSurfaceReadSummary()),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetSheetSchema.class,
              "GET_SHEET_SCHEMA",
              "Infer a simple schema from a rectangular sheet window."
                  + " rowCount * columnCount must not exceed "
                  + InspectionQuery.MAX_WINDOW_CELLS
                  + "."
                  + " The first row is treated as the header; dataRowCount is 0 when all header"
                  + " cells are blank."
                  + " dominantType is null when all data cells are blank, or when two or more types"
                  + " tie for the highest count."
                  + " Formula cells contribute their evaluated result type (NUMBER, STRING, etc.)"
                  + " to dominantType and observedTypes, not FORMULA."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetNamedRangeSurface.class,
              "GET_NAMED_RANGE_SURFACE",
              GridGrindContractText.namedRangeSurfaceReadSummary()),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzeFormulaHealth.class,
              "ANALYZE_FORMULA_HEALTH",
              GridGrindContractText.formulaHealthReadSummary()),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzeDataValidationHealth.class,
              "ANALYZE_DATA_VALIDATION_HEALTH",
              "Return analysis.checkedValidationCount, a severity summary,"
                  + " and findings such as unsupported, overlapping, or"
                  + " broken-formula rules."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzeConditionalFormattingHealth.class,
              "ANALYZE_CONDITIONAL_FORMATTING_HEALTH",
              "Return analysis.checkedConditionalFormattingBlockCount,"
                  + " a severity summary, and conditional-formatting findings such as broken formulas,"
                  + " unsupported loaded rules, invalid target ranges, or priority collisions."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzeAutofilterHealth.class,
              "ANALYZE_AUTOFILTER_HEALTH",
              "Return analysis.checkedAutofilterCount, a severity summary,"
                  + " and autofilter findings such as invalid ranges, blank header rows,"
                  + " or ownership mismatches between sheet-level filters and tables."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzeTableHealth.class,
              "ANALYZE_TABLE_HEALTH",
              "Return analysis.checkedTableCount, a severity summary,"
                  + " and table findings such as overlaps, broken references,"
                  + " blank or duplicate headers, and unresolved styles."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzePivotTableHealth.class,
              "ANALYZE_PIVOT_TABLE_HEALTH",
              "Return analysis.checkedPivotTableCount, a severity summary,"
                  + " and pivot-table findings such as missing cache parts, broken sources,"
                  + " duplicate or synthetic names, and unsupported persisted detail."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzeHyperlinkHealth.class,
              "ANALYZE_HYPERLINK_HEALTH",
              "Return analysis.checkedHyperlinkCount, a severity summary,"
                  + " and hyperlink findings such as malformed, missing, unresolved, or broken"
                  + " targets."
                  + " Relative FILE targets are resolved against the workbook's persisted path"
                  + " when one exists."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzeNamedRangeHealth.class,
              "ANALYZE_NAMED_RANGE_HEALTH",
              GridGrindContractText.namedRangeHealthReadSummary()),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzeWorkbookFindings.class,
              "ANALYZE_WORKBOOK_FINDINGS",
              GridGrindContractText.workbookFindingsReadSummary()));

  private GridGrindProtocolCatalogInspectionQueryTypeDescriptors() {}
}
