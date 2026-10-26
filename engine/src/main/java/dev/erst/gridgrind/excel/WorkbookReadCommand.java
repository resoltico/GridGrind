package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.plan.CellReadProjection;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.vocabulary.ExcelAddressLists;
import dev.erst.gridgrind.core.vocabulary.ExcelReadLimits;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.List;
import java.util.Objects;

/** Workbook-core read commands executed after mutations and before persistence. */
public sealed interface WorkbookReadCommand
    permits WorkbookReadCommand.Introspection,
        WorkbookReadCommand.Surface,
        WorkbookReadCommand.Analysis {

  /** Stable caller-provided identifier used to correlate one read command with its result. */
  StepId stepId();

  /** Marker for fact-only workbook reads. */
  sealed interface Introspection extends WorkbookReadCommand
      permits GetWorkbookSummary,
          GetPackageSecurity,
          GetWorkbookProtection,
          GetCustomXmlMappings,
          ExportCustomXmlMapping,
          GetNamedRanges,
          GetSheetSummary,
          GetArrayFormulas,
          GetCells,
          GetWindow,
          GetMergedRegions,
          GetHyperlinks,
          GetComments,
          GetDrawingObjects,
          GetCharts,
          GetPivotTables,
          GetDrawingObjectPayload,
          GetSheetLayout,
          GetPrintLayout,
          GetDataValidations,
          GetConditionalFormatting,
          GetAutofilters,
          GetTables {}

  /** Marker for derived factual surface-summary reads. */
  sealed interface Surface extends WorkbookReadCommand
      permits GetFormulaSurface, GetSheetSchema, GetNamedRangeSurface {}

  /** Marker for derived analysis commands. */
  sealed interface Analysis extends WorkbookReadCommand
      permits AnalyzeFormulaHealth,
          AnalyzeDataValidationHealth,
          AnalyzeConditionalFormattingHealth,
          AnalyzeAutofilterHealth,
          AnalyzeTableHealth,
          AnalyzePivotTableHealth,
          AnalyzeHyperlinkHealth,
          AnalyzeNamedRangeHealth,
          AnalyzeWorkbookFindings {}

  /** Returns workbook-level summary facts. */
  record GetWorkbookSummary(StepId stepId) implements Introspection {
    public GetWorkbookSummary {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
    }
  }

  /** Returns OOXML package-encryption and package-signature facts for the current workbook. */
  record GetPackageSecurity(StepId stepId) implements Introspection {
    public GetPackageSecurity {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
    }
  }

  /** Returns workbook-level protection facts. */
  record GetWorkbookProtection(StepId stepId) implements Introspection {
    public GetWorkbookProtection {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
    }
  }

  /** Returns factual workbook custom-XML mapping metadata. */
  record GetCustomXmlMappings(StepId stepId) implements Introspection {
    public GetCustomXmlMappings {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
    }
  }

  /** Exports XML for one existing workbook custom-XML mapping. */
  record ExportCustomXmlMapping(
      StepId stepId, CustomXmlMappingLocator mapping, boolean validateSchema, String encoding)
      implements Introspection {
    public ExportCustomXmlMapping {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(mapping, "mapping must not be null");
      encoding = requireNonBlank(encoding, "encoding");
    }
  }

  /** Returns named ranges selected by exact selector or workbook-wide selection. */
  record GetNamedRanges(StepId stepId, ExcelNamedRangeSelection selection)
      implements Introspection {
    public GetNamedRanges {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns structural summary facts for one sheet. */
  record GetSheetSummary(StepId stepId, String sheetName) implements Introspection {
    public GetSheetSummary {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns array-formula groups across the selected sheets. */
  record GetArrayFormulas(StepId stepId, ExcelSheetSelection selection) implements Introspection {
    public GetArrayFormulas {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns exact cell snapshots for the provided ordered A1 addresses. */
  record GetCells(
      StepId stepId, String sheetName, List<String> addresses, CellReadProjection projection)
      implements Introspection {
    public GetCells {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      addresses = copyAddresses(addresses);
      Objects.requireNonNull(projection, "projection must not be null");
    }
  }

  /**
   * Maximum number of factual cells permitted in one cell-returning read surface. Exact-address
   * reads enforce this as an address-count cap, while rectangular reads enforce it on {@code
   * rowCount * columnCount}. Enforced in both the protocol and engine layers to prevent
   * out-of-memory failures during cell-grid serialization. See docs/LIMITATIONS.md LIM-001.
   */
  int MAX_READ_CELLS = ExcelReadLimits.MAX_READ_CELLS; // LIM-001

  /** Returns a rectangular window of cell snapshots anchored at the provided top-left address. */
  record GetWindow(
      StepId stepId,
      String sheetName,
      ExcelReadWindow window,
      CellReadProjection projection,
      boolean includeBlanks)
      implements Introspection {
    public GetWindow {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(window, "window must not be null");
      Objects.requireNonNull(projection, "projection must not be null");
    }
  }

  /** Returns the exact merged regions defined on one sheet. */
  record GetMergedRegions(StepId stepId, String sheetName) implements Introspection {
    public GetMergedRegions {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns hyperlink metadata for selected cells on one sheet. */
  record GetHyperlinks(StepId stepId, String sheetName, ExcelCellSelection selection)
      implements Introspection {
    public GetHyperlinks {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns comment metadata for selected cells on one sheet. */
  record GetComments(StepId stepId, String sheetName, ExcelCellSelection selection)
      implements Introspection {
    public GetComments {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns factual drawing-object metadata for one sheet. */
  record GetDrawingObjects(StepId stepId, String sheetName) implements Introspection {
    public GetDrawingObjects {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns factual chart metadata selected by sheet-wide or exact chart selection. */
  record GetCharts(StepId stepId, ExcelChartSelection selection) implements Introspection {
    public GetCharts {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns the extracted binary payload for one existing drawing object on one sheet. */
  record GetDrawingObjectPayload(StepId stepId, String sheetName, String objectName)
      implements Introspection {
    public GetDrawingObjectPayload {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      objectName = requireNonBlank(objectName, "objectName");
    }
  }

  /** Returns layout metadata such as pane state, zoom, and row and column sizing. */
  record GetSheetLayout(StepId stepId, String sheetName) implements Introspection {
    public GetSheetLayout {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns supported print-layout metadata for one sheet. */
  record GetPrintLayout(StepId stepId, String sheetName) implements Introspection {
    public GetPrintLayout {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns data-validation metadata for the selected ranges on one sheet. */
  record GetDataValidations(StepId stepId, String sheetName, ExcelRangeSelection selection)
      implements Introspection {
    public GetDataValidations {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns conditional-formatting metadata for the selected ranges on one sheet. */
  record GetConditionalFormatting(StepId stepId, String sheetName, ExcelRangeSelection selection)
      implements Introspection {
    public GetConditionalFormatting {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns sheet- and table-owned autofilter metadata for one sheet. */
  record GetAutofilters(StepId stepId, String sheetName) implements Introspection {
    public GetAutofilters {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Returns factual table metadata selected by workbook-global table name or all tables. */
  record GetTables(StepId stepId, TableSelector selection) implements Introspection {
    public GetTables {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Returns factual pivot-table metadata selected by workbook-global pivot name or all pivots. */
  record GetPivotTables(StepId stepId, PivotTableSelector selection) implements Introspection {
    public GetPivotTables {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Groups formula usage patterns across one or more sheets. */
  record GetFormulaSurface(StepId stepId, ExcelSheetSelection selection) implements Surface {
    public GetFormulaSurface {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Infers a simple column schema from a rectangular window on one sheet. */
  record GetSheetSchema(
      StepId stepId, String sheetName, ExcelReadWindow window, CellReadProjection projection)
      implements Surface {
    public GetSheetSchema {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(window, "window must not be null");
      Objects.requireNonNull(projection, "projection must not be null");
    }
  }

  /** Summarizes the scope and backing kind of selected named ranges. */
  record GetNamedRangeSurface(StepId stepId, ExcelNamedRangeSelection selection)
      implements Surface {
    public GetNamedRangeSurface {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports formula findings such as error results and volatile usage. */
  record AnalyzeFormulaHealth(StepId stepId, ExcelSheetSelection selection) implements Analysis {
    public AnalyzeFormulaHealth {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports data-validation findings such as malformed or overlapping rules. */
  record AnalyzeDataValidationHealth(StepId stepId, ExcelSheetSelection selection)
      implements Analysis {
    public AnalyzeDataValidationHealth {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports conditional-formatting findings such as broken formulas or priority collisions. */
  record AnalyzeConditionalFormattingHealth(StepId stepId, ExcelSheetSelection selection)
      implements Analysis {
    public AnalyzeConditionalFormattingHealth {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports autofilter findings such as invalid ranges or table-ownership mismatches. */
  record AnalyzeAutofilterHealth(StepId stepId, ExcelSheetSelection selection) implements Analysis {
    public AnalyzeAutofilterHealth {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports table findings such as overlaps, broken ranges, or invalid headers. */
  record AnalyzeTableHealth(StepId stepId, TableSelector selection) implements Analysis {
    public AnalyzeTableHealth {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports pivot-table findings such as broken caches, names, and sources. */
  record AnalyzePivotTableHealth(StepId stepId, PivotTableSelector selection) implements Analysis {
    public AnalyzePivotTableHealth {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports hyperlink findings such as malformed targets and missing document destinations. */
  record AnalyzeHyperlinkHealth(StepId stepId, ExcelSheetSelection selection) implements Analysis {
    public AnalyzeHyperlinkHealth {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Reports named-range findings such as broken references and scope shadowing. */
  record AnalyzeNamedRangeHealth(StepId stepId, ExcelNamedRangeSelection selection)
      implements Analysis {
    public AnalyzeNamedRangeHealth {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
    }
  }

  /** Runs the first analysis family across the whole workbook and aggregates their findings. */
  record AnalyzeWorkbookFindings(StepId stepId) implements Analysis {
    public AnalyzeWorkbookFindings {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static List<String> copyAddresses(List<String> addresses) {
    return ExcelAddressLists.copyNonEmptyDistinctAddresses(addresses);
  }
}
