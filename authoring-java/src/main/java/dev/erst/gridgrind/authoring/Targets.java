package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import java.util.List;
import java.util.Objects;

/** Selector builders and common target-scoped step builders for the Java authoring API. */
@SuppressWarnings("PMD.ExcessivePublicCount")
public final class Targets {
  private Targets() {}

  /** Returns one workbook-scoped fluent target. */
  public static WorkbookRef workbook() {
    return new WorkbookRef(new WorkbookSelector.Current());
  }

  /** Returns one exact sheet target by sheet name. */
  public static SheetRef sheet(String name) {
    return new SheetRef(new SheetSelector.ByName(name));
  }

  /** Returns one selector that targets every sheet in the workbook. */
  public static SheetSelector.All allSheets() {
    return new SheetSelector.All();
  }

  /** Returns one selector that targets the named sheets. */
  public static SheetSelector.ByNames sheets(String... names) {
    return new SheetSelector.ByNames(List.of(names));
  }

  /** Returns one exact cell target by sheet name and A1 address. */
  public static CellRef cell(String sheetName, String address) {
    return new CellRef(new CellSelector.ByAddress(sheetName, address));
  }

  /** Returns one selector that targets several exact cells on one sheet. */
  public static CellSelector.ByAddresses cells(String sheetName, String... addresses) {
    return new CellSelector.ByAddresses(sheetName, List.of(addresses));
  }

  /** Returns one selector that targets every used cell on one sheet. */
  public static CellSelector.AllUsedInSheet allUsedCells(String sheetName) {
    return new CellSelector.AllUsedInSheet(sheetName);
  }

  /** Returns one selector that targets several fully qualified sheet-plus-address cells. */
  public static CellSelector.ByQualifiedAddresses qualifiedCells(
      CellSelector.QualifiedAddress... cells) {
    return new CellSelector.ByQualifiedAddresses(List.of(cells));
  }

  /** Returns one qualified cell reference for multi-sheet selectors. */
  public static CellSelector.QualifiedAddress qualifiedCell(String sheetName, String address) {
    return new CellSelector.QualifiedAddress(sheetName, address);
  }

  /** Returns one exact range target by sheet name and A1 range. */
  public static RangeRef range(String sheetName, String range) {
    return new RangeRef(new RangeSelector.ByRange(sheetName, range));
  }

  /** Returns one selector that targets several exact ranges on one sheet. */
  public static RangeSelector.ByRanges ranges(String sheetName, String... ranges) {
    return new RangeSelector.ByRanges(sheetName, List.of(ranges));
  }

  /** Returns one selector that targets every persisted range on one sheet. */
  public static RangeSelector.AllOnSheet allRanges(String sheetName) {
    return new RangeSelector.AllOnSheet(sheetName);
  }

  /** Returns one rectangular window target from a top-left cell plus dimensions. */
  public static WindowRef window(
      String sheetName, String topLeftAddress, int rowCount, int columnCount) {
    return new WindowRef(
        new RangeSelector.RectangularWindow(sheetName, topLeftAddress, rowCount, columnCount));
  }

  /** Returns one row-span selector on a sheet. */
  public static RowBandSelector.Span rows(String sheetName, int firstRowIndex, int lastRowIndex) {
    return new RowBandSelector.Span(sheetName, firstRowIndex, lastRowIndex);
  }

  /** Returns one row-insertion selector before the given zero-based row index. */
  public static RowBandSelector.Insertion insertRowsBefore(
      String sheetName, int beforeRowIndex, int rowCount) {
    return new RowBandSelector.Insertion(sheetName, beforeRowIndex, rowCount);
  }

  /** Returns one column-span selector on a sheet. */
  public static ColumnBandSelector.Span columns(
      String sheetName, int firstColumnIndex, int lastColumnIndex) {
    return new ColumnBandSelector.Span(sheetName, firstColumnIndex, lastColumnIndex);
  }

  /** Returns one column-insertion selector before the given zero-based column index. */
  public static ColumnBandSelector.Insertion insertColumnsBefore(
      String sheetName, int beforeColumnIndex, int columnCount) {
    return new ColumnBandSelector.Insertion(sheetName, beforeColumnIndex, columnCount);
  }

  /** Returns one exact table target by table name. */
  public static TableRef table(String name) {
    return new TableRef(new TableSelector.ByName(name));
  }

  /** Returns one exact table target by table name plus sheet ownership. */
  public static TableRef tableOnSheet(String name, String sheetName) {
    return new TableRef(new TableSelector.ByNameOnSheet(name, sheetName));
  }

  /** Returns one selector that targets every table in the workbook. */
  public static TableSelector.All allTables() {
    return new TableSelector.All();
  }

  /** Returns one selector that targets the named tables. */
  public static TableSelector.ByNames tables(String... names) {
    return new TableSelector.ByNames(List.of(names));
  }

  /** Returns one exact named-range target by name. */
  public static NamedRangeRef namedRange(String name) {
    return new NamedRangeRef(new NamedRangeSelector.ByName(name));
  }

  /** Returns one exact sheet-scoped named-range target. */
  public static NamedRangeRef namedRangeOnSheet(String name, String sheetName) {
    return new NamedRangeRef(new NamedRangeSelector.SheetScope(name, sheetName));
  }

  /** Returns one exact workbook-scoped named-range target. */
  public static NamedRangeRef workbookNamedRange(String name) {
    return new NamedRangeRef(new NamedRangeSelector.WorkbookScope(name));
  }

  /** Returns one selector that targets every named range. */
  public static NamedRangeSelector.All allNamedRanges() {
    return new NamedRangeSelector.All();
  }

  /** Returns one selector that targets the named ranges. */
  public static NamedRangeSelector.ByNames namedRanges(String... names) {
    return new NamedRangeSelector.ByNames(List.of(names));
  }

  /** Returns one selector that targets any of the provided exact named-range selectors. */
  public static NamedRangeSelector.AnyOf anyNamedRange(NamedRangeSelector.Ref... selectors) {
    return new NamedRangeSelector.AnyOf(List.of(selectors));
  }

  /** Returns one exact chart target by sheet name and chart name. */
  public static ChartRef chart(String sheetName, String chartName) {
    return new ChartRef(new ChartSelector.ByName(sheetName, chartName));
  }

  /** Returns one selector that targets every chart on a sheet. */
  public static ChartSelector.AllOnSheet chartsOnSheet(String sheetName) {
    return new ChartSelector.AllOnSheet(sheetName);
  }

  /** Returns one exact drawing-object target by sheet name and object name. */
  public static DrawingObjectSelector.ByName drawingObject(String sheetName, String objectName) {
    return new DrawingObjectSelector.ByName(sheetName, objectName);
  }

  /** Returns one selector that targets every drawing object on a sheet. */
  public static DrawingObjectSelector.AllOnSheet drawingObjectsOnSheet(String sheetName) {
    return new DrawingObjectSelector.AllOnSheet(sheetName);
  }

  /** Returns one exact pivot-table target by pivot-table name. */
  public static PivotTableRef pivotTable(String name) {
    return new PivotTableRef(new PivotTableSelector.ByName(name));
  }

  /** Returns one exact pivot-table target by pivot-table name plus sheet ownership. */
  public static PivotTableRef pivotTableOnSheet(String name, String sheetName) {
    return new PivotTableRef(new PivotTableSelector.ByNameOnSheet(name, sheetName));
  }

  /** Returns one selector that targets every pivot table in the workbook. */
  public static PivotTableSelector.All allPivotTables() {
    return new PivotTableSelector.All();
  }

  /** Returns one selector that targets the named pivot tables. */
  public static PivotTableSelector.ByNames pivotTables(String... names) {
    return new PivotTableSelector.ByNames(List.of(names));
  }

  /** Workbook-scoped fluent target. */
  public record WorkbookRef(WorkbookSelector selector) implements SelectorTarget {
    public WorkbookRef {
      Objects.requireNonNull(selector, "selector must not be null");
    }

    /** Returns one workbook-summary inspection step. */
    public PlannedInspection summary() {
      return new PlannedInspection(selector, Queries.workbookSummary());
    }

    /** Returns one package-security inspection step. */
    public PlannedInspection packageSecurity() {
      return new PlannedInspection(selector, Queries.packageSecurity());
    }

    /** Returns one workbook-protection inspection step. */
    public PlannedInspection protection() {
      return new PlannedInspection(selector, Queries.workbookProtection());
    }

    /** Returns one aggregate workbook-findings analysis step. */
    public PlannedInspection findings() {
      return new PlannedInspection(selector, Queries.workbookFindings());
    }

    /** Returns one workbook-protection mutation step. */
    public PlannedMutation protect(
        dev.erst.gridgrind.contract.dto.WorkbookProtectionInput protection) {
      return new PlannedMutation(selector, new MutationAction.SetWorkbookProtection(protection));
    }

    /** Returns one workbook-protection clearing mutation step. */
    public PlannedMutation clearProtection() {
      return new PlannedMutation(selector, new MutationAction.ClearWorkbookProtection());
    }
  }

  /** Sheet-scoped fluent target. */
  public record SheetRef(SheetSelector.ByName selector) implements SelectorTarget {
    public SheetRef {
      Objects.requireNonNull(selector, "selector must not be null");
    }

    /** Returns one ensure-sheet mutation step. */
    public PlannedMutation ensureExists() {
      return new PlannedMutation(selector, new MutationAction.EnsureSheet());
    }

    /** Returns one rename-sheet mutation step. */
    public PlannedMutation renameTo(String newSheetName) {
      return new PlannedMutation(selector, new MutationAction.RenameSheet(newSheetName));
    }

    /** Returns one delete-sheet mutation step. */
    public PlannedMutation delete() {
      return new PlannedMutation(selector, new MutationAction.DeleteSheet());
    }

    /** Returns one sheet-zoom mutation step. */
    public PlannedMutation setZoom(int zoomPercent) {
      return new PlannedMutation(selector, new MutationAction.SetSheetZoom(zoomPercent));
    }

    /** Returns one sheet-presentation mutation step. */
    public PlannedMutation setPresentation(SheetPresentationInput presentation) {
      return new PlannedMutation(selector, new MutationAction.SetSheetPresentation(presentation));
    }

    /** Returns one print-layout mutation step for this sheet. */
    public PlannedMutation setPrintLayout(PrintLayoutInput printLayout) {
      return new PlannedMutation(selector, new MutationAction.SetPrintLayout(printLayout));
    }

    /** Returns one print-layout clearing mutation step for this sheet. */
    public PlannedMutation clearPrintLayout() {
      return new PlannedMutation(selector, new MutationAction.ClearPrintLayout());
    }

    /** Returns one sheet-summary inspection step. */
    public PlannedInspection summary() {
      return new PlannedInspection(selector, Queries.sheetSummary());
    }

    /** Returns one sheet-layout inspection step. */
    public PlannedInspection layout() {
      return new PlannedInspection(selector, Queries.sheetLayout());
    }

    /** Returns one print-layout inspection step. */
    public PlannedInspection printLayout() {
      return new PlannedInspection(selector, Queries.printLayout());
    }

    /** Returns one merged-regions inspection step. */
    public PlannedInspection mergedRegions() {
      return new PlannedInspection(selector, Queries.mergedRegions());
    }

    /** Returns one autofilter inspection step. */
    public PlannedInspection autofilters() {
      return new PlannedInspection(selector, Queries.autofilters());
    }

    /** Returns one chart inventory inspection step for this sheet. */
    public PlannedInspection charts() {
      return new PlannedInspection(selector, Queries.charts());
    }

    /** Returns one drawing-object inventory inspection step for this sheet. */
    public PlannedInspection drawingObjects() {
      return new PlannedInspection(
          new DrawingObjectSelector.AllOnSheet(selector.name()), Queries.drawingObjects());
    }

    /** Returns one formula-surface inspection step for this sheet. */
    public PlannedInspection formulaSurface() {
      return new PlannedInspection(selector, Queries.formulaSurface());
    }

    /** Returns one formula-health analysis step for this sheet. */
    public PlannedInspection formulaHealth() {
      return new PlannedInspection(selector, Queries.formulaHealth());
    }
  }

  /** Exact A1 cell fluent target. */
  public record CellRef(CellSelector.ByAddress selector) implements SelectorTarget {
    public CellRef {
      Objects.requireNonNull(selector, "selector must not be null");
    }

    /** Returns one set-cell mutation step. */
    public PlannedMutation set(CellInput value) {
      return new PlannedMutation(selector, new MutationAction.SetCell(value));
    }

    /** Returns one set-hyperlink mutation step. */
    public PlannedMutation setHyperlink(HyperlinkTarget hyperlink) {
      return new PlannedMutation(selector, new MutationAction.SetHyperlink(hyperlink));
    }

    /** Returns one clear-hyperlink mutation step. */
    public PlannedMutation clearHyperlink() {
      return new PlannedMutation(selector, new MutationAction.ClearHyperlink());
    }

    /** Returns one set-comment mutation step. */
    public PlannedMutation setComment(CommentInput comment) {
      return new PlannedMutation(selector, new MutationAction.SetComment(comment));
    }

    /** Returns one clear-comment mutation step. */
    public PlannedMutation clearComment() {
      return new PlannedMutation(selector, new MutationAction.ClearComment());
    }

    /** Returns one exact-cell inspection step. */
    public PlannedInspection read() {
      return new PlannedInspection(selector, Queries.cells());
    }

    /** Returns one hyperlink inspection step for this cell. */
    public PlannedInspection hyperlinks() {
      return new PlannedInspection(selector, Queries.hyperlinks());
    }

    /** Returns one comment inspection step for this cell. */
    public PlannedInspection comments() {
      return new PlannedInspection(selector, Queries.comments());
    }

    /** Returns one effective-value assertion step. */
    public PlannedAssertion valueEquals(ExpectedCellValue expectedValue) {
      return new PlannedAssertion(selector, Checks.cellValue(expectedValue));
    }

    /** Returns one rendered display-value assertion step. */
    public PlannedAssertion displayValueEquals(String displayValue) {
      return new PlannedAssertion(selector, Checks.displayValue(displayValue));
    }

    /** Returns one formula-text assertion step. */
    public PlannedAssertion formulaEquals(String formula) {
      return new PlannedAssertion(selector, Checks.formulaText(formula));
    }

    /** Returns one cell-style assertion step. */
    public PlannedAssertion styleEquals(GridGrindResponse.CellStyleReport style) {
      return new PlannedAssertion(selector, Checks.cellStyle(style));
    }
  }

  /** Exact A1 range fluent target. */
  public record RangeRef(RangeSelector.ByRange selector) implements SelectorTarget {
    public RangeRef {
      Objects.requireNonNull(selector, "selector must not be null");
    }

    /** Returns one set-range mutation step. */
    public PlannedMutation setRows(List<List<CellInput>> rows) {
      return new PlannedMutation(selector, new MutationAction.SetRange(rows));
    }

    /** Returns one clear-range mutation step. */
    public PlannedMutation clear() {
      return new PlannedMutation(selector, new MutationAction.ClearRange());
    }

    /** Returns one merge-cells mutation step. */
    public PlannedMutation merge() {
      return new PlannedMutation(selector, new MutationAction.MergeCells());
    }

    /** Returns one unmerge-cells mutation step. */
    public PlannedMutation unmerge() {
      return new PlannedMutation(selector, new MutationAction.UnmergeCells());
    }

    /** Returns one apply-style mutation step. */
    public PlannedMutation applyStyle(dev.erst.gridgrind.contract.dto.CellStyleInput style) {
      return new PlannedMutation(selector, new MutationAction.ApplyStyle(style));
    }

    /** Returns one set-data-validation mutation step. */
    public PlannedMutation setDataValidation(DataValidationInput validation) {
      return new PlannedMutation(selector, new MutationAction.SetDataValidation(validation));
    }

    /** Returns one data-validations inspection step. */
    public PlannedInspection dataValidations() {
      return new PlannedInspection(selector, Queries.dataValidations());
    }

    /** Returns one conditional-formatting inspection step. */
    public PlannedInspection conditionalFormatting() {
      return new PlannedInspection(selector, Queries.conditionalFormatting());
    }
  }

  /** Window-scoped fluent target. */
  public record WindowRef(RangeSelector.RectangularWindow selector) implements SelectorTarget {
    public WindowRef {
      Objects.requireNonNull(selector, "selector must not be null");
    }

    /** Returns one rectangular-window inspection step. */
    public PlannedInspection read() {
      return new PlannedInspection(selector, Queries.window());
    }

    /** Returns one sheet-schema inspection step for this window. */
    public PlannedInspection schema() {
      return new PlannedInspection(selector, Queries.sheetSchema());
    }
  }

  /** Table-scoped fluent target. */
  public record TableRef(TableSelector selector) implements SelectorTarget {
    public TableRef {
      Objects.requireNonNull(selector, "selector must not be null");
    }

    /** Returns one exact table-row target by zero-based table row index. */
    public TableRowRef row(int rowIndex) {
      return new TableRowRef(new TableRowSelector.ByIndex(selector, rowIndex));
    }

    /** Returns one exact table-row target by logical key-column match. */
    public TableRowRef rowByKey(String columnName, CellInput expectedValue) {
      return new TableRowRef(new TableRowSelector.ByKeyCell(selector, columnName, expectedValue));
    }

    /** Returns one set-table mutation step. */
    public PlannedMutation define(TableInput table) {
      return new PlannedMutation(selector, new MutationAction.SetTable(table));
    }

    /** Returns one delete-table mutation step. */
    public PlannedMutation delete() {
      return new PlannedMutation(selector, new MutationAction.DeleteTable());
    }

    /** Returns one table inspection step. */
    public PlannedInspection inspect() {
      return new PlannedInspection(selector, Queries.tables());
    }

    /** Returns one table-health analysis step. */
    public PlannedInspection analyzeHealth() {
      return new PlannedInspection(selector, Queries.tableHealth());
    }

    /** Returns one table presence assertion step. */
    public PlannedAssertion present() {
      return new PlannedAssertion(selector, Checks.present());
    }

    /** Returns one table absence assertion step. */
    public PlannedAssertion absent() {
      return new PlannedAssertion(selector, Checks.absent());
    }
  }

  /** Table-row fluent target used to address one logical cell by column name. */
  public record TableRowRef(TableRowSelector selector) {
    public TableRowRef {
      Objects.requireNonNull(selector, "selector must not be null");
    }

    /** Returns one exact logical table-cell target by column name. */
    public TableCellRef cell(String columnName) {
      return new TableCellRef(new TableCellSelector.ByColumnName(selector, columnName));
    }
  }

  /** Table-cell fluent target compiled to one canonical table-cell selector. */
  public record TableCellRef(TableCellSelector.ByColumnName selector) implements SelectorTarget {
    public TableCellRef {
      Objects.requireNonNull(selector, "selector must not be null");
    }

    /** Returns one set-cell mutation step against the logical table cell. */
    public PlannedMutation set(CellInput value) {
      return new PlannedMutation(selector, new MutationAction.SetCell(value));
    }

    /** Returns one set-hyperlink mutation step against the logical table cell. */
    public PlannedMutation setHyperlink(HyperlinkTarget hyperlink) {
      return new PlannedMutation(selector, new MutationAction.SetHyperlink(hyperlink));
    }

    /** Returns one clear-hyperlink mutation step against the logical table cell. */
    public PlannedMutation clearHyperlink() {
      return new PlannedMutation(selector, new MutationAction.ClearHyperlink());
    }

    /** Returns one set-comment mutation step against the logical table cell. */
    public PlannedMutation setComment(CommentInput comment) {
      return new PlannedMutation(selector, new MutationAction.SetComment(comment));
    }

    /** Returns one clear-comment mutation step against the logical table cell. */
    public PlannedMutation clearComment() {
      return new PlannedMutation(selector, new MutationAction.ClearComment());
    }

    /** Returns one exact-cell inspection step against the logical table cell. */
    public PlannedInspection read() {
      return new PlannedInspection(selector, Queries.cells());
    }

    /** Returns one hyperlink inspection step against the logical table cell. */
    public PlannedInspection hyperlinks() {
      return new PlannedInspection(selector, Queries.hyperlinks());
    }

    /** Returns one comment inspection step against the logical table cell. */
    public PlannedInspection comments() {
      return new PlannedInspection(selector, Queries.comments());
    }

    /** Returns one effective-value assertion step against the logical table cell. */
    public PlannedAssertion valueEquals(ExpectedCellValue expectedValue) {
      return new PlannedAssertion(selector, Checks.cellValue(expectedValue));
    }

    /** Returns one rendered display-value assertion step against the logical table cell. */
    public PlannedAssertion displayValueEquals(String displayValue) {
      return new PlannedAssertion(selector, Checks.displayValue(displayValue));
    }

    /** Returns one formula-text assertion step against the logical table cell. */
    public PlannedAssertion formulaEquals(String formula) {
      return new PlannedAssertion(selector, Checks.formulaText(formula));
    }

    /** Returns one cell-style assertion step against the logical table cell. */
    public PlannedAssertion styleEquals(GridGrindResponse.CellStyleReport style) {
      return new PlannedAssertion(selector, Checks.cellStyle(style));
    }
  }

  /** Named-range fluent target. */
  public record NamedRangeRef(NamedRangeSelector selector) implements SelectorTarget {
    public NamedRangeRef {
      Objects.requireNonNull(selector, "selector must not be null");
    }

    /** Returns one set-named-range mutation step. */
    public PlannedMutation define(NamedRangeScope scope, NamedRangeTarget target) {
      String name = exactNamedRangeName(selector);
      return new PlannedMutation(selector, new MutationAction.SetNamedRange(name, scope, target));
    }

    private static String exactNamedRangeName(NamedRangeSelector selector) {
      if (selector instanceof NamedRangeSelector.ByName byName) {
        return byName.name();
      }
      if (selector instanceof NamedRangeSelector.WorkbookScope workbookScope) {
        return workbookScope.name();
      }
      if (selector instanceof NamedRangeSelector.SheetScope sheetScope) {
        return sheetScope.name();
      }
      throw new IllegalStateException("define() requires an exact named-range selector");
    }

    /** Returns one delete-named-range mutation step. */
    public PlannedMutation delete() {
      return new PlannedMutation(selector, new MutationAction.DeleteNamedRange());
    }

    /** Returns one named-range inspection step. */
    public PlannedInspection inspect() {
      return new PlannedInspection(selector, Queries.namedRanges());
    }

    /** Returns one named-range-surface inspection step. */
    public PlannedInspection surface() {
      return new PlannedInspection(selector, Queries.namedRangeSurface());
    }

    /** Returns one named-range-health analysis step. */
    public PlannedInspection analyzeHealth() {
      return new PlannedInspection(selector, Queries.namedRangeHealth());
    }

    /** Returns one named-range presence assertion step. */
    public PlannedAssertion present() {
      return new PlannedAssertion(selector, Checks.present());
    }

    /** Returns one named-range absence assertion step. */
    public PlannedAssertion absent() {
      return new PlannedAssertion(selector, Checks.absent());
    }
  }

  /** Chart-scoped fluent target. */
  public record ChartRef(ChartSelector selector) implements SelectorTarget {
    public ChartRef {
      Objects.requireNonNull(selector, "selector must not be null");
    }

    /** Returns one set-chart mutation step on the provided sheet. */
    public PlannedMutation defineOnSheet(String sheetName, ChartInput chart) {
      return new PlannedMutation(
          new SheetSelector.ByName(sheetName), new MutationAction.SetChart(chart));
    }

    /** Returns one chart inventory inspection step on the owning sheet. */
    public PlannedInspection inspectOnSheet() {
      return switch (selector) {
        case ChartSelector.ByName byName ->
            new PlannedInspection(new SheetSelector.ByName(byName.sheetName()), Queries.charts());
        case ChartSelector.AllOnSheet allOnSheet ->
            new PlannedInspection(
                new SheetSelector.ByName(allOnSheet.sheetName()), Queries.charts());
      };
    }

    /** Returns one chart presence assertion step. */
    public PlannedAssertion present() {
      return new PlannedAssertion(selector, Checks.present());
    }

    /** Returns one chart absence assertion step. */
    public PlannedAssertion absent() {
      return new PlannedAssertion(selector, Checks.absent());
    }
  }

  /** Pivot-table-scoped fluent target. */
  public record PivotTableRef(PivotTableSelector selector) implements SelectorTarget {
    public PivotTableRef {
      Objects.requireNonNull(selector, "selector must not be null");
    }

    /** Returns one set-pivot-table mutation step. */
    public PlannedMutation define(PivotTableInput pivotTable) {
      return new PlannedMutation(selector, new MutationAction.SetPivotTable(pivotTable));
    }

    /** Returns one delete-pivot-table mutation step. */
    public PlannedMutation delete() {
      return new PlannedMutation(selector, new MutationAction.DeletePivotTable());
    }

    /** Returns one pivot-table inspection step. */
    public PlannedInspection inspect() {
      return new PlannedInspection(selector, Queries.pivotTables());
    }

    /** Returns one pivot-table-health analysis step. */
    public PlannedInspection analyzeHealth() {
      return new PlannedInspection(selector, Queries.pivotTableHealth());
    }

    /** Returns one pivot-table presence assertion step. */
    public PlannedAssertion present() {
      return new PlannedAssertion(selector, Checks.present());
    }

    /** Returns one pivot-table absence assertion step. */
    public PlannedAssertion absent() {
      return new PlannedAssertion(selector, Checks.absent());
    }
  }
}
