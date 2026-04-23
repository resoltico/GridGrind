package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

/** Selector builders and focused target-scoped step builders for the fluent Java authoring API. */
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

  /** Returns one exact cell target by sheet name and A1 address. */
  public static CellRef cell(String sheetName, String address) {
    return new CellRef(new CellSelector.ByAddress(sheetName, address));
  }

  /** Returns one exact range target by sheet name and A1 range. */
  public static RangeRef range(String sheetName, String range) {
    return new RangeRef(new RangeSelector.ByRange(sheetName, range));
  }

  /** Returns one rectangular window target from a top-left cell plus dimensions. */
  public static WindowRef window(
      String sheetName, String topLeftAddress, int rowCount, int columnCount) {
    return new WindowRef(
        new RangeSelector.RectangularWindow(sheetName, topLeftAddress, rowCount, columnCount));
  }

  /** Returns one exact table target by table name. */
  public static TableRef table(String name) {
    return new TableRef(new TableSelector.ByName(name));
  }

  /** Returns one exact table target by table name plus sheet ownership. */
  public static TableRef tableOnSheet(String name, String sheetName) {
    return new TableRef(new TableSelector.ByNameOnSheet(name, sheetName));
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

  /** Returns one exact chart target by sheet name and chart name. */
  public static ChartRef chart(String sheetName, String chartName) {
    return new ChartRef(new ChartSelector.ByName(sheetName, chartName));
  }

  /** Returns one exact pivot-table target by pivot-table name. */
  public static PivotTableRef pivotTable(String name) {
    return new PivotTableRef(new PivotTableSelector.ByName(name));
  }

  /** Returns one exact pivot-table target by pivot-table name plus sheet ownership. */
  public static PivotTableRef pivotTableOnSheet(String name, String sheetName) {
    return new PivotTableRef(new PivotTableSelector.ByNameOnSheet(name, sheetName));
  }

  /** Workbook-scoped fluent target. */
  public static final class WorkbookRef {
    private final WorkbookSelector selector;

    WorkbookRef(WorkbookSelector selector) {
      this.selector = Objects.requireNonNull(selector, "selector must not be null");
    }

    WorkbookSelector selector() {
      return selector;
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
  }

  /** Sheet-scoped fluent target. */
  public static final class SheetRef {
    private final SheetSelector.ByName selector;

    SheetRef(SheetSelector.ByName selector) {
      this.selector = Objects.requireNonNull(selector, "selector must not be null");
    }

    SheetSelector.ByName selector() {
      return selector;
    }

    /** Returns one sheet-creation mutation step. */
    public PlannedMutation ensureExists() {
      return new PlannedMutation(selector, new MutationAction.EnsureSheet());
    }

    /** Returns one sheet-rename mutation step. */
    public PlannedMutation renameTo(String newSheetName) {
      return new PlannedMutation(selector, new MutationAction.RenameSheet(newSheetName));
    }

    /** Returns one sheet-delete mutation step. */
    public PlannedMutation delete() {
      return new PlannedMutation(selector, new MutationAction.DeleteSheet());
    }

    /** Returns one sheet-zoom mutation step. */
    public PlannedMutation setZoom(int zoomPercent) {
      return new PlannedMutation(selector, new MutationAction.SetSheetZoom(zoomPercent));
    }

    /** Returns one print-layout clearing mutation step. */
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

    /** Returns one formula-surface inspection step. */
    public PlannedInspection formulaSurface() {
      return new PlannedInspection(selector, Queries.formulaSurface());
    }

    /** Returns one formula-health analysis step. */
    public PlannedInspection formulaHealth() {
      return new PlannedInspection(selector, Queries.formulaHealth());
    }
  }

  /** Exact A1 cell fluent target. */
  public static final class CellRef {
    private final CellSelector.ByAddress selector;

    CellRef(CellSelector.ByAddress selector) {
      this.selector = Objects.requireNonNull(selector, "selector must not be null");
    }

    CellSelector.ByAddress selector() {
      return selector;
    }

    /** Returns one set-cell mutation step. */
    public PlannedMutation set(Values.CellValue value) {
      return new PlannedMutation(selector, new MutationAction.SetCell(Values.toCellInput(value)));
    }

    /** Returns one set-hyperlink mutation step. */
    public PlannedMutation setHyperlink(Links.Target hyperlink) {
      return new PlannedMutation(
          selector, new MutationAction.SetHyperlink(Links.toHyperlinkTarget(hyperlink)));
    }

    /** Returns one clear-hyperlink mutation step. */
    public PlannedMutation clearHyperlink() {
      return new PlannedMutation(selector, new MutationAction.ClearHyperlink());
    }

    /** Returns one set-comment mutation step. */
    public PlannedMutation setComment(Values.Comment comment) {
      return new PlannedMutation(
          selector, new MutationAction.SetComment(Values.toCommentInput(comment)));
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
    public PlannedAssertion valueEquals(Values.ExpectedValue expectedValue) {
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
  }

  /** Exact A1 range fluent target. */
  public static final class RangeRef {
    private final RangeSelector.ByRange selector;

    RangeRef(RangeSelector.ByRange selector) {
      this.selector = Objects.requireNonNull(selector, "selector must not be null");
    }

    RangeSelector.ByRange selector() {
      return selector;
    }

    /** Returns one set-range mutation step. */
    public PlannedMutation setRows(List<List<Values.CellValue>> rows) {
      Objects.requireNonNull(rows, "rows must not be null");
      return new PlannedMutation(
          selector,
          new MutationAction.SetRange(
              rows.stream()
                  .map(
                      row ->
                          row.stream()
                              .map(Values::toCellInput)
                              .collect(Collectors.toUnmodifiableList()))
                  .collect(Collectors.toUnmodifiableList())));
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
  public static final class WindowRef {
    private final RangeSelector.RectangularWindow selector;

    WindowRef(RangeSelector.RectangularWindow selector) {
      this.selector = Objects.requireNonNull(selector, "selector must not be null");
    }

    RangeSelector.RectangularWindow selector() {
      return selector;
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
  public static final class TableRef {
    private final TableSelector selector;

    TableRef(TableSelector selector) {
      this.selector = Objects.requireNonNull(selector, "selector must not be null");
    }

    TableSelector selector() {
      return selector;
    }

    /** Returns one exact table-row target by zero-based table row index. */
    public TableRowRef row(int rowIndex) {
      return new TableRowRef(new TableRowSelector.ByIndex(selector, rowIndex));
    }

    /** Returns one exact table-row target by logical key-column match. */
    public TableRowRef rowByKey(String columnName, Values.CellValue expectedValue) {
      return new TableRowRef(
          new TableRowSelector.ByKeyCell(selector, columnName, Values.toCellInput(expectedValue)));
    }

    /** Returns one set-table mutation step. */
    public PlannedMutation define(Tables.Definition table) {
      return new PlannedMutation(selector, new MutationAction.SetTable(Tables.toTableInput(table)));
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
  public static final class TableRowRef {
    private final TableRowSelector selector;

    TableRowRef(TableRowSelector selector) {
      this.selector = Objects.requireNonNull(selector, "selector must not be null");
    }

    TableRowSelector selector() {
      return selector;
    }

    /** Returns one exact logical table-cell target by column name. */
    public TableCellRef cell(String columnName) {
      return new TableCellRef(new TableCellSelector.ByColumnName(selector, columnName));
    }
  }

  /** Table-cell fluent target compiled to one canonical table-cell selector. */
  public static final class TableCellRef {
    private final TableCellSelector.ByColumnName selector;

    TableCellRef(TableCellSelector.ByColumnName selector) {
      this.selector = Objects.requireNonNull(selector, "selector must not be null");
    }

    TableCellSelector.ByColumnName selector() {
      return selector;
    }

    /** Returns one set-cell mutation step against the logical table cell. */
    public PlannedMutation set(Values.CellValue value) {
      return new PlannedMutation(selector, new MutationAction.SetCell(Values.toCellInput(value)));
    }

    /** Returns one set-hyperlink mutation step against the logical table cell. */
    public PlannedMutation setHyperlink(Links.Target hyperlink) {
      return new PlannedMutation(
          selector, new MutationAction.SetHyperlink(Links.toHyperlinkTarget(hyperlink)));
    }

    /** Returns one clear-hyperlink mutation step against the logical table cell. */
    public PlannedMutation clearHyperlink() {
      return new PlannedMutation(selector, new MutationAction.ClearHyperlink());
    }

    /** Returns one set-comment mutation step against the logical table cell. */
    public PlannedMutation setComment(Values.Comment comment) {
      return new PlannedMutation(
          selector, new MutationAction.SetComment(Values.toCommentInput(comment)));
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
    public PlannedAssertion valueEquals(Values.ExpectedValue expectedValue) {
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
  }

  /** Named-range fluent target. */
  public static final class NamedRangeRef {
    private final NamedRangeSelector selector;

    NamedRangeRef(NamedRangeSelector selector) {
      this.selector = Objects.requireNonNull(selector, "selector must not be null");
    }

    NamedRangeSelector selector() {
      return selector;
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
  public static final class ChartRef {
    private final ChartSelector selector;

    ChartRef(ChartSelector selector) {
      this.selector = Objects.requireNonNull(selector, "selector must not be null");
    }

    ChartSelector selector() {
      return selector;
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
  public static final class PivotTableRef {
    private final PivotTableSelector selector;

    PivotTableRef(PivotTableSelector selector) {
      this.selector = Objects.requireNonNull(selector, "selector must not be null");
    }

    PivotTableSelector selector() {
      return selector;
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
