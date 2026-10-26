package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.ResolvedTarget;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.vocabulary.ExcelColumnSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelRowSpan;
import java.util.List;

/** Applies executable core actions directly at the workbook boundary. */
public final class ResolvedMutationExecutor {
  private ResolvedMutationExecutor() {}

  /**
   * Applies one core action and invalidates formula and package state after successful mutation.
   */
  @SuppressWarnings(
      "PMD.NcssCount") // One exhaustive dispatch over the canonical mutation variants.
  public static void apply(ExcelWorkbook workbook, ResolvedStep step) {
    ResolvedTarget target = step.target();
    switch (step.action()) {
      case WorkbookMutationAction.EnsureSheet _ -> workbook.getOrCreateSheet(sheet(target));
      case WorkbookMutationAction.RenameSheet action ->
          workbook.sheets().renameSheet(sheet(target), action.newSheetName().value());
      case WorkbookMutationAction.DeleteSheet _ -> workbook.sheets().deleteSheet(sheet(target));
      case WorkbookMutationAction.MoveSheet action ->
          workbook.sheets().moveSheet(sheet(target), action.targetIndex());
      case WorkbookMutationAction.CopySheet action ->
          workbook
              .sheets()
              .copySheet(sheet(target), action.newSheetName().value(), action.position());
      case WorkbookMutationAction.SetActiveSheet _ ->
          workbook.sheets().setActiveSheet(sheet(target));
      case WorkbookMutationAction.SetSelectedSheets _ ->
          workbook
              .sheets()
              .setSelectedSheets(
                  ((SheetSelector.ByNames) target)
                      .names().stream()
                          .map(dev.erst.gridgrind.core.vocabulary.SheetName::value)
                          .toList());
      case WorkbookMutationAction.SetSheetVisibility action ->
          workbook.sheets().setSheetVisibility(sheet(target), action.visibility());
      case ResolvedMutationAction.SetSheetProtection action ->
          workbook
              .sheets()
              .setSheetProtection(sheet(target), action.protection(), action.password());
      case WorkbookMutationAction.ClearSheetProtection _ ->
          workbook.sheets().clearSheetProtection(sheet(target));
      case ResolvedMutationAction.SetWorkbookProtection action ->
          workbook.protection().setWorkbookProtection(action.protection());
      case WorkbookMutationAction.ClearWorkbookProtection _ ->
          workbook.protection().clearWorkbookProtection();
      case WorkbookMutationAction.MergeCells _ ->
          workbook.sheet(sheet(target)).layout().mergeCells(range(target));
      case WorkbookMutationAction.UnmergeCells _ ->
          workbook.sheet(sheet(target)).layout().unmergeCells(range(target));
      case WorkbookMutationAction.SetColumnWidth action ->
          workbook
              .sheet(sheet(target))
              .columns()
              .setWidth(
                  ((ColumnBandSelector.Span) target).firstColumnIndex(),
                  ((ColumnBandSelector.Span) target).lastColumnIndex(),
                  action.widthCharacters());
      case WorkbookMutationAction.SetRowHeight action ->
          workbook
              .sheet(sheet(target))
              .rows()
              .setHeight(
                  ((RowBandSelector.Span) target).firstRowIndex(),
                  ((RowBandSelector.Span) target).lastRowIndex(),
                  action.heightPoints());
      case WorkbookMutationAction.InsertRows _ ->
          workbook
              .sheet(sheet(target))
              .rows()
              .insertRows(
                  ((RowBandSelector.Insertion) target).beforeRowIndex(),
                  ((RowBandSelector.Insertion) target).rowCount());
      case WorkbookMutationAction.DeleteRows _ ->
          workbook.sheet(sheet(target)).rows().deleteRows(rows(target));
      case WorkbookMutationAction.ShiftRows action ->
          workbook.sheet(sheet(target)).rows().shiftRows(rows(target), action.delta());
      case WorkbookMutationAction.InsertColumns _ ->
          workbook
              .sheet(sheet(target))
              .columns()
              .insert(
                  ((ColumnBandSelector.Insertion) target).beforeColumnIndex(),
                  ((ColumnBandSelector.Insertion) target).columnCount());
      case WorkbookMutationAction.DeleteColumns _ ->
          workbook.sheet(sheet(target)).columns().delete(columns(target));
      case WorkbookMutationAction.ShiftColumns action ->
          workbook.sheet(sheet(target)).columns().shift(columns(target), action.delta());
      case WorkbookMutationAction.SetRowVisibility action ->
          workbook.sheet(sheet(target)).rows().setVisibility(rows(target), action.hidden());
      case WorkbookMutationAction.SetColumnVisibility action ->
          workbook.sheet(sheet(target)).columns().setVisibility(columns(target), action.hidden());
      case WorkbookMutationAction.GroupRows action ->
          workbook.sheet(sheet(target)).rows().group(rows(target), action.collapsed());
      case WorkbookMutationAction.UngroupRows _ ->
          workbook.sheet(sheet(target)).rows().ungroup(rows(target));
      case WorkbookMutationAction.GroupColumns action ->
          workbook.sheet(sheet(target)).columns().group(columns(target), action.collapsed());
      case WorkbookMutationAction.UngroupColumns _ ->
          workbook.sheet(sheet(target)).columns().ungroup(columns(target));
      case WorkbookMutationAction.SetSheetPane action ->
          workbook.sheet(sheet(target)).layout().setPane(action.pane());
      case WorkbookMutationAction.SetSheetZoom action ->
          workbook.sheet(sheet(target)).layout().setZoom(action.zoomPercent());
      case WorkbookMutationAction.SetSheetPresentation action ->
          workbook.sheet(sheet(target)).layout().setPresentation(action.presentation());
      case ResolvedMutationAction.SetPrintLayout action ->
          workbook.sheet(sheet(target)).layout().setPrintLayout(action.printLayout());
      case WorkbookMutationAction.ClearPrintLayout _ ->
          workbook.sheet(sheet(target)).layout().clearPrintLayout();
      case WorkbookMutationAction.AutoSizeColumns _ ->
          workbook.sheet(sheet(target)).columns().autoSize();
      case ResolvedMutationAction.SetCell action ->
          workbook.sheet(sheet(target)).cells().setCell(address(target), action.value());
      case ResolvedMutationAction.SetRange action ->
          workbook.sheet(sheet(target)).cells().setRange(range(target), action.rows());
      case CellMutationAction.ClearRange _ ->
          workbook.sheet(sheet(target)).cells().clearRange(range(target));
      case ResolvedMutationAction.SetArrayFormula action ->
          workbook.sheet(sheet(target)).cells().setArrayFormula(range(target), action.formula());
      case CellMutationAction.ClearArrayFormula _ ->
          workbook.sheet(sheet(target)).cells().clearArrayFormula(address(target));
      case ResolvedMutationAction.AppendRow action ->
          workbook
              .sheet(sheet(target))
              .cells()
              .appendRow(action.values().toArray(ResolvedCellWrite[]::new));
      case CellMutationAction.SetHyperlink action ->
          workbook
              .sheet(sheet(target))
              .annotations()
              .setHyperlink(address(target), action.target());
      case CellMutationAction.ClearHyperlink _ ->
          workbook.sheet(sheet(target)).annotations().clearHyperlink(address(target));
      case ResolvedMutationAction.SetComment action ->
          workbook.sheet(sheet(target)).annotations().setComment(address(target), action.comment());
      case CellMutationAction.ClearComment _ ->
          workbook.sheet(sheet(target)).annotations().clearComment(address(target));
      case CellMutationAction.ApplyStyle action ->
          workbook.sheet(sheet(target)).cells().applyStyle(range(target), action.style());
      case ResolvedMutationAction.ImportCustomXmlMapping action ->
          workbook.customXml().importCustomXmlMapping(action.mapping());
      case StructuredMutationAction.SetNamedRange action -> workbook.names().setNamedRange(action);
      case StructuredMutationAction.DeleteNamedRange _ -> {
        var scoped = (NamedRangeSelector.ScopedExact) target;
        switch (scoped) {
          case NamedRangeSelector.WorkbookScope named ->
              workbook
                  .names()
                  .deleteNamedRange(
                      named.name().value(),
                      new dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport.Workbook());
          case NamedRangeSelector.SheetScope named ->
              workbook
                  .names()
                  .deleteNamedRange(
                      named.name().value(),
                      new dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport.Sheet(
                          new dev.erst.gridgrind.core.vocabulary.SheetLabel(
                              named.sheetName().value())));
        }
      }
      case ResolvedMutationAction.SetPicture action ->
          workbook.sheet(sheet(target)).drawings().setPicture(action.picture());
      case ResolvedMutationAction.SetChart action ->
          workbook.sheet(sheet(target)).drawings().setChart(action.chart());
      case ResolvedMutationAction.SetShape action ->
          workbook.sheet(sheet(target)).drawings().setShape(action.shape());
      case ResolvedMutationAction.SetSignatureLine action ->
          workbook.sheet(sheet(target)).drawings().setSignatureLine(action.signatureLine());
      case ResolvedMutationAction.SetEmbeddedObject action ->
          workbook.sheet(sheet(target)).drawings().setEmbeddedObject(action.embeddedObject());
      case DrawingMutationAction.SetDrawingObjectAnchor action ->
          workbook
              .sheet(sheet(target))
              .drawings()
              .setDrawingObjectAnchor(
                  ((DrawingObjectSelector.ByName) target).objectName(),
                  (dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec.TwoCell)
                      action.anchor());
      case DrawingMutationAction.DeleteDrawingObject _ ->
          workbook
              .sheet(sheet(target))
              .drawings()
              .deleteDrawingObject(((DrawingObjectSelector.ByName) target).objectName());
      case ResolvedMutationAction.SetDataValidation action ->
          workbook
              .sheet(sheet(target))
              .metadata()
              .setDataValidation(range(target), action.validation());
      case StructuredMutationAction.ClearDataValidations _ ->
          workbook.sheet(sheet(target)).metadata().clearDataValidations(selection(target));
      case StructuredMutationAction.SetConditionalFormatting action ->
          workbook
              .sheet(sheet(target))
              .metadata()
              .setConditionalFormatting(ranges(target), action.conditionalFormatting());
      case StructuredMutationAction.ClearConditionalFormatting _ ->
          workbook.sheet(sheet(target)).metadata().clearConditionalFormatting(selection(target));
      case StructuredMutationAction.SetAutofilter action ->
          workbook
              .sheet(sheet(target))
              .metadata()
              .setAutofilter(range(target), action.criteria(), action.sortState());
      case StructuredMutationAction.ClearAutofilter _ ->
          workbook.sheet(sheet(target)).metadata().clearAutofilter();
      case ResolvedMutationAction.SetTable action -> workbook.tables().setTable(action.table());
      case StructuredMutationAction.SetPivotTable action ->
          workbook.pivots().setPivotTable(action.pivotTable());
      case StructuredMutationAction.DeleteTable _ ->
          workbook
              .tables()
              .deleteTable(((TableSelector.ByNameOnSheet) target).name().value(), sheet(target));
      case StructuredMutationAction.DeletePivotTable _ ->
          workbook
              .pivots()
              .deletePivotTable(((PivotTableSelector.ByNameOnSheet) target).name(), sheet(target));
    }
    workbook.markPackageMutated();
    workbook.invalidateFormulaRuntime();
  }

  private static String sheet(ResolvedTarget target) {
    return switch (target) {
      case SheetSelector.ByName sheet -> sheet.name().value();
      case CellSelector.ByAddress cell -> cell.sheetName().value();
      case RangeSelector.ByRange range -> range.sheetName().value();
      case RangeSelector.ByRanges ranges -> ranges.sheetName().value();
      case RangeSelector.AllOnSheet all -> all.sheetName().value();
      case RowBandSelector.Span rows -> rows.sheetName().value();
      case RowBandSelector.Insertion rows -> rows.sheetName().value();
      case ColumnBandSelector.Span columns -> columns.sheetName().value();
      case ColumnBandSelector.Insertion columns -> columns.sheetName().value();
      case DrawingObjectSelector.ByName object -> object.sheetName().value();
      case TableSelector.ByNameOnSheet table -> table.sheetName().value();
      case PivotTableSelector.ByNameOnSheet pivot -> pivot.sheetName().value();
      default -> throw new IllegalArgumentException("Target has no single sheet: " + target);
    };
  }

  private static String address(ResolvedTarget target) {
    return ((CellSelector.ByAddress) target).address().value();
  }

  private static String range(ResolvedTarget target) {
    return ((RangeSelector.ByRange) target).range().value();
  }

  private static ExcelRowSpan rows(ResolvedTarget target) {
    var rows = (RowBandSelector.Span) target;
    return new ExcelRowSpan(rows.firstRowIndex(), rows.lastRowIndex());
  }

  private static ExcelColumnSpan columns(ResolvedTarget target) {
    var columns = (ColumnBandSelector.Span) target;
    return new ExcelColumnSpan(columns.firstColumnIndex(), columns.lastColumnIndex());
  }

  private static List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges(ResolvedTarget target) {
    return switch (target) {
      case RangeSelector.ByRange range -> List.of(range.range());
      case RangeSelector.ByRanges ranges -> ranges.ranges();
      default ->
          throw new IllegalArgumentException(
              "SET_CONDITIONAL_FORMATTING requires explicit range targets");
    };
  }

  private static ExcelRangeSelection selection(ResolvedTarget target) {
    return target instanceof RangeSelector.AllOnSheet
        ? new ExcelRangeSelection.All()
        : new ExcelRangeSelection.Selected(
            ranges(target).stream()
                .map(dev.erst.gridgrind.core.vocabulary.RangeRef::value)
                .toList());
  }
}
