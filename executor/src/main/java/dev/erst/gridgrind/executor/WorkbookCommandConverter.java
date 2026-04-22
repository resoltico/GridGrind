package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.ArrayFormulaInput;
import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnInput;
import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortConditionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateInput;
import dev.erst.gridgrind.contract.dto.CellAlignmentInput;
import dev.erst.gridgrind.contract.dto.CellBorderInput;
import dev.erst.gridgrind.contract.dto.CellBorderSideInput;
import dev.erst.gridgrind.contract.dto.CellFillInput;
import dev.erst.gridgrind.contract.dto.CellFontInput;
import dev.erst.gridgrind.contract.dto.CellGradientFillInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellProtectionInput;
import dev.erst.gridgrind.contract.dto.CellStyleInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ColorInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdInput;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderSideInput;
import dev.erst.gridgrind.contract.dto.DifferentialStyleInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.FontHeightInput;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.PaneInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.PrintAreaInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.PrintScalingInput;
import dev.erst.gridgrind.contract.dto.PrintSetupInput;
import dev.erst.gridgrind.contract.dto.PrintTitleColumnsInput;
import dev.erst.gridgrind.contract.dto.PrintTitleRowsInput;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.excel.ExcelArrayFormulaDefinition;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterColumn;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterCriterion;
import dev.erst.gridgrind.excel.ExcelAutofilterSortCondition;
import dev.erst.gridgrind.excel.ExcelAutofilterSortState;
import dev.erst.gridgrind.excel.ExcelBinaryData;
import dev.erst.gridgrind.excel.ExcelBorder;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelCellAlignment;
import dev.erst.gridgrind.excel.ExcelCellFill;
import dev.erst.gridgrind.excel.ExcelCellFont;
import dev.erst.gridgrind.excel.ExcelCellProtection;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelColor;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelCommentAnchor;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThreshold;
import dev.erst.gridgrind.excel.ExcelCustomXmlImportDefinition;
import dev.erst.gridgrind.excel.ExcelCustomXmlMappingLocator;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorAlert;
import dev.erst.gridgrind.excel.ExcelDataValidationPrompt;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.excel.ExcelDifferentialBorder;
import dev.erst.gridgrind.excel.ExcelDifferentialBorderSide;
import dev.erst.gridgrind.excel.ExcelDifferentialStyle;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelGradientFill;
import dev.erst.gridgrind.excel.ExcelGradientStop;
import dev.erst.gridgrind.excel.ExcelHeaderFooterText;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelIgnoredError;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.ExcelPivotTableDefinition;
import dev.erst.gridgrind.excel.ExcelPrintLayout;
import dev.erst.gridgrind.excel.ExcelPrintMargins;
import dev.erst.gridgrind.excel.ExcelPrintSetup;
import dev.erst.gridgrind.excel.ExcelRichText;
import dev.erst.gridgrind.excel.ExcelRichTextRun;
import dev.erst.gridgrind.excel.ExcelShapeDefinition;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetDefaults;
import dev.erst.gridgrind.excel.ExcelSheetDisplay;
import dev.erst.gridgrind.excel.ExcelSheetOutlineSummary;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelSheetPresentation;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelSignatureLineDefinition;
import dev.erst.gridgrind.excel.ExcelTableColumnDefinition;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSettings;
import dev.erst.gridgrind.excel.WorkbookCommand;
import java.util.Base64;

/**
 * Converts contract mutation steps and style inputs into workbook-core commands.
 *
 * <p>This translation seam intentionally spans the full mutation surface on both sides.
 */
@SuppressWarnings("PMD.ExcessiveImports")
final class WorkbookCommandConverter {
  private WorkbookCommandConverter() {}

  /** Converts one protocol mutation step into the matching workbook-core command. */
  static WorkbookCommand toCommand(MutationStep step) {
    return toCommand(step.target(), step.action());
  }

  /**
   * Converts one protocol mutation action plus selector into the matching workbook-core command.
   */
  static WorkbookCommand toCommand(Selector target, MutationAction action) {
    return switch (action) {
      case MutationAction.EnsureSheet _ ->
          new WorkbookCommand.CreateSheet(sheetByName(target, action).name());
      case MutationAction.RenameSheet renameSheet ->
          new WorkbookCommand.RenameSheet(
              sheetByName(target, action).name(), renameSheet.newSheetName());
      case MutationAction.DeleteSheet _ ->
          new WorkbookCommand.DeleteSheet(sheetByName(target, action).name());
      case MutationAction.MoveSheet moveSheet ->
          new WorkbookCommand.MoveSheet(
              sheetByName(target, action).name(), moveSheet.targetIndex());
      case MutationAction.CopySheet copySheet ->
          new WorkbookCommand.CopySheet(
              sheetByName(target, action).name(),
              copySheet.newSheetName(),
              toExcelSheetCopyPosition(copySheet.position()));
      case MutationAction.SetActiveSheet _ ->
          new WorkbookCommand.SetActiveSheet(sheetByName(target, action).name());
      case MutationAction.SetSelectedSheets _ ->
          new WorkbookCommand.SetSelectedSheets(sheetByNames(target, action).names());
      case MutationAction.SetSheetVisibility setSheetVisibility ->
          new WorkbookCommand.SetSheetVisibility(
              sheetByName(target, action).name(), setSheetVisibility.visibility());
      case MutationAction.SetSheetProtection setSheetProtection ->
          new WorkbookCommand.SetSheetProtection(
              sheetByName(target, action).name(),
              toExcelSheetProtectionSettings(setSheetProtection.protection()),
              setSheetProtection.password());
      case MutationAction.ClearSheetProtection _ ->
          new WorkbookCommand.ClearSheetProtection(sheetByName(target, action).name());
      case MutationAction.SetWorkbookProtection setWorkbookProtection ->
          new WorkbookCommand.SetWorkbookProtection(
              toExcelWorkbookProtectionSettings(setWorkbookProtection.protection()));
      case MutationAction.ClearWorkbookProtection _ ->
          new WorkbookCommand.ClearWorkbookProtection();
      case MutationAction.MergeCells _ -> {
        RangeSelector.ByRange selector = rangeByRange(target, action);
        yield new WorkbookCommand.MergeCells(selector.sheetName(), selector.range());
      }
      case MutationAction.UnmergeCells _ -> {
        RangeSelector.ByRange selector = rangeByRange(target, action);
        yield new WorkbookCommand.UnmergeCells(selector.sheetName(), selector.range());
      }
      case MutationAction.SetColumnWidth setColumnWidth -> {
        ColumnBandSelector.Span selector = columnSpan(target, action);
        yield new WorkbookCommand.SetColumnWidth(
            selector.sheetName(),
            selector.firstColumnIndex(),
            selector.lastColumnIndex(),
            setColumnWidth.widthCharacters());
      }
      case MutationAction.SetRowHeight setRowHeight -> {
        RowBandSelector.Span selector = rowSpan(target, action);
        yield new WorkbookCommand.SetRowHeight(
            selector.sheetName(),
            selector.firstRowIndex(),
            selector.lastRowIndex(),
            setRowHeight.heightPoints());
      }
      case MutationAction.InsertRows _ -> {
        RowBandSelector.Insertion selector = rowInsertion(target, action);
        yield new WorkbookCommand.InsertRows(
            selector.sheetName(), selector.beforeRowIndex(), selector.rowCount());
      }
      case MutationAction.DeleteRows _ ->
          new WorkbookCommand.DeleteRows(
              SelectorConverter.toSheetName(rowSpan(target, action)),
              SelectorConverter.toExcelRowSpan(rowSpan(target, action)));
      case MutationAction.ShiftRows shiftRows ->
          new WorkbookCommand.ShiftRows(
              SelectorConverter.toSheetName(rowSpan(target, action)),
              SelectorConverter.toExcelRowSpan(rowSpan(target, action)),
              shiftRows.delta());
      case MutationAction.InsertColumns _ -> {
        ColumnBandSelector.Insertion selector = columnInsertion(target, action);
        yield new WorkbookCommand.InsertColumns(
            selector.sheetName(), selector.beforeColumnIndex(), selector.columnCount());
      }
      case MutationAction.DeleteColumns _ ->
          new WorkbookCommand.DeleteColumns(
              SelectorConverter.toSheetName(columnSpan(target, action)),
              SelectorConverter.toExcelColumnSpan(columnSpan(target, action)));
      case MutationAction.ShiftColumns shiftColumns ->
          new WorkbookCommand.ShiftColumns(
              SelectorConverter.toSheetName(columnSpan(target, action)),
              SelectorConverter.toExcelColumnSpan(columnSpan(target, action)),
              shiftColumns.delta());
      case MutationAction.SetRowVisibility setRowVisibility ->
          new WorkbookCommand.SetRowVisibility(
              SelectorConverter.toSheetName(rowSpan(target, action)),
              SelectorConverter.toExcelRowSpan(rowSpan(target, action)),
              setRowVisibility.hidden());
      case MutationAction.SetColumnVisibility setColumnVisibility ->
          new WorkbookCommand.SetColumnVisibility(
              SelectorConverter.toSheetName(columnSpan(target, action)),
              SelectorConverter.toExcelColumnSpan(columnSpan(target, action)),
              setColumnVisibility.hidden());
      case MutationAction.GroupRows groupRows ->
          new WorkbookCommand.GroupRows(
              SelectorConverter.toSheetName(rowSpan(target, action)),
              SelectorConverter.toExcelRowSpan(rowSpan(target, action)),
              groupRows.collapsed());
      case MutationAction.UngroupRows _ ->
          new WorkbookCommand.UngroupRows(
              SelectorConverter.toSheetName(rowSpan(target, action)),
              SelectorConverter.toExcelRowSpan(rowSpan(target, action)));
      case MutationAction.GroupColumns groupColumns ->
          new WorkbookCommand.GroupColumns(
              SelectorConverter.toSheetName(columnSpan(target, action)),
              SelectorConverter.toExcelColumnSpan(columnSpan(target, action)),
              groupColumns.collapsed());
      case MutationAction.UngroupColumns _ ->
          new WorkbookCommand.UngroupColumns(
              SelectorConverter.toSheetName(columnSpan(target, action)),
              SelectorConverter.toExcelColumnSpan(columnSpan(target, action)));
      case MutationAction.SetSheetPane setSheetPane ->
          new WorkbookCommand.SetSheetPane(
              sheetByName(target, action).name(), toExcelSheetPane(setSheetPane.pane()));
      case MutationAction.SetSheetZoom setSheetZoom ->
          new WorkbookCommand.SetSheetZoom(
              sheetByName(target, action).name(), setSheetZoom.zoomPercent());
      case MutationAction.SetSheetPresentation setSheetPresentation ->
          new WorkbookCommand.SetSheetPresentation(
              sheetByName(target, action).name(),
              toExcelSheetPresentation(setSheetPresentation.presentation()));
      case MutationAction.SetPrintLayout setPrintLayout ->
          new WorkbookCommand.SetPrintLayout(
              sheetByName(target, action).name(), toExcelPrintLayout(setPrintLayout.printLayout()));
      case MutationAction.ClearPrintLayout _ ->
          new WorkbookCommand.ClearPrintLayout(sheetByName(target, action).name());
      case MutationAction.SetCell setCell -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(cellByAddress(target, action));
        yield new WorkbookCommand.SetCell(
            cellTarget.sheetName(), cellTarget.address(), toExcelCellValue(setCell.value()));
      }
      case MutationAction.SetRange setRange -> {
        RangeSelector.ByRange selector = rangeByRange(target, action);
        yield new WorkbookCommand.SetRange(
            selector.sheetName(),
            selector.range(),
            setRange.rows().stream()
                .map(row -> row.stream().map(WorkbookCommandConverter::toExcelCellValue).toList())
                .toList());
      }
      case MutationAction.ClearRange _ -> {
        RangeSelector.ByRange selector = rangeByRange(target, action);
        yield new WorkbookCommand.ClearRange(selector.sheetName(), selector.range());
      }
      case MutationAction.SetArrayFormula setArrayFormula -> {
        RangeSelector.ByRange selector = rangeByRange(target, action);
        yield new WorkbookCommand.SetArrayFormula(
            selector.sheetName(),
            selector.range(),
            toExcelArrayFormulaDefinition(setArrayFormula.formula()));
      }
      case MutationAction.ClearArrayFormula _ -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(cellByAddress(target, action));
        yield new WorkbookCommand.ClearArrayFormula(cellTarget.sheetName(), cellTarget.address());
      }
      case MutationAction.ImportCustomXmlMapping importCustomXmlMapping ->
          new WorkbookCommand.ImportCustomXmlMapping(
              toExcelCustomXmlImportDefinition(importCustomXmlMapping.mapping()));
      case MutationAction.SetHyperlink setHyperlink -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(cellByAddress(target, action));
        yield new WorkbookCommand.SetHyperlink(
            cellTarget.sheetName(), cellTarget.address(), toExcelHyperlink(setHyperlink.target()));
      }
      case MutationAction.ClearHyperlink _ -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(cellByAddress(target, action));
        yield new WorkbookCommand.ClearHyperlink(cellTarget.sheetName(), cellTarget.address());
      }
      case MutationAction.SetComment setComment -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(cellByAddress(target, action));
        yield new WorkbookCommand.SetComment(
            cellTarget.sheetName(), cellTarget.address(), toExcelComment(setComment.comment()));
      }
      case MutationAction.ClearComment _ -> {
        SelectorConverter.SingleCellTarget cellTarget =
            SelectorConverter.toSingleCellTarget(cellByAddress(target, action));
        yield new WorkbookCommand.ClearComment(cellTarget.sheetName(), cellTarget.address());
      }
      case MutationAction.SetPicture setPicture ->
          new WorkbookCommand.SetPicture(
              sheetByName(target, action).name(), toExcelPictureDefinition(setPicture.picture()));
      case MutationAction.SetSignatureLine setSignatureLine ->
          new WorkbookCommand.SetSignatureLine(
              sheetByName(target, action).name(),
              toExcelSignatureLineDefinition(setSignatureLine.signatureLine()));
      case MutationAction.SetChart setChart ->
          new WorkbookCommand.SetChart(
              sheetByName(target, action).name(), toExcelChartDefinition(setChart.chart()));
      case MutationAction.SetPivotTable setPivotTable -> {
        ensurePivotTableIdentity(target, setPivotTable);
        yield new WorkbookCommand.SetPivotTable(
            toExcelPivotTableDefinition(setPivotTable.pivotTable()));
      }
      case MutationAction.SetShape setShape ->
          new WorkbookCommand.SetShape(
              sheetByName(target, action).name(), toExcelShapeDefinition(setShape.shape()));
      case MutationAction.SetEmbeddedObject setEmbeddedObject ->
          new WorkbookCommand.SetEmbeddedObject(
              sheetByName(target, action).name(),
              toExcelEmbeddedObjectDefinition(setEmbeddedObject.embeddedObject()));
      case MutationAction.SetDrawingObjectAnchor setDrawingObjectAnchor -> {
        DrawingObjectSelector.ByName selector = drawingObjectByName(target, action);
        yield new WorkbookCommand.SetDrawingObjectAnchor(
            selector.sheetName(),
            selector.objectName(),
            toExcelDrawingAnchor(setDrawingObjectAnchor.anchor()));
      }
      case MutationAction.DeleteDrawingObject _ -> {
        DrawingObjectSelector.ByName selector = drawingObjectByName(target, action);
        yield new WorkbookCommand.DeleteDrawingObject(selector.sheetName(), selector.objectName());
      }
      case MutationAction.ApplyStyle applyStyle -> {
        RangeSelector.ByRange selector = rangeByRange(target, action);
        yield new WorkbookCommand.ApplyStyle(
            selector.sheetName(), selector.range(), toExcelCellStyle(applyStyle.style()));
      }
      case MutationAction.SetDataValidation setDataValidation -> {
        RangeSelector.ByRange selector = rangeByRange(target, action);
        yield new WorkbookCommand.SetDataValidation(
            selector.sheetName(),
            selector.range(),
            toExcelDataValidationDefinition(setDataValidation.validation()));
      }
      case MutationAction.ClearDataValidations _ -> {
        SelectorConverter.SheetLocalRangeSelection selection =
            SelectorConverter.toSheetLocalRangeSelection((RangeSelector) target);
        yield new WorkbookCommand.ClearDataValidations(
            selection.sheetName(), selection.selection());
      }
      case MutationAction.SetConditionalFormatting setConditionalFormatting ->
          new WorkbookCommand.SetConditionalFormatting(
              sheetByName(target, action).name(),
              toExcelConditionalFormattingBlock(setConditionalFormatting.conditionalFormatting()));
      case MutationAction.ClearConditionalFormatting _ -> {
        SelectorConverter.SheetLocalRangeSelection selection =
            SelectorConverter.toSheetLocalRangeSelection((RangeSelector) target);
        yield new WorkbookCommand.ClearConditionalFormatting(
            selection.sheetName(), selection.selection());
      }
      case MutationAction.SetAutofilter setAutofilter -> {
        RangeSelector.ByRange selector = rangeByRange(target, action);
        yield new WorkbookCommand.SetAutofilter(
            selector.sheetName(),
            selector.range(),
            setAutofilter.criteria().stream()
                .map(WorkbookCommandConverter::toExcelAutofilterFilterColumn)
                .toList(),
            toExcelAutofilterSortState(setAutofilter.sortState()));
      }
      case MutationAction.ClearAutofilter _ ->
          new WorkbookCommand.ClearAutofilter(sheetByName(target, action).name());
      case MutationAction.SetTable setTable -> {
        ensureTableIdentity(target, setTable);
        yield new WorkbookCommand.SetTable(toExcelTableDefinition(setTable.table()));
      }
      case MutationAction.DeleteTable _ -> {
        TableSelector.ByNameOnSheet selector = tableByNameOnSheet(target, action);
        yield new WorkbookCommand.DeleteTable(selector.name(), selector.sheetName());
      }
      case MutationAction.DeletePivotTable _ -> {
        PivotTableSelector.ByNameOnSheet selector = pivotTableByNameOnSheet(target, action);
        yield new WorkbookCommand.DeletePivotTable(selector.name(), selector.sheetName());
      }
      case MutationAction.SetNamedRange setNamedRange -> {
        ensureNamedRangeIdentity(target, setNamedRange);
        yield new WorkbookCommand.SetNamedRange(
            new ExcelNamedRangeDefinition(
                setNamedRange.name(),
                toExcelNamedRangeScope(setNamedRange.scope()),
                toExcelNamedRangeTarget(setNamedRange.target())));
      }
      case MutationAction.DeleteNamedRange _ -> {
        NamedRangeSelector.ScopedExact selector = namedRangeScopedExact(target, action);
        yield new WorkbookCommand.DeleteNamedRange(
            toExcelNamedRangeName(selector), toExcelNamedRangeScope(selector));
      }
      case MutationAction.AppendRow appendRow ->
          new WorkbookCommand.AppendRow(
              sheetByName(target, action).name(),
              appendRow.values().stream().map(WorkbookCommandConverter::toExcelCellValue).toList());
      case MutationAction.AutoSizeColumns _ ->
          new WorkbookCommand.AutoSizeColumns(sheetByName(target, action).name());
    };
  }

  private static SheetSelector.ByName sheetByName(Selector target, MutationAction action) {
    return requireTarget(target, SheetSelector.ByName.class, action.actionType());
  }

  private static SheetSelector.ByNames sheetByNames(Selector target, MutationAction action) {
    return requireTarget(target, SheetSelector.ByNames.class, action.actionType());
  }

  private static RangeSelector.ByRange rangeByRange(Selector target, MutationAction action) {
    return requireTarget(target, RangeSelector.ByRange.class, action.actionType());
  }

  private static RowBandSelector.Span rowSpan(Selector target, MutationAction action) {
    return requireTarget(target, RowBandSelector.Span.class, action.actionType());
  }

  private static RowBandSelector.Insertion rowInsertion(Selector target, MutationAction action) {
    return requireTarget(target, RowBandSelector.Insertion.class, action.actionType());
  }

  private static ColumnBandSelector.Span columnSpan(Selector target, MutationAction action) {
    return requireTarget(target, ColumnBandSelector.Span.class, action.actionType());
  }

  private static ColumnBandSelector.Insertion columnInsertion(
      Selector target, MutationAction action) {
    return requireTarget(target, ColumnBandSelector.Insertion.class, action.actionType());
  }

  private static CellSelector.ByAddress cellByAddress(Selector target, MutationAction action) {
    return requireTarget(target, CellSelector.ByAddress.class, action.actionType());
  }

  private static DrawingObjectSelector.ByName drawingObjectByName(
      Selector target, MutationAction action) {
    return requireTarget(target, DrawingObjectSelector.ByName.class, action.actionType());
  }

  private static TableSelector.ByNameOnSheet tableByNameOnSheet(
      Selector target, MutationAction action) {
    return requireTarget(target, TableSelector.ByNameOnSheet.class, action.actionType());
  }

  private static PivotTableSelector.ByNameOnSheet pivotTableByNameOnSheet(
      Selector target, MutationAction action) {
    return requireTarget(target, PivotTableSelector.ByNameOnSheet.class, action.actionType());
  }

  private static NamedRangeSelector.ScopedExact namedRangeScopedExact(
      Selector target, MutationAction action) {
    return requireTarget(target, NamedRangeSelector.ScopedExact.class, action.actionType());
  }

  private static void ensureTableIdentity(Selector target, MutationAction.SetTable action) {
    TableSelector.ByNameOnSheet selector = tableByNameOnSheet(target, action);
    if (!selector.name().equals(action.table().name())
        || !selector.sheetName().equals(action.table().sheetName())) {
      throw new IllegalArgumentException(
          action.actionType() + " target must match table.name and table.sheetName");
    }
  }

  private static void ensurePivotTableIdentity(
      Selector target, MutationAction.SetPivotTable action) {
    PivotTableSelector.ByNameOnSheet selector = pivotTableByNameOnSheet(target, action);
    if (!selector.name().equals(action.pivotTable().name())
        || !selector.sheetName().equals(action.pivotTable().sheetName())) {
      throw new IllegalArgumentException(
          action.actionType() + " target must match pivotTable.name and pivotTable.sheetName");
    }
  }

  private static void ensureNamedRangeIdentity(
      Selector target, MutationAction.SetNamedRange action) {
    NamedRangeSelector.ScopedExact selector = namedRangeScopedExact(target, action);
    if (!toExcelNamedRangeName(selector).equals(action.name())
        || !toExcelNamedRangeScope(selector).equals(toExcelNamedRangeScope(action.scope()))) {
      throw new IllegalArgumentException(
          action.actionType() + " target must match action name and scope");
    }
  }

  private static <T extends Selector> T requireTarget(
      Selector target, Class<T> expectedType, String actionType) {
    if (expectedType.isInstance(target)) {
      return expectedType.cast(target);
    }
    throw new IllegalArgumentException(
        actionType
            + " requires target type "
            + expectedType.getSimpleName()
            + " but got "
            + target.getClass().getSimpleName());
  }

  static ExcelCellValue toExcelCellValue(CellInput value) {
    return switch (value) {
      case CellInput.Blank _ -> ExcelCellValue.blank();
      case CellInput.Text text -> ExcelCellValue.text(inlineText(text.source(), "cell text"));
      case CellInput.RichText richText -> ExcelCellValue.richText(toExcelRichText(richText));
      case CellInput.Numeric numeric -> ExcelCellValue.number(numeric.number());
      case CellInput.BooleanValue booleanValue -> ExcelCellValue.bool(booleanValue.bool());
      case CellInput.Date date -> ExcelCellValue.date(date.date());
      case CellInput.DateTime dateTime -> ExcelCellValue.dateTime(dateTime.dateTime());
      case CellInput.Formula formula ->
          ExcelCellValue.formula(inlineText(formula.source(), "formula"));
    };
  }

  static ExcelArrayFormulaDefinition toExcelArrayFormulaDefinition(ArrayFormulaInput input) {
    return new ExcelArrayFormulaDefinition(inlineText(input.source(), "array formula"));
  }

  static ExcelCustomXmlImportDefinition toExcelCustomXmlImportDefinition(
      CustomXmlImportInput input) {
    return new ExcelCustomXmlImportDefinition(
        toExcelCustomXmlMappingLocator(input.locator()), inlineText(input.xml(), "custom XML"));
  }

  static ExcelCustomXmlMappingLocator toExcelCustomXmlMappingLocator(
      dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator locator) {
    return new ExcelCustomXmlMappingLocator(locator.mapId(), locator.name());
  }

  static ExcelRichText toExcelRichText(CellInput.RichText richText) {
    return new ExcelRichText(
        richText.runs().stream().map(WorkbookCommandConverter::toExcelRichTextRun).toList());
  }

  static ExcelRichTextRun toExcelRichTextRun(RichTextRunInput run) {
    return new ExcelRichTextRun(
        inlineText(run.source(), "rich-text run"), toExcelCellFont(run.font()));
  }

  static ExcelHyperlink toExcelHyperlink(HyperlinkTarget target) {
    return switch (target) {
      case HyperlinkTarget.Url url -> new ExcelHyperlink.Url(url.target());
      case HyperlinkTarget.Email email -> new ExcelHyperlink.Email(email.email());
      case HyperlinkTarget.File file -> new ExcelHyperlink.File(file.path());
      case HyperlinkTarget.Document document -> new ExcelHyperlink.Document(document.target());
    };
  }

  static ExcelComment toExcelComment(CommentInput comment) {
    return new ExcelComment(
        inlineText(comment.text(), "comment text"),
        comment.author(),
        comment.visible(),
        comment.runs() == null
            ? null
            : new ExcelRichText(
                comment.runs().stream().map(WorkbookCommandConverter::toExcelRichTextRun).toList()),
        comment.anchor() == null
            ? null
            : new ExcelCommentAnchor(
                comment.anchor().firstColumn(),
                comment.anchor().firstRow(),
                comment.anchor().lastColumn(),
                comment.anchor().lastRow()));
  }

  static ExcelPictureDefinition toExcelPictureDefinition(PictureInput picture) {
    return new ExcelPictureDefinition(
        picture.name(),
        new ExcelBinaryData(inlineBinary(picture.image().source(), "picture payload")),
        picture.image().format(),
        toExcelDrawingAnchor(picture.anchor()),
        picture.description() == null
            ? null
            : inlineText(picture.description(), "picture description"));
  }

  static ExcelChartDefinition toExcelChartDefinition(ChartInput chart) {
    return new ExcelChartDefinition(
        chart.name(),
        toExcelDrawingAnchor(chart.anchor()),
        toExcelChartTitle(chart.title()),
        toExcelChartLegend(chart.legend()),
        chart.displayBlanksAs(),
        chart.plotOnlyVisibleCells(),
        chart.plots().stream().map(WorkbookCommandConverter::toExcelChartPlot).toList());
  }

  static ExcelSignatureLineDefinition toExcelSignatureLineDefinition(
      SignatureLineInput signatureLine) {
    return new ExcelSignatureLineDefinition(
        signatureLine.name(),
        toExcelDrawingAnchor(signatureLine.anchor()),
        signatureLine.allowComments(),
        signatureLine.signingInstructions(),
        signatureLine.suggestedSigner(),
        signatureLine.suggestedSigner2(),
        signatureLine.suggestedSignerEmail(),
        signatureLine.caption(),
        signatureLine.invalidStamp(),
        signatureLine.plainSignature() == null ? null : signatureLine.plainSignature().format(),
        signatureLine.plainSignature() == null
            ? null
            : new ExcelBinaryData(
                inlineBinary(
                    signatureLine.plainSignature().source(), "signature-line plain signature")));
  }

  static ExcelShapeDefinition toExcelShapeDefinition(ShapeInput shape) {
    return new ExcelShapeDefinition(
        shape.name(),
        shape.kind(),
        toExcelDrawingAnchor(shape.anchor()),
        shape.presetGeometryToken(),
        shape.text() == null ? null : inlineText(shape.text(), "shape text"));
  }

  static ExcelEmbeddedObjectDefinition toExcelEmbeddedObjectDefinition(
      EmbeddedObjectInput embeddedObject) {
    return new ExcelEmbeddedObjectDefinition(
        embeddedObject.name(),
        embeddedObject.label(),
        embeddedObject.fileName(),
        embeddedObject.command(),
        new ExcelBinaryData(inlineBinary(embeddedObject.payload(), "embedded-object payload")),
        embeddedObject.previewImage().format(),
        new ExcelBinaryData(
            inlineBinary(embeddedObject.previewImage().source(), "embedded-object preview image")),
        toExcelDrawingAnchor(embeddedObject.anchor()));
  }

  static ExcelDrawingAnchor.TwoCell toExcelDrawingAnchor(DrawingAnchorInput anchor) {
    return switch (anchor) {
      case DrawingAnchorInput.TwoCell twoCell ->
          new ExcelDrawingAnchor.TwoCell(
              new ExcelDrawingMarker(
                  twoCell.from().columnIndex(),
                  twoCell.from().rowIndex(),
                  twoCell.from().dx(),
                  twoCell.from().dy()),
              new ExcelDrawingMarker(
                  twoCell.to().columnIndex(),
                  twoCell.to().rowIndex(),
                  twoCell.to().dx(),
                  twoCell.to().dy()),
              twoCell.behavior());
    };
  }

  private static ExcelChartDefinition.Title toExcelChartTitle(ChartInput.Title title) {
    return switch (title) {
      case ChartInput.Title.None _ -> new ExcelChartDefinition.Title.None();
      case ChartInput.Title.Text text ->
          new ExcelChartDefinition.Title.Text(inlineText(text.source(), "chart title"));
      case ChartInput.Title.Formula formula ->
          new ExcelChartDefinition.Title.Formula(formula.formula());
    };
  }

  private static ExcelChartDefinition.Legend toExcelChartLegend(ChartInput.Legend legend) {
    return switch (legend) {
      case ChartInput.Legend.Hidden _ -> new ExcelChartDefinition.Legend.Hidden();
      case ChartInput.Legend.Visible visible ->
          new ExcelChartDefinition.Legend.Visible(visible.position());
    };
  }

  private static ExcelChartDefinition.Series toExcelChartSeries(ChartInput.Series series) {
    return new ExcelChartDefinition.Series(
        toExcelChartTitle(series.title()),
        toExcelChartDataSource(series.categories()),
        toExcelChartDataSource(series.values()),
        series.smooth(),
        series.markerStyle(),
        series.markerSize(),
        series.explosion());
  }

  private static ExcelChartDefinition.DataSource toExcelChartDataSource(
      ChartInput.DataSource source) {
    return switch (source) {
      case ChartInput.DataSource.Reference reference ->
          new ExcelChartDefinition.DataSource.Reference(reference.formula());
      case ChartInput.DataSource.StringLiteral literal ->
          new ExcelChartDefinition.DataSource.StringLiteral(literal.values());
      case ChartInput.DataSource.NumericLiteral literal ->
          new ExcelChartDefinition.DataSource.NumericLiteral(literal.values());
    };
  }

  private static ExcelChartDefinition.Axis toExcelChartAxis(ChartInput.Axis axis) {
    return new ExcelChartDefinition.Axis(
        axis.kind(), axis.position(), axis.crosses(), axis.visible());
  }

  private static ExcelChartDefinition.Plot toExcelChartPlot(ChartInput.Plot plot) {
    return switch (plot) {
      case ChartInput.Area area ->
          new ExcelChartDefinition.Area(
              area.varyColors(),
              area.grouping(),
              area.axes().stream().map(WorkbookCommandConverter::toExcelChartAxis).toList(),
              area.series().stream().map(WorkbookCommandConverter::toExcelChartSeries).toList());
      case ChartInput.Area3D area3D ->
          new ExcelChartDefinition.Area3D(
              area3D.varyColors(),
              area3D.grouping(),
              area3D.gapDepth(),
              area3D.axes().stream().map(WorkbookCommandConverter::toExcelChartAxis).toList(),
              area3D.series().stream().map(WorkbookCommandConverter::toExcelChartSeries).toList());
      case ChartInput.Bar bar ->
          new ExcelChartDefinition.Bar(
              bar.varyColors(),
              bar.barDirection(),
              bar.grouping(),
              bar.gapWidth(),
              bar.overlap(),
              bar.axes().stream().map(WorkbookCommandConverter::toExcelChartAxis).toList(),
              bar.series().stream().map(WorkbookCommandConverter::toExcelChartSeries).toList());
      case ChartInput.Bar3D bar3D ->
          new ExcelChartDefinition.Bar3D(
              bar3D.varyColors(),
              bar3D.barDirection(),
              bar3D.grouping(),
              bar3D.gapDepth(),
              bar3D.gapWidth(),
              bar3D.shape(),
              bar3D.axes().stream().map(WorkbookCommandConverter::toExcelChartAxis).toList(),
              bar3D.series().stream().map(WorkbookCommandConverter::toExcelChartSeries).toList());
      case ChartInput.Doughnut doughnut ->
          new ExcelChartDefinition.Doughnut(
              doughnut.varyColors(),
              doughnut.firstSliceAngle(),
              doughnut.holeSize(),
              doughnut.series().stream()
                  .map(WorkbookCommandConverter::toExcelChartSeries)
                  .toList());
      case ChartInput.Line line ->
          new ExcelChartDefinition.Line(
              line.varyColors(),
              line.grouping(),
              line.axes().stream().map(WorkbookCommandConverter::toExcelChartAxis).toList(),
              line.series().stream().map(WorkbookCommandConverter::toExcelChartSeries).toList());
      case ChartInput.Line3D line3D ->
          new ExcelChartDefinition.Line3D(
              line3D.varyColors(),
              line3D.grouping(),
              line3D.gapDepth(),
              line3D.axes().stream().map(WorkbookCommandConverter::toExcelChartAxis).toList(),
              line3D.series().stream().map(WorkbookCommandConverter::toExcelChartSeries).toList());
      case ChartInput.Pie pie ->
          new ExcelChartDefinition.Pie(
              pie.varyColors(),
              pie.firstSliceAngle(),
              pie.series().stream().map(WorkbookCommandConverter::toExcelChartSeries).toList());
      case ChartInput.Pie3D pie3D ->
          new ExcelChartDefinition.Pie3D(
              pie3D.varyColors(),
              pie3D.series().stream().map(WorkbookCommandConverter::toExcelChartSeries).toList());
      case ChartInput.Radar radar ->
          new ExcelChartDefinition.Radar(
              radar.varyColors(),
              radar.style(),
              radar.axes().stream().map(WorkbookCommandConverter::toExcelChartAxis).toList(),
              radar.series().stream().map(WorkbookCommandConverter::toExcelChartSeries).toList());
      case ChartInput.Scatter scatter ->
          new ExcelChartDefinition.Scatter(
              scatter.varyColors(),
              scatter.style(),
              scatter.axes().stream().map(WorkbookCommandConverter::toExcelChartAxis).toList(),
              scatter.series().stream().map(WorkbookCommandConverter::toExcelChartSeries).toList());
      case ChartInput.Surface surface ->
          new ExcelChartDefinition.Surface(
              surface.varyColors(),
              surface.wireframe(),
              surface.axes().stream().map(WorkbookCommandConverter::toExcelChartAxis).toList(),
              surface.series().stream().map(WorkbookCommandConverter::toExcelChartSeries).toList());
      case ChartInput.Surface3D surface3D ->
          new ExcelChartDefinition.Surface3D(
              surface3D.varyColors(),
              surface3D.wireframe(),
              surface3D.axes().stream().map(WorkbookCommandConverter::toExcelChartAxis).toList(),
              surface3D.series().stream()
                  .map(WorkbookCommandConverter::toExcelChartSeries)
                  .toList());
    };
  }

  static ExcelCellStyle toExcelCellStyle(CellStyleInput style) {
    return new ExcelCellStyle(
        style.numberFormat(),
        toExcelCellAlignment(style.alignment()),
        toExcelCellFont(style.font()),
        toExcelCellFill(style.fill()),
        toExcelBorder(style.border()),
        toExcelCellProtection(style.protection()));
  }

  static ExcelCellAlignment toExcelCellAlignment(CellAlignmentInput alignment) {
    if (alignment == null) {
      return null;
    }
    return new ExcelCellAlignment(
        alignment.wrapText(),
        alignment.horizontalAlignment(),
        alignment.verticalAlignment(),
        alignment.textRotation(),
        alignment.indentation());
  }

  static ExcelCellFont toExcelCellFont(CellFontInput font) {
    if (font == null) {
      return null;
    }
    return new ExcelCellFont(
        font.bold(),
        font.italic(),
        font.fontName(),
        toExcelFontHeight(font.fontHeight()),
        toExcelColor(
            font.fontColor(), font.fontColorTheme(), font.fontColorIndexed(), font.fontColorTint()),
        font.underline(),
        font.strikeout());
  }

  static ExcelCellFill toExcelCellFill(CellFillInput fill) {
    if (fill == null) {
      return null;
    }
    return new ExcelCellFill(
        fill.pattern(),
        toExcelColor(
            fill.foregroundColor(),
            fill.foregroundColorTheme(),
            fill.foregroundColorIndexed(),
            fill.foregroundColorTint()),
        toExcelColor(
            fill.backgroundColor(),
            fill.backgroundColorTheme(),
            fill.backgroundColorIndexed(),
            fill.backgroundColorTint()),
        toExcelGradientFill(fill.gradient()));
  }

  static ExcelCellProtection toExcelCellProtection(CellProtectionInput protection) {
    if (protection == null) {
      return null;
    }
    return new ExcelCellProtection(protection.locked(), protection.hiddenFormula());
  }

  static ExcelFontHeight toExcelFontHeight(FontHeightInput fontHeight) {
    if (fontHeight == null) {
      return null;
    }
    return switch (fontHeight) {
      case FontHeightInput.Points points -> ExcelFontHeight.fromPoints(points.points());
      case FontHeightInput.Twips twips -> new ExcelFontHeight(twips.twips());
    };
  }

  static ExcelBorder toExcelBorder(CellBorderInput border) {
    if (border == null) {
      return null;
    }
    return new ExcelBorder(
        toExcelBorderSide(border.all()),
        toExcelBorderSide(border.top()),
        toExcelBorderSide(border.right()),
        toExcelBorderSide(border.bottom()),
        toExcelBorderSide(border.left()));
  }

  static ExcelBorderSide toExcelBorderSide(CellBorderSideInput side) {
    return side == null
        ? null
        : new ExcelBorderSide(
            side.style(),
            toExcelColor(side.color(), side.colorTheme(), side.colorIndexed(), side.colorTint()));
  }

  static ExcelDataValidationDefinition toExcelDataValidationDefinition(
      DataValidationInput validation) {
    return new ExcelDataValidationDefinition(
        toExcelDataValidationRule(validation.rule()),
        validation.allowBlank(),
        validation.suppressDropDownArrow(),
        toExcelDataValidationPrompt(validation.prompt()),
        toExcelDataValidationErrorAlert(validation.errorAlert()));
  }

  static ExcelDataValidationRule toExcelDataValidationRule(DataValidationRuleInput rule) {
    return switch (rule) {
      case DataValidationRuleInput.ExplicitList explicitList ->
          new ExcelDataValidationRule.ExplicitList(explicitList.values());
      case DataValidationRuleInput.FormulaList formulaList ->
          new ExcelDataValidationRule.FormulaList(formulaList.formula());
      case DataValidationRuleInput.WholeNumber wholeNumber ->
          new ExcelDataValidationRule.WholeNumber(
              wholeNumber.operator(), wholeNumber.formula1(), wholeNumber.formula2());
      case DataValidationRuleInput.DecimalNumber decimalNumber ->
          new ExcelDataValidationRule.DecimalNumber(
              decimalNumber.operator(), decimalNumber.formula1(), decimalNumber.formula2());
      case DataValidationRuleInput.DateRule dateRule ->
          new ExcelDataValidationRule.DateRule(
              dateRule.operator(), dateRule.formula1(), dateRule.formula2());
      case DataValidationRuleInput.TimeRule timeRule ->
          new ExcelDataValidationRule.TimeRule(
              timeRule.operator(), timeRule.formula1(), timeRule.formula2());
      case DataValidationRuleInput.TextLength textLength ->
          new ExcelDataValidationRule.TextLength(
              textLength.operator(), textLength.formula1(), textLength.formula2());
      case DataValidationRuleInput.CustomFormula customFormula ->
          new ExcelDataValidationRule.CustomFormula(customFormula.formula());
    };
  }

  static ExcelDataValidationPrompt toExcelDataValidationPrompt(DataValidationPromptInput prompt) {
    return prompt == null
        ? null
        : new ExcelDataValidationPrompt(
            inlineText(prompt.title(), "validation prompt title"),
            inlineText(prompt.text(), "validation prompt text"),
            prompt.showPromptBox());
  }

  static ExcelDataValidationErrorAlert toExcelDataValidationErrorAlert(
      DataValidationErrorAlertInput errorAlert) {
    return errorAlert == null
        ? null
        : new ExcelDataValidationErrorAlert(
            errorAlert.style(),
            inlineText(errorAlert.title(), "validation error title"),
            inlineText(errorAlert.text(), "validation error text"),
            errorAlert.showErrorBox());
  }

  static ExcelConditionalFormattingBlockDefinition toExcelConditionalFormattingBlock(
      ConditionalFormattingBlockInput block) {
    return new ExcelConditionalFormattingBlockDefinition(
        block.ranges(),
        block.rules().stream()
            .map(WorkbookCommandConverter::toExcelConditionalFormattingRule)
            .toList());
  }

  static ExcelConditionalFormattingRule toExcelConditionalFormattingRule(
      ConditionalFormattingRuleInput rule) {
    return switch (rule) {
      case ConditionalFormattingRuleInput.FormulaRule formulaRule ->
          new ExcelConditionalFormattingRule.FormulaRule(
              formulaRule.formula(),
              formulaRule.stopIfTrue(),
              toExcelDifferentialStyle(formulaRule.style()));
      case ConditionalFormattingRuleInput.CellValueRule cellValueRule ->
          new ExcelConditionalFormattingRule.CellValueRule(
              cellValueRule.operator(),
              cellValueRule.formula1(),
              cellValueRule.formula2(),
              cellValueRule.stopIfTrue(),
              toExcelDifferentialStyle(cellValueRule.style()));
      case ConditionalFormattingRuleInput.ColorScaleRule colorScaleRule ->
          new ExcelConditionalFormattingRule.ColorScaleRule(
              colorScaleRule.thresholds().stream()
                  .map(WorkbookCommandConverter::toExcelConditionalFormattingThreshold)
                  .toList(),
              colorScaleRule.colors().stream().map(WorkbookCommandConverter::toExcelColor).toList(),
              colorScaleRule.stopIfTrue());
      case ConditionalFormattingRuleInput.DataBarRule dataBarRule ->
          new ExcelConditionalFormattingRule.DataBarRule(
              toExcelColor(dataBarRule.color()),
              dataBarRule.iconOnly(),
              dataBarRule.widthMin(),
              dataBarRule.widthMax(),
              toExcelConditionalFormattingThreshold(dataBarRule.minThreshold()),
              toExcelConditionalFormattingThreshold(dataBarRule.maxThreshold()),
              dataBarRule.stopIfTrue());
      case ConditionalFormattingRuleInput.IconSetRule iconSetRule ->
          new ExcelConditionalFormattingRule.IconSetRule(
              iconSetRule.iconSet(),
              iconSetRule.iconOnly(),
              iconSetRule.reversed(),
              iconSetRule.thresholds().stream()
                  .map(WorkbookCommandConverter::toExcelConditionalFormattingThreshold)
                  .toList(),
              iconSetRule.stopIfTrue());
      case ConditionalFormattingRuleInput.Top10Rule top10Rule ->
          new ExcelConditionalFormattingRule.Top10Rule(
              top10Rule.rank(),
              top10Rule.percent(),
              top10Rule.bottom(),
              top10Rule.stopIfTrue(),
              toExcelDifferentialStyle(top10Rule.style()));
    };
  }

  static ExcelDifferentialStyle toExcelDifferentialStyle(DifferentialStyleInput style) {
    if (style == null) {
      return null;
    }
    return new ExcelDifferentialStyle(
        style.numberFormat(),
        style.bold(),
        style.italic(),
        toExcelFontHeight(style.fontHeight()),
        style.fontColor(),
        style.underline(),
        style.strikeout(),
        style.fillColor(),
        toExcelDifferentialBorder(style.border()));
  }

  static ExcelDifferentialBorder toExcelDifferentialBorder(DifferentialBorderInput border) {
    if (border == null) {
      return null;
    }
    return new ExcelDifferentialBorder(
        toExcelDifferentialBorderSide(border.all()),
        toExcelDifferentialBorderSide(border.top()),
        toExcelDifferentialBorderSide(border.right()),
        toExcelDifferentialBorderSide(border.bottom()),
        toExcelDifferentialBorderSide(border.left()));
  }

  private static ExcelDifferentialBorderSide toExcelDifferentialBorderSide(
      DifferentialBorderSideInput side) {
    return side == null ? null : new ExcelDifferentialBorderSide(side.style(), side.color());
  }

  static ExcelTableDefinition toExcelTableDefinition(TableInput table) {
    return new ExcelTableDefinition(
        table.name(),
        table.sheetName(),
        table.range(),
        table.showTotalsRow(),
        table.hasAutofilter(),
        toExcelTableStyle(table.style()),
        inlineText(table.comment(), "table comment"),
        table.published(),
        table.insertRow(),
        table.insertRowShift(),
        table.headerRowCellStyle(),
        table.dataCellStyle(),
        table.totalsRowCellStyle(),
        table.columns().stream()
            .map(
                column ->
                    new ExcelTableColumnDefinition(
                        column.columnIndex(),
                        column.uniqueName(),
                        column.totalsRowLabel(),
                        column.totalsRowFunction(),
                        column.calculatedColumnFormula()))
            .toList());
  }

  static ExcelPivotTableDefinition toExcelPivotTableDefinition(PivotTableInput pivotTable) {
    return new ExcelPivotTableDefinition(
        pivotTable.name(),
        pivotTable.sheetName(),
        toExcelPivotTableSource(pivotTable.source()),
        new ExcelPivotTableDefinition.Anchor(pivotTable.anchor().topLeftAddress()),
        pivotTable.rowLabels(),
        pivotTable.columnLabels(),
        pivotTable.reportFilters(),
        pivotTable.dataFields().stream()
            .map(WorkbookCommandConverter::toExcelPivotTableDataField)
            .toList());
  }

  private static ExcelPivotTableDefinition.Source toExcelPivotTableSource(
      PivotTableInput.Source source) {
    return switch (source) {
      case PivotTableInput.Source.Range range ->
          new ExcelPivotTableDefinition.Source.Range(range.sheetName(), range.range());
      case PivotTableInput.Source.NamedRange namedRange ->
          new ExcelPivotTableDefinition.Source.NamedRange(namedRange.name());
      case PivotTableInput.Source.Table table ->
          new ExcelPivotTableDefinition.Source.Table(table.name());
    };
  }

  private static ExcelPivotTableDefinition.DataField toExcelPivotTableDataField(
      PivotTableInput.DataField dataField) {
    return new ExcelPivotTableDefinition.DataField(
        dataField.sourceColumnName(),
        dataField.function(),
        dataField.displayName(),
        dataField.valueFormat());
  }

  static ExcelTableStyle toExcelTableStyle(TableStyleInput style) {
    return switch (style) {
      case TableStyleInput.None _ -> new ExcelTableStyle.None();
      case TableStyleInput.Named named ->
          new ExcelTableStyle.Named(
              named.name(),
              named.showFirstColumn(),
              named.showLastColumn(),
              named.showRowStripes(),
              named.showColumnStripes());
    };
  }

  static ExcelNamedRangeScope toExcelNamedRangeScope(NamedRangeScope scope) {
    return switch (scope) {
      case NamedRangeScope.Workbook _ -> new ExcelNamedRangeScope.WorkbookScope();
      case NamedRangeScope.Sheet sheet -> new ExcelNamedRangeScope.SheetScope(sheet.sheetName());
    };
  }

  static ExcelNamedRangeScope toExcelNamedRangeScope(NamedRangeSelector.ScopedExact selector) {
    return switch (selector) {
      case NamedRangeSelector.WorkbookScope _ -> new ExcelNamedRangeScope.WorkbookScope();
      case NamedRangeSelector.SheetScope sheet ->
          new ExcelNamedRangeScope.SheetScope(sheet.sheetName());
    };
  }

  static String toExcelNamedRangeName(NamedRangeSelector.ScopedExact selector) {
    return switch (selector) {
      case NamedRangeSelector.WorkbookScope workbook -> workbook.name();
      case NamedRangeSelector.SheetScope sheet -> sheet.name();
    };
  }

  static ExcelNamedRangeTarget toExcelNamedRangeTarget(NamedRangeTarget target) {
    return target.formula() != null
        ? new ExcelNamedRangeTarget(target.formula())
        : new ExcelNamedRangeTarget(target.sheetName(), target.range());
  }

  /** Converts a protocol sheet copy-position variant into the engine position type. */
  static ExcelSheetCopyPosition toExcelSheetCopyPosition(SheetCopyPosition position) {
    return switch (position) {
      case SheetCopyPosition.AppendAtEnd _ -> new ExcelSheetCopyPosition.AppendAtEnd();
      case SheetCopyPosition.AtIndex atIndex ->
          new ExcelSheetCopyPosition.AtIndex(atIndex.targetIndex());
    };
  }

  private static ExcelSheetProtectionSettings toExcelSheetProtectionSettings(
      SheetProtectionSettings settings) {
    return new ExcelSheetProtectionSettings(
        settings.autoFilterLocked(),
        settings.deleteColumnsLocked(),
        settings.deleteRowsLocked(),
        settings.formatCellsLocked(),
        settings.formatColumnsLocked(),
        settings.formatRowsLocked(),
        settings.insertColumnsLocked(),
        settings.insertHyperlinksLocked(),
        settings.insertRowsLocked(),
        settings.objectsLocked(),
        settings.pivotTablesLocked(),
        settings.scenariosLocked(),
        settings.selectLockedCellsLocked(),
        settings.selectUnlockedCellsLocked(),
        settings.sortLocked());
  }

  private static ExcelWorkbookProtectionSettings toExcelWorkbookProtectionSettings(
      WorkbookProtectionInput protection) {
    return new ExcelWorkbookProtectionSettings(
        protection.structureLocked(),
        protection.windowsLocked(),
        protection.revisionsLocked(),
        protection.workbookPassword(),
        protection.revisionsPassword());
  }

  private static ExcelSheetPane toExcelSheetPane(PaneInput pane) {
    return switch (pane) {
      case PaneInput.None _ -> new ExcelSheetPane.None();
      case PaneInput.Frozen frozen ->
          new ExcelSheetPane.Frozen(
              frozen.splitColumn(), frozen.splitRow(), frozen.leftmostColumn(), frozen.topRow());
      case PaneInput.Split split ->
          new ExcelSheetPane.Split(
              split.xSplitPosition(),
              split.ySplitPosition(),
              split.leftmostColumn(),
              split.topRow(),
              split.activePane());
    };
  }

  private static ExcelPrintLayout toExcelPrintLayout(PrintLayoutInput printLayout) {
    return new ExcelPrintLayout(
        toExcelPrintArea(printLayout.printArea()),
        printLayout.orientation(),
        toExcelPrintScaling(printLayout.scaling()),
        toExcelPrintTitleRows(printLayout.repeatingRows()),
        toExcelPrintTitleColumns(printLayout.repeatingColumns()),
        new ExcelHeaderFooterText(
            inlineText(printLayout.header().left(), "header left"),
            inlineText(printLayout.header().center(), "header center"),
            inlineText(printLayout.header().right(), "header right")),
        new ExcelHeaderFooterText(
            inlineText(printLayout.footer().left(), "footer left"),
            inlineText(printLayout.footer().center(), "footer center"),
            inlineText(printLayout.footer().right(), "footer right")),
        toExcelPrintSetup(printLayout.setup()));
  }

  private static String inlineText(TextSourceInput source, String fieldName) {
    if (source instanceof TextSourceInput.Inline inline) {
      return inline.text();
    }
    throw new IllegalStateException(fieldName + " must be resolved to INLINE before conversion");
  }

  private static byte[] inlineBinary(BinarySourceInput source, String fieldName) {
    if (source instanceof BinarySourceInput.InlineBase64 inline) {
      return Base64.getDecoder().decode(inline.base64Data());
    }
    throw new IllegalStateException(
        fieldName + " must be resolved to INLINE_BASE64 before conversion");
  }

  private static ExcelColor toExcelColor(String rgb, Integer theme, Integer indexed, Double tint) {
    if (rgb == null && theme == null && indexed == null) {
      return null;
    }
    return new ExcelColor(rgb, theme, indexed, tint);
  }

  private static ExcelGradientFill toExcelGradientFill(CellGradientFillInput gradient) {
    if (gradient == null) {
      return null;
    }
    return new ExcelGradientFill(
        gradient.type(),
        gradient.degree(),
        gradient.left(),
        gradient.right(),
        gradient.top(),
        gradient.bottom(),
        gradient.stops().stream()
            .map(stop -> new ExcelGradientStop(stop.position(), toExcelColor(stop.color())))
            .toList());
  }

  private static ExcelColor toExcelColor(ColorInput color) {
    if (color == null) {
      return null;
    }
    return new ExcelColor(color.rgb(), color.theme(), color.indexed(), color.tint());
  }

  private static ExcelConditionalFormattingThreshold toExcelConditionalFormattingThreshold(
      ConditionalFormattingThresholdInput threshold) {
    return new ExcelConditionalFormattingThreshold(
        threshold.type(), threshold.formula(), threshold.value());
  }

  private static ExcelPrintSetup toExcelPrintSetup(PrintSetupInput setup) {
    return new ExcelPrintSetup(
        new ExcelPrintMargins(
            setup.margins().left(),
            setup.margins().right(),
            setup.margins().top(),
            setup.margins().bottom(),
            setup.margins().header(),
            setup.margins().footer()),
        setup.printGridlines(),
        setup.horizontallyCentered(),
        setup.verticallyCentered(),
        setup.paperSize(),
        setup.draft(),
        setup.blackAndWhite(),
        setup.copies(),
        setup.useFirstPageNumber(),
        setup.firstPageNumber(),
        setup.rowBreaks(),
        setup.columnBreaks());
  }

  private static ExcelSheetPresentation toExcelSheetPresentation(
      SheetPresentationInput presentation) {
    return new ExcelSheetPresentation(
        new ExcelSheetDisplay(
            presentation.display().displayGridlines(),
            presentation.display().displayZeros(),
            presentation.display().displayRowColHeadings(),
            presentation.display().displayFormulas(),
            presentation.display().rightToLeft()),
        toExcelColor(presentation.tabColor()),
        new ExcelSheetOutlineSummary(
            presentation.outlineSummary().rowSumsBelow(),
            presentation.outlineSummary().rowSumsRight()),
        new ExcelSheetDefaults(
            presentation.sheetDefaults().defaultColumnWidth(),
            presentation.sheetDefaults().defaultRowHeightPoints()),
        presentation.ignoredErrors().stream()
            .map(
                ignoredError ->
                    new ExcelIgnoredError(ignoredError.range(), ignoredError.errorTypes()))
            .toList());
  }

  private static ExcelAutofilterFilterColumn toExcelAutofilterFilterColumn(
      AutofilterFilterColumnInput column) {
    return new ExcelAutofilterFilterColumn(
        column.columnId(),
        column.showButton(),
        toExcelAutofilterFilterCriterion(column.criterion()));
  }

  private static ExcelAutofilterFilterCriterion toExcelAutofilterFilterCriterion(
      AutofilterFilterCriterionInput criterion) {
    return switch (criterion) {
      case AutofilterFilterCriterionInput.Values values ->
          new ExcelAutofilterFilterCriterion.Values(values.values(), values.includeBlank());
      case AutofilterFilterCriterionInput.Custom custom ->
          new ExcelAutofilterFilterCriterion.Custom(
              custom.and(),
              custom.conditions().stream()
                  .map(
                      condition ->
                          new ExcelAutofilterFilterCriterion.CustomCondition(
                              condition.operator(), condition.value()))
                  .toList());
      case AutofilterFilterCriterionInput.Dynamic dynamic ->
          new ExcelAutofilterFilterCriterion.Dynamic(
              dynamic.type(), dynamic.value(), dynamic.maxValue());
      case AutofilterFilterCriterionInput.Top10 top10 ->
          new ExcelAutofilterFilterCriterion.Top10(top10.value(), top10.top(), top10.percent());
      case AutofilterFilterCriterionInput.Color color ->
          new ExcelAutofilterFilterCriterion.Color(color.cellColor(), toExcelColor(color.color()));
      case AutofilterFilterCriterionInput.Icon icon ->
          new ExcelAutofilterFilterCriterion.Icon(icon.iconSet(), icon.iconId());
    };
  }

  private static ExcelAutofilterSortState toExcelAutofilterSortState(
      AutofilterSortStateInput sortState) {
    if (sortState == null) {
      return null;
    }
    return new ExcelAutofilterSortState(
        sortState.range(),
        sortState.caseSensitive(),
        sortState.columnSort(),
        sortState.sortMethod(),
        sortState.conditions().stream()
            .map(WorkbookCommandConverter::toExcelAutofilterSortCondition)
            .toList());
  }

  private static ExcelAutofilterSortCondition toExcelAutofilterSortCondition(
      AutofilterSortConditionInput condition) {
    return new ExcelAutofilterSortCondition(
        condition.range(),
        condition.descending(),
        condition.sortBy(),
        toExcelColor(condition.color()),
        condition.iconId());
  }

  private static ExcelPrintLayout.Area toExcelPrintArea(PrintAreaInput printArea) {
    return switch (printArea) {
      case PrintAreaInput.None _ -> new ExcelPrintLayout.Area.None();
      case PrintAreaInput.Range range -> new ExcelPrintLayout.Area.Range(range.range());
    };
  }

  private static ExcelPrintLayout.Scaling toExcelPrintScaling(PrintScalingInput scaling) {
    return switch (scaling) {
      case PrintScalingInput.Automatic _ -> new ExcelPrintLayout.Scaling.Automatic();
      case PrintScalingInput.Fit fit ->
          new ExcelPrintLayout.Scaling.Fit(fit.widthPages(), fit.heightPages());
    };
  }

  private static ExcelPrintLayout.TitleRows toExcelPrintTitleRows(
      PrintTitleRowsInput repeatingRows) {
    return switch (repeatingRows) {
      case PrintTitleRowsInput.None _ -> new ExcelPrintLayout.TitleRows.None();
      case PrintTitleRowsInput.Band band ->
          new ExcelPrintLayout.TitleRows.Band(band.firstRowIndex(), band.lastRowIndex());
    };
  }

  private static ExcelPrintLayout.TitleColumns toExcelPrintTitleColumns(
      PrintTitleColumnsInput repeatingColumns) {
    return switch (repeatingColumns) {
      case PrintTitleColumnsInput.None _ -> new ExcelPrintLayout.TitleColumns.None();
      case PrintTitleColumnsInput.Band band ->
          new ExcelPrintLayout.TitleColumns.Band(band.firstColumnIndex(), band.lastColumnIndex());
    };
  }
}
