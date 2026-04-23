package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.ArrayFormulaInput;
import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateInput;
import dev.erst.gridgrind.contract.dto.CellAlignmentInput;
import dev.erst.gridgrind.contract.dto.CellBorderInput;
import dev.erst.gridgrind.contract.dto.CellBorderSideInput;
import dev.erst.gridgrind.contract.dto.CellFillInput;
import dev.erst.gridgrind.contract.dto.CellFontInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellProtectionInput;
import dev.erst.gridgrind.contract.dto.CellStyleInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderInput;
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
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
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
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.excel.ExcelArrayFormulaDefinition;
import dev.erst.gridgrind.excel.ExcelAutofilterFilterColumn;
import dev.erst.gridgrind.excel.ExcelAutofilterSortState;
import dev.erst.gridgrind.excel.ExcelBorder;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelCellAlignment;
import dev.erst.gridgrind.excel.ExcelCellFill;
import dev.erst.gridgrind.excel.ExcelCellFont;
import dev.erst.gridgrind.excel.ExcelCellProtection;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import dev.erst.gridgrind.excel.ExcelCustomXmlImportDefinition;
import dev.erst.gridgrind.excel.ExcelCustomXmlMappingLocator;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorAlert;
import dev.erst.gridgrind.excel.ExcelDataValidationPrompt;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.excel.ExcelDifferentialBorder;
import dev.erst.gridgrind.excel.ExcelDifferentialStyle;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.ExcelPivotTableDefinition;
import dev.erst.gridgrind.excel.ExcelPrintLayout;
import dev.erst.gridgrind.excel.ExcelRichText;
import dev.erst.gridgrind.excel.ExcelRichTextRun;
import dev.erst.gridgrind.excel.ExcelShapeDefinition;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelSheetPresentation;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelSignatureLineDefinition;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSettings;
import dev.erst.gridgrind.excel.WorkbookCommand;

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
    return WorkbookCommandSelectorSupport.sheetByName(target, action);
  }

  private static SheetSelector.ByNames sheetByNames(Selector target, MutationAction action) {
    return WorkbookCommandSelectorSupport.sheetByNames(target, action);
  }

  private static RangeSelector.ByRange rangeByRange(Selector target, MutationAction action) {
    return WorkbookCommandSelectorSupport.rangeByRange(target, action);
  }

  private static RowBandSelector.Span rowSpan(Selector target, MutationAction action) {
    return WorkbookCommandSelectorSupport.rowSpan(target, action);
  }

  private static RowBandSelector.Insertion rowInsertion(Selector target, MutationAction action) {
    return WorkbookCommandSelectorSupport.rowInsertion(target, action);
  }

  private static ColumnBandSelector.Span columnSpan(Selector target, MutationAction action) {
    return WorkbookCommandSelectorSupport.columnSpan(target, action);
  }

  private static ColumnBandSelector.Insertion columnInsertion(
      Selector target, MutationAction action) {
    return WorkbookCommandSelectorSupport.columnInsertion(target, action);
  }

  private static CellSelector.ByAddress cellByAddress(Selector target, MutationAction action) {
    return WorkbookCommandSelectorSupport.cellByAddress(target, action);
  }

  private static DrawingObjectSelector.ByName drawingObjectByName(
      Selector target, MutationAction action) {
    return WorkbookCommandSelectorSupport.drawingObjectByName(target, action);
  }

  private static TableSelector.ByNameOnSheet tableByNameOnSheet(
      Selector target, MutationAction action) {
    return WorkbookCommandSelectorSupport.tableByNameOnSheet(target, action);
  }

  private static PivotTableSelector.ByNameOnSheet pivotTableByNameOnSheet(
      Selector target, MutationAction action) {
    return WorkbookCommandSelectorSupport.pivotTableByNameOnSheet(target, action);
  }

  private static NamedRangeSelector.ScopedExact namedRangeScopedExact(
      Selector target, MutationAction action) {
    return WorkbookCommandSelectorSupport.namedRangeScopedExact(target, action);
  }

  private static void ensureTableIdentity(Selector target, MutationAction.SetTable action) {
    WorkbookCommandSelectorSupport.ensureTableIdentity(target, action);
  }

  private static void ensurePivotTableIdentity(
      Selector target, MutationAction.SetPivotTable action) {
    WorkbookCommandSelectorSupport.ensurePivotTableIdentity(target, action);
  }

  private static void ensureNamedRangeIdentity(
      Selector target, MutationAction.SetNamedRange action) {
    WorkbookCommandSelectorSupport.ensureNamedRangeIdentity(target, action);
  }

  static ExcelCellValue toExcelCellValue(CellInput value) {
    return WorkbookCommandCellInputConverter.toExcelCellValue(value);
  }

  static ExcelArrayFormulaDefinition toExcelArrayFormulaDefinition(ArrayFormulaInput input) {
    return WorkbookCommandCellInputConverter.toExcelArrayFormulaDefinition(input);
  }

  static ExcelCustomXmlImportDefinition toExcelCustomXmlImportDefinition(
      CustomXmlImportInput input) {
    return WorkbookCommandStructuredInputConverter.toExcelCustomXmlImportDefinition(input);
  }

  static ExcelCustomXmlMappingLocator toExcelCustomXmlMappingLocator(
      dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator locator) {
    return WorkbookCommandStructuredInputConverter.toExcelCustomXmlMappingLocator(locator);
  }

  static ExcelRichText toExcelRichText(CellInput.RichText richText) {
    return WorkbookCommandCellInputConverter.toExcelRichText(richText);
  }

  static ExcelRichTextRun toExcelRichTextRun(RichTextRunInput run) {
    return WorkbookCommandCellInputConverter.toExcelRichTextRun(run);
  }

  static ExcelHyperlink toExcelHyperlink(HyperlinkTarget target) {
    return WorkbookCommandCellInputConverter.toExcelHyperlink(target);
  }

  static ExcelComment toExcelComment(CommentInput comment) {
    return WorkbookCommandCellInputConverter.toExcelComment(comment);
  }

  static ExcelPictureDefinition toExcelPictureDefinition(PictureInput picture) {
    return WorkbookCommandStructuredInputConverter.toExcelPictureDefinition(picture);
  }

  static ExcelChartDefinition toExcelChartDefinition(ChartInput chart) {
    return WorkbookCommandStructuredInputConverter.toExcelChartDefinition(chart);
  }

  static ExcelSignatureLineDefinition toExcelSignatureLineDefinition(
      SignatureLineInput signatureLine) {
    return WorkbookCommandStructuredInputConverter.toExcelSignatureLineDefinition(signatureLine);
  }

  static ExcelShapeDefinition toExcelShapeDefinition(ShapeInput shape) {
    return WorkbookCommandStructuredInputConverter.toExcelShapeDefinition(shape);
  }

  static ExcelEmbeddedObjectDefinition toExcelEmbeddedObjectDefinition(
      EmbeddedObjectInput embeddedObject) {
    return WorkbookCommandStructuredInputConverter.toExcelEmbeddedObjectDefinition(embeddedObject);
  }

  static ExcelDrawingAnchor.TwoCell toExcelDrawingAnchor(DrawingAnchorInput anchor) {
    return WorkbookCommandStructuredInputConverter.toExcelDrawingAnchor(anchor);
  }

  static ExcelCellStyle toExcelCellStyle(CellStyleInput style) {
    return WorkbookCommandCellInputConverter.toExcelCellStyle(style);
  }

  static ExcelCellAlignment toExcelCellAlignment(CellAlignmentInput alignment) {
    return WorkbookCommandCellInputConverter.toExcelCellAlignment(alignment);
  }

  static ExcelCellFont toExcelCellFont(CellFontInput font) {
    return WorkbookCommandCellInputConverter.toExcelCellFont(font);
  }

  static ExcelCellFill toExcelCellFill(CellFillInput fill) {
    return WorkbookCommandCellInputConverter.toExcelCellFill(fill);
  }

  static ExcelCellProtection toExcelCellProtection(CellProtectionInput protection) {
    return WorkbookCommandCellInputConverter.toExcelCellProtection(protection);
  }

  static ExcelFontHeight toExcelFontHeight(FontHeightInput fontHeight) {
    return WorkbookCommandCellInputConverter.toExcelFontHeight(fontHeight);
  }

  static ExcelBorder toExcelBorder(CellBorderInput border) {
    return WorkbookCommandCellInputConverter.toExcelBorder(border);
  }

  static ExcelBorderSide toExcelBorderSide(CellBorderSideInput side) {
    return WorkbookCommandCellInputConverter.toExcelBorderSide(side);
  }

  static ExcelDataValidationDefinition toExcelDataValidationDefinition(
      DataValidationInput validation) {
    return WorkbookCommandStructuredInputConverter.toExcelDataValidationDefinition(validation);
  }

  static ExcelDataValidationRule toExcelDataValidationRule(DataValidationRuleInput rule) {
    return WorkbookCommandStructuredInputConverter.toExcelDataValidationRule(rule);
  }

  static ExcelDataValidationPrompt toExcelDataValidationPrompt(DataValidationPromptInput prompt) {
    return WorkbookCommandStructuredInputConverter.toExcelDataValidationPrompt(prompt);
  }

  static ExcelDataValidationErrorAlert toExcelDataValidationErrorAlert(
      DataValidationErrorAlertInput errorAlert) {
    return WorkbookCommandStructuredInputConverter.toExcelDataValidationErrorAlert(errorAlert);
  }

  static ExcelConditionalFormattingBlockDefinition toExcelConditionalFormattingBlock(
      ConditionalFormattingBlockInput block) {
    return WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingBlock(block);
  }

  static ExcelConditionalFormattingRule toExcelConditionalFormattingRule(
      ConditionalFormattingRuleInput rule) {
    return WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingRule(rule);
  }

  static ExcelDifferentialStyle toExcelDifferentialStyle(DifferentialStyleInput style) {
    return WorkbookCommandStructuredInputConverter.toExcelDifferentialStyle(style);
  }

  static ExcelDifferentialBorder toExcelDifferentialBorder(DifferentialBorderInput border) {
    return WorkbookCommandStructuredInputConverter.toExcelDifferentialBorder(border);
  }

  static ExcelTableDefinition toExcelTableDefinition(TableInput table) {
    return WorkbookCommandStructuredInputConverter.toExcelTableDefinition(table);
  }

  static ExcelPivotTableDefinition toExcelPivotTableDefinition(PivotTableInput pivotTable) {
    return WorkbookCommandStructuredInputConverter.toExcelPivotTableDefinition(pivotTable);
  }

  static ExcelTableStyle toExcelTableStyle(TableStyleInput style) {
    return WorkbookCommandStructuredInputConverter.toExcelTableStyle(style);
  }

  static ExcelNamedRangeScope toExcelNamedRangeScope(NamedRangeScope scope) {
    return WorkbookCommandStructuredInputConverter.toExcelNamedRangeScope(scope);
  }

  static ExcelNamedRangeScope toExcelNamedRangeScope(NamedRangeSelector.ScopedExact selector) {
    return WorkbookCommandStructuredInputConverter.toExcelNamedRangeScope(selector);
  }

  static String toExcelNamedRangeName(NamedRangeSelector.ScopedExact selector) {
    return WorkbookCommandStructuredInputConverter.toExcelNamedRangeName(selector);
  }

  static ExcelNamedRangeTarget toExcelNamedRangeTarget(NamedRangeTarget target) {
    return WorkbookCommandStructuredInputConverter.toExcelNamedRangeTarget(target);
  }

  /** Converts a protocol sheet copy-position variant into the engine position type. */
  static ExcelSheetCopyPosition toExcelSheetCopyPosition(SheetCopyPosition position) {
    return WorkbookCommandStructuredInputConverter.toExcelSheetCopyPosition(position);
  }

  private static ExcelSheetProtectionSettings toExcelSheetProtectionSettings(
      SheetProtectionSettings settings) {
    return WorkbookCommandStructuredInputConverter.toExcelSheetProtectionSettings(settings);
  }

  private static ExcelWorkbookProtectionSettings toExcelWorkbookProtectionSettings(
      WorkbookProtectionInput protection) {
    return WorkbookCommandStructuredInputConverter.toExcelWorkbookProtectionSettings(protection);
  }

  private static ExcelSheetPane toExcelSheetPane(PaneInput pane) {
    return WorkbookCommandStructuredInputConverter.toExcelSheetPane(pane);
  }

  private static ExcelPrintLayout toExcelPrintLayout(PrintLayoutInput printLayout) {
    return WorkbookCommandStructuredInputConverter.toExcelPrintLayout(printLayout);
  }

  private static ExcelSheetPresentation toExcelSheetPresentation(
      SheetPresentationInput presentation) {
    return WorkbookCommandStructuredInputConverter.toExcelSheetPresentation(presentation);
  }

  private static ExcelAutofilterFilterColumn toExcelAutofilterFilterColumn(
      AutofilterFilterColumnInput column) {
    return WorkbookCommandStructuredInputConverter.toExcelAutofilterFilterColumn(column);
  }

  private static ExcelAutofilterSortState toExcelAutofilterSortState(
      AutofilterSortStateInput sortState) {
    return WorkbookCommandStructuredInputConverter.toExcelAutofilterSortState(sortState);
  }
}
