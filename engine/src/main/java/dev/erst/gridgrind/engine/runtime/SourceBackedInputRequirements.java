package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ChartPlotInput;
import dev.erst.gridgrind.contract.dto.ChartSeriesInput;
import dev.erst.gridgrind.contract.dto.ChartTitleInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextInput;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.List;
import java.util.Objects;

/**
 * Detects whether any authored request surface depends on bound standard input.
 *
 * <p>This seam intentionally spans the full source-backed request vocabulary, so import count
 * tracks protocol coverage rather than accidental coupling.
 */
@SuppressWarnings("PMD.ExcessiveImports")
final class SourceBackedInputRequirements {
  private SourceBackedInputRequirements() {}

  static boolean requiresStandardInput(WorkbookPlan plan) {
    Objects.requireNonNull(plan, "plan must not be null");
    return plan.steps().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
  }

  static boolean requiresStandardInput(WorkbookStep step) {
    return requiresStandardInput(step.target())
        || (step instanceof MutationStep mutationStep
            && requiresStandardInput(mutationStep.action()));
  }

  static boolean requiresStandardInput(MutationAction action) {
    return switch (action) {
      case CellMutationAction cellAction ->
          switch (cellAction) {
            case CellMutationAction.SetCell setCell -> requiresStandardInput(setCell.value());
            case CellMutationAction.SetRange setRange ->
                setRange.rows().toCellInputRows().stream()
                    .flatMap(List::stream)
                    .anyMatch(SourceBackedInputRequirements::requiresStandardInput);
            case CellMutationAction.ClearRange _ -> false;
            case CellMutationAction.SetArrayFormula _ -> false;
            case CellMutationAction.ClearArrayFormula _ -> false;
            case CellMutationAction.SetHyperlink _ -> false;
            case CellMutationAction.ClearHyperlink _ -> false;
            case CellMutationAction.SetComment setComment ->
                requiresStandardInput(setComment.comment());
            case CellMutationAction.ClearComment _ -> false;
            case CellMutationAction.ApplyStyle _ -> false;
            case CellMutationAction.AppendRow appendRow ->
                appendRow.values().toCellInputs().stream()
                    .anyMatch(SourceBackedInputRequirements::requiresStandardInput);
          };
      case DrawingMutationAction drawingAction ->
          switch (drawingAction) {
            case DrawingMutationAction.SetPicture setPicture ->
                requiresStandardInput(setPicture.picture());
            case DrawingMutationAction.SetSignatureLine setSignatureLine ->
                requiresStandardInput(setSignatureLine.signatureLine());
            case DrawingMutationAction.SetChart setChart -> requiresStandardInput(setChart.chart());
            case DrawingMutationAction.SetShape setShape -> requiresStandardInput(setShape.shape());
            case DrawingMutationAction.SetEmbeddedObject setEmbeddedObject ->
                requiresStandardInput(setEmbeddedObject.embeddedObject());
            case DrawingMutationAction.SetDrawingObjectAnchor _ -> false;
            case DrawingMutationAction.DeleteDrawingObject _ -> false;
          };
      case StructuredMutationAction structuredAction ->
          switch (structuredAction) {
            case StructuredMutationAction.ImportCustomXmlMapping importCustomXmlMapping ->
                requiresStandardInput(importCustomXmlMapping.mapping());
            case StructuredMutationAction.SetPivotTable _ -> false;
            case StructuredMutationAction.SetDataValidation setDataValidation ->
                requiresStandardInput(setDataValidation.validation());
            case StructuredMutationAction.ClearDataValidations _ -> false;
            case StructuredMutationAction.SetConditionalFormatting _ -> false;
            case StructuredMutationAction.ClearConditionalFormatting _ -> false;
            case StructuredMutationAction.SetAutofilter _ -> false;
            case StructuredMutationAction.ClearAutofilter _ -> false;
            case StructuredMutationAction.SetTable setTable ->
                requiresStandardInput(setTable.table());
            case StructuredMutationAction.DeleteTable _ -> false;
            case StructuredMutationAction.DeletePivotTable _ -> false;
            case StructuredMutationAction.SetNamedRange _ -> false;
            case StructuredMutationAction.DeleteNamedRange _ -> false;
          };
      case WorkbookMutationAction workbookAction ->
          switch (workbookAction) {
            case WorkbookMutationAction.EnsureSheet _ -> false;
            case WorkbookMutationAction.RenameSheet _ -> false;
            case WorkbookMutationAction.DeleteSheet _ -> false;
            case WorkbookMutationAction.MoveSheet _ -> false;
            case WorkbookMutationAction.CopySheet _ -> false;
            case WorkbookMutationAction.SetActiveSheet _ -> false;
            case WorkbookMutationAction.SetSelectedSheets _ -> false;
            case WorkbookMutationAction.SetSheetVisibility _ -> false;
            case WorkbookMutationAction.SetSheetProtection _ -> false;
            case WorkbookMutationAction.ClearSheetProtection _ -> false;
            case WorkbookMutationAction.SetWorkbookProtection _ -> false;
            case WorkbookMutationAction.ClearWorkbookProtection _ -> false;
            case WorkbookMutationAction.MergeCells _ -> false;
            case WorkbookMutationAction.UnmergeCells _ -> false;
            case WorkbookMutationAction.SetColumnWidth _ -> false;
            case WorkbookMutationAction.SetRowHeight _ -> false;
            case WorkbookMutationAction.InsertRows _ -> false;
            case WorkbookMutationAction.DeleteRows _ -> false;
            case WorkbookMutationAction.ShiftRows _ -> false;
            case WorkbookMutationAction.InsertColumns _ -> false;
            case WorkbookMutationAction.DeleteColumns _ -> false;
            case WorkbookMutationAction.ShiftColumns _ -> false;
            case WorkbookMutationAction.SetRowVisibility _ -> false;
            case WorkbookMutationAction.SetColumnVisibility _ -> false;
            case WorkbookMutationAction.GroupRows _ -> false;
            case WorkbookMutationAction.UngroupRows _ -> false;
            case WorkbookMutationAction.GroupColumns _ -> false;
            case WorkbookMutationAction.UngroupColumns _ -> false;
            case WorkbookMutationAction.SetSheetPane _ -> false;
            case WorkbookMutationAction.SetSheetZoom _ -> false;
            case WorkbookMutationAction.SetSheetPresentation _ -> false;
            case WorkbookMutationAction.SetPrintLayout setPrintLayout ->
                requiresStandardInput(setPrintLayout.printLayout());
            case WorkbookMutationAction.ClearPrintLayout _ -> false;
            case WorkbookMutationAction.AutoSizeColumns _ -> false;
          };
    };
  }

  static boolean requiresStandardInput(CellInput value) {
    return switch (value) {
      case CellInput.Text text -> requiresStandardInput(text.source());
      case CellInput.RichText richText ->
          richText.runs().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case CellInput.Formula formula -> requiresStandardInput(formula.source());
      case CellInput.RawFormula rawFormula -> requiresStandardInput(rawFormula.source());
      default -> false;
    };
  }

  static boolean requiresStandardInput(RichTextRunInput run) {
    return requiresStandardInput(run.source());
  }

  static boolean requiresStandardInput(CommentInput comment) {
    return requiresStandardInput(comment.text())
        || (comment.runs().isPresent()
            && comment.runs().orElseThrow().stream()
                .anyMatch(SourceBackedInputRequirements::requiresStandardInput));
  }

  static boolean requiresStandardInput(PictureInput picture) {
    return requiresStandardInput(picture.image())
        || picture.description().stream().anyMatch(TextSourceInput.StandardInput.class::isInstance);
  }

  static boolean requiresStandardInput(SignatureLineInput signatureLine) {
    return signatureLine.plainSignature().isPresent()
        && requiresStandardInput(signatureLine.plainSignature().orElseThrow());
  }

  static boolean requiresStandardInput(PictureDataInput pictureData) {
    return requiresStandardInput(pictureData.source());
  }

  static boolean requiresStandardInput(ChartInput chart) {
    return requiresStandardInput(chart.title())
        || chart.plots().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
  }

  static boolean requiresStandardInput(ChartTitleInput title) {
    return title instanceof ChartTitleInput.Text text && requiresStandardInput(text.source());
  }

  static boolean requiresStandardInput(ChartPlotInput plot) {
    return switch (plot) {
      case ChartPlotInput.Area area ->
          area.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotInput.Area3D area3D ->
          area3D.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotInput.Bar bar ->
          bar.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotInput.Bar3D bar3D ->
          bar3D.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotInput.Doughnut doughnut ->
          doughnut.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotInput.Line line ->
          line.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotInput.Line3D line3D ->
          line3D.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotInput.Pie pie ->
          pie.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotInput.Pie3D pie3D ->
          pie3D.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotInput.Radar radar ->
          radar.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotInput.Scatter scatter ->
          scatter.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotInput.Surface surface ->
          surface.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotInput.Surface3D surface3D ->
          surface3D.series().stream()
              .anyMatch(SourceBackedInputRequirements::requiresStandardInput);
    };
  }

  static boolean requiresStandardInput(ChartSeriesInput series) {
    return requiresStandardInput(series.title());
  }

  static boolean requiresStandardInput(ShapeInput shape) {
    return switch (shape) {
      case ShapeInput.SimpleShape simpleShape ->
          simpleShape.text().isPresent() && requiresStandardInput(simpleShape.text().orElseThrow());
      case ShapeInput.Connector _ -> false;
    };
  }

  static boolean requiresStandardInput(EmbeddedObjectInput embeddedObject) {
    return requiresStandardInput(embeddedObject.payload())
        || requiresStandardInput(embeddedObject.previewImage());
  }

  static boolean requiresStandardInput(DataValidationInput validation) {
    return (validation.prompt().isPresent()
            && requiresStandardInput(validation.prompt().orElseThrow()))
        || (validation.errorAlert().isPresent()
            && requiresStandardInput(validation.errorAlert().orElseThrow()));
  }

  static boolean requiresStandardInput(DataValidationPromptInput prompt) {
    return requiresStandardInput(prompt.title()) || requiresStandardInput(prompt.text());
  }

  static boolean requiresStandardInput(DataValidationErrorAlertInput alert) {
    return requiresStandardInput(alert.title()) || requiresStandardInput(alert.text());
  }

  static boolean requiresStandardInput(TableInput table) {
    return requiresStandardInput(table.comment());
  }

  static boolean requiresStandardInput(PrintLayoutInput printLayout) {
    return requiresStandardInput(printLayout.header())
        || requiresStandardInput(printLayout.footer());
  }

  static boolean requiresStandardInput(CustomXmlImportInput input) {
    return requiresStandardInput(input.xml());
  }

  static boolean requiresStandardInput(HeaderFooterTextInput text) {
    return requiresStandardInput(text.left())
        || requiresStandardInput(text.center())
        || requiresStandardInput(text.right());
  }

  static boolean requiresStandardInput(TextSourceInput source) {
    return source instanceof TextSourceInput.StandardInput;
  }

  static boolean requiresStandardInput(BinarySourceInput source) {
    return source instanceof BinarySourceInput.StandardInput;
  }

  static boolean requiresStandardInput(Selector selector) {
    return switch (selector) {
      case TableCellSelector tableCellSelector ->
          switch (tableCellSelector) {
            case TableCellSelector.ByColumnName tableCell -> requiresStandardInput(tableCell.row());
          };
      case TableRowSelector tableRowSelector ->
          switch (tableRowSelector) {
            case TableRowSelector.AllRows _ -> false;
            case TableRowSelector.ByIndex _ -> false;
            case TableRowSelector.ByKeyCell byKeyCell ->
                requiresStandardInput(byKeyCell.expectedValue());
          };
      case WorkbookSelector _ -> false;
      case SheetSelector _ -> false;
      case CellSelector _ -> false;
      case RangeSelector _ -> false;
      case RowBandSelector _ -> false;
      case ColumnBandSelector _ -> false;
      case DrawingObjectSelector _ -> false;
      case ChartSelector _ -> false;
      case TableSelector _ -> false;
      case PivotTableSelector _ -> false;
      case NamedRangeSelector _ -> false;
    };
  }
}
