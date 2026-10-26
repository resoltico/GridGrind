package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartSeriesSpec;
import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.customxml.CustomXmlImportSpec;
import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectSpec;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.drawing.PictureSpec;
import dev.erst.gridgrind.core.model.drawing.ShapeSpec;
import dev.erst.gridgrind.core.model.drawing.SignatureLineSpec;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextSpec;
import dev.erst.gridgrind.core.model.print.PrintLayoutSpec;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.MutationAction;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
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

  static boolean requiresStandardInput(CellValueSpec value) {
    return switch (value) {
      case CellValueSpec.Text text -> requiresStandardInput(text.value());
      case CellValueSpec.RichText richText ->
          richText.runs().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case CellValueSpec.Formula formula -> requiresStandardInput(formula.value());
      case CellValueSpec.RawFormula rawFormula -> requiresStandardInput(rawFormula.value());
      default -> false;
    };
  }

  static boolean requiresStandardInput(RichTextRunSpec run) {
    return requiresStandardInput(run.text());
  }

  static boolean requiresStandardInput(CommentSpec comment) {
    return requiresStandardInput(comment.text())
        || (comment.runs().isPresent()
            && comment.runs().orElseThrow().stream()
                .anyMatch(SourceBackedInputRequirements::requiresStandardInput));
  }

  static boolean requiresStandardInput(PictureSpec picture) {
    return requiresStandardInput(picture.image())
        || picture.description().stream().anyMatch(TextSourceSpec.Stdin.class::isInstance);
  }

  static boolean requiresStandardInput(SignatureLineSpec signatureLine) {
    return signatureLine.plainSignature().isPresent()
        && requiresStandardInput(signatureLine.plainSignature().orElseThrow());
  }

  static boolean requiresStandardInput(PictureDataSpec pictureData) {
    return requiresStandardInput(pictureData.source());
  }

  static boolean requiresStandardInput(ChartSpec chart) {
    return requiresStandardInput(chart.title())
        || chart.plots().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
  }

  static boolean requiresStandardInput(ChartTitleSpec title) {
    return title instanceof ChartTitleSpec.Text text && requiresStandardInput(text.source());
  }

  static boolean requiresStandardInput(ChartPlotSpec plot) {
    return switch (plot) {
      case ChartPlotSpec.Area area ->
          area.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotSpec.Area3D area3D ->
          area3D.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotSpec.Bar bar ->
          bar.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotSpec.Bar3D bar3D ->
          bar3D.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotSpec.Doughnut doughnut ->
          doughnut.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotSpec.Line line ->
          line.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotSpec.Line3D line3D ->
          line3D.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotSpec.Pie pie ->
          pie.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotSpec.Pie3D pie3D ->
          pie3D.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotSpec.Radar radar ->
          radar.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotSpec.Scatter scatter ->
          scatter.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotSpec.Surface surface ->
          surface.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartPlotSpec.Surface3D surface3D ->
          surface3D.series().stream()
              .anyMatch(SourceBackedInputRequirements::requiresStandardInput);
    };
  }

  static boolean requiresStandardInput(ChartSeriesSpec series) {
    return requiresStandardInput(series.title());
  }

  static boolean requiresStandardInput(ShapeSpec shape) {
    return switch (shape) {
      case ShapeSpec.SimpleShape simpleShape ->
          simpleShape.text().isPresent() && requiresStandardInput(simpleShape.text().orElseThrow());
      case ShapeSpec.Connector _ -> false;
    };
  }

  static boolean requiresStandardInput(EmbeddedObjectSpec embeddedObject) {
    return requiresStandardInput(embeddedObject.payload())
        || requiresStandardInput(embeddedObject.previewImage());
  }

  static boolean requiresStandardInput(DataValidationSpec validation) {
    return (validation.prompt().isPresent()
            && requiresStandardInput(validation.prompt().orElseThrow()))
        || (validation.errorAlert().isPresent()
            && requiresStandardInput(validation.errorAlert().orElseThrow()));
  }

  static boolean requiresStandardInput(DataValidationPromptSpec prompt) {
    return requiresStandardInput(prompt.title()) || requiresStandardInput(prompt.text());
  }

  static boolean requiresStandardInput(DataValidationErrorAlertSpec alert) {
    return requiresStandardInput(alert.title()) || requiresStandardInput(alert.text());
  }

  static boolean requiresStandardInput(TableSpec table) {
    return requiresStandardInput(table.comment());
  }

  static boolean requiresStandardInput(PrintLayoutSpec printLayout) {
    return requiresStandardInput(printLayout.header())
        || requiresStandardInput(printLayout.footer());
  }

  static boolean requiresStandardInput(CustomXmlImportSpec input) {
    return requiresStandardInput(input.xml());
  }

  static boolean requiresStandardInput(HeaderFooterTextSpec text) {
    return requiresStandardInput(text.left())
        || requiresStandardInput(text.center())
        || requiresStandardInput(text.right());
  }

  static boolean requiresStandardInput(TextSourceSpec source) {
    return source instanceof TextSourceSpec.Stdin;
  }

  static boolean requiresStandardInput(BinarySourceSpec source) {
    return source instanceof BinarySourceSpec.Stdin;
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
