package dev.erst.gridgrind.executor;

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
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
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
      case CellMutationAction.SetCell setCell -> requiresStandardInput(setCell.value());
      case CellMutationAction.SetRange setRange ->
          setRange.rows().stream()
              .flatMap(List::stream)
              .anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case CellMutationAction.SetComment setComment -> requiresStandardInput(setComment.comment());
      case DrawingMutationAction.SetPicture setPicture ->
          requiresStandardInput(setPicture.picture());
      case DrawingMutationAction.SetSignatureLine setSignatureLine ->
          requiresStandardInput(setSignatureLine.signatureLine());
      case DrawingMutationAction.SetChart setChart -> requiresStandardInput(setChart.chart());
      case DrawingMutationAction.SetShape setShape -> requiresStandardInput(setShape.shape());
      case DrawingMutationAction.SetEmbeddedObject setEmbeddedObject ->
          requiresStandardInput(setEmbeddedObject.embeddedObject());
      case StructuredMutationAction.SetDataValidation setDataValidation ->
          requiresStandardInput(setDataValidation.validation());
      case StructuredMutationAction.SetTable setTable -> requiresStandardInput(setTable.table());
      case CellMutationAction.AppendRow appendRow ->
          appendRow.values().stream()
              .anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case WorkbookMutationAction.SetPrintLayout setPrintLayout ->
          requiresStandardInput(setPrintLayout.printLayout());
      case StructuredMutationAction.ImportCustomXmlMapping importCustomXmlMapping ->
          requiresStandardInput(importCustomXmlMapping.mapping());
      default -> false;
    };
  }

  static boolean requiresStandardInput(CellInput value) {
    if (value instanceof CellInput.Text text) {
      return requiresStandardInput(text.source());
    }
    if (value instanceof CellInput.RichText richText) {
      return richText.runs().stream()
          .anyMatch(SourceBackedInputRequirements::requiresStandardInput);
    }
    return value instanceof CellInput.Formula formula && requiresStandardInput(formula.source());
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
        || picture.description() instanceof TextSourceInput.StandardInput;
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
    return shape.text() != null && requiresStandardInput(shape.text());
  }

  static boolean requiresStandardInput(EmbeddedObjectInput embeddedObject) {
    return requiresStandardInput(embeddedObject.payload())
        || requiresStandardInput(embeddedObject.previewImage());
  }

  static boolean requiresStandardInput(DataValidationInput validation) {
    return (validation.prompt() != null && requiresStandardInput(validation.prompt()))
        || (validation.errorAlert() != null && requiresStandardInput(validation.errorAlert()));
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
      case TableCellSelector.ByColumnName tableCell -> requiresStandardInput(tableCell.row());
      case TableRowSelector.ByKeyCell byKeyCell -> requiresStandardInput(byKeyCell.expectedValue());
      default -> false;
    };
  }
}
