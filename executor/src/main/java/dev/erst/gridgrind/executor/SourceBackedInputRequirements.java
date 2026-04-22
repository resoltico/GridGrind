package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
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

/** Detects whether any authored request surface depends on bound standard input. */
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
      case MutationAction.SetCell setCell -> requiresStandardInput(setCell.value());
      case MutationAction.SetRange setRange ->
          setRange.rows().stream()
              .flatMap(List::stream)
              .anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case MutationAction.SetComment setComment -> requiresStandardInput(setComment.comment());
      case MutationAction.SetPicture setPicture -> requiresStandardInput(setPicture.picture());
      case MutationAction.SetSignatureLine setSignatureLine ->
          requiresStandardInput(setSignatureLine.signatureLine());
      case MutationAction.SetChart setChart -> requiresStandardInput(setChart.chart());
      case MutationAction.SetShape setShape -> requiresStandardInput(setShape.shape());
      case MutationAction.SetEmbeddedObject setEmbeddedObject ->
          requiresStandardInput(setEmbeddedObject.embeddedObject());
      case MutationAction.SetDataValidation setDataValidation ->
          requiresStandardInput(setDataValidation.validation());
      case MutationAction.SetTable setTable -> requiresStandardInput(setTable.table());
      case MutationAction.AppendRow appendRow ->
          appendRow.values().stream()
              .anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case MutationAction.SetPrintLayout setPrintLayout ->
          requiresStandardInput(setPrintLayout.printLayout());
      case MutationAction.ImportCustomXmlMapping importCustomXmlMapping ->
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
        || (comment.runs() != null
            && comment.runs().stream()
                .anyMatch(SourceBackedInputRequirements::requiresStandardInput));
  }

  static boolean requiresStandardInput(PictureInput picture) {
    return requiresStandardInput(picture.image())
        || picture.description() instanceof TextSourceInput.StandardInput;
  }

  static boolean requiresStandardInput(SignatureLineInput signatureLine) {
    return signatureLine.plainSignature() != null
        && requiresStandardInput(signatureLine.plainSignature());
  }

  static boolean requiresStandardInput(PictureDataInput pictureData) {
    return requiresStandardInput(pictureData.source());
  }

  static boolean requiresStandardInput(ChartInput chart) {
    return requiresStandardInput(chart.title())
        || chart.plots().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
  }

  static boolean requiresStandardInput(ChartInput.Title title) {
    return title instanceof ChartInput.Title.Text text && requiresStandardInput(text.source());
  }

  static boolean requiresStandardInput(ChartInput.Plot plot) {
    return switch (plot) {
      case ChartInput.Area area ->
          area.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartInput.Area3D area3D ->
          area3D.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartInput.Bar bar ->
          bar.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartInput.Bar3D bar3D ->
          bar3D.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartInput.Doughnut doughnut ->
          doughnut.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartInput.Line line ->
          line.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartInput.Line3D line3D ->
          line3D.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartInput.Pie pie ->
          pie.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartInput.Pie3D pie3D ->
          pie3D.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartInput.Radar radar ->
          radar.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartInput.Scatter scatter ->
          scatter.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartInput.Surface surface ->
          surface.series().stream().anyMatch(SourceBackedInputRequirements::requiresStandardInput);
      case ChartInput.Surface3D surface3D ->
          surface3D.series().stream()
              .anyMatch(SourceBackedInputRequirements::requiresStandardInput);
    };
  }

  static boolean requiresStandardInput(ChartInput.Series series) {
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
