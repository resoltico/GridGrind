package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartSeriesSpec;
import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.chart.ResolvedChartTitle;
import dev.erst.gridgrind.core.model.customxml.CustomXmlImportSpec;
import dev.erst.gridgrind.core.model.customxml.ResolvedCustomXmlImport;
import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectSpec;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.drawing.PictureSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.model.drawing.ResolvedShape;
import dev.erst.gridgrind.core.model.drawing.ResolvedSignatureLine;
import dev.erst.gridgrind.core.model.drawing.ShapeSpec;
import dev.erst.gridgrind.core.model.drawing.SignatureLineSpec;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaSpec;
import dev.erst.gridgrind.core.model.formula.ResolvedArrayFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedRawFormula;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintLayoutSpec;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import dev.erst.gridgrind.core.model.protection.ResolvedWorkbookProtection;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionSpec;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;
import dev.erst.gridgrind.core.model.style.ResolvedRichTextRun;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.MutationAction;
import dev.erst.gridgrind.core.plan.PreparedMutationAction;
import dev.erst.gridgrind.core.plan.PreparedMutationStep;
import dev.erst.gridgrind.core.plan.PreparedStep;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedPlan;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.util.Base64;
import java.util.Optional;

/** Materializes only source-bearing action operands; source-free actions retain their identity. */
@SuppressWarnings(
    "PMD.ExcessiveImports") // The closed source-bearing action surface spans these core operand
// families.
final class MutationResolution {
  private MutationResolution() {}

  static ResolvedPlan preparePlan(dev.erst.gridgrind.core.plan.WorkbookPlan request) {
    return new ResolvedPlan(
        request, request.steps().stream().map(MutationResolution::prepareStep).toList());
  }

  static PreparedStep prepareStep(dev.erst.gridgrind.core.plan.WorkbookStep step) {
    return switch (step) {
      case dev.erst.gridgrind.core.plan.MutationStep mutation ->
          new PreparedMutationStep(mutation, prepare(mutation.action()));
      case dev.erst.gridgrind.core.plan.InspectionStep inspection -> inspection;
      case dev.erst.gridgrind.core.plan.AssertionStep assertion -> assertion;
    };
  }

  static PreparedMutationAction prepare(MutationAction action) {
    return switch (action) {
      case ResolvedMutationAction shared -> shared;
      case CellMutationAction.SetCell cell ->
          new ResolvedMutationAction.SetCell(MutationResolution.toResolvedCellWrite(cell.value()));
      case CellMutationAction.SetRange range ->
          new ResolvedMutationAction.SetRange(
              range.rows().toCellInputRows().stream()
                  .map(row -> row.stream().map(MutationResolution::toResolvedCellWrite).toList())
                  .toList());
      case CellMutationAction.AppendRow row ->
          new ResolvedMutationAction.AppendRow(
              row.values().toCellInputs().stream()
                  .map(MutationResolution::toResolvedCellWrite)
                  .toList());
      case CellMutationAction.SetArrayFormula formula ->
          new ResolvedMutationAction.SetArrayFormula(
              MutationResolution.toResolvedArrayFormula(formula.formula()));
      case CellMutationAction.SetComment comment ->
          new ResolvedMutationAction.SetComment(
              MutationResolution.toResolvedComment(comment.comment()));
      case DrawingMutationAction.SetChart chart ->
          new ResolvedMutationAction.SetChart(MutationResolution.toResolvedChart(chart.chart()));
      case DrawingMutationAction.SetPicture picture ->
          new ResolvedMutationAction.SetPicture(
              MutationResolution.toResolvedPicture(picture.picture()));
      case DrawingMutationAction.SetEmbeddedObject object ->
          new ResolvedMutationAction.SetEmbeddedObject(
              MutationResolution.toResolvedEmbeddedObject(object.embeddedObject()));
      case DrawingMutationAction.SetShape shape ->
          new ResolvedMutationAction.SetShape(MutationResolution.toResolvedShape(shape.shape()));
      case DrawingMutationAction.SetSignatureLine signature ->
          new ResolvedMutationAction.SetSignatureLine(
              MutationResolution.toResolvedSignatureLine(signature.signatureLine()));
      case StructuredMutationAction.ImportCustomXmlMapping mapping ->
          new ResolvedMutationAction.ImportCustomXmlMapping(
              MutationResolution.toResolvedCustomXmlImport(mapping.mapping()));
      case StructuredMutationAction.SetTable table ->
          new ResolvedMutationAction.SetTable(MutationResolution.toResolvedTable(table.table()));
      case StructuredMutationAction.SetDataValidation validation ->
          new ResolvedMutationAction.SetDataValidation(
              MutationResolution.toResolvedDataValidation(validation.validation()));
      case WorkbookMutationAction.SetPrintLayout layout ->
          new ResolvedMutationAction.SetPrintLayout(
              MutationResolution.toResolvedPrintLayout(layout.printLayout()));
      case WorkbookMutationAction.SetSheetProtection protection -> protection;
      case WorkbookMutationAction.SetWorkbookProtection protection -> protection;
    };
  }

  static ResolvedMutationAction finish(
      PreparedMutationAction action, ExecutionInputBindings bindings) {
    return switch (action) {
      case ResolvedMutationAction resolved -> resolved;
      case WorkbookMutationAction.SetSheetProtection protection ->
          new ResolvedMutationAction.SetSheetProtection(
              protection.protection(), protection.passwordRef().map(bindings::resolveSecret));
      case WorkbookMutationAction.SetWorkbookProtection protection ->
          new ResolvedMutationAction.SetWorkbookProtection(
              toResolvedWorkbookProtection(protection.protection(), bindings));
    };
  }

  static ResolvedCellWrite toResolvedCellWrite(CellValueSpec value) {
    return switch (value) {
      case CellValueSpec.Blank _ -> new CellValue.Blank();
      case CellValueSpec.Text text ->
          new CellValue.Text(MutationResolution.inlineText(text.value(), "cell text"));
      case CellValueSpec.RichText richText -> toResolvedRichText(richText);
      case CellValueSpec.NumberValue numberValue -> new CellValue.NumberValue(numberValue.value());
      case CellValueSpec.BooleanValue booleanValue ->
          new CellValue.BooleanValue(booleanValue.value());
      case CellValueSpec.ErrorValue errorValue -> new CellValue.ErrorValue(errorValue.value());
      case CellValueSpec.Date date -> date;
      case CellValueSpec.DateTime dateTime -> dateTime;
      case CellValueSpec.Formula formula ->
          new ResolvedFormula(MutationResolution.inlineText(formula.value(), "formula"));
      case CellValueSpec.RawFormula rawFormula ->
          new ResolvedRawFormula(MutationResolution.inlineText(rawFormula.value(), "raw formula"));
    };
  }

  static ResolvedArrayFormula toResolvedArrayFormula(ArrayFormulaSpec input) {
    return new ResolvedArrayFormula(MutationResolution.inlineText(input.source(), "array formula"));
  }

  static ResolvedRichText toResolvedRichText(CellValueSpec.RichText richText) {
    return new ResolvedRichText(
        richText.runs().stream().map(MutationResolution::toResolvedRichTextRun).toList());
  }

  static ResolvedRichTextRun toResolvedRichTextRun(RichTextRunSpec run) {
    return new ResolvedRichTextRun(
        MutationResolution.inlineText(run.text(), "rich-text run"), run.font());
  }

  static ResolvedComment toResolvedComment(CommentSpec comment) {
    return new ResolvedComment(
        MutationResolution.inlineText(comment.text(), "comment text"),
        comment.author(),
        comment.visible(),
        commentRuns(comment),
        comment.anchor());
  }

  private static Optional<ResolvedRichText> commentRuns(CommentSpec comment) {
    return comment.runs().isEmpty()
        ? Optional.empty()
        : Optional.of(
            new ResolvedRichText(
                comment.runs().orElseThrow().stream()
                    .map(MutationResolution::toResolvedRichTextRun)
                    .toList()));
  }

  static ResolvedDataValidation toResolvedDataValidation(DataValidationSpec validation) {
    return new ResolvedDataValidation(
        validation.rule(),
        validation.allowBlank(),
        validation.suppressDropDownArrow(),
        validation.prompt().flatMap(MutationResolution::toResolvedValidationPrompt),
        validation.errorAlert().flatMap(MutationResolution::toResolvedValidationErrorAlert));
  }

  static Optional<ResolvedValidationPrompt> toResolvedValidationPrompt(
      DataValidationPromptSpec prompt) {
    return prompt == null
        ? Optional.empty()
        : Optional.of(
            new ResolvedValidationPrompt(
                MutationResolution.inlineText(prompt.title(), "validation prompt title"),
                MutationResolution.inlineText(prompt.text(), "validation prompt text"),
                prompt.showPromptBox()));
  }

  static Optional<ResolvedValidationErrorAlert> toResolvedValidationErrorAlert(
      DataValidationErrorAlertSpec errorAlert) {
    return errorAlert == null
        ? Optional.empty()
        : Optional.of(
            new ResolvedValidationErrorAlert(
                errorAlert.style(),
                MutationResolution.inlineText(errorAlert.title(), "validation error title"),
                MutationResolution.inlineText(errorAlert.text(), "validation error text"),
                errorAlert.showErrorBox()));
  }

  static ResolvedCustomXmlImport toResolvedCustomXmlImport(CustomXmlImportSpec input) {
    return new ResolvedCustomXmlImport(
        input.locator(), MutationResolution.inlineText(input.xml(), "custom XML"));
  }

  static ResolvedPicture toResolvedPicture(PictureSpec picture) {
    return new ResolvedPicture(
        picture.name(),
        new BinaryContent(
            MutationResolution.inlineBinary(picture.image().source(), "picture payload")),
        picture.image().format(),
        twoCellAnchor(picture.anchor()),
        picture
            .description()
            .map(description -> MutationResolution.inlineText(description, "picture description")));
  }

  static ResolvedChart toResolvedChart(ChartSpec chart) {
    return new ResolvedChart(
        chart.name(),
        twoCellAnchor(chart.anchor()),
        chartTitle(chart.title()),
        chart.legend(),
        chart.displayBlanksAs(),
        chart.plotOnlyVisibleCells(),
        chart.plots().stream().map(MutationResolution::chartPlot).toList());
  }

  static ResolvedSignatureLine toResolvedSignatureLine(SignatureLineSpec signatureLine) {
    return new ResolvedSignatureLine(
        signatureLine.name(),
        twoCellAnchor(signatureLine.anchor()),
        signatureLine.allowComments(),
        signatureLine.signingInstructions(),
        signatureLine.suggestedSigner(),
        signatureLine.suggestedSigner2(),
        signatureLine.suggestedSignerEmail(),
        signatureLine.caption(),
        signatureLine.invalidStamp(),
        signatureLine.plainSignature().map(PictureDataSpec::format),
        signatureLine
            .plainSignature()
            .map(
                plainSignature ->
                    new BinaryContent(
                        MutationResolution.inlineBinary(
                            plainSignature.source(), "signature-line plain signature"))));
  }

  static ResolvedShape toResolvedShape(ShapeSpec shape) {
    return switch (shape) {
      case ShapeSpec.SimpleShape simpleShape ->
          new ResolvedShape.SimpleShape(
              simpleShape.name(),
              twoCellAnchor(simpleShape.anchor()),
              simpleShape.presetGeometryToken(),
              simpleShape.text().map(text -> MutationResolution.inlineText(text, "shape text")));
      case ShapeSpec.Connector connector -> connector;
    };
  }

  static ResolvedEmbeddedObject toResolvedEmbeddedObject(EmbeddedObjectSpec embeddedObject) {
    return new ResolvedEmbeddedObject(
        embeddedObject.name(),
        embeddedObject.label(),
        embeddedObject.fileName(),
        embeddedObject.command(),
        new BinaryContent(
            MutationResolution.inlineBinary(embeddedObject.payload(), "embedded-object payload")),
        embeddedObject.previewImage().format(),
        new BinaryContent(
            MutationResolution.inlineBinary(
                embeddedObject.previewImage().source(), "embedded-object preview image")),
        twoCellAnchor(embeddedObject.anchor()));
  }

  static DrawingAnchorSpec.TwoCell twoCellAnchor(DrawingAnchorSpec anchor) {
    return switch (anchor) {
      case DrawingAnchorSpec.TwoCell twoCell -> twoCell;
    };
  }

  private static ResolvedChartTitle chartTitle(ChartTitleSpec title) {
    return switch (title) {
      case ChartTitleSpec.None none -> none;
      case ChartTitleSpec.Text text ->
          new ResolvedChartTitle.Text(MutationResolution.inlineText(text.source(), "chart title"));
      case ChartTitleSpec.Formula formula -> formula;
    };
  }

  private static ResolvedChart.Series chartSeries(ChartSeriesSpec series) {
    return new ResolvedChart.Series(
        chartTitle(series.title()),
        series.categories(),
        series.values(),
        series.smooth(),
        series.markerStyle(),
        series.markerSize(),
        series.explosion());
  }

  private static ResolvedChart.Plot chartPlot(ChartPlotSpec plot) {
    return switch (plot) {
      case ChartPlotSpec.Area area ->
          new ResolvedChart.Area(
              area.varyColors(),
              area.grouping(),
              area.axes(),
              area.series().stream().map(MutationResolution::chartSeries).toList());
      case ChartPlotSpec.Area3D area3D ->
          new ResolvedChart.Area3D(
              area3D.varyColors(),
              area3D.grouping(),
              area3D.gapDepth(),
              area3D.axes(),
              area3D.series().stream().map(MutationResolution::chartSeries).toList());
      case ChartPlotSpec.Bar bar ->
          new ResolvedChart.Bar(
              bar.varyColors(),
              bar.barDirection(),
              bar.grouping(),
              bar.gapWidth(),
              bar.overlap(),
              bar.axes(),
              bar.series().stream().map(MutationResolution::chartSeries).toList());
      case ChartPlotSpec.Bar3D bar3D ->
          new ResolvedChart.Bar3D(
              bar3D.varyColors(),
              bar3D.barDirection(),
              bar3D.grouping(),
              bar3D.gapDepth(),
              bar3D.gapWidth(),
              bar3D.shape(),
              bar3D.axes(),
              bar3D.series().stream().map(MutationResolution::chartSeries).toList());
      case ChartPlotSpec.Doughnut doughnut ->
          new ResolvedChart.Doughnut(
              doughnut.varyColors(),
              doughnut.firstSliceAngle(),
              doughnut.holeSize(),
              doughnut.series().stream().map(MutationResolution::chartSeries).toList());
      case ChartPlotSpec.Line line ->
          new ResolvedChart.Line(
              line.varyColors(),
              line.grouping(),
              line.axes(),
              line.series().stream().map(MutationResolution::chartSeries).toList());
      case ChartPlotSpec.Line3D line3D ->
          new ResolvedChart.Line3D(
              line3D.varyColors(),
              line3D.grouping(),
              line3D.gapDepth(),
              line3D.axes(),
              line3D.series().stream().map(MutationResolution::chartSeries).toList());
      case ChartPlotSpec.Pie pie ->
          new ResolvedChart.Pie(
              pie.varyColors(),
              pie.firstSliceAngle(),
              pie.series().stream().map(MutationResolution::chartSeries).toList());
      case ChartPlotSpec.Pie3D pie3D ->
          new ResolvedChart.Pie3D(
              pie3D.varyColors(),
              pie3D.series().stream().map(MutationResolution::chartSeries).toList());
      case ChartPlotSpec.Radar radar ->
          new ResolvedChart.Radar(
              radar.varyColors(),
              radar.style(),
              radar.axes(),
              radar.series().stream().map(MutationResolution::chartSeries).toList());
      case ChartPlotSpec.Scatter scatter ->
          new ResolvedChart.Scatter(
              scatter.varyColors(),
              scatter.style(),
              scatter.axes(),
              scatter.series().stream().map(MutationResolution::chartSeries).toList());
      case ChartPlotSpec.Surface surface ->
          new ResolvedChart.Surface(
              surface.varyColors(),
              surface.wireframe(),
              surface.axes(),
              surface.series().stream().map(MutationResolution::chartSeries).toList());
      case ChartPlotSpec.Surface3D surface3D ->
          new ResolvedChart.Surface3D(
              surface3D.varyColors(),
              surface3D.wireframe(),
              surface3D.axes(),
              surface3D.series().stream().map(MutationResolution::chartSeries).toList());
    };
  }

  static String inlineText(TextSourceSpec source, String fieldName) {
    if (source instanceof TextSourceSpec.Inline inline) {
      return inline.text();
    }
    throw new IllegalStateException(fieldName + " must be resolved to INLINE before conversion");
  }

  static byte[] inlineBinary(BinarySourceSpec source, String fieldName) {
    if (source instanceof BinarySourceSpec.InlineBase64 inline) {
      return Base64.getDecoder().decode(inline.base64Data());
    }
    throw new IllegalStateException(
        fieldName + " must be resolved to INLINE_BASE64 before conversion");
  }

  static ResolvedTable toResolvedTable(TableSpec table) {
    return new ResolvedTable(
        table.name(),
        table.sheetName(),
        table.range(),
        table.showTotalsRow(),
        table.hasAutofilter(),
        table.style(),
        MutationResolution.inlineText(table.comment(), "table comment"),
        table.published(),
        table.insertRow(),
        table.insertRowShift(),
        table.headerRowCellStyle(),
        table.dataCellStyle(),
        table.totalsRowCellStyle(),
        table.columns());
  }

  static ResolvedWorkbookProtection toResolvedWorkbookProtection(
      WorkbookProtectionSpec protection, ExecutionInputBindings bindings) {
    return new ResolvedWorkbookProtection(
        protection.structureLocked(),
        protection.windowsLocked(),
        protection.revisionsLocked(),
        resolveOptionalSecret(protection.workbookPasswordRef(), bindings),
        resolveOptionalSecret(protection.revisionsPasswordRef(), bindings));
  }

  private static Optional<String> resolveOptionalSecret(
      Optional<dev.erst.gridgrind.core.source.SecretReference> reference,
      ExecutionInputBindings bindings) {
    if (reference.isEmpty()) {
      return Optional.empty();
    }
    if (bindings == null) {
      throw new IllegalStateException(
          "secret-bearing workbook protection requires execution bindings");
    }
    return Optional.of(bindings.resolveSecret(reference.orElseThrow()));
  }

  static ResolvedPrintLayout toResolvedPrintLayout(PrintLayoutSpec printLayout) {
    return new ResolvedPrintLayout(
        printLayout.printArea(),
        printLayout.orientation(),
        printLayout.scaling(),
        printLayout.repeatingRows(),
        printLayout.repeatingColumns(),
        new HeaderFooterTextReport(
            MutationResolution.inlineText(printLayout.header().left(), "header left"),
            MutationResolution.inlineText(printLayout.header().center(), "header center"),
            MutationResolution.inlineText(printLayout.header().right(), "header right")),
        new HeaderFooterTextReport(
            MutationResolution.inlineText(printLayout.footer().left(), "footer left"),
            MutationResolution.inlineText(printLayout.footer().center(), "footer center"),
            MutationResolution.inlineText(printLayout.footer().right(), "footer right")),
        printLayout.setup());
  }
}
