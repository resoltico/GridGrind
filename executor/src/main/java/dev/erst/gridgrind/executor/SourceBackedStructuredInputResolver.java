package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.SourceBackedResolutionIdentitySupport.sameOptionalReference;
import static dev.erst.gridgrind.executor.SourceBackedResolutionIdentitySupport.sameReference;

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
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/** Resolves authored structured inputs whose nested values can be bound from external sources. */
final class SourceBackedStructuredInputResolver {
  private SourceBackedStructuredInputResolver() {}

  static List<RichTextRunInput> resolveRuns(
      List<RichTextRunInput> runs, ExecutionInputBindings bindings) throws IOException {
    List<RichTextRunInput> resolvedRuns = new ArrayList<>(runs.size());
    boolean changed = false;
    for (RichTextRunInput run : runs) {
      RichTextRunInput resolvedRun = resolveRichTextRun(run, bindings);
      resolvedRuns.add(resolvedRun);
      changed |= !sameReference(resolvedRun, run);
    }
    return changed ? List.copyOf(resolvedRuns) : runs;
  }

  private static RichTextRunInput resolveRichTextRun(
      RichTextRunInput run, ExecutionInputBindings bindings) throws IOException {
    TextSourceInput resolvedSource =
        SourceBackedPlanResolver.resolveTextSource(run.source(), bindings, false, "rich-text run");
    String resolvedText = ((TextSourceInput.Inline) resolvedSource).text();
    if (resolvedText.isEmpty()) {
      throw new IllegalArgumentException("rich-text run must not be empty");
    }
    return sameReference(resolvedSource, run.source())
        ? run
        : new RichTextRunInput(resolvedSource, run.font());
  }

  static CommentInput resolveComment(CommentInput comment, ExecutionInputBindings bindings)
      throws IOException {
    TextSourceInput resolvedText =
        SourceBackedPlanResolver.resolveTextSource(comment.text(), bindings, true, "comment text");
    Optional<List<RichTextRunInput>> resolvedRuns =
        comment.runs().isEmpty()
            ? Optional.empty()
            : Optional.of(resolveRuns(comment.runs().orElseThrow(), bindings));
    return sameReference(resolvedText, comment.text())
            && sameOptionalReference(resolvedRuns, comment.runs())
        ? comment
        : new CommentInput(
            resolvedText, comment.author(), comment.visible(), resolvedRuns, comment.anchor());
  }

  static CustomXmlImportInput resolveCustomXmlImport(
      CustomXmlImportInput input, ExecutionInputBindings bindings) throws IOException {
    TextSourceInput resolvedXml =
        SourceBackedPlanResolver.resolveTextSource(input.xml(), bindings, true, "custom XML");
    return sameReference(resolvedXml, input.xml())
        ? input
        : new CustomXmlImportInput(input.locator(), resolvedXml);
  }

  static SignatureLineInput resolveSignatureLine(
      SignatureLineInput signatureLine, ExecutionInputBindings bindings) throws IOException {
    if (signatureLine.plainSignature().isEmpty()) {
      return signatureLine;
    }
    PictureDataInput resolvedPlainSignature =
        resolvePictureData(signatureLine.plainSignature().orElseThrow(), bindings);
    return sameOptionalReference(
            Optional.of(resolvedPlainSignature), signatureLine.plainSignature())
        ? signatureLine
        : new SignatureLineInput(
            signatureLine.name(),
            signatureLine.anchor(),
            signatureLine.allowComments(),
            signatureLine.signingInstructions(),
            signatureLine.suggestedSigner(),
            signatureLine.suggestedSigner2(),
            signatureLine.suggestedSignerEmail(),
            signatureLine.caption(),
            signatureLine.invalidStamp(),
            Optional.of(resolvedPlainSignature));
  }

  static PictureInput resolvePicture(PictureInput picture, ExecutionInputBindings bindings)
      throws IOException {
    PictureDataInput resolvedImage = resolvePictureData(picture.image(), bindings);
    TextSourceInput resolvedDescription =
        picture.description() == null
            ? null
            : SourceBackedPlanResolver.resolveTextSource(
                picture.description(), bindings, true, "picture description");
    return sameReference(resolvedImage, picture.image())
            && sameReference(resolvedDescription, picture.description())
        ? picture
        : new PictureInput(picture.name(), resolvedImage, picture.anchor(), resolvedDescription);
  }

  private static PictureDataInput resolvePictureData(
      PictureDataInput image, ExecutionInputBindings bindings) throws IOException {
    return sameReference(
            SourceBackedPlanResolver.resolveBinarySource(
                image.source(), bindings, "picture payload"),
            image.source())
        ? image
        : new PictureDataInput(
            image.format(),
            SourceBackedPlanResolver.resolveBinarySource(
                image.source(), bindings, "picture payload"));
  }

  static EmbeddedObjectInput resolveEmbeddedObject(
      EmbeddedObjectInput embeddedObject, ExecutionInputBindings bindings) throws IOException {
    BinarySourceInput resolvedPayload =
        SourceBackedPlanResolver.resolveBinarySource(
            embeddedObject.payload(), bindings, "embedded-object payload");
    PictureDataInput resolvedPreview = resolvePictureData(embeddedObject.previewImage(), bindings);
    return sameReference(resolvedPayload, embeddedObject.payload())
            && sameReference(resolvedPreview, embeddedObject.previewImage())
        ? embeddedObject
        : new EmbeddedObjectInput(
            embeddedObject.name(),
            embeddedObject.label(),
            embeddedObject.fileName(),
            embeddedObject.command(),
            resolvedPayload,
            resolvedPreview,
            embeddedObject.anchor());
  }

  static ChartInput resolveChart(ChartInput chart, ExecutionInputBindings bindings)
      throws IOException {
    ChartTitleInput resolvedTitle = resolveChartTitle(chart.title(), bindings);
    List<ChartPlotInput> resolvedPlots = new ArrayList<>(chart.plots().size());
    boolean changed = !sameReference(resolvedTitle, chart.title());
    for (ChartPlotInput plot : chart.plots()) {
      ChartPlotInput resolvedPlot = resolveChartPlot(plot, bindings);
      resolvedPlots.add(resolvedPlot);
      changed |= !sameReference(resolvedPlot, plot);
    }
    return changed
        ? new ChartInput(
            chart.name(),
            chart.anchor(),
            resolvedTitle,
            chart.legend(),
            chart.displayBlanksAs(),
            chart.plotOnlyVisibleCells(),
            List.copyOf(resolvedPlots))
        : chart;
  }

  private static ChartTitleInput resolveChartTitle(
      ChartTitleInput title, ExecutionInputBindings bindings) throws IOException {
    return switch (title) {
      case ChartTitleInput.None none -> none;
      case ChartTitleInput.Formula formula -> formula;
      case ChartTitleInput.Text text ->
          sameReference(
                  SourceBackedPlanResolver.resolveTextSource(
                      text.source(), bindings, true, "chart title"),
                  text.source())
              ? text
              : new ChartTitleInput.Text(
                  SourceBackedPlanResolver.resolveTextSource(
                      text.source(), bindings, true, "chart title"));
    };
  }

  private static ChartPlotInput resolveChartPlot(
      ChartPlotInput plot, ExecutionInputBindings bindings) throws IOException {
    return switch (plot) {
      case ChartPlotInput.Area area -> {
        List<ChartSeriesInput> resolvedSeries = resolveChartSeries(area.series(), bindings);
        yield sameReference(resolvedSeries, area.series())
            ? area
            : new ChartPlotInput.Area(
                area.varyColors(), area.grouping(), area.axes(), resolvedSeries);
      }
      case ChartPlotInput.Area3D area3D -> {
        List<ChartSeriesInput> resolvedSeries = resolveChartSeries(area3D.series(), bindings);
        yield sameReference(resolvedSeries, area3D.series())
            ? area3D
            : new ChartPlotInput.Area3D(
                area3D.varyColors(),
                area3D.grouping(),
                area3D.gapDepth(),
                area3D.axes(),
                resolvedSeries);
      }
      case ChartPlotInput.Bar bar -> {
        List<ChartSeriesInput> resolvedSeries = resolveChartSeries(bar.series(), bindings);
        yield sameReference(resolvedSeries, bar.series())
            ? bar
            : new ChartPlotInput.Bar(
                bar.varyColors(),
                bar.barDirection(),
                bar.grouping(),
                bar.gapWidth(),
                bar.overlap(),
                bar.axes(),
                resolvedSeries);
      }
      case ChartPlotInput.Bar3D bar3D -> {
        List<ChartSeriesInput> resolvedSeries = resolveChartSeries(bar3D.series(), bindings);
        yield sameReference(resolvedSeries, bar3D.series())
            ? bar3D
            : new ChartPlotInput.Bar3D(
                bar3D.varyColors(),
                bar3D.barDirection(),
                bar3D.grouping(),
                bar3D.gapDepth(),
                bar3D.gapWidth(),
                bar3D.shape(),
                bar3D.axes(),
                resolvedSeries);
      }
      case ChartPlotInput.Doughnut doughnut -> {
        List<ChartSeriesInput> resolvedSeries = resolveChartSeries(doughnut.series(), bindings);
        yield sameReference(resolvedSeries, doughnut.series())
            ? doughnut
            : new ChartPlotInput.Doughnut(
                doughnut.varyColors(),
                doughnut.firstSliceAngle(),
                doughnut.holeSize(),
                resolvedSeries);
      }
      case ChartPlotInput.Line line -> {
        List<ChartSeriesInput> resolvedSeries = resolveChartSeries(line.series(), bindings);
        yield sameReference(resolvedSeries, line.series())
            ? line
            : new ChartPlotInput.Line(
                line.varyColors(), line.grouping(), line.axes(), resolvedSeries);
      }
      case ChartPlotInput.Line3D line3D -> {
        List<ChartSeriesInput> resolvedSeries = resolveChartSeries(line3D.series(), bindings);
        yield sameReference(resolvedSeries, line3D.series())
            ? line3D
            : new ChartPlotInput.Line3D(
                line3D.varyColors(),
                line3D.grouping(),
                line3D.gapDepth(),
                line3D.axes(),
                resolvedSeries);
      }
      case ChartPlotInput.Pie pie -> {
        List<ChartSeriesInput> resolvedSeries = resolveChartSeries(pie.series(), bindings);
        yield sameReference(resolvedSeries, pie.series())
            ? pie
            : new ChartPlotInput.Pie(pie.varyColors(), pie.firstSliceAngle(), resolvedSeries);
      }
      case ChartPlotInput.Pie3D pie3D -> {
        List<ChartSeriesInput> resolvedSeries = resolveChartSeries(pie3D.series(), bindings);
        yield sameReference(resolvedSeries, pie3D.series())
            ? pie3D
            : new ChartPlotInput.Pie3D(pie3D.varyColors(), resolvedSeries);
      }
      case ChartPlotInput.Radar radar -> {
        List<ChartSeriesInput> resolvedSeries = resolveChartSeries(radar.series(), bindings);
        yield sameReference(resolvedSeries, radar.series())
            ? radar
            : new ChartPlotInput.Radar(
                radar.varyColors(), radar.style(), radar.axes(), resolvedSeries);
      }
      case ChartPlotInput.Scatter scatter -> {
        List<ChartSeriesInput> resolvedSeries = resolveChartSeries(scatter.series(), bindings);
        yield sameReference(resolvedSeries, scatter.series())
            ? scatter
            : new ChartPlotInput.Scatter(
                scatter.varyColors(), scatter.style(), scatter.axes(), resolvedSeries);
      }
      case ChartPlotInput.Surface surface -> {
        List<ChartSeriesInput> resolvedSeries = resolveChartSeries(surface.series(), bindings);
        yield sameReference(resolvedSeries, surface.series())
            ? surface
            : new ChartPlotInput.Surface(
                surface.varyColors(), surface.wireframe(), surface.axes(), resolvedSeries);
      }
      case ChartPlotInput.Surface3D surface3D -> {
        List<ChartSeriesInput> resolvedSeries = resolveChartSeries(surface3D.series(), bindings);
        yield sameReference(resolvedSeries, surface3D.series())
            ? surface3D
            : new ChartPlotInput.Surface3D(
                surface3D.varyColors(), surface3D.wireframe(), surface3D.axes(), resolvedSeries);
      }
    };
  }

  private static List<ChartSeriesInput> resolveChartSeries(
      List<ChartSeriesInput> series, ExecutionInputBindings bindings) throws IOException {
    List<ChartSeriesInput> resolvedSeries = new ArrayList<>(series.size());
    boolean changed = false;
    for (ChartSeriesInput value : series) {
      ChartTitleInput resolvedTitle = resolveChartTitle(value.title(), bindings);
      ChartSeriesInput resolved = resolvedChartSeries(value, resolvedTitle);
      resolvedSeries.add(resolved);
      changed |= !sameReference(resolved, value);
    }
    return changed ? List.copyOf(resolvedSeries) : series;
  }

  private static ChartSeriesInput resolvedChartSeries(
      ChartSeriesInput series, ChartTitleInput resolvedTitle) {
    return sameReference(resolvedTitle, series.title())
        ? series
        : new ChartSeriesInput(
            resolvedTitle,
            series.categories(),
            series.values(),
            series.smooth(),
            series.markerStyle(),
            series.markerSize(),
            series.explosion());
  }

  static ShapeInput resolveShape(ShapeInput shape, ExecutionInputBindings bindings)
      throws IOException {
    TextSourceInput resolvedText =
        shape.text() == null
            ? null
            : SourceBackedPlanResolver.resolveTextSource(
                shape.text(), bindings, true, "shape text");
    return sameReference(resolvedText, shape.text())
        ? shape
        : new ShapeInput(
            shape.name(), shape.kind(), shape.anchor(), shape.presetGeometryToken(), resolvedText);
  }

  static DataValidationInput resolveDataValidation(
      DataValidationInput validation, ExecutionInputBindings bindings) throws IOException {
    DataValidationPromptInput resolvedPrompt =
        validation.prompt() == null ? null : resolvePrompt(validation.prompt(), bindings);
    DataValidationErrorAlertInput resolvedAlert =
        validation.errorAlert() == null
            ? null
            : resolveErrorAlert(validation.errorAlert(), bindings);
    return sameReference(resolvedPrompt, validation.prompt())
            && sameReference(resolvedAlert, validation.errorAlert())
        ? validation
        : new DataValidationInput(
            validation.rule(),
            validation.allowBlank(),
            validation.suppressDropDownArrow(),
            resolvedPrompt,
            resolvedAlert);
  }

  private static DataValidationPromptInput resolvePrompt(
      DataValidationPromptInput prompt, ExecutionInputBindings bindings) throws IOException {
    TextSourceInput resolvedTitle =
        SourceBackedPlanResolver.resolveTextSource(
            prompt.title(), bindings, true, "validation prompt title");
    TextSourceInput resolvedText =
        SourceBackedPlanResolver.resolveTextSource(
            prompt.text(), bindings, true, "validation prompt text");
    return sameReference(resolvedTitle, prompt.title())
            && sameReference(resolvedText, prompt.text())
        ? prompt
        : new DataValidationPromptInput(resolvedTitle, resolvedText, prompt.showPromptBox());
  }

  private static DataValidationErrorAlertInput resolveErrorAlert(
      DataValidationErrorAlertInput alert, ExecutionInputBindings bindings) throws IOException {
    TextSourceInput resolvedTitle =
        SourceBackedPlanResolver.resolveTextSource(
            alert.title(), bindings, true, "validation error title");
    TextSourceInput resolvedText =
        SourceBackedPlanResolver.resolveTextSource(
            alert.text(), bindings, true, "validation error text");
    return sameReference(resolvedTitle, alert.title()) && sameReference(resolvedText, alert.text())
        ? alert
        : new DataValidationErrorAlertInput(
            alert.style(), resolvedTitle, resolvedText, alert.showErrorBox());
  }

  static TableInput resolveTable(TableInput table, ExecutionInputBindings bindings)
      throws IOException {
    TextSourceInput resolvedComment =
        SourceBackedPlanResolver.resolveTextSource(
            table.comment(), bindings, false, "table comment");
    return sameReference(resolvedComment, table.comment())
        ? table
        : new TableInput(
            table.name(),
            table.sheetName(),
            table.range(),
            table.showTotalsRow(),
            table.hasAutofilter(),
            table.style(),
            resolvedComment,
            table.published(),
            table.insertRow(),
            table.insertRowShift(),
            table.headerRowCellStyle(),
            table.dataCellStyle(),
            table.totalsRowCellStyle(),
            table.columns());
  }

  static PrintLayoutInput resolvePrintLayout(
      PrintLayoutInput printLayout, ExecutionInputBindings bindings) throws IOException {
    HeaderFooterTextInput resolvedHeader = resolveHeaderFooter(printLayout.header(), bindings);
    HeaderFooterTextInput resolvedFooter = resolveHeaderFooter(printLayout.footer(), bindings);
    return sameReference(resolvedHeader, printLayout.header())
            && sameReference(resolvedFooter, printLayout.footer())
        ? printLayout
        : new PrintLayoutInput(
            printLayout.printArea(),
            printLayout.orientation(),
            printLayout.scaling(),
            printLayout.repeatingRows(),
            printLayout.repeatingColumns(),
            resolvedHeader,
            resolvedFooter,
            printLayout.setup());
  }

  static HeaderFooterTextInput resolveHeaderFooter(
      HeaderFooterTextInput text, ExecutionInputBindings bindings) throws IOException {
    TextSourceInput resolvedLeft =
        SourceBackedPlanResolver.resolveTextSource(
            text.left(), bindings, false, "header/footer left");
    TextSourceInput resolvedCenter =
        SourceBackedPlanResolver.resolveTextSource(
            text.center(), bindings, false, "header/footer center");
    TextSourceInput resolvedRight =
        SourceBackedPlanResolver.resolveTextSource(
            text.right(), bindings, false, "header/footer right");
    return sameReference(resolvedLeft, text.left())
            && sameReference(resolvedCenter, text.center())
            && sameReference(resolvedRight, text.right())
        ? text
        : new HeaderFooterTextInput(resolvedLeft, resolvedCenter, resolvedRight);
  }
}
