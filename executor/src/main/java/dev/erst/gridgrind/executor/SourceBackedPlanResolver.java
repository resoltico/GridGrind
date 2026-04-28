package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.SourceBackedResolutionIdentitySupport.sameOptionalReference;
import static dev.erst.gridgrind.executor.SourceBackedResolutionIdentitySupport.sameReference;

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
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * Resolves source-backed authored text and binary fields into inline canonical contract values.
 *
 * <p>This resolver intentionally spans the authored payload families that can carry file-backed or
 * stdin-backed content.
 */
@SuppressWarnings("PMD.ExcessiveImports")
public final class SourceBackedPlanResolver {
  private SourceBackedPlanResolver() {}

  /** Returns true when any authored mutation value requires stdin bytes for resolution. */
  public static boolean requiresStandardInput(WorkbookPlan plan) {
    return SourceBackedInputRequirements.requiresStandardInput(plan);
  }

  static WorkbookPlan resolve(WorkbookPlan plan, ExecutionInputBindings bindings)
      throws IOException {
    Objects.requireNonNull(plan, "plan must not be null");
    Objects.requireNonNull(bindings, "bindings must not be null");
    List<WorkbookStep> resolvedSteps = new ArrayList<>(plan.steps().size());
    for (WorkbookStep step : plan.steps()) {
      resolvedSteps.add(resolveStep(step, bindings));
    }
    return new WorkbookPlan(
        plan.protocolVersion(),
        plan.planId(),
        plan.source(),
        plan.persistence(),
        plan.execution(),
        plan.formulaEnvironment(),
        resolvedSteps);
  }

  private static WorkbookStep resolveStep(WorkbookStep step, ExecutionInputBindings bindings)
      throws IOException {
    return switch (step) {
      case MutationStep mutationStep -> {
        Selector resolvedTarget = resolveSelector(mutationStep.target(), bindings);
        MutationAction resolvedAction = resolveAction(mutationStep.action(), bindings);
        yield sameReference(resolvedTarget, mutationStep.target())
                && sameReference(resolvedAction, mutationStep.action())
            ? mutationStep
            : new MutationStep(mutationStep.stepId(), resolvedTarget, resolvedAction);
      }
      case AssertionStep assertionStep -> {
        Selector resolvedTarget = resolveSelector(assertionStep.target(), bindings);
        yield sameReference(resolvedTarget, assertionStep.target())
            ? assertionStep
            : new AssertionStep(assertionStep.stepId(), resolvedTarget, assertionStep.assertion());
      }
      case InspectionStep inspectionStep -> {
        Selector resolvedTarget = resolveSelector(inspectionStep.target(), bindings);
        yield sameReference(resolvedTarget, inspectionStep.target())
            ? inspectionStep
            : new InspectionStep(inspectionStep.stepId(), resolvedTarget, inspectionStep.query());
      }
    };
  }

  private static Selector resolveSelector(Selector selector, ExecutionInputBindings bindings)
      throws IOException {
    return switch (selector) {
      case TableCellSelector.ByColumnName tableCell -> {
        TableRowSelector resolvedRow =
            (TableRowSelector) resolveSelector(tableCell.row(), bindings);
        yield sameReference(resolvedRow, tableCell.row())
            ? tableCell
            : new TableCellSelector.ByColumnName(resolvedRow, tableCell.columnName());
      }
      case TableRowSelector.ByKeyCell byKeyCell -> {
        CellInput resolvedValue = resolveCellInput(byKeyCell.expectedValue(), bindings);
        yield sameReference(resolvedValue, byKeyCell.expectedValue())
            ? byKeyCell
            : new TableRowSelector.ByKeyCell(
                byKeyCell.table(), byKeyCell.columnName(), resolvedValue);
      }
      default -> selector;
    };
  }

  private static MutationAction resolveAction(
      MutationAction action, ExecutionInputBindings bindings) throws IOException {
    return switch (action) {
      case MutationAction.SetCell setCell -> {
        CellInput resolvedValue = resolveCellInput(setCell.value(), bindings);
        yield sameReference(resolvedValue, setCell.value())
            ? setCell
            : new MutationAction.SetCell(resolvedValue);
      }
      case MutationAction.SetRange setRange -> {
        List<List<CellInput>> resolvedRows = resolveRows(setRange.rows(), bindings);
        yield sameReference(resolvedRows, setRange.rows())
            ? setRange
            : new MutationAction.SetRange(resolvedRows);
      }
      case MutationAction.SetComment setComment -> {
        CommentInput resolvedComment = resolveComment(setComment.comment(), bindings);
        yield sameReference(resolvedComment, setComment.comment())
            ? setComment
            : new MutationAction.SetComment(resolvedComment);
      }
      case MutationAction.SetPicture setPicture -> {
        PictureInput resolvedPicture = resolvePicture(setPicture.picture(), bindings);
        yield sameReference(resolvedPicture, setPicture.picture())
            ? setPicture
            : new MutationAction.SetPicture(resolvedPicture);
      }
      case MutationAction.SetSignatureLine setSignatureLine -> {
        SignatureLineInput resolvedSignatureLine =
            resolveSignatureLine(setSignatureLine.signatureLine(), bindings);
        yield sameReference(resolvedSignatureLine, setSignatureLine.signatureLine())
            ? setSignatureLine
            : new MutationAction.SetSignatureLine(resolvedSignatureLine);
      }
      case MutationAction.SetChart setChart -> {
        ChartInput resolvedChart = resolveChart(setChart.chart(), bindings);
        yield sameReference(resolvedChart, setChart.chart())
            ? setChart
            : new MutationAction.SetChart(resolvedChart);
      }
      case MutationAction.SetShape setShape -> {
        ShapeInput resolvedShape = resolveShape(setShape.shape(), bindings);
        yield sameReference(resolvedShape, setShape.shape())
            ? setShape
            : new MutationAction.SetShape(resolvedShape);
      }
      case MutationAction.SetEmbeddedObject setEmbeddedObject -> {
        EmbeddedObjectInput resolvedEmbeddedObject =
            resolveEmbeddedObject(setEmbeddedObject.embeddedObject(), bindings);
        yield sameReference(resolvedEmbeddedObject, setEmbeddedObject.embeddedObject())
            ? setEmbeddedObject
            : new MutationAction.SetEmbeddedObject(resolvedEmbeddedObject);
      }
      case MutationAction.SetDataValidation setDataValidation -> {
        DataValidationInput resolvedValidation =
            resolveDataValidation(setDataValidation.validation(), bindings);
        yield sameReference(resolvedValidation, setDataValidation.validation())
            ? setDataValidation
            : new MutationAction.SetDataValidation(resolvedValidation);
      }
      case MutationAction.SetTable setTable -> {
        TableInput resolvedTable = resolveTable(setTable.table(), bindings);
        yield sameReference(resolvedTable, setTable.table())
            ? setTable
            : new MutationAction.SetTable(resolvedTable);
      }
      case MutationAction.AppendRow appendRow -> {
        List<CellInput> resolvedValues = resolveCells(appendRow.values(), bindings);
        yield sameReference(resolvedValues, appendRow.values())
            ? appendRow
            : new MutationAction.AppendRow(resolvedValues);
      }
      case MutationAction.SetPrintLayout setPrintLayout -> {
        PrintLayoutInput resolvedPrintLayout =
            resolvePrintLayout(setPrintLayout.printLayout(), bindings);
        yield sameReference(resolvedPrintLayout, setPrintLayout.printLayout())
            ? setPrintLayout
            : new MutationAction.SetPrintLayout(resolvedPrintLayout);
      }
      case MutationAction.ImportCustomXmlMapping importCustomXmlMapping -> {
        CustomXmlImportInput resolvedImport =
            resolveCustomXmlImport(importCustomXmlMapping.mapping(), bindings);
        yield sameReference(resolvedImport, importCustomXmlMapping.mapping())
            ? importCustomXmlMapping
            : new MutationAction.ImportCustomXmlMapping(resolvedImport);
      }
      default -> action;
    };
  }

  private static CellInput resolveCellInput(CellInput value, ExecutionInputBindings bindings)
      throws IOException {
    return switch (value) {
      case CellInput.Blank blank -> blank;
      case CellInput.Text text ->
          sameReference(
                  resolveTextSource(text.source(), bindings, true, "cell text"), text.source())
              ? text
              : new CellInput.Text(resolveTextSource(text.source(), bindings, true, "cell text"));
      case CellInput.RichText richText ->
          sameReference(resolveRuns(richText.runs(), bindings), richText.runs())
              ? richText
              : new CellInput.RichText(resolveRuns(richText.runs(), bindings));
      case CellInput.Numeric numeric -> numeric;
      case CellInput.BooleanValue booleanValue -> booleanValue;
      case CellInput.Date date -> date;
      case CellInput.DateTime dateTime -> dateTime;
      case CellInput.Formula formula ->
          sameReference(resolveFormulaSource(formula.source(), bindings), formula.source())
              ? formula
              : new CellInput.Formula(resolveFormulaSource(formula.source(), bindings));
    };
  }

  private static RichTextRunInput resolveRichTextRun(
      RichTextRunInput run, ExecutionInputBindings bindings) throws IOException {
    TextSourceInput resolvedSource =
        resolveTextSource(run.source(), bindings, false, "rich-text run");
    String resolvedText = ((TextSourceInput.Inline) resolvedSource).text();
    if (resolvedText.isEmpty()) {
      throw new IllegalArgumentException("rich-text run must not be empty");
    }
    return sameReference(resolvedSource, run.source())
        ? run
        : new RichTextRunInput(resolvedSource, run.font());
  }

  private static CommentInput resolveComment(CommentInput comment, ExecutionInputBindings bindings)
      throws IOException {
    TextSourceInput resolvedText =
        resolveTextSource(comment.text(), bindings, true, "comment text");
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

  private static CustomXmlImportInput resolveCustomXmlImport(
      CustomXmlImportInput input, ExecutionInputBindings bindings) throws IOException {
    TextSourceInput resolvedXml = resolveTextSource(input.xml(), bindings, true, "custom XML");
    return sameReference(resolvedXml, input.xml())
        ? input
        : new CustomXmlImportInput(input.locator(), resolvedXml);
  }

  private static PictureInput resolvePicture(PictureInput picture, ExecutionInputBindings bindings)
      throws IOException {
    PictureDataInput resolvedImage = resolvePictureData(picture.image(), bindings);
    TextSourceInput resolvedDescription =
        picture.description() == null
            ? null
            : resolveTextSource(picture.description(), bindings, true, "picture description");
    return sameReference(resolvedImage, picture.image())
            && sameReference(resolvedDescription, picture.description())
        ? picture
        : new PictureInput(picture.name(), resolvedImage, picture.anchor(), resolvedDescription);
  }

  private static SignatureLineInput resolveSignatureLine(
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

  private static PictureDataInput resolvePictureData(
      PictureDataInput pictureData, ExecutionInputBindings bindings) throws IOException {
    BinarySourceInput resolvedSource =
        resolveBinarySource(pictureData.source(), bindings, "picture payload");
    return sameReference(resolvedSource, pictureData.source())
        ? pictureData
        : new PictureDataInput(pictureData.format(), resolvedSource);
  }

  private static EmbeddedObjectInput resolveEmbeddedObject(
      EmbeddedObjectInput embeddedObject, ExecutionInputBindings bindings) throws IOException {
    BinarySourceInput resolvedPayload =
        resolveBinarySource(embeddedObject.payload(), bindings, "embedded-object payload");
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

  private static ChartInput resolveChart(ChartInput chart, ExecutionInputBindings bindings)
      throws IOException {
    ChartInput.Title resolvedTitle = resolveChartTitle(chart.title(), bindings);
    List<ChartInput.Plot> resolvedPlots = new ArrayList<>(chart.plots().size());
    boolean changed = !sameReference(resolvedTitle, chart.title());
    for (ChartInput.Plot plot : chart.plots()) {
      ChartInput.Plot resolvedPlot = resolveChartPlot(plot, bindings);
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

  private static ChartInput.Title resolveChartTitle(
      ChartInput.Title title, ExecutionInputBindings bindings) throws IOException {
    return switch (title) {
      case ChartInput.Title.None none -> none;
      case ChartInput.Title.Formula formula -> formula;
      case ChartInput.Title.Text text ->
          sameReference(
                  resolveTextSource(text.source(), bindings, true, "chart title"), text.source())
              ? text
              : new ChartInput.Title.Text(
                  resolveTextSource(text.source(), bindings, true, "chart title"));
    };
  }

  private static ChartInput.Plot resolveChartPlot(
      ChartInput.Plot plot, ExecutionInputBindings bindings) throws IOException {
    return switch (plot) {
      case ChartInput.Area area -> {
        List<ChartInput.Series> resolvedSeries = resolveChartSeries(area.series(), bindings);
        yield resolvedSeries.equals(area.series())
            ? area
            : new ChartInput.Area(area.varyColors(), area.grouping(), area.axes(), resolvedSeries);
      }
      case ChartInput.Area3D area3D -> {
        List<ChartInput.Series> resolvedSeries = resolveChartSeries(area3D.series(), bindings);
        yield resolvedSeries.equals(area3D.series())
            ? area3D
            : new ChartInput.Area3D(
                area3D.varyColors(),
                area3D.grouping(),
                area3D.gapDepth(),
                area3D.axes(),
                resolvedSeries);
      }
      case ChartInput.Bar bar -> {
        List<ChartInput.Series> resolvedSeries = resolveChartSeries(bar.series(), bindings);
        yield resolvedSeries.equals(bar.series())
            ? bar
            : new ChartInput.Bar(
                bar.varyColors(),
                bar.barDirection(),
                bar.grouping(),
                bar.gapWidth(),
                bar.overlap(),
                bar.axes(),
                resolvedSeries);
      }
      case ChartInput.Bar3D bar3D -> {
        List<ChartInput.Series> resolvedSeries = resolveChartSeries(bar3D.series(), bindings);
        yield resolvedSeries.equals(bar3D.series())
            ? bar3D
            : new ChartInput.Bar3D(
                bar3D.varyColors(),
                bar3D.barDirection(),
                bar3D.grouping(),
                bar3D.gapDepth(),
                bar3D.gapWidth(),
                bar3D.shape(),
                bar3D.axes(),
                resolvedSeries);
      }
      case ChartInput.Doughnut doughnut -> {
        List<ChartInput.Series> resolvedSeries = resolveChartSeries(doughnut.series(), bindings);
        yield resolvedSeries.equals(doughnut.series())
            ? doughnut
            : new ChartInput.Doughnut(
                doughnut.varyColors(),
                doughnut.firstSliceAngle(),
                doughnut.holeSize(),
                resolvedSeries);
      }
      case ChartInput.Line line -> {
        List<ChartInput.Series> resolvedSeries = resolveChartSeries(line.series(), bindings);
        yield resolvedSeries.equals(line.series())
            ? line
            : new ChartInput.Line(line.varyColors(), line.grouping(), line.axes(), resolvedSeries);
      }
      case ChartInput.Line3D line3D -> {
        List<ChartInput.Series> resolvedSeries = resolveChartSeries(line3D.series(), bindings);
        yield resolvedSeries.equals(line3D.series())
            ? line3D
            : new ChartInput.Line3D(
                line3D.varyColors(),
                line3D.grouping(),
                line3D.gapDepth(),
                line3D.axes(),
                resolvedSeries);
      }
      case ChartInput.Pie pie -> {
        List<ChartInput.Series> resolvedSeries = resolveChartSeries(pie.series(), bindings);
        yield resolvedSeries.equals(pie.series())
            ? pie
            : new ChartInput.Pie(pie.varyColors(), pie.firstSliceAngle(), resolvedSeries);
      }
      case ChartInput.Pie3D pie3D -> {
        List<ChartInput.Series> resolvedSeries = resolveChartSeries(pie3D.series(), bindings);
        yield resolvedSeries.equals(pie3D.series())
            ? pie3D
            : new ChartInput.Pie3D(pie3D.varyColors(), resolvedSeries);
      }
      case ChartInput.Radar radar -> {
        List<ChartInput.Series> resolvedSeries = resolveChartSeries(radar.series(), bindings);
        yield resolvedSeries.equals(radar.series())
            ? radar
            : new ChartInput.Radar(radar.varyColors(), radar.style(), radar.axes(), resolvedSeries);
      }
      case ChartInput.Scatter scatter -> {
        List<ChartInput.Series> resolvedSeries = resolveChartSeries(scatter.series(), bindings);
        yield resolvedSeries.equals(scatter.series())
            ? scatter
            : new ChartInput.Scatter(
                scatter.varyColors(), scatter.style(), scatter.axes(), resolvedSeries);
      }
      case ChartInput.Surface surface -> {
        List<ChartInput.Series> resolvedSeries = resolveChartSeries(surface.series(), bindings);
        yield resolvedSeries.equals(surface.series())
            ? surface
            : new ChartInput.Surface(
                surface.varyColors(), surface.wireframe(), surface.axes(), resolvedSeries);
      }
      case ChartInput.Surface3D surface3D -> {
        List<ChartInput.Series> resolvedSeries = resolveChartSeries(surface3D.series(), bindings);
        yield resolvedSeries.equals(surface3D.series())
            ? surface3D
            : new ChartInput.Surface3D(
                surface3D.varyColors(), surface3D.wireframe(), surface3D.axes(), resolvedSeries);
      }
    };
  }

  private static List<ChartInput.Series> resolveChartSeries(
      List<ChartInput.Series> series, ExecutionInputBindings bindings) throws IOException {
    List<ChartInput.Series> resolvedSeries = new ArrayList<>(series.size());
    boolean changed = false;
    for (ChartInput.Series value : series) {
      ChartInput.Title resolvedTitle = resolveChartTitle(value.title(), bindings);
      ChartInput.Series resolved = resolvedChartSeries(value, resolvedTitle);
      resolvedSeries.add(resolved);
      changed |= !sameReference(resolved, value);
    }
    return changed ? List.copyOf(resolvedSeries) : series;
  }

  private static ChartInput.Series resolvedChartSeries(
      ChartInput.Series series, ChartInput.Title resolvedTitle) {
    return sameReference(resolvedTitle, series.title())
        ? series
        : new ChartInput.Series(
            resolvedTitle,
            series.categories(),
            series.values(),
            series.smooth(),
            series.markerStyle(),
            series.markerSize(),
            series.explosion());
  }

  private static ShapeInput resolveShape(ShapeInput shape, ExecutionInputBindings bindings)
      throws IOException {
    TextSourceInput resolvedText =
        shape.text() == null ? null : resolveTextSource(shape.text(), bindings, true, "shape text");
    return sameReference(resolvedText, shape.text())
        ? shape
        : new ShapeInput(
            shape.name(), shape.kind(), shape.anchor(), shape.presetGeometryToken(), resolvedText);
  }

  private static DataValidationInput resolveDataValidation(
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
        resolveTextSource(prompt.title(), bindings, true, "validation prompt title");
    TextSourceInput resolvedText =
        resolveTextSource(prompt.text(), bindings, true, "validation prompt text");
    return sameReference(resolvedTitle, prompt.title())
            && sameReference(resolvedText, prompt.text())
        ? prompt
        : new DataValidationPromptInput(resolvedTitle, resolvedText, prompt.showPromptBox());
  }

  private static DataValidationErrorAlertInput resolveErrorAlert(
      DataValidationErrorAlertInput alert, ExecutionInputBindings bindings) throws IOException {
    TextSourceInput resolvedTitle =
        resolveTextSource(alert.title(), bindings, true, "validation error title");
    TextSourceInput resolvedText =
        resolveTextSource(alert.text(), bindings, true, "validation error text");
    return sameReference(resolvedTitle, alert.title()) && sameReference(resolvedText, alert.text())
        ? alert
        : new DataValidationErrorAlertInput(
            alert.style(), resolvedTitle, resolvedText, alert.showErrorBox());
  }

  private static TableInput resolveTable(TableInput table, ExecutionInputBindings bindings)
      throws IOException {
    TextSourceInput resolvedComment =
        resolveTextSource(table.comment(), bindings, false, "table comment");
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

  private static PrintLayoutInput resolvePrintLayout(
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

  private static HeaderFooterTextInput resolveHeaderFooter(
      HeaderFooterTextInput text, ExecutionInputBindings bindings) throws IOException {
    TextSourceInput resolvedLeft =
        resolveTextSource(text.left(), bindings, false, "header/footer left");
    TextSourceInput resolvedCenter =
        resolveTextSource(text.center(), bindings, false, "header/footer center");
    TextSourceInput resolvedRight =
        resolveTextSource(text.right(), bindings, false, "header/footer right");
    return sameReference(resolvedLeft, text.left())
            && sameReference(resolvedCenter, text.center())
            && sameReference(resolvedRight, text.right())
        ? text
        : new HeaderFooterTextInput(resolvedLeft, resolvedCenter, resolvedRight);
  }

  private static String resolveFormulaText(TextSourceInput source, ExecutionInputBindings bindings)
      throws IOException {
    String formula = resolveText(source, bindings, true, "formula");
    return formula.startsWith("=") ? formula.substring(1) : formula;
  }

  private static TextSourceInput resolveTextSource(
      TextSourceInput source,
      ExecutionInputBindings bindings,
      boolean requireNonBlank,
      String inputKind)
      throws IOException {
    String resolvedText = resolveText(source, bindings, requireNonBlank, inputKind);
    return source instanceof TextSourceInput.Inline
        ? source
        : new TextSourceInput.Inline(resolvedText);
  }

  static TextSourceInput resolveFormulaSource(
      TextSourceInput source, ExecutionInputBindings bindings) throws IOException {
    String resolvedText = resolveFormulaText(source, bindings);
    if (source instanceof TextSourceInput.Inline inline) {
      if (inline.text().equals(resolvedText)) {
        return source;
      }
      return new TextSourceInput.Inline(resolvedText);
    }
    return new TextSourceInput.Inline(resolvedText);
  }

  private static BinarySourceInput resolveBinarySource(
      BinarySourceInput source, ExecutionInputBindings bindings, String inputKind)
      throws IOException {
    String resolvedBase64 = resolveBinaryBase64(source, bindings, inputKind);
    return source instanceof BinarySourceInput.InlineBase64 inline
            && inline.base64Data().equals(resolvedBase64)
        ? source
        : new BinarySourceInput.InlineBase64(resolvedBase64);
  }

  private static String resolveText(
      TextSourceInput source,
      ExecutionInputBindings bindings,
      boolean requireNonBlank,
      String inputKind)
      throws IOException {
    Objects.requireNonNull(source, "source must not be null");
    String text =
        switch (source) {
          case TextSourceInput.Inline inline -> inline.text();
          case TextSourceInput.Utf8File file -> readUtf8File(file.path(), bindings, inputKind);
          case TextSourceInput.StandardInput _ -> readStandardInputText(bindings, inputKind);
        };
    if (requireNonBlank && text.isBlank()) {
      throw new IllegalArgumentException(inputKind + " must not be blank");
    }
    return text;
  }

  private static String readUtf8File(String path, ExecutionInputBindings bindings, String inputKind)
      throws IOException {
    Path resolved =
        SourceBackedPathResolver.resolvePath(path, bindings.workingDirectory(), inputKind);
    try {
      return Files.readString(resolved, StandardCharsets.UTF_8);
    } catch (java.nio.file.NoSuchFileException exception) {
      throw new InputSourceNotFoundException(
          inputKind + " file does not exist: " + resolved,
          inputKind,
          resolved.toString(),
          exception);
    } catch (IOException exception) {
      throw new InputSourceReadException(
          "Failed to read " + inputKind + " file: " + resolved,
          inputKind,
          resolved.toString(),
          exception);
    }
  }

  private static String readStandardInputText(ExecutionInputBindings bindings, String inputKind)
      throws IOException {
    byte[] bytes = standardInputBytes(bindings, inputKind);
    return new String(bytes, StandardCharsets.UTF_8);
  }

  private static String resolveBinaryBase64(
      BinarySourceInput source, ExecutionInputBindings bindings, String inputKind)
      throws IOException {
    Objects.requireNonNull(source, "source must not be null");
    byte[] bytes =
        switch (source) {
          case BinarySourceInput.InlineBase64 inline ->
              Base64.getDecoder().decode(inline.base64Data());
          case BinarySourceInput.File file -> readBinaryFile(file.path(), bindings, inputKind);
          case BinarySourceInput.StandardInput _ -> standardInputBytes(bindings, inputKind);
        };
    if (bytes.length == 0) {
      throw new IllegalArgumentException(inputKind + " must not be empty");
    }
    return Base64.getEncoder().encodeToString(bytes);
  }

  private static byte[] readBinaryFile(
      String path, ExecutionInputBindings bindings, String inputKind) throws IOException {
    Path resolved =
        SourceBackedPathResolver.resolvePath(path, bindings.workingDirectory(), inputKind);
    try {
      return Files.readAllBytes(resolved);
    } catch (java.nio.file.NoSuchFileException exception) {
      throw new InputSourceNotFoundException(
          inputKind + " file does not exist: " + resolved,
          inputKind,
          resolved.toString(),
          exception);
    } catch (IOException exception) {
      throw new InputSourceReadException(
          "Failed to read " + inputKind + " file: " + resolved,
          inputKind,
          resolved.toString(),
          exception);
    }
  }

  private static byte[] standardInputBytes(ExecutionInputBindings bindings, String inputKind)
      throws InputSourceUnavailableException {
    return bindings
        .standardInputBytes()
        .orElseThrow(
            () ->
                new InputSourceUnavailableException(
                    inputKind + " requires STANDARD_INPUT but no standard-input bytes were bound",
                    inputKind));
  }

  private static List<List<CellInput>> resolveRows(
      List<List<CellInput>> rows, ExecutionInputBindings bindings) throws IOException {
    List<List<CellInput>> resolvedRows = new ArrayList<>(rows.size());
    boolean changed = false;
    for (List<CellInput> row : rows) {
      List<CellInput> resolvedRow = resolveCells(row, bindings);
      resolvedRows.add(resolvedRow);
      changed |= !sameReference(resolvedRow, row);
    }
    return changed ? List.copyOf(resolvedRows) : rows;
  }

  private static List<CellInput> resolveCells(
      List<CellInput> values, ExecutionInputBindings bindings) throws IOException {
    List<CellInput> resolvedValues = new ArrayList<>(values.size());
    boolean changed = false;
    for (CellInput value : values) {
      CellInput resolvedValue = resolveCellInput(value, bindings);
      resolvedValues.add(resolvedValue);
      changed |= !sameReference(resolvedValue, value);
    }
    return changed ? List.copyOf(resolvedValues) : values;
  }

  private static List<RichTextRunInput> resolveRuns(
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
}
