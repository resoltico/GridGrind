package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
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
import java.nio.file.InvalidPathException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;
import java.util.Objects;

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
    Objects.requireNonNull(plan, "plan must not be null");
    return plan.steps().stream().anyMatch(SourceBackedPlanResolver::requiresStandardInput);
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

  private static boolean requiresStandardInput(WorkbookStep step) {
    return requiresStandardInput(step.target())
        || (step instanceof MutationStep mutationStep
            && requiresStandardInput(mutationStep.action()));
  }

  private static boolean requiresStandardInput(MutationAction action) {
    return switch (action) {
      case MutationAction.SetCell setCell -> requiresStandardInput(setCell.value());
      case MutationAction.SetRange setRange ->
          setRange.rows().stream()
              .flatMap(List::stream)
              .anyMatch(SourceBackedPlanResolver::requiresStandardInput);
      case MutationAction.SetComment setComment -> requiresStandardInput(setComment.comment());
      case MutationAction.SetPicture setPicture -> requiresStandardInput(setPicture.picture());
      case MutationAction.SetChart setChart -> requiresStandardInput(setChart.chart());
      case MutationAction.SetShape setShape -> requiresStandardInput(setShape.shape());
      case MutationAction.SetEmbeddedObject setEmbeddedObject ->
          requiresStandardInput(setEmbeddedObject.embeddedObject());
      case MutationAction.SetDataValidation setDataValidation ->
          requiresStandardInput(setDataValidation.validation());
      case MutationAction.SetTable setTable -> requiresStandardInput(setTable.table());
      case MutationAction.AppendRow appendRow ->
          appendRow.values().stream().anyMatch(SourceBackedPlanResolver::requiresStandardInput);
      case MutationAction.SetPrintLayout setPrintLayout ->
          requiresStandardInput(setPrintLayout.printLayout());
      default -> false;
    };
  }

  private static boolean requiresStandardInput(CellInput value) {
    if (value instanceof CellInput.Text text) {
      return requiresStandardInput(text.source());
    }
    if (value instanceof CellInput.RichText richText) {
      return richText.runs().stream().anyMatch(SourceBackedPlanResolver::requiresStandardInput);
    }
    return value instanceof CellInput.Formula formula && requiresStandardInput(formula.source());
  }

  private static boolean requiresStandardInput(RichTextRunInput run) {
    return requiresStandardInput(run.source());
  }

  private static boolean requiresStandardInput(CommentInput comment) {
    return requiresStandardInput(comment.text())
        || (comment.runs() != null
            && comment.runs().stream().anyMatch(SourceBackedPlanResolver::requiresStandardInput));
  }

  private static boolean requiresStandardInput(PictureInput picture) {
    return requiresStandardInput(picture.image())
        || picture.description() instanceof TextSourceInput.StandardInput;
  }

  private static boolean requiresStandardInput(PictureDataInput pictureData) {
    return requiresStandardInput(pictureData.source());
  }

  private static boolean requiresStandardInput(ChartInput chart) {
    return requiresStandardInput(chart.title());
  }

  private static boolean requiresStandardInput(ChartInput.Title title) {
    return title instanceof ChartInput.Title.Text text && requiresStandardInput(text.source());
  }

  private static boolean requiresStandardInput(ShapeInput shape) {
    return shape.text() != null && requiresStandardInput(shape.text());
  }

  private static boolean requiresStandardInput(EmbeddedObjectInput embeddedObject) {
    return requiresStandardInput(embeddedObject.payload())
        || requiresStandardInput(embeddedObject.previewImage());
  }

  private static boolean requiresStandardInput(DataValidationInput validation) {
    return (validation.prompt() != null && requiresStandardInput(validation.prompt()))
        || (validation.errorAlert() != null && requiresStandardInput(validation.errorAlert()));
  }

  private static boolean requiresStandardInput(DataValidationPromptInput prompt) {
    return requiresStandardInput(prompt.title()) || requiresStandardInput(prompt.text());
  }

  private static boolean requiresStandardInput(DataValidationErrorAlertInput alert) {
    return requiresStandardInput(alert.title()) || requiresStandardInput(alert.text());
  }

  private static boolean requiresStandardInput(TableInput table) {
    return requiresStandardInput(table.comment());
  }

  private static boolean requiresStandardInput(PrintLayoutInput printLayout) {
    return requiresStandardInput(printLayout.header())
        || requiresStandardInput(printLayout.footer());
  }

  private static boolean requiresStandardInput(HeaderFooterTextInput text) {
    return requiresStandardInput(text.left())
        || requiresStandardInput(text.center())
        || requiresStandardInput(text.right());
  }

  private static boolean requiresStandardInput(TextSourceInput source) {
    return source instanceof TextSourceInput.StandardInput;
  }

  private static boolean requiresStandardInput(BinarySourceInput source) {
    return source instanceof BinarySourceInput.StandardInput;
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

  private static boolean requiresStandardInput(Selector selector) {
    return switch (selector) {
      case TableCellSelector.ByColumnName tableCell -> requiresStandardInput(tableCell.row());
      case TableRowSelector.ByKeyCell byKeyCell -> requiresStandardInput(byKeyCell.expectedValue());
      default -> false;
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
      default -> action;
    };
  }

  /**
   * Preserves authored step and selector object identity when source-backed resolution makes no
   * semantic change, which avoids rebuilding canonical records unnecessarily.
   */
  @SuppressWarnings("PMD.CompareObjectsWithEquals")
  private static boolean sameReference(Object left, Object right) {
    return left == right;
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
    List<RichTextRunInput> resolvedRuns =
        comment.runs() == null ? null : resolveRuns(comment.runs(), bindings);
    return sameReference(resolvedText, comment.text())
            && sameReference(resolvedRuns, comment.runs())
        ? comment
        : new CommentInput(
            resolvedText, comment.author(), comment.visible(), resolvedRuns, comment.anchor());
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
    return switch (chart) {
      case ChartInput.Bar bar -> {
        ChartInput.Title resolvedTitle = resolveChartTitle(bar.title(), bindings);
        yield sameReference(resolvedTitle, bar.title())
            ? bar
            : new ChartInput.Bar(
                bar.name(),
                bar.anchor(),
                resolvedTitle,
                bar.legend(),
                bar.displayBlanksAs(),
                bar.plotOnlyVisibleCells(),
                bar.varyColors(),
                bar.barDirection(),
                bar.series());
      }
      case ChartInput.Line line -> {
        ChartInput.Title resolvedTitle = resolveChartTitle(line.title(), bindings);
        yield sameReference(resolvedTitle, line.title())
            ? line
            : new ChartInput.Line(
                line.name(),
                line.anchor(),
                resolvedTitle,
                line.legend(),
                line.displayBlanksAs(),
                line.plotOnlyVisibleCells(),
                line.varyColors(),
                line.series());
      }
      case ChartInput.Pie pie -> {
        ChartInput.Title resolvedTitle = resolveChartTitle(pie.title(), bindings);
        yield sameReference(resolvedTitle, pie.title())
            ? pie
            : new ChartInput.Pie(
                pie.name(),
                pie.anchor(),
                resolvedTitle,
                pie.legend(),
                pie.displayBlanksAs(),
                pie.plotOnlyVisibleCells(),
                pie.varyColors(),
                pie.firstSliceAngle(),
                pie.series());
      }
    };
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
    Path resolved = resolvePath(path, bindings.workingDirectory(), inputKind);
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
    Path resolved = resolvePath(path, bindings.workingDirectory(), inputKind);
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
    if (!bindings.hasStandardInput()) {
      throw new InputSourceUnavailableException(
          inputKind + " requires STANDARD_INPUT but no standard-input bytes were bound", inputKind);
    }
    return bindings.standardInputBytes();
  }

  private static Path resolvePath(String rawPath, Path workingDirectory, String inputKind)
      throws InputSourceReadException {
    try {
      Path candidate = Path.of(rawPath);
      Path resolved =
          candidate.isAbsolute()
              ? candidate.toAbsolutePath().normalize()
              : workingDirectory.resolve(candidate).normalize();
      if (Files.isDirectory(resolved)) {
        throw new InputSourceReadException(
            inputKind + " path must resolve to a file, not a directory: " + resolved,
            inputKind,
            resolved.toString(),
            null);
      }
      return resolved;
    } catch (InvalidPathException exception) {
      throw new InputSourceReadException(
          "Invalid " + inputKind + " path: " + rawPath, inputKind, rawPath, exception);
    }
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
