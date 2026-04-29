package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.SourceBackedResolutionIdentitySupport.sameReference;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.selector.Selector;
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

/**
 * Resolves source-backed authored text and binary fields into inline canonical contract values.
 *
 * <p>This resolver intentionally spans the authored payload families that can carry file-backed or
 * stdin-backed content.
 */
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
        Selector resolvedTarget =
            SourceBackedSelectorResolver.resolve(mutationStep.target(), bindings);
        MutationAction resolvedAction = resolveAction(mutationStep.action(), bindings);
        yield sameReference(resolvedTarget, mutationStep.target())
                && sameReference(resolvedAction, mutationStep.action())
            ? mutationStep
            : new MutationStep(mutationStep.stepId(), resolvedTarget, resolvedAction);
      }
      case AssertionStep assertionStep -> {
        Selector resolvedTarget =
            SourceBackedSelectorResolver.resolve(assertionStep.target(), bindings);
        yield sameReference(resolvedTarget, assertionStep.target())
            ? assertionStep
            : new AssertionStep(assertionStep.stepId(), resolvedTarget, assertionStep.assertion());
      }
      case InspectionStep inspectionStep -> {
        Selector resolvedTarget =
            SourceBackedSelectorResolver.resolve(inspectionStep.target(), bindings);
        yield sameReference(resolvedTarget, inspectionStep.target())
            ? inspectionStep
            : new InspectionStep(inspectionStep.stepId(), resolvedTarget, inspectionStep.query());
      }
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
        var resolvedComment =
            SourceBackedStructuredInputResolver.resolveComment(setComment.comment(), bindings);
        yield sameReference(resolvedComment, setComment.comment())
            ? setComment
            : new MutationAction.SetComment(resolvedComment);
      }
      case MutationAction.SetPicture setPicture -> {
        var resolvedPicture =
            SourceBackedStructuredInputResolver.resolvePicture(setPicture.picture(), bindings);
        yield sameReference(resolvedPicture, setPicture.picture())
            ? setPicture
            : new MutationAction.SetPicture(resolvedPicture);
      }
      case MutationAction.SetSignatureLine setSignatureLine -> {
        var resolvedSignatureLine =
            SourceBackedStructuredInputResolver.resolveSignatureLine(
                setSignatureLine.signatureLine(), bindings);
        yield sameReference(resolvedSignatureLine, setSignatureLine.signatureLine())
            ? setSignatureLine
            : new MutationAction.SetSignatureLine(resolvedSignatureLine);
      }
      case MutationAction.SetChart setChart -> {
        var resolvedChart =
            SourceBackedStructuredInputResolver.resolveChart(setChart.chart(), bindings);
        yield sameReference(resolvedChart, setChart.chart())
            ? setChart
            : new MutationAction.SetChart(resolvedChart);
      }
      case MutationAction.SetShape setShape -> {
        var resolvedShape =
            SourceBackedStructuredInputResolver.resolveShape(setShape.shape(), bindings);
        yield sameReference(resolvedShape, setShape.shape())
            ? setShape
            : new MutationAction.SetShape(resolvedShape);
      }
      case MutationAction.SetEmbeddedObject setEmbeddedObject -> {
        var resolvedEmbeddedObject =
            SourceBackedStructuredInputResolver.resolveEmbeddedObject(
                setEmbeddedObject.embeddedObject(), bindings);
        yield sameReference(resolvedEmbeddedObject, setEmbeddedObject.embeddedObject())
            ? setEmbeddedObject
            : new MutationAction.SetEmbeddedObject(resolvedEmbeddedObject);
      }
      case MutationAction.SetDataValidation setDataValidation -> {
        var resolvedValidation =
            SourceBackedStructuredInputResolver.resolveDataValidation(
                setDataValidation.validation(), bindings);
        yield sameReference(resolvedValidation, setDataValidation.validation())
            ? setDataValidation
            : new MutationAction.SetDataValidation(resolvedValidation);
      }
      case MutationAction.SetTable setTable -> {
        var resolvedTable =
            SourceBackedStructuredInputResolver.resolveTable(setTable.table(), bindings);
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
        var resolvedPrintLayout =
            SourceBackedStructuredInputResolver.resolvePrintLayout(
                setPrintLayout.printLayout(), bindings);
        yield sameReference(resolvedPrintLayout, setPrintLayout.printLayout())
            ? setPrintLayout
            : new MutationAction.SetPrintLayout(resolvedPrintLayout);
      }
      case MutationAction.ImportCustomXmlMapping importCustomXmlMapping -> {
        var resolvedImport =
            SourceBackedStructuredInputResolver.resolveCustomXmlImport(
                importCustomXmlMapping.mapping(), bindings);
        yield sameReference(resolvedImport, importCustomXmlMapping.mapping())
            ? importCustomXmlMapping
            : new MutationAction.ImportCustomXmlMapping(resolvedImport);
      }
      default -> action;
    };
  }

  static CellInput resolveCellInput(CellInput value, ExecutionInputBindings bindings)
      throws IOException {
    return switch (value) {
      case CellInput.Blank blank -> blank;
      case CellInput.Text text ->
          sameReference(
                  resolveTextSource(text.source(), bindings, true, "cell text"), text.source())
              ? text
              : new CellInput.Text(resolveTextSource(text.source(), bindings, true, "cell text"));
      case CellInput.RichText richText ->
          sameReference(
                  SourceBackedStructuredInputResolver.resolveRuns(richText.runs(), bindings),
                  richText.runs())
              ? richText
              : new CellInput.RichText(
                  SourceBackedStructuredInputResolver.resolveRuns(richText.runs(), bindings));
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

  private static String resolveFormulaText(TextSourceInput source, ExecutionInputBindings bindings)
      throws IOException {
    String formula = resolveText(source, bindings, true, "formula");
    return formula.startsWith("=") ? formula.substring(1) : formula;
  }

  static TextSourceInput resolveTextSource(
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

  static BinarySourceInput resolveBinarySource(
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
}
