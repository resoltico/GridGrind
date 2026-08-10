package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.SourceBackedResolutionIdentitySupport.sameReference;

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
      resolvedSteps.add(resolveStepUnchecked(step, bindings));
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

  /** Resolves one already-bound step so phase-four preflight can collect sibling input failures. */
  static WorkbookStep resolveStep(WorkbookStep step, ExecutionInputBindings bindings)
      throws IOException {
    InputResolutionFailures failures = new InputResolutionFailures();
    WorkbookStep resolved =
        resolveStepUnchecked(step, bindings.collectingInputResolutionFailures(failures));
    failures.throwIfAny();
    return resolved;
  }

  private static WorkbookStep resolveStepUnchecked(
      WorkbookStep step, ExecutionInputBindings bindings) throws IOException {
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
    return SourceBackedMutationActionResolver.resolve(action, bindings);
  }

  static CellInput resolveCellInput(CellInput value, ExecutionInputBindings bindings)
      throws IOException {
    return switch (value) {
      case CellInput.Blank blank -> blank;
      case CellInput.Text text -> {
        TextSourceInput resolvedSource =
            resolveTextSource(text.source(), bindings, true, "cell text");
        yield sameReference(resolvedSource, text.source())
            ? text
            : new CellInput.Text(resolvedSource);
      }
      case CellInput.RichText richText -> {
        var resolvedRuns =
            SourceBackedStructuredInputResolver.resolveRuns(richText.runs(), bindings);
        yield sameReference(resolvedRuns, richText.runs())
            ? richText
            : new CellInput.RichText(resolvedRuns);
      }
      case CellInput.NumberValue numberValue -> numberValue;
      case CellInput.BooleanValue booleanValue -> booleanValue;
      case CellInput.ErrorValue errorValue -> errorValue;
      case CellInput.Date date -> date;
      case CellInput.DateTime dateTime -> dateTime;
      case CellInput.Formula formula -> {
        TextSourceInput resolvedSource = resolveFormulaSource(formula.source(), bindings);
        yield sameReference(resolvedSource, formula.source())
            ? formula
            : new CellInput.Formula(resolvedSource);
      }
    };
  }

  static TextSourceInput resolveTextSource(
      TextSourceInput source,
      ExecutionInputBindings bindings,
      boolean requireNonBlank,
      String inputKind)
      throws IOException {
    return resolveOrCollect(
        source,
        bindings,
        () -> {
          String resolvedText = resolveText(source, bindings, requireNonBlank, inputKind);
          return source instanceof TextSourceInput.Inline
              ? source
              : new TextSourceInput.Inline(resolvedText);
        });
  }

  static TextSourceInput resolveFormulaSource(
      TextSourceInput source, ExecutionInputBindings bindings) throws IOException {
    TextSourceInput resolvedSource = resolveTextSource(source, bindings, true, "formula");
    if (!(resolvedSource instanceof TextSourceInput.Inline inline)) {
      return source;
    }
    String resolvedText = inline.text();
    if (resolvedText.startsWith("=")) {
      resolvedText = resolvedText.substring(1);
    }
    if (source instanceof TextSourceInput.Inline originalInline) {
      if (originalInline.text().equals(resolvedText)) {
        return source;
      }
      return new TextSourceInput.Inline(resolvedText);
    }
    return new TextSourceInput.Inline(resolvedText);
  }

  static BinarySourceInput resolveBinarySource(
      BinarySourceInput source, ExecutionInputBindings bindings, String inputKind)
      throws IOException {
    return resolveOrCollect(
        source,
        bindings,
        () -> {
          String resolvedBase64 = resolveBinaryBase64(source, bindings, inputKind);
          return source instanceof BinarySourceInput.InlineBase64 inline
                  && inline.base64Data().equals(resolvedBase64)
              ? source
              : new BinarySourceInput.InlineBase64(resolvedBase64);
        });
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

  private static <T> T resolveOrCollect(
      T source, ExecutionInputBindings bindings, SourceResolution<T> resolution)
      throws IOException {
    try {
      return resolution.resolve();
    } catch (IOException | RuntimeException exception) {
      if (bindings.collectInputResolutionFailure(exception)) {
        return source;
      }
      throw exception;
    }
  }

  /** One source-resolution operation whose checked failure can join the current batch. */
  @FunctionalInterface
  private interface SourceResolution<T> {
    /** Resolves one authored source value. */
    T resolve() throws IOException;
  }

  static List<List<CellInput>> resolveRows(
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

  static List<CellInput> resolveCells(List<CellInput> values, ExecutionInputBindings bindings)
      throws IOException {
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
