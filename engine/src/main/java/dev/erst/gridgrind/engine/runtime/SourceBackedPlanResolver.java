package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.SourceBackedResolutionIdentitySupport.sameReference;

import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationAction;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
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
    if (!bindings.hasRequestPathAccess()) {
      try (RequestPathAccess access =
          new RequestPathAccess(
              bindings.workingDirectory(), bindings.tempFileFactory(), bindings.executionGrant())) {
        return resolveBound(plan, bindings.withRequestPathAccess(access));
      }
    }
    return resolveBound(plan, bindings);
  }

  private static WorkbookPlan resolveBound(WorkbookPlan plan, ExecutionInputBindings bindings)
      throws IOException {
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

  static CellValueSpec resolveCellInput(CellValueSpec value, ExecutionInputBindings bindings)
      throws IOException {
    return switch (value) {
      case CellValueSpec.Blank blank -> blank;
      case CellValueSpec.Text text -> {
        TextSourceSpec resolvedSource =
            resolveTextSource(text.value(), bindings, true, "cell text");
        yield sameReference(resolvedSource, text.value())
            ? text
            : new CellValueSpec.Text(resolvedSource);
      }
      case CellValueSpec.RichText richText -> {
        var resolvedRuns =
            SourceBackedStructuredInputResolver.resolveRuns(richText.runs(), bindings);
        yield sameReference(resolvedRuns, richText.runs())
            ? richText
            : new CellValueSpec.RichText(resolvedRuns);
      }
      case CellValueSpec.NumberValue numberValue -> numberValue;
      case CellValueSpec.BooleanValue booleanValue -> booleanValue;
      case CellValueSpec.ErrorValue errorValue -> errorValue;
      case CellValueSpec.Date date -> date;
      case CellValueSpec.DateTime dateTime -> dateTime;
      case CellValueSpec.Formula formula -> {
        TextSourceSpec resolvedSource = resolveFormulaSource(formula.value(), bindings);
        yield sameReference(resolvedSource, formula.value())
            ? formula
            : new CellValueSpec.Formula(resolvedSource);
      }
      case CellValueSpec.RawFormula rawFormula -> {
        TextSourceSpec resolvedSource = resolveRawFormulaSource(rawFormula.value(), bindings);
        yield sameReference(resolvedSource, rawFormula.value())
            ? rawFormula
            : new CellValueSpec.RawFormula(resolvedSource);
      }
    };
  }

  static TextSourceSpec resolveTextSource(
      TextSourceSpec source,
      ExecutionInputBindings bindings,
      boolean requireNonBlank,
      String inputKind)
      throws IOException {
    if (!bindings.hasRequestPathAccess()) {
      try (RequestPathAccess access =
          new RequestPathAccess(
              bindings.workingDirectory(), bindings.tempFileFactory(), bindings.executionGrant())) {
        return resolveTextSource(
            source, bindings.withRequestPathAccess(access), requireNonBlank, inputKind);
      }
    }
    return resolveOrCollect(
        source,
        bindings,
        () -> {
          String resolvedText = resolveText(source, bindings, requireNonBlank, inputKind);
          return source instanceof TextSourceSpec.Inline
              ? source
              : new TextSourceSpec.Inline(resolvedText);
        });
  }

  static TextSourceSpec resolveFormulaSource(TextSourceSpec source, ExecutionInputBindings bindings)
      throws IOException {
    TextSourceSpec resolvedSource = resolveTextSource(source, bindings, true, "formula");
    if (!(resolvedSource instanceof TextSourceSpec.Inline inline)) {
      return source;
    }
    String resolvedText = inline.text();
    new CellValueSpec.Formula(new TextSourceSpec.Inline(resolvedText));
    return source instanceof TextSourceSpec.Inline
        ? source
        : new TextSourceSpec.Inline(resolvedText);
  }

  static TextSourceSpec resolveRawFormulaSource(
      TextSourceSpec source, ExecutionInputBindings bindings) throws IOException {
    TextSourceSpec resolvedSource = resolveTextSource(source, bindings, true, "raw formula");
    if (!(resolvedSource instanceof TextSourceSpec.Inline inline)) {
      return source;
    }
    TextSourceSpec.Inline validated =
        (TextSourceSpec.Inline)
            new CellValueSpec.RawFormula(new TextSourceSpec.Inline(inline.text())).value();
    return source instanceof TextSourceSpec.Inline ? source : validated;
  }

  static BinarySourceSpec resolveBinarySource(
      BinarySourceSpec source, ExecutionInputBindings bindings, String inputKind)
      throws IOException {
    if (!bindings.hasRequestPathAccess()) {
      try (RequestPathAccess access =
          new RequestPathAccess(
              bindings.workingDirectory(), bindings.tempFileFactory(), bindings.executionGrant())) {
        return resolveBinarySource(source, bindings.withRequestPathAccess(access), inputKind);
      }
    }
    return resolveOrCollect(
        source,
        bindings,
        () -> {
          String resolvedBase64 = resolveBinaryBase64(source, bindings, inputKind);
          return source instanceof BinarySourceSpec.InlineBase64 inline
                  && inline.base64Data().equals(resolvedBase64)
              ? source
              : new BinarySourceSpec.InlineBase64(resolvedBase64);
        });
  }

  private static String resolveText(
      TextSourceSpec source,
      ExecutionInputBindings bindings,
      boolean requireNonBlank,
      String inputKind)
      throws IOException {
    Objects.requireNonNull(source, "source must not be null");
    String text =
        switch (source) {
          case TextSourceSpec.Inline inline -> inline.text();
          case TextSourceSpec.Utf8File file -> readUtf8File(file.path(), bindings, inputKind);
          case TextSourceSpec.Stdin _ -> readStandardInputText(bindings, inputKind);
        };
    if (requireNonBlank && text.isBlank()) {
      throw new IllegalArgumentException(inputKind + " must not be blank");
    }
    return text;
  }

  private static String readUtf8File(
      String path, ExecutionInputBindings bindings, String inputKind) // LIM-030
      throws IOException {
    try {
      return Files.readString(
          bindings
              .requestPathAccess()
              .materializeRead(path, inputKind, "gridgrind-request-input-", ".utf8"),
          StandardCharsets.UTF_8);
    } catch (java.nio.file.InvalidPathException exception) {
      throw new InputSourceReadException(
          "Invalid " + inputKind + " path: " + path, inputKind, path, exception);
    } catch (IOException exception) {
      throw inputFileFailure(path, inputKind, exception);
    }
  }

  private static String readStandardInputText(ExecutionInputBindings bindings, String inputKind)
      throws IOException {
    byte[] bytes = standardInputBytes(bindings, inputKind);
    return new String(bytes, StandardCharsets.UTF_8);
  }

  private static String resolveBinaryBase64(
      BinarySourceSpec source, ExecutionInputBindings bindings, String inputKind)
      throws IOException {
    Objects.requireNonNull(source, "source must not be null");
    byte[] bytes =
        switch (source) {
          case BinarySourceSpec.InlineBase64 inline ->
              Base64.getDecoder().decode(inline.base64Data());
          case BinarySourceSpec.File file -> readBinaryFile(file.path(), bindings, inputKind);
          case BinarySourceSpec.Stdin _ -> standardInputBytes(bindings, inputKind);
        };
    if (bytes.length == 0) {
      throw new IllegalArgumentException(inputKind + " must not be empty");
    }
    return Base64.getEncoder().encodeToString(bytes);
  }

  private static byte[] readBinaryFile(
      String path, ExecutionInputBindings bindings, String inputKind) throws IOException {
    try {
      return Files.readAllBytes(
          bindings
              .requestPathAccess()
              .materializeRead(path, inputKind, "gridgrind-request-input-", ".bin"));
    } catch (java.nio.file.InvalidPathException exception) {
      throw new InputSourceReadException(
          "Invalid " + inputKind + " path: " + path, inputKind, path, exception);
    } catch (IOException exception) {
      throw inputFileFailure(path, inputKind, exception);
    }
  }

  static IOException inputFileFailure(String path, String inputKind, IOException exception) {
    if (exception instanceof java.nio.file.NoSuchFileException) {
      return new InputSourceNotFoundException(
          inputKind + " file does not exist: " + path, inputKind, path, exception);
    }
    if (exception instanceof UnsafePathAccessException) {
      return exception;
    }
    return new InputSourceReadException(
        "Failed to read " + inputKind + " file: " + path, inputKind, path, exception);
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
      if (bindings.collectInputResolutionFailure(exception, source)) {
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

  static List<List<CellValueSpec>> resolveRows(
      List<List<CellValueSpec>> rows, ExecutionInputBindings bindings) throws IOException {
    List<List<CellValueSpec>> resolvedRows = new ArrayList<>(rows.size());
    boolean changed = false;
    for (List<CellValueSpec> row : rows) {
      List<CellValueSpec> resolvedRow = resolveCells(row, bindings);
      resolvedRows.add(resolvedRow);
      changed |= !sameReference(resolvedRow, row);
    }
    return changed ? List.copyOf(resolvedRows) : rows;
  }

  static List<CellValueSpec> resolveCells(
      List<CellValueSpec> values, ExecutionInputBindings bindings) throws IOException {
    List<CellValueSpec> resolvedValues = new ArrayList<>(values.size());
    boolean changed = false;
    for (CellValueSpec value : values) {
      CellValueSpec resolvedValue = resolveCellInput(value, bindings);
      resolvedValues.add(resolvedValue);
      changed |= !sameReference(resolvedValue, value);
    }
    return changed ? List.copyOf(resolvedValues) : values;
  }
}
