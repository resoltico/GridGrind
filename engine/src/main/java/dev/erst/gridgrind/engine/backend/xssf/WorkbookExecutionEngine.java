package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.plan.InspectionResult;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/** Public workbook-core execution facade for applying commands and running workbook reads. */
public final class WorkbookExecutionEngine {
  private final WorkbookReadExecutor readExecutor;

  /** Creates the production workbook execution facade. */
  public WorkbookExecutionEngine() {
    this.readExecutor = new WorkbookReadExecutor();
  }

  /** Applies executable core mutations in order. */
  public XssfWorkbook apply(
      XssfWorkbook workbook, dev.erst.gridgrind.core.plan.ResolvedStep... steps) {
    return apply(workbook, Arrays.asList(steps));
  }

  /** Applies executable core mutations from an ordered source. */
  public XssfWorkbook apply(
      XssfWorkbook workbook, Iterable<dev.erst.gridgrind.core.plan.ResolvedStep> steps) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(steps, "steps must not be null");
    for (var step : steps) ResolvedMutationExecutor.apply(workbook, step);
    return workbook;
  }

  /** Executes one or more read commands against one workbook instance. */
  public List<InspectionResult> read(
      XssfWorkbook workbook, dev.erst.gridgrind.core.plan.ResolvedInspectionStep... commands) {
    Objects.requireNonNull(commands, "commands must not be null");
    return read(workbook, new WorkbookLocation.UnsavedWorkbook(), Arrays.asList(commands));
  }

  /** Executes read commands with explicit workbook-location context. */
  public List<InspectionResult> read(
      XssfWorkbook workbook,
      WorkbookLocation workbookLocation,
      dev.erst.gridgrind.core.plan.ResolvedInspectionStep... commands) {
    Objects.requireNonNull(commands, "commands must not be null");
    return read(workbook, workbookLocation, Arrays.asList(commands));
  }

  /** Executes read commands from any iterable source using an unsaved workbook location. */
  public List<InspectionResult> read(
      XssfWorkbook workbook,
      Iterable<dev.erst.gridgrind.core.plan.ResolvedInspectionStep> commands) {
    return read(workbook, new WorkbookLocation.UnsavedWorkbook(), commands);
  }

  /** Executes read commands from any iterable source with explicit workbook-location context. */
  public List<InspectionResult> read(
      XssfWorkbook workbook,
      WorkbookLocation workbookLocation,
      Iterable<dev.erst.gridgrind.core.plan.ResolvedInspectionStep> commands) {
    return readExecutor.apply(workbook, workbookLocation, commands);
  }
}
