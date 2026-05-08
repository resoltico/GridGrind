package dev.erst.gridgrind.excel;

import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/** Public workbook-core execution facade for applying commands and running workbook reads. */
public final class WorkbookExecutionEngine {
  private final WorkbookCommandExecutor commandExecutor;
  private final WorkbookReadExecutor readExecutor;

  /** Creates the production workbook execution facade. */
  public WorkbookExecutionEngine() {
    this(new WorkbookCommandExecutor(), new WorkbookReadExecutor());
  }

  WorkbookExecutionEngine(
      WorkbookCommandExecutor commandExecutor, WorkbookReadExecutor readExecutor) {
    this.commandExecutor =
        Objects.requireNonNull(commandExecutor, "commandExecutor must not be null");
    this.readExecutor = Objects.requireNonNull(readExecutor, "readExecutor must not be null");
  }

  /** Applies one or more commands in order to one workbook instance. */
  public ExcelWorkbook apply(ExcelWorkbook workbook, WorkbookCommand... commands) {
    Objects.requireNonNull(commands, "commands must not be null");
    return apply(workbook, Arrays.asList(commands));
  }

  /** Applies commands from any iterable source in order to one workbook instance. */
  public ExcelWorkbook apply(ExcelWorkbook workbook, Iterable<WorkbookCommand> commands) {
    return commandExecutor.apply(workbook, commands);
  }

  /** Executes one or more read commands against one workbook instance. */
  public List<WorkbookReadResult> read(ExcelWorkbook workbook, WorkbookReadCommand... commands) {
    Objects.requireNonNull(commands, "commands must not be null");
    return read(workbook, new WorkbookLocation.UnsavedWorkbook(), Arrays.asList(commands));
  }

  /** Executes read commands with explicit workbook-location context. */
  public List<WorkbookReadResult> read(
      ExcelWorkbook workbook, WorkbookLocation workbookLocation, WorkbookReadCommand... commands) {
    Objects.requireNonNull(commands, "commands must not be null");
    return read(workbook, workbookLocation, Arrays.asList(commands));
  }

  /** Executes read commands from any iterable source using an unsaved workbook location. */
  public List<WorkbookReadResult> read(
      ExcelWorkbook workbook, Iterable<WorkbookReadCommand> commands) {
    return read(workbook, new WorkbookLocation.UnsavedWorkbook(), commands);
  }

  /** Executes read commands from any iterable source with explicit workbook-location context. */
  public List<WorkbookReadResult> read(
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      Iterable<WorkbookReadCommand> commands) {
    return readExecutor.apply(workbook, workbookLocation, commands);
  }
}
