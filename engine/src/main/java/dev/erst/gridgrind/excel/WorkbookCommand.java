package dev.erst.gridgrind.excel;

/**
 * Workbook-core commands that can be executed deterministically against an {@link ExcelWorkbook}.
 */
public sealed interface WorkbookCommand
    permits WorkbookSheetCommand,
        WorkbookStructureCommand,
        WorkbookLayoutCommand,
        WorkbookCellCommand,
        WorkbookMetadataCommand,
        WorkbookAnnotationCommand,
        WorkbookDrawingCommand,
        WorkbookFormattingCommand,
        WorkbookTabularCommand {

  /** Returns the canonical operation-style discriminator for diagnostics and telemetry. */
  default String commandType() {
    return WorkbookCommandTypeResolver.commandType(this);
  }
}
