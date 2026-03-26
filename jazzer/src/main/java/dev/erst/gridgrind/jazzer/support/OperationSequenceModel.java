package dev.erst.gridgrind.jazzer.support;

import com.code_intelligence.jazzer.api.FuzzedDataProvider;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.protocol.GridGrindRequest;
import dev.erst.gridgrind.protocol.WorkbookOperation;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Builds bounded protocol requests and workbook command sequences for Jazzer harnesses. */
public final class OperationSequenceModel {
  private OperationSequenceModel() {}

  /** Returns a bounded protocol workflow plus any owned local scratch paths it created. */
  public static GeneratedProtocolWorkflow nextProtocolWorkflow(FuzzedDataProvider data)
      throws IOException {
    Objects.requireNonNull(data, "data must not be null");

    String primarySheet = FuzzDataDecoders.nextSheetName(data, true);
    String secondarySheet = FuzzDataDecoders.nextSheetName(data, true);
    List<WorkbookOperation> operations = new ArrayList<>();

    int operationCount = data.consumeInt(0, 8);
    for (int index = 0; index < operationCount; index++) {
      operations.add(nextOperation(data, primarySheet, secondarySheet));
    }

    List<GridGrindRequest.SheetInspectionRequest> sheets = new ArrayList<>();
    if (data.consumeBoolean()) {
      sheets.add(
          new GridGrindRequest.SheetInspectionRequest(
              data.consumeBoolean() ? primarySheet : secondarySheet,
              List.of(FuzzDataDecoders.nextNonBlankCellAddress(data, data.consumeBoolean())),
              null,
              null));
    }

    WorkflowStorage workflowStorage = nextWorkflowStorage(primarySheet, secondarySheet, data);
    return new GeneratedProtocolWorkflow(
        new GridGrindRequest(
            workflowStorage.source(),
            workflowStorage.persistence(),
            List.copyOf(operations),
            new GridGrindRequest.WorkbookAnalysisRequest(List.copyOf(sheets))),
        List.of(workflowStorage.cleanupRoot()));
  }

  /** Returns a bounded sequence of workbook-core commands. */
  public static List<WorkbookCommand> nextWorkbookCommands(FuzzedDataProvider data) {
    Objects.requireNonNull(data, "data must not be null");

    String primarySheet = FuzzDataDecoders.nextSheetName(data, true);
    String secondarySheet = FuzzDataDecoders.nextSheetName(data, true);
    List<WorkbookCommand> commands = new ArrayList<>();
    int commandCount = data.consumeInt(1, 10);
    for (int index = 0; index < commandCount; index++) {
      commands.add(nextCommand(data, primarySheet, secondarySheet));
    }
    return List.copyOf(commands);
  }

  private static WorkbookOperation nextOperation(
      FuzzedDataProvider data, String primarySheet, String secondarySheet) {
    String targetSheet = data.consumeBoolean() ? primarySheet : secondarySheet;
    boolean validAddress = data.consumeBoolean();
    boolean validRange = data.consumeBoolean();
    IndexSpan columnSpan = nextIndexSpan(data, 3);
    IndexSpan rowSpan = nextIndexSpan(data, 3);
    FreezePaneArguments freezePaneArguments = nextFreezePaneArguments(data);

    return switch (data.consumeInt(0, 15)) {
      case 0 -> new WorkbookOperation.EnsureSheet(targetSheet);
      case 1 -> new WorkbookOperation.RenameSheet(targetSheet, primarySheet + "Renamed");
      case 2 -> new WorkbookOperation.DeleteSheet(targetSheet);
      case 3 -> new WorkbookOperation.MoveSheet(targetSheet, data.consumeInt(0, 2));
      case 4 ->
          new WorkbookOperation.MergeCells(
              targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
      case 5 ->
          new WorkbookOperation.UnmergeCells(
              targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
      case 6 ->
          new WorkbookOperation.SetColumnWidth(
              targetSheet,
              columnSpan.first(),
              columnSpan.last(),
              data.consumeRegularDouble(1.0d, 20.0d));
      case 7 ->
          new WorkbookOperation.SetRowHeight(
              targetSheet,
              rowSpan.first(),
              rowSpan.last(),
              data.consumeRegularDouble(5.0d, 40.0d));
      case 8 ->
          new WorkbookOperation.FreezePanes(
              targetSheet,
              freezePaneArguments.splitColumn(),
              freezePaneArguments.splitRow(),
              freezePaneArguments.leftmostColumn(),
              freezePaneArguments.topRow());
      case 9 ->
          new WorkbookOperation.SetCell(
              targetSheet,
              FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
              FuzzDataDecoders.nextCellInput(data));
      case 10 -> {
        String range = validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false);
        yield new WorkbookOperation.SetRange(
            targetSheet, range, FuzzDataDecoders.nextProtocolMatrix(data, 2, 2));
      }
      case 11 ->
          new WorkbookOperation.ClearRange(
              targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
      case 12 ->
          new WorkbookOperation.ApplyStyle(
              targetSheet,
              validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false),
              WorkbookStyleInputs.nextStyleInput(data));
      case 13 ->
          new WorkbookOperation.AppendRow(
              targetSheet,
              List.of(FuzzDataDecoders.nextCellInput(data), FuzzDataDecoders.nextCellInput(data)));
      case 14 -> new WorkbookOperation.AutoSizeColumns(targetSheet);
      default ->
          data.consumeBoolean()
              ? new WorkbookOperation.EvaluateFormulas()
              : new WorkbookOperation.ForceFormulaRecalculationOnOpen();
    };
  }

  private static WorkbookCommand nextCommand(
      FuzzedDataProvider data, String primarySheet, String secondarySheet) {
    String targetSheet = data.consumeBoolean() ? primarySheet : secondarySheet;
    boolean validAddress = data.consumeBoolean();
    boolean validRange = data.consumeBoolean();
    IndexSpan columnSpan = nextIndexSpan(data, 3);
    IndexSpan rowSpan = nextIndexSpan(data, 3);
    FreezePaneArguments freezePaneArguments = nextFreezePaneArguments(data);

    return switch (data.consumeInt(0, 15)) {
      case 0 -> new WorkbookCommand.CreateSheet(targetSheet);
      case 1 -> new WorkbookCommand.RenameSheet(targetSheet, primarySheet + "Renamed");
      case 2 -> new WorkbookCommand.DeleteSheet(targetSheet);
      case 3 -> new WorkbookCommand.MoveSheet(targetSheet, data.consumeInt(0, 2));
      case 4 ->
          new WorkbookCommand.MergeCells(
              targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
      case 5 ->
          new WorkbookCommand.UnmergeCells(
              targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
      case 6 ->
          new WorkbookCommand.SetColumnWidth(
              targetSheet,
              columnSpan.first(),
              columnSpan.last(),
              data.consumeRegularDouble(1.0d, 20.0d));
      case 7 ->
          new WorkbookCommand.SetRowHeight(
              targetSheet,
              rowSpan.first(),
              rowSpan.last(),
              data.consumeRegularDouble(5.0d, 40.0d));
      case 8 ->
          new WorkbookCommand.FreezePanes(
              targetSheet,
              freezePaneArguments.splitColumn(),
              freezePaneArguments.splitRow(),
              freezePaneArguments.leftmostColumn(),
              freezePaneArguments.topRow());
      case 9 ->
          new WorkbookCommand.SetCell(
              targetSheet,
              FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
              FuzzDataDecoders.nextExcelCellValue(data));
      case 10 -> {
        String range = validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false);
        yield new WorkbookCommand.SetRange(
            targetSheet, range, FuzzDataDecoders.nextExcelMatrix(data, 2, 2));
      }
      case 11 ->
          new WorkbookCommand.ClearRange(
              targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
      case 12 ->
          new WorkbookCommand.ApplyStyle(
              targetSheet,
              validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false),
              FuzzDataDecoders.nextStyle(data));
      case 13 ->
          new WorkbookCommand.AppendRow(
              targetSheet,
              List.of(
                  FuzzDataDecoders.nextExcelCellValue(data),
                  FuzzDataDecoders.nextExcelCellValue(data)));
      case 14 -> new WorkbookCommand.AutoSizeColumns(targetSheet);
      default ->
          data.consumeBoolean()
              ? new WorkbookCommand.EvaluateAllFormulas()
              : new WorkbookCommand.ForceFormulaRecalculationOnOpen();
    };
  }

  private static IndexSpan nextIndexSpan(FuzzedDataProvider data, int upperBound) {
    int first = data.consumeInt(0, upperBound - 1);
    return new IndexSpan(first, data.consumeInt(first, upperBound));
  }

  private static FreezePaneArguments nextFreezePaneArguments(FuzzedDataProvider data) {
    return switch (data.consumeInt(0, 2)) {
      case 0 -> {
        int splitColumn = data.consumeInt(1, 3);
        yield new FreezePaneArguments(
            splitColumn, 0, data.consumeInt(splitColumn, splitColumn + 2), 0);
      }
      case 1 -> {
        int splitRow = data.consumeInt(1, 3);
        yield new FreezePaneArguments(0, splitRow, 0, data.consumeInt(splitRow, splitRow + 2));
      }
      default -> {
        int splitColumn = data.consumeInt(1, 3);
        int splitRow = data.consumeInt(1, 3);
        yield new FreezePaneArguments(
            splitColumn,
            splitRow,
            data.consumeInt(splitColumn, splitColumn + 2),
            data.consumeInt(splitRow, splitRow + 2));
      }
    };
  }

  private static WorkflowStorage nextWorkflowStorage(
      String primarySheet, String secondarySheet, FuzzedDataProvider data) throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-jazzer-workflow-");
    Path sourcePath = directory.resolve("source.xlsx");
    Path saveAsPath = directory.resolve("output.xlsx");

    return switch (data.consumeInt(0, 4)) {
      case 0 ->
          new WorkflowStorage(
              new GridGrindRequest.WorkbookSource.New(),
              new GridGrindRequest.WorkbookPersistence.None(),
              directory);
      case 1 ->
          new WorkflowStorage(
              new GridGrindRequest.WorkbookSource.New(),
              new GridGrindRequest.WorkbookPersistence.SaveAs(saveAsPath.toString()),
              directory);
      case 2 -> {
        writeExistingWorkbook(sourcePath, primarySheet, secondarySheet, data);
        yield new WorkflowStorage(
            new GridGrindRequest.WorkbookSource.ExistingFile(sourcePath.toString()),
            new GridGrindRequest.WorkbookPersistence.None(),
            directory);
      }
      case 3 -> {
        writeExistingWorkbook(sourcePath, primarySheet, secondarySheet, data);
        yield new WorkflowStorage(
            new GridGrindRequest.WorkbookSource.ExistingFile(sourcePath.toString()),
            new GridGrindRequest.WorkbookPersistence.OverwriteSource(),
            directory);
      }
      default -> {
        writeExistingWorkbook(sourcePath, primarySheet, secondarySheet, data);
        yield new WorkflowStorage(
            new GridGrindRequest.WorkbookSource.ExistingFile(sourcePath.toString()),
            new GridGrindRequest.WorkbookPersistence.SaveAs(saveAsPath.toString()),
            directory);
      }
    };
  }

  private static void writeExistingWorkbook(
      Path sourcePath, String primarySheet, String secondarySheet, FuzzedDataProvider data)
      throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet(primarySheet).createRow(0).createCell(0).setCellValue("seed");
      if (data.consumeBoolean()) {
        workbook.createSheet(secondarySheet).createRow(0).createCell(0).setCellValue("seed");
      }
      Files.createDirectories(sourcePath.getParent());
      try (OutputStream outputStream = Files.newOutputStream(sourcePath)) {
        workbook.write(outputStream);
      }
    }
  }

  private record IndexSpan(int first, int last) {}

  private record FreezePaneArguments(
      int splitColumn, int splitRow, int leftmostColumn, int topRow) {}

  private record WorkflowStorage(
      GridGrindRequest.WorkbookSource source,
      GridGrindRequest.WorkbookPersistence persistence,
      Path cleanupRoot) {}
}
