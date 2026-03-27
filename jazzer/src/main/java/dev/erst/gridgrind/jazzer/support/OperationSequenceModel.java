package dev.erst.gridgrind.jazzer.support;

import com.code_intelligence.jazzer.api.FuzzedDataProvider;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.protocol.CommentInput;
import dev.erst.gridgrind.protocol.CellSelection;
import dev.erst.gridgrind.protocol.GridGrindRequest;
import dev.erst.gridgrind.protocol.HyperlinkTarget;
import dev.erst.gridgrind.protocol.NamedRangeSelection;
import dev.erst.gridgrind.protocol.NamedRangeScope;
import dev.erst.gridgrind.protocol.NamedRangeSelector;
import dev.erst.gridgrind.protocol.NamedRangeTarget;
import dev.erst.gridgrind.protocol.SheetSelection;
import dev.erst.gridgrind.protocol.WorkbookReadOperation;
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
    List<WorkbookReadOperation> reads = new ArrayList<>();
    String workbookNamedRange = nextNamedRangeName(data, true);
    String sheetNamedRange = nextNamedRangeName(data, true);

    int operationCount = data.consumeInt(0, 8);
    for (int index = 0; index < operationCount; index++) {
      operations.add(
          nextOperation(data, primarySheet, secondarySheet, workbookNamedRange, sheetNamedRange));
    }
    int readCount = data.consumeInt(0, 6);
    for (int index = 0; index < readCount; index++) {
      reads.add(
          nextRead(
              data, index, primarySheet, secondarySheet, workbookNamedRange, sheetNamedRange));
    }
    WorkflowStorage workflowStorage = nextWorkflowStorage(primarySheet, secondarySheet, data);
    return new GeneratedProtocolWorkflow(
        new GridGrindRequest(
            workflowStorage.source(),
            workflowStorage.persistence(),
            List.copyOf(operations),
            List.copyOf(reads)),
        List.of(workflowStorage.cleanupRoot()));
  }

  /** Returns a bounded sequence of workbook-core commands. */
  public static List<WorkbookCommand> nextWorkbookCommands(FuzzedDataProvider data) {
    Objects.requireNonNull(data, "data must not be null");

    String primarySheet = FuzzDataDecoders.nextSheetName(data, true);
    String secondarySheet = FuzzDataDecoders.nextSheetName(data, true);
    List<WorkbookCommand> commands = new ArrayList<>();
    String workbookNamedRange = nextNamedRangeName(data, true);
    String sheetNamedRange = nextNamedRangeName(data, true);
    int commandCount = data.consumeInt(1, 10);
    for (int index = 0; index < commandCount; index++) {
      commands.add(
          nextCommand(data, primarySheet, secondarySheet, workbookNamedRange, sheetNamedRange));
    }
    return List.copyOf(commands);
  }

  private static WorkbookOperation nextOperation(
      FuzzedDataProvider data,
      String primarySheet,
      String secondarySheet,
      String workbookNamedRange,
      String sheetNamedRange) {
    String targetSheet = data.consumeBoolean() ? primarySheet : secondarySheet;
    boolean validAddress = data.consumeBoolean();
    boolean validRange = data.consumeBoolean();
    boolean validName = data.consumeBoolean();
    IndexSpan columnSpan = nextIndexSpan(data, 3);
    IndexSpan rowSpan = nextIndexSpan(data, 3);
    FreezePaneArguments freezePaneArguments = nextFreezePaneArguments(data);
    String namedRangeName = data.consumeBoolean() ? workbookNamedRange : sheetNamedRange;

    return switch (data.consumeInt(0, 20)) {
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
          new WorkbookOperation.SetHyperlink(
              targetSheet,
              FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
              nextHyperlinkTarget(data));
      case 13 ->
          new WorkbookOperation.ClearHyperlink(
              targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress));
      case 14 ->
          new WorkbookOperation.SetComment(
              targetSheet,
              FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
              nextCommentInput(data));
      case 15 ->
          new WorkbookOperation.ClearComment(
              targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress));
      case 16 ->
          new WorkbookOperation.ApplyStyle(
              targetSheet,
              validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false),
              WorkbookStyleInputs.nextStyleInput(data));
      case 17 ->
          new WorkbookOperation.SetNamedRange(
              validName ? namedRangeName : nextNamedRangeName(data, false),
              data.consumeBoolean()
                  ? new NamedRangeScope.Workbook()
                  : new NamedRangeScope.Sheet(targetSheet),
              new NamedRangeTarget(
                  targetSheet, validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false)));
      case 18 ->
          new WorkbookOperation.DeleteNamedRange(
              validName ? namedRangeName : nextNamedRangeName(data, false),
              data.consumeBoolean()
                  ? new NamedRangeScope.Workbook()
                  : new NamedRangeScope.Sheet(targetSheet));
      case 19 ->
          new WorkbookOperation.AppendRow(
              targetSheet,
              List.of(FuzzDataDecoders.nextCellInput(data), FuzzDataDecoders.nextCellInput(data)));
      case 20 -> new WorkbookOperation.AutoSizeColumns(targetSheet);
      default ->
          data.consumeBoolean()
              ? new WorkbookOperation.EvaluateFormulas()
              : new WorkbookOperation.ForceFormulaRecalculationOnOpen();
    };
  }

  private static WorkbookCommand nextCommand(
      FuzzedDataProvider data,
      String primarySheet,
      String secondarySheet,
      String workbookNamedRange,
      String sheetNamedRange) {
    String targetSheet = data.consumeBoolean() ? primarySheet : secondarySheet;
    boolean validAddress = data.consumeBoolean();
    boolean validRange = data.consumeBoolean();
    boolean validName = data.consumeBoolean();
    IndexSpan columnSpan = nextIndexSpan(data, 3);
    IndexSpan rowSpan = nextIndexSpan(data, 3);
    FreezePaneArguments freezePaneArguments = nextFreezePaneArguments(data);
    String namedRangeName = data.consumeBoolean() ? workbookNamedRange : sheetNamedRange;

    return switch (data.consumeInt(0, 20)) {
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
          new WorkbookCommand.SetHyperlink(
              targetSheet,
              FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
              nextExcelHyperlink(data));
      case 13 ->
          new WorkbookCommand.ClearHyperlink(
              targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress));
      case 14 ->
          new WorkbookCommand.SetComment(
              targetSheet,
              FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
              nextExcelComment(data));
      case 15 ->
          new WorkbookCommand.ClearComment(
              targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress));
      case 16 ->
          new WorkbookCommand.ApplyStyle(
              targetSheet,
              validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false),
              FuzzDataDecoders.nextStyle(data));
      case 17 ->
          new WorkbookCommand.SetNamedRange(
              new ExcelNamedRangeDefinition(
                  validName ? namedRangeName : nextNamedRangeName(data, false),
                  data.consumeBoolean()
                      ? new ExcelNamedRangeScope.WorkbookScope()
                      : new ExcelNamedRangeScope.SheetScope(targetSheet),
                  new ExcelNamedRangeTarget(
                      targetSheet,
                      validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false))));
      case 18 ->
          new WorkbookCommand.DeleteNamedRange(
              validName ? namedRangeName : nextNamedRangeName(data, false),
              data.consumeBoolean()
                  ? new ExcelNamedRangeScope.WorkbookScope()
                  : new ExcelNamedRangeScope.SheetScope(targetSheet));
      case 19 ->
          new WorkbookCommand.AppendRow(
              targetSheet,
              List.of(
                  FuzzDataDecoders.nextExcelCellValue(data),
                  FuzzDataDecoders.nextExcelCellValue(data)));
      case 20 -> new WorkbookCommand.AutoSizeColumns(targetSheet);
      default ->
          data.consumeBoolean()
              ? new WorkbookCommand.EvaluateAllFormulas()
              : new WorkbookCommand.ForceFormulaRecalculationOnOpen();
    };
  }

  private static WorkbookReadOperation nextRead(
      FuzzedDataProvider data,
      int index,
      String primarySheet,
      String secondarySheet,
      String workbookNamedRange,
      String sheetNamedRange) {
    String requestId = "read-" + index;
    String targetSheet = data.consumeBoolean() ? primarySheet : secondarySheet;
    boolean validAddress = data.consumeBoolean();
    boolean validRange = data.consumeBoolean();
    boolean validName = data.consumeBoolean();

    return switch (data.consumeInt(0, 11)) {
      case 0 -> new WorkbookReadOperation.GetWorkbookSummary(requestId);
      case 1 ->
          new WorkbookReadOperation.GetNamedRanges(
              requestId,
              nextNamedRangeSelection(
                  data, targetSheet, workbookNamedRange, sheetNamedRange, validName));
      case 2 -> new WorkbookReadOperation.GetSheetSummary(requestId, targetSheet);
      case 3 ->
          new WorkbookReadOperation.GetCells(
              requestId,
              targetSheet,
              nextReadAddresses(data, validAddress));
      case 4 ->
          new WorkbookReadOperation.GetWindow(
              requestId,
              targetSheet,
              FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
              data.consumeInt(1, 4),
              data.consumeInt(1, 4));
      case 5 -> new WorkbookReadOperation.GetMergedRegions(requestId, targetSheet);
      case 6 ->
          new WorkbookReadOperation.GetHyperlinks(
              requestId, targetSheet, nextCellSelection(data, validAddress));
      case 7 ->
          new WorkbookReadOperation.GetComments(
              requestId, targetSheet, nextCellSelection(data, validAddress));
      case 8 -> new WorkbookReadOperation.GetSheetLayout(requestId, targetSheet);
      case 9 ->
          new WorkbookReadOperation.AnalyzeFormulaSurface(
              requestId, nextSheetSelection(data, primarySheet, secondarySheet));
      case 10 ->
          new WorkbookReadOperation.AnalyzeSheetSchema(
              requestId,
              targetSheet,
              validAddress ? "A1" : FuzzDataDecoders.nextNonBlankCellAddress(data, false),
              data.consumeInt(1, 5),
              data.consumeInt(1, 4));
      default ->
          new WorkbookReadOperation.AnalyzeNamedRangeSurface(
              requestId,
              nextNamedRangeSelection(
                  data, targetSheet, workbookNamedRange, sheetNamedRange, validName));
    };
  }

  private static NamedRangeSelection nextNamedRangeSelection(
      FuzzedDataProvider data,
      String sheetName,
      String workbookNamedRange,
      String sheetNamedRange,
      boolean validName) {
    return switch (data.consumeInt(0, 2)) {
      case 0 -> new NamedRangeSelection.All();
      default ->
          new NamedRangeSelection.Selected(
              List.of(
                  data.consumeBoolean()
                      ? new NamedRangeSelector.WorkbookScope(
                          validName ? workbookNamedRange : nextNamedRangeName(data, false))
                      : new NamedRangeSelector.SheetScope(
                          validName ? sheetNamedRange : nextNamedRangeName(data, false), sheetName)));
    };
  }

  private static CellSelection nextCellSelection(FuzzedDataProvider data, boolean validAddress) {
    return switch (data.consumeInt(0, 1)) {
      case 0 -> new CellSelection.AllUsedCells();
      default ->
          new CellSelection.Selected(
              nextReadAddresses(data, validAddress));
    };
  }

  private static SheetSelection nextSheetSelection(
      FuzzedDataProvider data, String primarySheet, String secondarySheet) {
    return switch (data.consumeInt(0, 1)) {
      case 0 -> new SheetSelection.All();
      default ->
          new SheetSelection.Selected(
              data.consumeBoolean()
                  ? List.of(primarySheet, secondarySheet)
                  : List.of(secondarySheet, primarySheet));
    };
  }

  private static List<String> nextReadAddresses(FuzzedDataProvider data, boolean validAddress) {
    String first = FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress);
    String second = FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress);
    if (first.equals(second)) {
      second = validAddress ? ("A1".equals(first) ? "B2" : "A1") : "ZZZ999999";
    }
    return List.of(first, second);
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
      var primary = workbook.createSheet(primarySheet);
      primary.createRow(0).createCell(0).setCellValue("seed");
      primary.getRow(0).createCell(1).setCellFormula("1+1");
      if (data.consumeBoolean()) {
        workbook.createSheet(secondarySheet).createRow(0).createCell(0).setCellValue("seed");
      }
      Files.createDirectories(sourcePath.getParent());
      try (OutputStream outputStream = Files.newOutputStream(sourcePath)) {
        workbook.write(outputStream);
      }
    }
  }

  private static HyperlinkTarget nextHyperlinkTarget(FuzzedDataProvider data) {
    return switch (data.consumeInt(0, 3)) {
      case 0 -> new HyperlinkTarget.Url("https://example.com/" + nextNamedRangeName(data, true));
      case 1 -> new HyperlinkTarget.Email(nextNamedRangeName(data, true) + "@example.com");
      case 2 -> new HyperlinkTarget.File("/tmp/" + nextNamedRangeName(data, true) + ".xlsx");
      default -> new HyperlinkTarget.Document("Budget!A1");
    };
  }

  private static ExcelHyperlink nextExcelHyperlink(FuzzedDataProvider data) {
    return switch (data.consumeInt(0, 3)) {
      case 0 -> new ExcelHyperlink.Url("https://example.com/" + nextNamedRangeName(data, true));
      case 1 -> new ExcelHyperlink.Email(nextNamedRangeName(data, true) + "@example.com");
      case 2 -> new ExcelHyperlink.File("/tmp/" + nextNamedRangeName(data, true) + ".xlsx");
      default -> new ExcelHyperlink.Document("Budget!A1");
    };
  }

  private static CommentInput nextCommentInput(FuzzedDataProvider data) {
    return new CommentInput(
        "Note " + nextNamedRangeName(data, true),
        "GridGrind",
        data.consumeBoolean());
  }

  private static ExcelComment nextExcelComment(FuzzedDataProvider data) {
    return new ExcelComment("Note " + nextNamedRangeName(data, true), "GridGrind", data.consumeBoolean());
  }

  private static String nextNamedRangeName(FuzzedDataProvider data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");
    if (!valid) {
      return switch (data.consumeInt(0, 4)) {
        case 0 -> "";
        case 1 -> "A1";
        case 2 -> "R1C1";
        case 3 -> "_xlnm.Print_Area";
        default -> "1Budget";
      };
    }
    return switch (data.consumeInt(0, 4)) {
      case 0 -> "BudgetTotal";
      case 1 -> "LocalItem";
      case 2 -> "Report_Value";
      case 3 -> "Summary.Total";
      default -> "Name" + data.consumeInt(1, 9);
    };
  }

  private record IndexSpan(int first, int last) {}

  private record FreezePaneArguments(
      int splitColumn, int splitRow, int leftmostColumn, int topRow) {}

  private record WorkflowStorage(
      GridGrindRequest.WorkbookSource source,
      GridGrindRequest.WorkbookPersistence persistence,
      Path cleanupRoot) {}
}
