package dev.erst.gridgrind.jazzer.support;

import com.code_intelligence.jazzer.api.FuzzedDataProvider;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorAlert;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.ExcelDataValidationPrompt;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.excel.ExcelDifferentialStyle;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelPaneRegion;
import dev.erst.gridgrind.excel.ExcelRangeSelection;
import dev.erst.gridgrind.excel.ExcelHeaderFooterText;
import dev.erst.gridgrind.excel.ExcelPrintLayout;
import dev.erst.gridgrind.excel.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableSelection;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.protocol.dto.CellSelection;
import dev.erst.gridgrind.protocol.dto.CommentInput;
import dev.erst.gridgrind.protocol.dto.ComparisonOperator;
import dev.erst.gridgrind.protocol.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.protocol.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.protocol.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.protocol.dto.DataValidationErrorStyle;
import dev.erst.gridgrind.protocol.dto.DataValidationInput;
import dev.erst.gridgrind.protocol.dto.DataValidationPromptInput;
import dev.erst.gridgrind.protocol.dto.DataValidationRuleInput;
import dev.erst.gridgrind.protocol.dto.DifferentialStyleInput;
import dev.erst.gridgrind.protocol.dto.GridGrindRequest;
import dev.erst.gridgrind.protocol.dto.HeaderFooterTextInput;
import dev.erst.gridgrind.protocol.dto.HyperlinkTarget;
import dev.erst.gridgrind.protocol.dto.NamedRangeSelection;
import dev.erst.gridgrind.protocol.dto.NamedRangeScope;
import dev.erst.gridgrind.protocol.dto.NamedRangeSelector;
import dev.erst.gridgrind.protocol.dto.NamedRangeTarget;
import dev.erst.gridgrind.protocol.dto.PaneInput;
import dev.erst.gridgrind.protocol.dto.PaneRegion;
import dev.erst.gridgrind.protocol.dto.PrintAreaInput;
import dev.erst.gridgrind.protocol.dto.PrintLayoutInput;
import dev.erst.gridgrind.protocol.dto.PrintOrientation;
import dev.erst.gridgrind.protocol.dto.PrintScalingInput;
import dev.erst.gridgrind.protocol.dto.PrintTitleColumnsInput;
import dev.erst.gridgrind.protocol.dto.PrintTitleRowsInput;
import dev.erst.gridgrind.protocol.dto.RangeSelection;
import dev.erst.gridgrind.protocol.dto.SheetSelection;
import dev.erst.gridgrind.protocol.dto.SheetCopyPosition;
import dev.erst.gridgrind.protocol.dto.SheetProtectionSettings;
import dev.erst.gridgrind.protocol.dto.SheetVisibility;
import dev.erst.gridgrind.protocol.dto.TableInput;
import dev.erst.gridgrind.protocol.dto.TableSelection;
import dev.erst.gridgrind.protocol.dto.TableStyleInput;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.stream.Stream;
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
    String namedRangeName = data.consumeBoolean() ? workbookNamedRange : sheetNamedRange;
    String tableName = nextTableName(data, validName, targetSheet);
    int selector = nextSelectorByte(data);

    return switch (selectorFamily(selector)) {
      case 0x0 -> switch (selectorSlot(selector)) {
        case 0x0 -> new WorkbookOperation.EnsureSheet(targetSheet);
        case 0x1 -> new WorkbookOperation.RenameSheet(targetSheet, primarySheet + "Renamed");
        case 0x2 -> new WorkbookOperation.DeleteSheet(targetSheet);
        case 0x3 -> new WorkbookOperation.MoveSheet(targetSheet, data.consumeInt(0, 2));
        case 0x4 ->
            new WorkbookOperation.CopySheet(
                targetSheet, nextCopySheetName(targetSheet), nextSheetCopyPosition(data));
        case 0x5 -> new WorkbookOperation.SetActiveSheet(targetSheet);
        case 0x6 ->
            new WorkbookOperation.SetSelectedSheets(
                nextSelectedSheetNames(data, primarySheet, secondarySheet));
        case 0x7 ->
            new WorkbookOperation.SetSheetVisibility(targetSheet, nextProtocolSheetVisibility(data));
        case 0x8 ->
            new WorkbookOperation.SetSheetProtection(
                targetSheet, nextProtocolSheetProtectionSettings(data));
        default -> new WorkbookOperation.ClearSheetProtection(targetSheet);
      };
      case 0x1 -> switch (selectorSlot(selector)) {
        case 0x0 ->
            new WorkbookOperation.MergeCells(
                targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
        case 0x1 ->
            new WorkbookOperation.UnmergeCells(
                targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
        case 0x2 ->
            new WorkbookOperation.SetColumnWidth(
                targetSheet,
                columnSpan.first(),
                columnSpan.last(),
                data.consumeRegularDouble(1.0d, 20.0d));
        case 0x3 ->
            new WorkbookOperation.SetRowHeight(
                targetSheet,
                rowSpan.first(),
                rowSpan.last(),
                data.consumeRegularDouble(5.0d, 40.0d));
        case 0x4 -> new WorkbookOperation.SetSheetPane(targetSheet, nextPaneInput(data));
        case 0x5 ->
            new WorkbookOperation.SetSheetZoom(targetSheet, data.consumeInt(10, 400));
        case 0x6 ->
            new WorkbookOperation.SetPrintLayout(targetSheet, nextPrintLayoutInput(data));
        default -> new WorkbookOperation.ClearPrintLayout(targetSheet);
      };
      case 0x2 -> switch (selectorSlot(selector)) {
        case 0x0 ->
            new WorkbookOperation.SetCell(
                targetSheet,
                FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
                FuzzDataDecoders.nextCellInput(data));
        case 0x1 -> {
          String range = validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false);
          yield new WorkbookOperation.SetRange(
              targetSheet, range, FuzzDataDecoders.nextProtocolMatrix(data, 2, 2));
        }
        case 0x2 ->
            new WorkbookOperation.ClearRange(
                targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
        default ->
            new WorkbookOperation.AppendRow(
                targetSheet,
                List.of(FuzzDataDecoders.nextCellInput(data), FuzzDataDecoders.nextCellInput(data)));
      };
      case 0x3 -> switch (selectorSlot(selector)) {
        case 0x0 ->
            new WorkbookOperation.SetHyperlink(
                targetSheet,
                FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
                nextHyperlinkTarget(data));
        case 0x1 ->
            new WorkbookOperation.ClearHyperlink(
                targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress));
        case 0x2 ->
            new WorkbookOperation.SetComment(
                targetSheet,
                FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
                nextCommentInput(data));
        case 0x3 ->
            new WorkbookOperation.ClearComment(
                targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress));
        default ->
            new WorkbookOperation.ApplyStyle(
                targetSheet,
                validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false),
                WorkbookStyleInputs.nextStyleInput(data));
      };
      case 0x4 -> switch (selectorSlot(selector)) {
        case 0x0 ->
            new WorkbookOperation.SetDataValidation(
                targetSheet,
                validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false),
                nextDataValidationInput(data));
        case 0x1 ->
            new WorkbookOperation.ClearDataValidations(
                targetSheet, nextRangeSelection(data, validRange));
        case 0x2 ->
            new WorkbookOperation.SetConditionalFormatting(
                targetSheet, nextConditionalFormattingInput(data, validRange));
        default ->
            new WorkbookOperation.ClearConditionalFormatting(
                targetSheet, nextRangeSelection(data, validRange));
      };
      case 0x5 -> switch (selectorSlot(selector)) {
        case 0x0 -> new WorkbookOperation.SetAutofilter(targetSheet, nextAutofilterRange(validRange));
        case 0x1 -> new WorkbookOperation.ClearAutofilter(targetSheet);
        case 0x2 ->
            new WorkbookOperation.SetTable(
                nextTableInput(data, targetSheet, tableName, validRange));
        default -> new WorkbookOperation.DeleteTable(tableName, targetSheet);
      };
      case 0x6 -> switch (selectorSlot(selector)) {
        case 0x0 ->
            new WorkbookOperation.SetNamedRange(
                validName ? namedRangeName : nextNamedRangeName(data, false),
                data.consumeBoolean()
                    ? new NamedRangeScope.Workbook()
                    : new NamedRangeScope.Sheet(targetSheet),
                new NamedRangeTarget(
                    targetSheet,
                    validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false)));
        default ->
            new WorkbookOperation.DeleteNamedRange(
                validName ? namedRangeName : nextNamedRangeName(data, false),
                data.consumeBoolean()
                    ? new NamedRangeScope.Workbook()
                    : new NamedRangeScope.Sheet(targetSheet));
      };
      default -> switch (selectorSlot(selector)) {
        case 0x0 -> new WorkbookOperation.AutoSizeColumns(targetSheet);
        case 0x1 -> new WorkbookOperation.EvaluateFormulas();
        default -> new WorkbookOperation.ForceFormulaRecalculationOnOpen();
      };
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
    String namedRangeName = data.consumeBoolean() ? workbookNamedRange : sheetNamedRange;
    String tableName = nextTableName(data, validName, targetSheet);
    int selector = nextSelectorByte(data);

    return switch (selectorFamily(selector)) {
      case 0x0 -> switch (selectorSlot(selector)) {
        case 0x0 -> new WorkbookCommand.CreateSheet(targetSheet);
        case 0x1 -> new WorkbookCommand.RenameSheet(targetSheet, primarySheet + "Renamed");
        case 0x2 -> new WorkbookCommand.DeleteSheet(targetSheet);
        case 0x3 -> new WorkbookCommand.MoveSheet(targetSheet, data.consumeInt(0, 2));
        case 0x4 ->
            new WorkbookCommand.CopySheet(
                targetSheet, nextCopySheetName(targetSheet), nextExcelSheetCopyPosition(data));
        case 0x5 -> new WorkbookCommand.SetActiveSheet(targetSheet);
        case 0x6 ->
            new WorkbookCommand.SetSelectedSheets(
                nextSelectedSheetNames(data, primarySheet, secondarySheet));
        case 0x7 -> new WorkbookCommand.SetSheetVisibility(targetSheet, nextSheetVisibility(data));
        case 0x8 ->
            new WorkbookCommand.SetSheetProtection(targetSheet, nextSheetProtectionSettings(data));
        default -> new WorkbookCommand.ClearSheetProtection(targetSheet);
      };
      case 0x1 -> switch (selectorSlot(selector)) {
        case 0x0 ->
            new WorkbookCommand.MergeCells(
                targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
        case 0x1 ->
            new WorkbookCommand.UnmergeCells(
                targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
        case 0x2 ->
            new WorkbookCommand.SetColumnWidth(
                targetSheet,
                columnSpan.first(),
                columnSpan.last(),
                data.consumeRegularDouble(1.0d, 20.0d));
        case 0x3 ->
            new WorkbookCommand.SetRowHeight(
                targetSheet,
                rowSpan.first(),
                rowSpan.last(),
                data.consumeRegularDouble(5.0d, 40.0d));
        case 0x4 -> new WorkbookCommand.SetSheetPane(targetSheet, nextExcelSheetPane(data));
        case 0x5 -> new WorkbookCommand.SetSheetZoom(targetSheet, data.consumeInt(10, 400));
        case 0x6 ->
            new WorkbookCommand.SetPrintLayout(targetSheet, nextExcelPrintLayout(data));
        default -> new WorkbookCommand.ClearPrintLayout(targetSheet);
      };
      case 0x2 -> switch (selectorSlot(selector)) {
        case 0x0 ->
            new WorkbookCommand.SetCell(
                targetSheet,
                FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
                FuzzDataDecoders.nextExcelCellValue(data));
        case 0x1 -> {
          String range = validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false);
          yield new WorkbookCommand.SetRange(
              targetSheet, range, FuzzDataDecoders.nextExcelMatrix(data, 2, 2));
        }
        case 0x2 ->
            new WorkbookCommand.ClearRange(
                targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
        default ->
            new WorkbookCommand.AppendRow(
                targetSheet,
                List.of(
                    FuzzDataDecoders.nextExcelCellValue(data),
                    FuzzDataDecoders.nextExcelCellValue(data)));
      };
      case 0x3 -> switch (selectorSlot(selector)) {
        case 0x0 ->
            new WorkbookCommand.SetHyperlink(
                targetSheet,
                FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
                nextExcelHyperlink(data));
        case 0x1 ->
            new WorkbookCommand.ClearHyperlink(
                targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress));
        case 0x2 ->
            new WorkbookCommand.SetComment(
                targetSheet,
                FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
                nextExcelComment(data));
        case 0x3 ->
            new WorkbookCommand.ClearComment(
                targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress));
        default ->
            new WorkbookCommand.ApplyStyle(
                targetSheet,
                validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false),
                FuzzDataDecoders.nextStyle(data));
      };
      case 0x4 -> switch (selectorSlot(selector)) {
        case 0x0 ->
            new WorkbookCommand.SetDataValidation(
                targetSheet,
                validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false),
                nextExcelDataValidationDefinition(data));
        case 0x1 ->
            new WorkbookCommand.ClearDataValidations(
                targetSheet, nextExcelRangeSelection(data, validRange));
        case 0x2 ->
            new WorkbookCommand.SetConditionalFormatting(
                targetSheet, nextExcelConditionalFormattingBlockDefinition(data, validRange));
        default ->
            new WorkbookCommand.ClearConditionalFormatting(
                targetSheet, nextExcelRangeSelection(data, validRange));
      };
      case 0x5 -> switch (selectorSlot(selector)) {
        case 0x0 -> new WorkbookCommand.SetAutofilter(targetSheet, nextAutofilterRange(validRange));
        case 0x1 -> new WorkbookCommand.ClearAutofilter(targetSheet);
        case 0x2 ->
            new WorkbookCommand.SetTable(
                nextExcelTableDefinition(data, targetSheet, tableName, validRange));
        default -> new WorkbookCommand.DeleteTable(tableName, targetSheet);
      };
      case 0x6 -> switch (selectorSlot(selector)) {
        case 0x0 ->
            new WorkbookCommand.SetNamedRange(
                new ExcelNamedRangeDefinition(
                    validName ? namedRangeName : nextNamedRangeName(data, false),
                    data.consumeBoolean()
                        ? new ExcelNamedRangeScope.WorkbookScope()
                        : new ExcelNamedRangeScope.SheetScope(targetSheet),
                    new ExcelNamedRangeTarget(
                        targetSheet,
                        validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false))));
        default ->
            new WorkbookCommand.DeleteNamedRange(
                validName ? namedRangeName : nextNamedRangeName(data, false),
                data.consumeBoolean()
                    ? new ExcelNamedRangeScope.WorkbookScope()
                    : new ExcelNamedRangeScope.SheetScope(targetSheet));
      };
      default -> switch (selectorSlot(selector)) {
        case 0x0 -> new WorkbookCommand.AutoSizeColumns(targetSheet);
        case 0x1 -> new WorkbookCommand.EvaluateAllFormulas();
        default -> new WorkbookCommand.ForceFormulaRecalculationOnOpen();
      };
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
    int selector = nextSelectorByte(data);

    return switch (selectorFamily(selector)) {
      case 0x0 -> switch (selectorSlot(selector)) {
        case 0x0 -> new WorkbookReadOperation.GetWorkbookSummary(requestId);
        case 0x1 -> new WorkbookReadOperation.GetSheetSummary(requestId, targetSheet);
        case 0x2 ->
            new WorkbookReadOperation.GetNamedRanges(
                requestId,
                nextNamedRangeSelection(
                    data, targetSheet, workbookNamedRange, sheetNamedRange, validName));
        default ->
            new WorkbookReadOperation.GetNamedRangeSurface(
                requestId,
                nextNamedRangeSelection(
                    data, targetSheet, workbookNamedRange, sheetNamedRange, validName));
      };
      case 0x1 -> switch (selectorSlot(selector)) {
        case 0x0 ->
            new WorkbookReadOperation.GetCells(
                requestId, targetSheet, nextReadAddresses(data, validAddress));
        case 0x1 ->
            new WorkbookReadOperation.GetWindow(
                requestId,
                targetSheet,
                FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
                data.consumeInt(1, 4),
                data.consumeInt(1, 4));
        case 0x2 -> new WorkbookReadOperation.GetSheetLayout(requestId, targetSheet);
        default ->
            new WorkbookReadOperation.GetSheetSchema(
                requestId,
                targetSheet,
                validAddress ? "A1" : FuzzDataDecoders.nextNonBlankCellAddress(data, false),
                data.consumeInt(1, 5),
                data.consumeInt(1, 4));
      };
      case 0x2 -> switch (selectorSlot(selector)) {
        case 0x0 -> new WorkbookReadOperation.GetMergedRegions(requestId, targetSheet);
        case 0x1 ->
            new WorkbookReadOperation.GetHyperlinks(
                requestId, targetSheet, nextCellSelection(data, validAddress));
        case 0x2 ->
            new WorkbookReadOperation.GetComments(
                requestId, targetSheet, nextCellSelection(data, validAddress));
        case 0x3 -> new WorkbookReadOperation.GetSheetLayout(requestId, targetSheet);
        case 0x4 -> new WorkbookReadOperation.GetPrintLayout(requestId, targetSheet);
        default ->
            new WorkbookReadOperation.GetFormulaSurface(
                requestId, nextSheetSelection(data, primarySheet, secondarySheet));
      };
      case 0x3 -> switch (selectorSlot(selector)) {
        case 0x0 ->
            new WorkbookReadOperation.GetDataValidations(
                requestId, targetSheet, nextRangeSelection(data, validRange));
        case 0x1 ->
            new WorkbookReadOperation.GetConditionalFormatting(
                requestId, targetSheet, nextRangeSelection(data, validRange));
        default -> new WorkbookReadOperation.GetAutofilters(requestId, targetSheet);
      };
      case 0x4 -> switch (selectorSlot(selector)) {
        case 0x0 ->
            new WorkbookReadOperation.GetTables(
                requestId, nextTableSelection(data, primarySheet, secondarySheet));
        case 0x1 ->
            new WorkbookReadOperation.AnalyzeTableHealth(
                requestId, nextTableSelection(data, primarySheet, secondarySheet));
        default ->
            new WorkbookReadOperation.AnalyzeAutofilterHealth(
                requestId, nextSheetSelection(data, primarySheet, secondarySheet));
      };
      case 0x5 -> switch (selectorSlot(selector)) {
        case 0x0 ->
            new WorkbookReadOperation.AnalyzeFormulaHealth(
                requestId, nextSheetSelection(data, primarySheet, secondarySheet));
        case 0x1 ->
            new WorkbookReadOperation.AnalyzeDataValidationHealth(
                requestId, nextSheetSelection(data, primarySheet, secondarySheet));
        case 0x2 ->
            new WorkbookReadOperation.AnalyzeConditionalFormattingHealth(
                requestId, nextSheetSelection(data, primarySheet, secondarySheet));
        default ->
            new WorkbookReadOperation.AnalyzeHyperlinkHealth(
                requestId, nextSheetSelection(data, primarySheet, secondarySheet));
      };
      default -> switch (selectorSlot(selector)) {
        case 0x0 ->
            new WorkbookReadOperation.AnalyzeNamedRangeHealth(
                requestId,
                nextNamedRangeSelection(
                    data, targetSheet, workbookNamedRange, sheetNamedRange, validName));
        default -> new WorkbookReadOperation.AnalyzeWorkbookFindings(requestId);
      };
    };
  }

  private static NamedRangeSelection nextNamedRangeSelection(
      FuzzedDataProvider data,
      String sheetName,
      String workbookNamedRange,
      String sheetNamedRange,
      boolean validName) {
    return switch (selectorSlot(nextSelectorByte(data))) {
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
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new CellSelection.AllUsedCells();
      default ->
          new CellSelection.Selected(
              nextReadAddresses(data, validAddress));
    };
  }

  private static SheetSelection nextSheetSelection(
      FuzzedDataProvider data, String primarySheet, String secondarySheet) {
    return switch (selectorSlot(nextSelectorByte(data))) {
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

  private static PaneInput nextPaneInput(FuzzedDataProvider data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new PaneInput.None();
      case 1 -> {
        int splitColumn = data.consumeInt(1, 3);
        yield new PaneInput.Frozen(splitColumn, 0, data.consumeInt(splitColumn, splitColumn + 2), 0);
      }
      case 2 -> {
        int splitRow = data.consumeInt(1, 3);
        yield new PaneInput.Frozen(0, splitRow, 0, data.consumeInt(splitRow, splitRow + 2));
      }
      case 3 -> {
        int splitColumn = data.consumeInt(1, 3);
        int splitRow = data.consumeInt(1, 3);
        yield new PaneInput.Frozen(
            splitColumn,
            splitRow,
            data.consumeInt(splitColumn, splitColumn + 2),
            data.consumeInt(splitRow, splitRow + 2));
      }
      default ->
          new PaneInput.Split(
              data.consumeInt(0, 2400),
              data.consumeInt(1, 2400),
              0,
              data.consumeInt(1, 4),
              nextProtocolPaneRegion(data));
    };
  }

  private static ExcelSheetPane nextExcelSheetPane(FuzzedDataProvider data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new ExcelSheetPane.None();
      case 1 -> {
        int splitColumn = data.consumeInt(1, 3);
        yield new ExcelSheetPane.Frozen(
            splitColumn, 0, data.consumeInt(splitColumn, splitColumn + 2), 0);
      }
      case 2 -> {
        int splitRow = data.consumeInt(1, 3);
        yield new ExcelSheetPane.Frozen(0, splitRow, 0, data.consumeInt(splitRow, splitRow + 2));
      }
      case 3 -> {
        int splitColumn = data.consumeInt(1, 3);
        int splitRow = data.consumeInt(1, 3);
        yield new ExcelSheetPane.Frozen(
            splitColumn,
            splitRow,
            data.consumeInt(splitColumn, splitColumn + 2),
            data.consumeInt(splitRow, splitRow + 2));
      }
      default ->
          new ExcelSheetPane.Split(
              data.consumeInt(0, 2400),
              data.consumeInt(1, 2400),
              0,
              data.consumeInt(1, 4),
              nextPaneRegion(data));
    };
  }

  private static PrintLayoutInput nextPrintLayoutInput(FuzzedDataProvider data) {
    return new PrintLayoutInput(
        data.consumeBoolean() ? new PrintAreaInput.Range("A1:C20") : new PrintAreaInput.None(),
        data.consumeBoolean() ? PrintOrientation.LANDSCAPE : PrintOrientation.PORTRAIT,
        data.consumeBoolean()
            ? new PrintScalingInput.Fit(data.consumeInt(0, 2), data.consumeInt(0, 2))
            : new PrintScalingInput.Automatic(),
        data.consumeBoolean()
            ? new PrintTitleRowsInput.Band(0, data.consumeInt(0, 2))
            : new PrintTitleRowsInput.None(),
        data.consumeBoolean()
            ? new PrintTitleColumnsInput.Band(0, data.consumeInt(0, 2))
            : new PrintTitleColumnsInput.None(),
        new HeaderFooterTextInput("L" + data.consumeInt(0, 9), "", "R" + data.consumeInt(0, 9)),
        new HeaderFooterTextInput("", "P" + data.consumeInt(0, 9), ""));
  }

  private static ExcelPrintLayout nextExcelPrintLayout(FuzzedDataProvider data) {
    return new ExcelPrintLayout(
        data.consumeBoolean()
            ? new ExcelPrintLayout.Area.Range("A1:C20")
            : new ExcelPrintLayout.Area.None(),
        data.consumeBoolean() ? ExcelPrintOrientation.LANDSCAPE : ExcelPrintOrientation.PORTRAIT,
        data.consumeBoolean()
            ? new ExcelPrintLayout.Scaling.Fit(data.consumeInt(0, 2), data.consumeInt(0, 2))
            : new ExcelPrintLayout.Scaling.Automatic(),
        data.consumeBoolean()
            ? new ExcelPrintLayout.TitleRows.Band(0, data.consumeInt(0, 2))
            : new ExcelPrintLayout.TitleRows.None(),
        data.consumeBoolean()
            ? new ExcelPrintLayout.TitleColumns.Band(0, data.consumeInt(0, 2))
            : new ExcelPrintLayout.TitleColumns.None(),
        new ExcelHeaderFooterText("L" + data.consumeInt(0, 9), "", "R" + data.consumeInt(0, 9)),
        new ExcelHeaderFooterText("", "P" + data.consumeInt(0, 9), ""));
  }

  private static PaneRegion nextProtocolPaneRegion(FuzzedDataProvider data) {
    return switch (selectorSlot(nextSelectorByte(data)) & 0x3) {
      case 0 -> PaneRegion.UPPER_LEFT;
      case 1 -> PaneRegion.UPPER_RIGHT;
      case 2 -> PaneRegion.LOWER_LEFT;
      default -> PaneRegion.LOWER_RIGHT;
    };
  }

  private static ExcelPaneRegion nextPaneRegion(FuzzedDataProvider data) {
    return switch (selectorSlot(nextSelectorByte(data)) & 0x3) {
      case 0 -> ExcelPaneRegion.UPPER_LEFT;
      case 1 -> ExcelPaneRegion.UPPER_RIGHT;
      case 2 -> ExcelPaneRegion.LOWER_LEFT;
      default -> ExcelPaneRegion.LOWER_RIGHT;
    };
  }

  private static SheetCopyPosition nextSheetCopyPosition(FuzzedDataProvider data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new SheetCopyPosition.AppendAtEnd();
      default -> new SheetCopyPosition.AtIndex(data.consumeInt(0, 2));
    };
  }

  private static ExcelSheetCopyPosition nextExcelSheetCopyPosition(FuzzedDataProvider data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new ExcelSheetCopyPosition.AppendAtEnd();
      default -> new ExcelSheetCopyPosition.AtIndex(data.consumeInt(0, 2));
    };
  }

  private static List<String> nextSelectedSheetNames(
      FuzzedDataProvider data, String primarySheet, String secondarySheet) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> List.of(primarySheet);
      case 1 -> List.of(secondarySheet);
      default -> List.of(primarySheet, secondarySheet);
    };
  }

  private static SheetVisibility nextProtocolSheetVisibility(FuzzedDataProvider data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> SheetVisibility.VISIBLE;
      case 1 -> SheetVisibility.HIDDEN;
      default -> SheetVisibility.VERY_HIDDEN;
    };
  }

  private static ExcelSheetVisibility nextSheetVisibility(FuzzedDataProvider data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> ExcelSheetVisibility.VISIBLE;
      case 1 -> ExcelSheetVisibility.HIDDEN;
      default -> ExcelSheetVisibility.VERY_HIDDEN;
    };
  }

  private static SheetProtectionSettings nextProtocolSheetProtectionSettings(
      FuzzedDataProvider data) {
    return new SheetProtectionSettings(
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean());
  }

  private static ExcelSheetProtectionSettings nextSheetProtectionSettings(
      FuzzedDataProvider data) {
    return new ExcelSheetProtectionSettings(
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean());
  }

  private static WorkflowStorage nextWorkflowStorage(
      String primarySheet, String secondarySheet, FuzzedDataProvider data) throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-jazzer-workflow-");
    Path sourcePath = directory.resolve("source.xlsx");
    Path saveAsPath = directory.resolve("output.xlsx");

    return switch (selectorSlot(nextSelectorByte(data))) {
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
      primary.createRow(0).createCell(0).setCellValue("HeaderA");
      primary.getRow(0).createCell(1).setCellValue("HeaderB");
      primary.createRow(1).createCell(0).setCellValue("seed");
      primary.getRow(1).createCell(1).setCellValue(2.0d);
      primary.createRow(4).createCell(4).setCellValue("Queue");
      primary.getRow(4).createCell(5).setCellValue("Owner");
      primary.createRow(5).createCell(4).setCellValue("seed");
      primary.getRow(5).createCell(5).setCellValue("GridGrind");
      if (data.consumeBoolean()) {
        var secondary = workbook.createSheet(secondarySheet);
        secondary.createRow(0).createCell(0).setCellValue("HeaderA");
        secondary.getRow(0).createCell(1).setCellValue("HeaderB");
        secondary.createRow(1).createCell(0).setCellValue("seed");
        secondary.getRow(1).createCell(1).setCellValue(3.0d);
      }
      Files.createDirectories(sourcePath.getParent());
      try (OutputStream outputStream = Files.newOutputStream(sourcePath)) {
        workbook.write(outputStream);
      }
    }
  }

  private static HyperlinkTarget nextHyperlinkTarget(FuzzedDataProvider data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new HyperlinkTarget.Url("https://example.com/" + nextNamedRangeName(data, true));
      case 1 -> new HyperlinkTarget.Email(nextNamedRangeName(data, true) + "@example.com");
      case 2 -> new HyperlinkTarget.File("/tmp/" + nextNamedRangeName(data, true) + ".xlsx");
      default -> new HyperlinkTarget.Document("Budget!A1");
    };
  }

  private static ExcelHyperlink nextExcelHyperlink(FuzzedDataProvider data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
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

  private static DataValidationInput nextDataValidationInput(FuzzedDataProvider data) {
    return new DataValidationInput(
        data.consumeBoolean()
            ? new DataValidationRuleInput.ExplicitList(List.of("Queued", "Done"))
            : new DataValidationRuleInput.WholeNumber(
                ComparisonOperator.GREATER_OR_EQUAL, "1", null),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean()
            ? new DataValidationPromptInput("Status", "Use an allowed value.", data.consumeBoolean())
            : null,
        data.consumeBoolean()
            ? new DataValidationErrorAlertInput(
                DataValidationErrorStyle.STOP,
                "Invalid",
                "Use an allowed value.",
                data.consumeBoolean())
            : null);
  }

  private static ExcelDataValidationDefinition nextExcelDataValidationDefinition(
      FuzzedDataProvider data) {
    return new ExcelDataValidationDefinition(
        data.consumeBoolean()
            ? new ExcelDataValidationRule.ExplicitList(List.of("Queued", "Done"))
            : new ExcelDataValidationRule.WholeNumber(
                ExcelComparisonOperator.GREATER_OR_EQUAL, "1", null),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean()
            ? new ExcelDataValidationPrompt("Status", "Use an allowed value.", data.consumeBoolean())
            : null,
        data.consumeBoolean()
            ? new ExcelDataValidationErrorAlert(
                ExcelDataValidationErrorStyle.STOP,
                "Invalid",
                "Use an allowed value.",
                data.consumeBoolean())
            : null);
  }

  private static ConditionalFormattingBlockInput nextConditionalFormattingInput(
      FuzzedDataProvider data, boolean validRange) {
    return new ConditionalFormattingBlockInput(
        List.of(validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)),
        List.of(
            data.consumeBoolean()
                ? new ConditionalFormattingRuleInput.FormulaRule(
                    "A1>0", data.consumeBoolean(), nextDifferentialStyleInput(data))
                : new ConditionalFormattingRuleInput.CellValueRule(
                    ComparisonOperator.GREATER_THAN,
                    "1",
                    null,
                    data.consumeBoolean(),
                    nextDifferentialStyleInput(data))));
  }

  private static ExcelConditionalFormattingBlockDefinition
      nextExcelConditionalFormattingBlockDefinition(FuzzedDataProvider data, boolean validRange) {
    return new ExcelConditionalFormattingBlockDefinition(
        List.of(validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)),
        List.of(
            data.consumeBoolean()
                ? new ExcelConditionalFormattingRule.FormulaRule(
                    "A1>0", data.consumeBoolean(), nextExcelDifferentialStyle(data))
                : new ExcelConditionalFormattingRule.CellValueRule(
                    ExcelComparisonOperator.GREATER_THAN,
                    "1",
                    null,
                    data.consumeBoolean(),
                    nextExcelDifferentialStyle(data))));
  }

  private static DifferentialStyleInput nextDifferentialStyleInput(FuzzedDataProvider data) {
    boolean includeNumberFormat = data.consumeBoolean();
    Boolean bold = data.consumeBoolean() ? Boolean.TRUE : null;
    Boolean italic = data.consumeBoolean() ? Boolean.TRUE : null;
    String fontColor = data.consumeBoolean() ? "#102030" : null;
    Boolean underline = data.consumeBoolean() ? Boolean.TRUE : null;
    Boolean strikeout = data.consumeBoolean() ? Boolean.TRUE : null;
    String fillColor = data.consumeBoolean() ? "#E0F0AA" : null;
    String numberFormat =
        includeNumberFormat
                || Stream.of(bold, italic, fontColor, underline, strikeout, fillColor)
                    .allMatch(Objects::isNull)
            ? "0.00"
            : null;
    return new DifferentialStyleInput(
        numberFormat,
        bold,
        italic,
        null,
        fontColor,
        underline,
        strikeout,
        fillColor,
        null);
  }

  private static ExcelDifferentialStyle nextExcelDifferentialStyle(FuzzedDataProvider data) {
    boolean includeNumberFormat = data.consumeBoolean();
    Boolean bold = data.consumeBoolean() ? Boolean.TRUE : null;
    Boolean italic = data.consumeBoolean() ? Boolean.TRUE : null;
    String fontColor = data.consumeBoolean() ? "#102030" : null;
    Boolean underline = data.consumeBoolean() ? Boolean.TRUE : null;
    Boolean strikeout = data.consumeBoolean() ? Boolean.TRUE : null;
    String fillColor = data.consumeBoolean() ? "#E0F0AA" : null;
    String numberFormat =
        includeNumberFormat
                || Stream.of(bold, italic, fontColor, underline, strikeout, fillColor)
                    .allMatch(Objects::isNull)
            ? "0.00"
            : null;
    return new ExcelDifferentialStyle(
        numberFormat,
        bold,
        italic,
        null,
        fontColor,
        underline,
        strikeout,
        fillColor,
        null);
  }

  private static RangeSelection nextRangeSelection(FuzzedDataProvider data, boolean validRange) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new RangeSelection.All();
      default ->
          new RangeSelection.Selected(
              List.of(validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)));
    };
  }

  private static ExcelRangeSelection nextExcelRangeSelection(
      FuzzedDataProvider data, boolean validRange) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new ExcelRangeSelection.All();
      default ->
          new ExcelRangeSelection.Selected(
              List.of(validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)));
    };
  }

  private static String nextAutofilterRange(boolean validRange) {
    return validRange ? "E1:F3" : "BadRange";
  }

  private static String nextCopySheetName(String sourceSheetName) {
    Objects.requireNonNull(sourceSheetName, "sourceSheetName must not be null");
    String base = sourceSheetName.length() <= 27 ? sourceSheetName : sourceSheetName.substring(0, 27);
    return base + "_B1";
  }

  private static TableInput nextTableInput(
      FuzzedDataProvider data, String sheetName, String tableName, boolean validRange) {
    return new TableInput(
        tableName,
        sheetName,
        validRange ? "A1:B3" : FuzzDataDecoders.nextNonBlankRange(data, false),
        data.consumeBoolean(),
        nextTableStyleInput(data));
  }

  private static ExcelTableDefinition nextExcelTableDefinition(
      FuzzedDataProvider data, String sheetName, String tableName, boolean validRange) {
    return new ExcelTableDefinition(
        tableName,
        sheetName,
        validRange ? "A1:B3" : FuzzDataDecoders.nextNonBlankRange(data, false),
        data.consumeBoolean(),
        nextExcelTableStyle(data));
  }

  private static TableStyleInput nextTableStyleInput(FuzzedDataProvider data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new TableStyleInput.None();
      default ->
          new TableStyleInput.Named(
              data.consumeBoolean() ? "TableStyleMedium2" : "TableStyleLight9",
              data.consumeBoolean(),
              data.consumeBoolean(),
              data.consumeBoolean(),
              data.consumeBoolean());
    };
  }

  private static ExcelTableStyle nextExcelTableStyle(FuzzedDataProvider data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new ExcelTableStyle.None();
      default ->
          new ExcelTableStyle.Named(
              data.consumeBoolean() ? "TableStyleMedium2" : "TableStyleLight9",
              data.consumeBoolean(),
              data.consumeBoolean(),
              data.consumeBoolean(),
              data.consumeBoolean());
    };
  }

  private static TableSelection nextTableSelection(
      FuzzedDataProvider data, String primarySheet, String secondarySheet) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new TableSelection.All();
      default ->
          new TableSelection.ByNames(
              List.of(
                  data.consumeBoolean()
                      ? nextTableName(data, true, primarySheet)
                      : nextTableName(data, true, secondarySheet)));
    };
  }

  private static String nextTableName(FuzzedDataProvider data, boolean valid, String sheetName) {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    if (!valid) {
      return nextNamedRangeName(data, false);
    }
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> sheetName + "Table";
      case 1 -> "BudgetTable";
      default -> "OpsTable";
    };
  }

  private static ExcelComment nextExcelComment(FuzzedDataProvider data) {
    return new ExcelComment("Note " + nextNamedRangeName(data, true), "GridGrind", data.consumeBoolean());
  }

  private static String nextNamedRangeName(FuzzedDataProvider data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");
    if (!valid) {
      return switch (selectorSlot(nextSelectorByte(data))) {
        case 0 -> "";
        case 1 -> "A1";
        case 2 -> "R1C1";
        case 3 -> "_xlnm.Print_Area";
        default -> "1Budget";
      };
    }
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> "BudgetTotal";
      case 1 -> "LocalItem";
      case 2 -> "Report_Value";
      case 3 -> "Summary.Total";
      default -> "Name" + data.consumeInt(1, 9);
    };
  }

  private static int nextSelectorByte(FuzzedDataProvider data) {
    return Byte.toUnsignedInt(data.consumeByte());
  }

  private static int selectorFamily(int selector) {
    return selector >>> 4;
  }

  private static int selectorSlot(int selector) {
    return selector & 0x0F;
  }

  private record IndexSpan(int first, int last) {}

  private record WorkflowStorage(
      GridGrindRequest.WorkbookSource source,
      GridGrindRequest.WorkbookPersistence persistence,
      Path cleanupRoot) {}
}
