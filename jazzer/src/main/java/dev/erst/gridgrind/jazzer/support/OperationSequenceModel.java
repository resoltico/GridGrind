package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.*;
import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.assertThat;
import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.inspect;
import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.mutate;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.AnalysisFindingCode;
import dev.erst.gridgrind.contract.dto.AnalysisSeverity;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelSheetPresentation;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSettings;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.WorkflowStorage;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Builds bounded protocol requests and workbook command sequences for Jazzer harnesses. */
public final class OperationSequenceModel {
  private OperationSequenceModel() {}

  /** Returns a bounded protocol workflow plus any owned local scratch paths it created. */
  public static GeneratedProtocolWorkflow nextProtocolWorkflow(GridGrindFuzzData data)
      throws IOException {
    Objects.requireNonNull(data, "data must not be null");

    String primarySheet = FuzzDataDecoders.nextSheetName(data, true);
    String secondarySheet = FuzzDataDecoders.nextSheetName(data, true);
    List<ProtocolStepSupport.PendingMutation> mutations = new ArrayList<>();
    List<ProtocolStepSupport.PendingAssertion> assertions = new ArrayList<>();
    List<InspectionStep> inspections = new ArrayList<>();
    String workbookNamedRange = nextNamedRangeName(data, true);
    String sheetNamedRange = nextNamedRangeName(data, true);
    String pivotTableName = nextPivotTableName(data, true);

    int operationCount = data.consumeInt(0, 8);
    for (int index = 0; index < operationCount; index++) {
      mutations.add(
          nextMutation(
              data,
              primarySheet,
              secondarySheet,
              workbookNamedRange,
              sheetNamedRange,
              pivotTableName));
    }
    int readCount = data.consumeInt(0, 6);
    for (int index = 0; index < readCount; index++) {
      inspections.add(
          nextInspection(
              data,
              index,
              primarySheet,
              secondarySheet,
              workbookNamedRange,
              sheetNamedRange,
              pivotTableName));
    }
    int assertionCount = data.consumeInt(0, 4);
    for (int index = 0; index < assertionCount; index++) {
      assertions.add(
          nextAssertion(
              data,
              index,
              primarySheet,
              secondarySheet,
              workbookNamedRange,
              sheetNamedRange,
              pivotTableName));
    }
    WorkflowStorage workflowStorage = nextWorkflowStorage(primarySheet, secondarySheet, data);
    ExecutionPolicyInput execution =
        nextExecutionPolicy(data, primarySheet, secondarySheet, validFormulaAddress(data));
    return new GeneratedProtocolWorkflow(
        ProtocolStepSupport.request(
            workflowStorage.source(),
            workflowStorage.persistence(),
            execution,
            null,
            mutations,
            assertions,
            inspections),
        List.of(workflowStorage.cleanupRoot()));
  }

  /** Returns a bounded sequence of workbook-core commands. */
  public static List<WorkbookCommand> nextWorkbookCommands(GridGrindFuzzData data) {
    Objects.requireNonNull(data, "data must not be null");

    String primarySheet = FuzzDataDecoders.nextSheetName(data, true);
    String secondarySheet = FuzzDataDecoders.nextSheetName(data, true);
    List<WorkbookCommand> commands = new ArrayList<>();
    String workbookNamedRange = nextNamedRangeName(data, true);
    String sheetNamedRange = nextNamedRangeName(data, true);
    String pivotTableName = nextPivotTableName(data, true);
    int commandCount = data.consumeInt(1, 10);
    for (int index = 0; index < commandCount; index++) {
      commands.add(
          nextCommand(
              data,
              primarySheet,
              secondarySheet,
              workbookNamedRange,
              sheetNamedRange,
              pivotTableName));
    }
    return List.copyOf(commands);
  }

  private static ProtocolStepSupport.PendingMutation nextMutation(
      GridGrindFuzzData data,
      String primarySheet,
      String secondarySheet,
      String workbookNamedRange,
      String sheetNamedRange,
      String pivotTableName) {
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
      case 0x0 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(new SheetSelector.ByName(targetSheet), new MutationAction.EnsureSheet());
            case 0x1 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.RenameSheet(primarySheet + "Renamed"));
            case 0x2 ->
                mutate(new SheetSelector.ByName(targetSheet), new MutationAction.DeleteSheet());
            case 0x3 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.MoveSheet(data.consumeInt(0, 2)));
            case 0x4 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.CopySheet(
                        nextCopySheetName(targetSheet), nextSheetCopyPosition(data)));
            case 0x5 ->
                mutate(new SheetSelector.ByName(targetSheet), new MutationAction.SetActiveSheet());
            case 0x6 ->
                mutate(
                    new SheetSelector.ByNames(
                        nextSelectedSheetNames(data, primarySheet, secondarySheet)),
                    new MutationAction.SetSelectedSheets());
            case 0x7 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetSheetVisibility(nextProtocolSheetVisibility(data)));
            case 0x8 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetSheetProtection(
                        nextProtocolSheetProtectionSettings(data)));
            default ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.ClearSheetProtection());
          };
      case 0x1 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new RangeSelector.ByRange(
                        targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange)),
                    new MutationAction.MergeCells());
            case 0x1 ->
                mutate(
                    new RangeSelector.ByRange(
                        targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange)),
                    new MutationAction.UnmergeCells());
            case 0x2 ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new MutationAction.SetColumnWidth(data.consumeRegularDouble(1.0d, 20.0d)));
            case 0x3 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new MutationAction.SetRowHeight(data.consumeRegularDouble(5.0d, 40.0d)));
            case 0x4 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetSheetPane(nextPaneInput(data)));
            case 0x5 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetSheetZoom(data.consumeInt(10, 400)));
            case 0x6 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetPrintLayout(nextPrintLayoutInput(data)));
            default ->
                mutate(
                    new SheetSelector.ByName(targetSheet), new MutationAction.ClearPrintLayout());
          };
      case 0x2 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new CellSelector.ByAddress(
                        targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress)),
                    new MutationAction.SetCell(FuzzDataDecoders.nextCellInput(data)));
            case 0x1 -> {
              String range = validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false);
              yield mutate(
                  new RangeSelector.ByRange(targetSheet, range),
                  new MutationAction.SetRange(FuzzDataDecoders.nextProtocolMatrix(data, 2, 2)));
            }
            case 0x2 ->
                mutate(
                    new RangeSelector.ByRange(
                        targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange)),
                    new MutationAction.ClearRange());
            default ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.AppendRow(
                        List.of(
                            FuzzDataDecoders.nextCellInput(data),
                            FuzzDataDecoders.nextCellInput(data))));
          };
      case 0x3 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new CellSelector.ByAddress(
                        targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress)),
                    new MutationAction.SetHyperlink(nextHyperlinkTarget(data)));
            case 0x1 ->
                mutate(
                    new CellSelector.ByAddress(
                        targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress)),
                    new MutationAction.ClearHyperlink());
            case 0x2 ->
                mutate(
                    new CellSelector.ByAddress(
                        targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress)),
                    new MutationAction.SetComment(nextCommentInput(data)));
            case 0x3 ->
                mutate(
                    new CellSelector.ByAddress(
                        targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress)),
                    new MutationAction.ClearComment());
            case 0x4 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetPicture(nextPictureInput(data)));
            case 0x5 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetChart(nextChartInput(data)));
            case 0x6 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetShape(nextShapeInput(data)));
            case 0x7 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetEmbeddedObject(nextEmbeddedObjectInput(data)));
            case 0x8 ->
                mutate(
                    new DrawingObjectSelector.ByName(targetSheet, nextDrawingObjectName(data)),
                    new MutationAction.SetDrawingObjectAnchor(nextDrawingAnchorInput(data)));
            case 0x9 ->
                mutate(
                    new DrawingObjectSelector.ByName(targetSheet, nextDrawingObjectName(data)),
                    new MutationAction.DeleteDrawingObject());
            default ->
                mutate(
                    new RangeSelector.ByRange(
                        targetSheet,
                        validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false)),
                    new MutationAction.ApplyStyle(WorkbookStyleInputs.nextStyleInput(data)));
          };
      case 0x4 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new RangeSelector.ByRange(
                        targetSheet,
                        validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)),
                    new MutationAction.SetDataValidation(nextDataValidationInput(data)));
            case 0x1 ->
                mutate(
                    nextRangeSelector(data, targetSheet, validRange),
                    new MutationAction.ClearDataValidations());
            case 0x2 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetConditionalFormatting(
                        nextConditionalFormattingInput(data, validRange)));
            default ->
                mutate(
                    nextRangeSelector(data, targetSheet, validRange),
                    new MutationAction.ClearConditionalFormatting());
          };
      case 0x5 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new RangeSelector.ByRange(targetSheet, nextAutofilterRange(validRange)),
                    new MutationAction.SetAutofilter());
            case 0x1 ->
                mutate(new SheetSelector.ByName(targetSheet), new MutationAction.ClearAutofilter());
            case 0x2 ->
                mutate(
                    new MutationAction.SetTable(
                        nextTableInput(data, targetSheet, tableName, validRange)));
            case 0x3 ->
                mutate(
                    new TableSelector.ByNameOnSheet(tableName, targetSheet),
                    new MutationAction.DeleteTable());
            case 0x4 ->
                mutate(
                    new MutationAction.SetPivotTable(
                        nextPivotTableInput(
                            data,
                            targetSheet,
                            pivotTableName,
                            workbookNamedRange,
                            tableName,
                            validName,
                            validRange)));
            default ->
                mutate(
                    new PivotTableSelector.ByNameOnSheet(
                        validName ? pivotTableName : nextPivotTableName(data, false), targetSheet),
                    new MutationAction.DeletePivotTable());
          };
      case 0x6 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new MutationAction.SetNamedRange(
                        validName ? namedRangeName : nextNamedRangeName(data, false),
                        data.consumeBoolean()
                            ? new NamedRangeScope.Workbook()
                            : new NamedRangeScope.Sheet(targetSheet),
                        new NamedRangeTarget(
                            targetSheet,
                            validRange
                                ? "A1:B2"
                                : FuzzDataDecoders.nextNonBlankRange(data, false))));
            default ->
                mutate(
                    data.consumeBoolean()
                        ? new NamedRangeSelector.WorkbookScope(
                            validName ? namedRangeName : nextNamedRangeName(data, false))
                        : new NamedRangeSelector.SheetScope(
                            validName ? namedRangeName : nextNamedRangeName(data, false),
                            targetSheet),
                    new MutationAction.DeleteNamedRange());
          };
      case 0x7 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new RowBandSelector.Insertion(
                        targetSheet, rowSpan.first(), rowSpan.last() - rowSpan.first() + 1),
                    new MutationAction.InsertRows());
            case 0x1 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new MutationAction.DeleteRows());
            case 0x2 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new MutationAction.ShiftRows(nextNonZeroDelta(data, 2)));
            case 0x3 ->
                mutate(
                    new ColumnBandSelector.Insertion(
                        targetSheet,
                        columnSpan.first(),
                        columnSpan.last() - columnSpan.first() + 1),
                    new MutationAction.InsertColumns());
            case 0x4 ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new MutationAction.DeleteColumns());
            case 0x5 ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new MutationAction.ShiftColumns(nextNonZeroDelta(data, 2)));
            case 0x6 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new MutationAction.SetRowVisibility(data.consumeBoolean()));
            case 0x7 ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new MutationAction.SetColumnVisibility(data.consumeBoolean()));
            case 0x8 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new MutationAction.GroupRows(data.consumeBoolean()));
            case 0x9 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new MutationAction.UngroupRows());
            case 0xA ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new MutationAction.GroupColumns(data.consumeBoolean()));
            default ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new MutationAction.UngroupColumns());
          };
      default ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(new SheetSelector.ByName(targetSheet), new MutationAction.AutoSizeColumns());
            case 0x1 ->
                mutate(
                    new WorkbookSelector.Current(),
                    new MutationAction.SetWorkbookProtection(
                        new WorkbookProtectionInput(true, false, false, null, null)));
            case 0x2 ->
                mutate(
                    new WorkbookSelector.Current(), new MutationAction.ClearWorkbookProtection());
            case 0x3 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetSheetPresentation(SheetPresentationInput.defaults()));
            default ->
                mutate(new SheetSelector.ByName(targetSheet), new MutationAction.AutoSizeColumns());
          };
    };
  }

  private static WorkbookCommand nextCommand(
      GridGrindFuzzData data,
      String primarySheet,
      String secondarySheet,
      String workbookNamedRange,
      String sheetNamedRange,
      String pivotTableName) {
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
      case 0x0 ->
          switch (selectorSlot(selector)) {
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
            case 0x7 ->
                new WorkbookCommand.SetSheetVisibility(targetSheet, nextSheetVisibility(data));
            case 0x8 ->
                new WorkbookCommand.SetSheetProtection(
                    targetSheet, nextSheetProtectionSettings(data));
            default -> new WorkbookCommand.ClearSheetProtection(targetSheet);
          };
      case 0x1 ->
          switch (selectorSlot(selector)) {
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
            case 0x6 -> new WorkbookCommand.SetPrintLayout(targetSheet, nextExcelPrintLayout(data));
            default -> new WorkbookCommand.ClearPrintLayout(targetSheet);
          };
      case 0x2 ->
          switch (selectorSlot(selector)) {
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
      case 0x3 ->
          switch (selectorSlot(selector)) {
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
            case 0x4 ->
                new WorkbookCommand.SetPicture(targetSheet, nextExcelPictureDefinition(data));
            case 0x5 -> new WorkbookCommand.SetChart(targetSheet, nextExcelChartDefinition(data));
            case 0x6 -> new WorkbookCommand.SetShape(targetSheet, nextExcelShapeDefinition(data));
            case 0x7 ->
                new WorkbookCommand.SetEmbeddedObject(
                    targetSheet, nextExcelEmbeddedObjectDefinition(data));
            case 0x8 ->
                new WorkbookCommand.SetDrawingObjectAnchor(
                    targetSheet, nextDrawingObjectName(data), nextExcelDrawingAnchor(data));
            case 0x9 ->
                new WorkbookCommand.DeleteDrawingObject(targetSheet, nextDrawingObjectName(data));
            default ->
                new WorkbookCommand.ApplyStyle(
                    targetSheet,
                    validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false),
                    FuzzDataDecoders.nextStyle(data));
          };
      case 0x4 ->
          switch (selectorSlot(selector)) {
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
      case 0x5 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new WorkbookCommand.SetAutofilter(targetSheet, nextAutofilterRange(validRange));
            case 0x1 -> new WorkbookCommand.ClearAutofilter(targetSheet);
            case 0x2 ->
                new WorkbookCommand.SetTable(
                    nextExcelTableDefinition(data, targetSheet, tableName, validRange));
            case 0x3 -> new WorkbookCommand.DeleteTable(tableName, targetSheet);
            case 0x4 ->
                new WorkbookCommand.SetPivotTable(
                    nextExcelPivotTableDefinition(
                        data,
                        targetSheet,
                        pivotTableName,
                        workbookNamedRange,
                        tableName,
                        validName,
                        validRange));
            default ->
                new WorkbookCommand.DeletePivotTable(
                    validName ? pivotTableName : nextPivotTableName(data, false), targetSheet);
          };
      case 0x6 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new WorkbookCommand.SetNamedRange(
                    new ExcelNamedRangeDefinition(
                        validName ? namedRangeName : nextNamedRangeName(data, false),
                        data.consumeBoolean()
                            ? new ExcelNamedRangeScope.WorkbookScope()
                            : new ExcelNamedRangeScope.SheetScope(targetSheet),
                        new ExcelNamedRangeTarget(
                            targetSheet,
                            validRange
                                ? "A1:B2"
                                : FuzzDataDecoders.nextNonBlankRange(data, false))));
            default ->
                new WorkbookCommand.DeleteNamedRange(
                    validName ? namedRangeName : nextNamedRangeName(data, false),
                    data.consumeBoolean()
                        ? new ExcelNamedRangeScope.WorkbookScope()
                        : new ExcelNamedRangeScope.SheetScope(targetSheet));
          };
      case 0x7 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new WorkbookCommand.InsertRows(
                    targetSheet, rowSpan.first(), rowSpan.last() - rowSpan.first() + 1);
            case 0x1 ->
                new WorkbookCommand.DeleteRows(
                    targetSheet, new ExcelRowSpan(rowSpan.first(), rowSpan.last()));
            case 0x2 ->
                new WorkbookCommand.ShiftRows(
                    targetSheet,
                    new ExcelRowSpan(rowSpan.first(), rowSpan.last()),
                    nextNonZeroDelta(data, 2));
            case 0x3 ->
                new WorkbookCommand.InsertColumns(
                    targetSheet, columnSpan.first(), columnSpan.last() - columnSpan.first() + 1);
            case 0x4 ->
                new WorkbookCommand.DeleteColumns(
                    targetSheet, new ExcelColumnSpan(columnSpan.first(), columnSpan.last()));
            case 0x5 ->
                new WorkbookCommand.ShiftColumns(
                    targetSheet,
                    new ExcelColumnSpan(columnSpan.first(), columnSpan.last()),
                    nextNonZeroDelta(data, 2));
            case 0x6 ->
                new WorkbookCommand.SetRowVisibility(
                    targetSheet,
                    new ExcelRowSpan(rowSpan.first(), rowSpan.last()),
                    data.consumeBoolean());
            case 0x7 ->
                new WorkbookCommand.SetColumnVisibility(
                    targetSheet,
                    new ExcelColumnSpan(columnSpan.first(), columnSpan.last()),
                    data.consumeBoolean());
            case 0x8 ->
                new WorkbookCommand.GroupRows(
                    targetSheet,
                    new ExcelRowSpan(rowSpan.first(), rowSpan.last()),
                    data.consumeBoolean());
            case 0x9 ->
                new WorkbookCommand.UngroupRows(
                    targetSheet, new ExcelRowSpan(rowSpan.first(), rowSpan.last()));
            case 0xA ->
                new WorkbookCommand.GroupColumns(
                    targetSheet,
                    new ExcelColumnSpan(columnSpan.first(), columnSpan.last()),
                    data.consumeBoolean());
            default ->
                new WorkbookCommand.UngroupColumns(
                    targetSheet, new ExcelColumnSpan(columnSpan.first(), columnSpan.last()));
          };
      default ->
          switch (selectorSlot(selector)) {
            case 0x0 -> new WorkbookCommand.AutoSizeColumns(targetSheet);
            case 0x1 ->
                new WorkbookCommand.SetWorkbookProtection(
                    new ExcelWorkbookProtectionSettings(true, false, false, null, null));
            case 0x2 -> new WorkbookCommand.ClearWorkbookProtection();
            case 0x3 ->
                new WorkbookCommand.SetSheetPresentation(
                    targetSheet, ExcelSheetPresentation.defaults());
            default -> new WorkbookCommand.AutoSizeColumns(targetSheet);
          };
    };
  }

  private static boolean validFormulaAddress(GridGrindFuzzData data) {
    return data.consumeBoolean();
  }

  private static ExecutionPolicyInput nextExecutionPolicy(
      GridGrindFuzzData data, String primarySheet, String secondarySheet, boolean validAddress) {
    return switch (selectorSlot(nextSelectorByte(data)) % 6) {
      case 0 -> null;
      case 1 -> ProtocolStepSupport.executionPolicy(ProtocolStepSupport.calculateAll());
      case 2 ->
          ProtocolStepSupport.executionPolicy(
              ProtocolStepSupport.calculateTargets(
                  nextQualifiedFormulaAddress(data, primarySheet, secondarySheet, validAddress)));
      case 3 -> ProtocolStepSupport.executionPolicy(ProtocolStepSupport.clearFormulaCaches());
      case 4 -> ProtocolStepSupport.executionPolicy(ProtocolStepSupport.markRecalculateOnOpen());
      default ->
          ProtocolStepSupport.executionPolicy(
              ProtocolStepSupport.calculateAllAndMarkRecalculateOnOpen());
    };
  }

  private static CellSelector.QualifiedAddress nextQualifiedFormulaAddress(
      GridGrindFuzzData data, String primarySheet, String secondarySheet, boolean validAddress) {
    return new CellSelector.QualifiedAddress(
        data.consumeBoolean() ? primarySheet : secondarySheet,
        nextFormulaTargetAddress(data, validAddress));
  }

  private static InspectionStep nextInspection(
      GridGrindFuzzData data,
      int index,
      String primarySheet,
      String secondarySheet,
      String workbookNamedRange,
      String sheetNamedRange,
      String pivotTableName) {
    String requestId = "read-" + index;
    String targetSheet = data.consumeBoolean() ? primarySheet : secondarySheet;
    boolean validAddress = data.consumeBoolean();
    boolean validRange = data.consumeBoolean();
    boolean validName = data.consumeBoolean();
    int selector = nextSelectorByte(data);

    return switch (selectorFamily(selector)) {
      case 0x0 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                inspect(
                    requestId,
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookSummary());
            case 0x1 ->
                inspect(
                    requestId,
                    new SheetSelector.ByName(targetSheet),
                    new InspectionQuery.GetSheetSummary());
            case 0x2 ->
                inspect(
                    requestId,
                    nextNamedRangeSelection(
                        data, targetSheet, workbookNamedRange, sheetNamedRange, validName),
                    new InspectionQuery.GetNamedRanges());
            default ->
                inspect(
                    requestId,
                    nextNamedRangeSelection(
                        data, targetSheet, workbookNamedRange, sheetNamedRange, validName),
                    new InspectionQuery.GetNamedRangeSurface());
          };
      case 0x1 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                inspect(
                    requestId,
                    new CellSelector.ByAddresses(
                        targetSheet, nextReadAddresses(data, validAddress)),
                    new InspectionQuery.GetCells());
            case 0x1 ->
                inspect(
                    requestId,
                    new RangeSelector.RectangularWindow(
                        targetSheet,
                        FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
                        data.consumeInt(1, 4),
                        data.consumeInt(1, 4)),
                    new InspectionQuery.GetWindow());
            case 0x2 ->
                inspect(
                    requestId,
                    new SheetSelector.ByName(targetSheet),
                    new InspectionQuery.GetSheetLayout());
            default ->
                inspect(
                    requestId,
                    new RangeSelector.RectangularWindow(
                        targetSheet,
                        validAddress ? "A1" : FuzzDataDecoders.nextNonBlankCellAddress(data, false),
                        data.consumeInt(1, 5),
                        data.consumeInt(1, 4)),
                    new InspectionQuery.GetSheetSchema());
          };
      case 0x2 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                inspect(
                    requestId,
                    new SheetSelector.ByName(targetSheet),
                    new InspectionQuery.GetMergedRegions());
            case 0x1 ->
                inspect(
                    requestId,
                    nextCellSelector(data, targetSheet, validAddress),
                    new InspectionQuery.GetHyperlinks());
            case 0x2 ->
                inspect(
                    requestId,
                    nextCellSelector(data, targetSheet, validAddress),
                    new InspectionQuery.GetComments());
            case 0x3 ->
                inspect(
                    requestId,
                    new DrawingObjectSelector.AllOnSheet(targetSheet),
                    new InspectionQuery.GetDrawingObjects());
            case 0x4 ->
                inspect(
                    requestId,
                    new DrawingObjectSelector.ByName(
                        targetSheet, nextDrawingBinaryObjectName(data)),
                    new InspectionQuery.GetDrawingObjectPayload());
            case 0x5 ->
                inspect(
                    requestId,
                    new ChartSelector.AllOnSheet(targetSheet),
                    new InspectionQuery.GetCharts());
            case 0x6 ->
                inspect(
                    requestId,
                    new SheetSelector.ByName(targetSheet),
                    new InspectionQuery.GetPrintLayout());
            default ->
                inspect(
                    requestId,
                    nextSheetSelector(data, primarySheet, secondarySheet),
                    new InspectionQuery.GetFormulaSurface());
          };
      case 0x3 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                inspect(
                    requestId,
                    nextRangeSelector(data, targetSheet, validRange),
                    new InspectionQuery.GetDataValidations());
            case 0x1 ->
                inspect(
                    requestId,
                    nextRangeSelector(data, targetSheet, validRange),
                    new InspectionQuery.GetConditionalFormatting());
            default ->
                inspect(
                    requestId,
                    new SheetSelector.ByName(targetSheet),
                    new InspectionQuery.GetAutofilters());
          };
      case 0x4 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                inspect(
                    requestId,
                    nextTableSelector(data, primarySheet, secondarySheet),
                    new InspectionQuery.GetTables());
            case 0x1 ->
                inspect(
                    requestId,
                    nextPivotTableSelector(
                        data, primarySheet, secondarySheet, pivotTableName, validName),
                    new InspectionQuery.GetPivotTables());
            case 0x2 ->
                inspect(
                    requestId,
                    nextTableSelector(data, primarySheet, secondarySheet),
                    new InspectionQuery.AnalyzeTableHealth());
            case 0x3 ->
                inspect(
                    requestId,
                    nextPivotTableSelector(
                        data, primarySheet, secondarySheet, pivotTableName, validName),
                    new InspectionQuery.AnalyzePivotTableHealth());
            default ->
                inspect(
                    requestId,
                    nextSheetSelector(data, primarySheet, secondarySheet),
                    new InspectionQuery.AnalyzeAutofilterHealth());
          };
      case 0x5 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                inspect(
                    requestId,
                    nextSheetSelector(data, primarySheet, secondarySheet),
                    new InspectionQuery.AnalyzeFormulaHealth());
            case 0x1 ->
                inspect(
                    requestId,
                    nextSheetSelector(data, primarySheet, secondarySheet),
                    new InspectionQuery.AnalyzeDataValidationHealth());
            case 0x2 ->
                inspect(
                    requestId,
                    nextSheetSelector(data, primarySheet, secondarySheet),
                    new InspectionQuery.AnalyzeConditionalFormattingHealth());
            default ->
                inspect(
                    requestId,
                    nextSheetSelector(data, primarySheet, secondarySheet),
                    new InspectionQuery.AnalyzeHyperlinkHealth());
          };
      default ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                inspect(
                    requestId,
                    nextNamedRangeSelection(
                        data, targetSheet, workbookNamedRange, sheetNamedRange, validName),
                    new InspectionQuery.AnalyzeNamedRangeHealth());
            default ->
                inspect(
                    requestId,
                    new WorkbookSelector.Current(),
                    new InspectionQuery.AnalyzeWorkbookFindings());
          };
    };
  }

  private static ProtocolStepSupport.PendingAssertion nextAssertion(
      GridGrindFuzzData data,
      int index,
      String primarySheet,
      String secondarySheet,
      String workbookNamedRange,
      String sheetNamedRange,
      String pivotTableName) {
    String stepId = "assert-" + index;
    String targetSheet = data.consumeBoolean() ? primarySheet : secondarySheet;
    return switch (selectorSlot(nextSelectorByte(data)) % 7) {
      case 0 -> assertThat(stepId, new WorkbookSelector.Current(), nextWorkbookAssertion(data));
      case 1 ->
          assertThat(
              stepId,
              new CellSelector.ByAddresses(
                  targetSheet,
                  List.of(
                      FuzzDataDecoders.nextNonBlankCellAddress(data, true),
                      data.consumeBoolean() ? "A1" : "C2")),
              nextCellAssertion(data));
      case 2 ->
          assertThat(
              stepId,
              nextNamedRangeSelection(data, targetSheet, workbookNamedRange, sheetNamedRange, true),
              nextNamedRangeAssertion(data));
      case 3 ->
          assertThat(
              stepId, new SheetSelector.ByName(targetSheet), nextSheetAssertion(data, targetSheet));
      case 4 ->
          assertThat(
              stepId,
              nextTableSelector(data, primarySheet, secondarySheet),
              nextTableAssertion(data));
      case 5 ->
          assertThat(
              stepId,
              nextPivotTableSelector(data, primarySheet, secondarySheet, pivotTableName, true),
              nextPivotTableAssertion(data));
      default ->
          assertThat(stepId, new ChartSelector.AllOnSheet(targetSheet), nextChartAssertion(data));
    };
  }

  private static NamedRangeSelector nextNamedRangeSelection(
      GridGrindFuzzData data,
      String sheetName,
      String workbookNamedRange,
      String sheetNamedRange,
      boolean validName) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new NamedRangeSelector.All();
      default ->
          new NamedRangeSelector.AnyOf(
              List.of(
                  data.consumeBoolean()
                      ? new NamedRangeSelector.WorkbookScope(
                          validName ? workbookNamedRange : nextNamedRangeName(data, false))
                      : new NamedRangeSelector.SheetScope(
                          validName ? sheetNamedRange : nextNamedRangeName(data, false),
                          sheetName)));
    };
  }

  private static CellSelector nextCellSelector(
      GridGrindFuzzData data, String sheetName, boolean validAddress) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new CellSelector.AllUsedInSheet(sheetName);
      default -> new CellSelector.ByAddresses(sheetName, nextReadAddresses(data, validAddress));
    };
  }

  private static SheetSelector nextSheetSelector(
      GridGrindFuzzData data, String primarySheet, String secondarySheet) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new SheetSelector.All();
      default ->
          new SheetSelector.ByNames(
              data.consumeBoolean()
                  ? List.of(primarySheet, secondarySheet)
                  : List.of(secondarySheet, primarySheet));
    };
  }

  private static List<String> nextReadAddresses(GridGrindFuzzData data, boolean validAddress) {
    String first = FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress);
    String second = FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress);
    if (first.equals(second)) {
      second = validAddress ? ("A1".equals(first) ? "B2" : "A1") : "ZZZ999999";
    }
    return List.of(first, second);
  }

  private static Assertion nextWorkbookAssertion(GridGrindFuzzData data) {
    Assertion primary =
        switch (selectorSlot(nextSelectorByte(data)) % 3) {
          case 0 ->
              new Assertion.WorkbookProtectionFacts(
                  new WorkbookProtectionReport(false, false, false, false, false));
          case 1 ->
              new Assertion.AnalysisMaxSeverity(
                  new InspectionQuery.AnalyzeWorkbookFindings(), nextMaximumSeverity(data));
          default ->
              new Assertion.AnalysisFindingAbsent(
                  new InspectionQuery.AnalyzeWorkbookFindings(),
                  AnalysisFindingCode.FORMULA_EXTERNAL_REFERENCE,
                  nextOptionalSeverity(data),
                  null);
        };
    Assertion alternate =
        new Assertion.AnalysisFindingPresent(
            new InspectionQuery.AnalyzeWorkbookFindings(),
            AnalysisFindingCode.NAMED_RANGE_BROKEN_REFERENCE,
            null,
            null);
    return maybeComposeAssertion(data, primary, alternate);
  }

  private static Assertion nextCellAssertion(GridGrindFuzzData data) {
    Assertion primary =
        switch (selectorSlot(nextSelectorByte(data)) % 4) {
          case 0 -> new Assertion.CellValue(nextExpectedCellValue(data));
          case 1 -> new Assertion.DisplayValue("display-" + data.consumeInt(0, 9));
          case 2 -> new Assertion.FormulaText(data.consumeBoolean() ? "B2*2" : "SUM(B2:B4)");
          default -> new Assertion.DisplayValue("Report");
        };
    Assertion alternate = new Assertion.CellValue(new ExpectedCellValue.Text("Jan"));
    return maybeComposeAssertion(data, primary, alternate);
  }

  private static Assertion nextNamedRangeAssertion(GridGrindFuzzData data) {
    Assertion primary =
        switch (selectorSlot(nextSelectorByte(data)) % 3) {
          case 0 -> new Assertion.Present();
          case 1 ->
              new Assertion.AnalysisMaxSeverity(
                  new InspectionQuery.AnalyzeNamedRangeHealth(), nextMaximumSeverity(data));
          default ->
              new Assertion.AnalysisFindingAbsent(
                  new InspectionQuery.AnalyzeNamedRangeHealth(),
                  AnalysisFindingCode.NAMED_RANGE_BROKEN_REFERENCE,
                  nextOptionalSeverity(data),
                  null);
        };
    Assertion alternate = new Assertion.Absent();
    return maybeComposeAssertion(data, primary, alternate);
  }

  private static Assertion nextSheetAssertion(GridGrindFuzzData data, String sheetName) {
    Assertion primary =
        switch (selectorSlot(nextSelectorByte(data)) % 6) {
          case 0 ->
              new Assertion.AnalysisMaxSeverity(
                  new InspectionQuery.AnalyzeFormulaHealth(), nextMaximumSeverity(data));
          case 1 ->
              new Assertion.AnalysisFindingAbsent(
                  new InspectionQuery.AnalyzeDataValidationHealth(),
                  AnalysisFindingCode.DATA_VALIDATION_BROKEN_FORMULA,
                  nextOptionalSeverity(data),
                  null);
          case 2 ->
              new Assertion.AnalysisFindingAbsent(
                  new InspectionQuery.AnalyzeConditionalFormattingHealth(),
                  AnalysisFindingCode.CONDITIONAL_FORMATTING_BROKEN_FORMULA,
                  nextOptionalSeverity(data),
                  null);
          case 3 ->
              new Assertion.AnalysisFindingAbsent(
                  new InspectionQuery.AnalyzeAutofilterHealth(),
                  AnalysisFindingCode.AUTOFILTER_INVALID_RANGE,
                  nextOptionalSeverity(data),
                  null);
          case 4 ->
              new Assertion.AnalysisFindingAbsent(
                  new InspectionQuery.AnalyzeHyperlinkHealth(),
                  AnalysisFindingCode.HYPERLINK_MALFORMED_TARGET,
                  nextOptionalSeverity(data),
                  null);
          default ->
              new Assertion.SheetStructureFacts(
                  new GridGrindResponse.SheetSummaryReport(
                      sheetName,
                      ExcelSheetVisibility.VISIBLE,
                      new GridGrindResponse.SheetProtectionReport.Unprotected(),
                      0,
                      -1,
                      -1));
        };
    Assertion alternate =
        new Assertion.AnalysisMaxSeverity(
            new InspectionQuery.AnalyzeFormulaHealth(), AnalysisSeverity.WARNING);
    return maybeComposeAssertion(data, primary, alternate);
  }

  private static Assertion nextTableAssertion(GridGrindFuzzData data) {
    Assertion primary =
        switch (selectorSlot(nextSelectorByte(data)) % 3) {
          case 0 -> new Assertion.Present();
          case 1 ->
              new Assertion.AnalysisMaxSeverity(
                  new InspectionQuery.AnalyzeTableHealth(), nextMaximumSeverity(data));
          default ->
              new Assertion.AnalysisFindingAbsent(
                  new InspectionQuery.AnalyzeTableHealth(),
                  AnalysisFindingCode.TABLE_BROKEN_REFERENCE,
                  nextOptionalSeverity(data),
                  null);
        };
    Assertion alternate = new Assertion.Absent();
    return maybeComposeAssertion(data, primary, alternate);
  }

  private static Assertion nextPivotTableAssertion(GridGrindFuzzData data) {
    Assertion primary =
        switch (selectorSlot(nextSelectorByte(data)) % 3) {
          case 0 -> new Assertion.Present();
          case 1 ->
              new Assertion.AnalysisMaxSeverity(
                  new InspectionQuery.AnalyzePivotTableHealth(), nextMaximumSeverity(data));
          default ->
              new Assertion.AnalysisFindingAbsent(
                  new InspectionQuery.AnalyzePivotTableHealth(),
                  AnalysisFindingCode.PIVOT_TABLE_BROKEN_SOURCE,
                  nextOptionalSeverity(data),
                  null);
        };
    Assertion alternate = new Assertion.Absent();
    return maybeComposeAssertion(data, primary, alternate);
  }

  private static Assertion nextChartAssertion(GridGrindFuzzData data) {
    Assertion primary = data.consumeBoolean() ? new Assertion.Present() : new Assertion.Absent();
    Assertion alternate = data.consumeBoolean() ? new Assertion.Absent() : new Assertion.Present();
    return maybeComposeAssertion(data, primary, alternate);
  }

  private static Assertion maybeComposeAssertion(
      GridGrindFuzzData data, Assertion primary, Assertion alternate) {
    return switch (selectorSlot(nextSelectorByte(data)) % 4) {
      case 0 -> primary;
      case 1 -> new Assertion.Not(primary);
      case 2 -> new Assertion.AllOf(List.of(primary, alternate));
      default -> new Assertion.AnyOf(List.of(primary, alternate));
    };
  }

  private static AnalysisSeverity nextMaximumSeverity(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 -> AnalysisSeverity.INFO;
      case 1 -> AnalysisSeverity.WARNING;
      default -> AnalysisSeverity.ERROR;
    };
  }

  private static AnalysisSeverity nextOptionalSeverity(GridGrindFuzzData data) {
    return data.consumeBoolean() ? nextMaximumSeverity(data) : null;
  }

  private static ExpectedCellValue nextExpectedCellValue(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 5) {
      case 0 -> new ExpectedCellValue.Blank();
      case 1 -> new ExpectedCellValue.Text("seed-" + data.consumeInt(0, 9));
      case 2 -> new ExpectedCellValue.NumericValue(data.consumeRegularDouble(0.0d, 10.0d));
      case 3 -> new ExpectedCellValue.BooleanValue(data.consumeBoolean());
      default -> new ExpectedCellValue.ErrorValue("#REF!");
    };
  }

  private static String nextFormulaTargetAddress(GridGrindFuzzData data, boolean validAddress) {
    return validAddress ? "C2" : FuzzDataDecoders.nextNonBlankCellAddress(data, false);
  }

  private static IndexSpan nextIndexSpan(GridGrindFuzzData data, int upperBound) {
    int first = data.consumeInt(0, upperBound - 1);
    return new IndexSpan(first, data.consumeInt(first, upperBound));
  }

  private static int nextNonZeroDelta(GridGrindFuzzData data, int upperBound) {
    int absoluteDelta = data.consumeInt(1, upperBound);
    return data.consumeBoolean() ? absoluteDelta : -absoluteDelta;
  }

  private static int nextSelectorByte(GridGrindFuzzData data) {
    return Byte.toUnsignedInt(data.consumeByte());
  }

  private static int selectorFamily(int selector) {
    return selector >>> 4;
  }

  private static int selectorSlot(int selector) {
    return selector & 0x0F;
  }

  private record IndexSpan(int first, int last) {}
}
