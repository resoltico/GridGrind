package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.assertThat;
import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.inspect;
import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.mutate;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.AnalysisFindingCode;
import dev.erst.gridgrind.contract.dto.AnalysisSeverity;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DifferentialStyleInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.DrawingMarkerInput;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextInput;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.PaneInput;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.PrintAreaInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.PrintScalingInput;
import dev.erst.gridgrind.contract.dto.PrintTitleColumnsInput;
import dev.erst.gridgrind.contract.dto.PrintTitleRowsInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
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
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.ExcelBinaryData;
import dev.erst.gridgrind.excel.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.ExcelColumnSpan;
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
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.ExcelHeaderFooterText;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelPaneRegion;
import dev.erst.gridgrind.excel.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.ExcelPictureFormat;
import dev.erst.gridgrind.excel.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.ExcelPivotTableDefinition;
import dev.erst.gridgrind.excel.ExcelPrintLayout;
import dev.erst.gridgrind.excel.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.ExcelRangeSelection;
import dev.erst.gridgrind.excel.ExcelRowSpan;
import dev.erst.gridgrind.excel.ExcelShapeDefinition;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelSheetPresentation;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSettings;
import dev.erst.gridgrind.excel.WorkbookCommand;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;
import java.util.Objects;
import java.util.stream.Stream;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Builds bounded protocol requests and workbook command sequences for Jazzer harnesses. */
public final class OperationSequenceModel {
  private static final String DRAWING_PICTURE_NAME = "OpsPicture";
  private static final String DRAWING_CHART_NAME = "OpsChart";
  private static final String PIVOT_TABLE_NAME = "OpsPivot";
  private static final String DRAWING_SHAPE_NAME = "OpsShape";
  private static final String DRAWING_CONNECTOR_NAME = "OpsConnector";
  private static final String DRAWING_EMBEDDED_OBJECT_NAME = "OpsEmbed";
  private static final String PNG_PIXEL_BASE64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=";

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

  private static PaneInput nextPaneInput(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new PaneInput.None();
      case 1 -> {
        int splitColumn = data.consumeInt(1, 3);
        yield new PaneInput.Frozen(
            splitColumn, 0, data.consumeInt(splitColumn, splitColumn + 2), 0);
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

  private static ExcelSheetPane nextExcelSheetPane(GridGrindFuzzData data) {
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

  private static PrintLayoutInput nextPrintLayoutInput(GridGrindFuzzData data) {
    return new PrintLayoutInput(
        data.consumeBoolean() ? new PrintAreaInput.Range("A1:C20") : new PrintAreaInput.None(),
        data.consumeBoolean() ? ExcelPrintOrientation.LANDSCAPE : ExcelPrintOrientation.PORTRAIT,
        data.consumeBoolean()
            ? new PrintScalingInput.Fit(data.consumeInt(0, 2), data.consumeInt(0, 2))
            : new PrintScalingInput.Automatic(),
        data.consumeBoolean()
            ? new PrintTitleRowsInput.Band(0, data.consumeInt(0, 2))
            : new PrintTitleRowsInput.None(),
        data.consumeBoolean()
            ? new PrintTitleColumnsInput.Band(0, data.consumeInt(0, 2))
            : new PrintTitleColumnsInput.None(),
        new HeaderFooterTextInput(
            TextSourceInput.inline("L" + data.consumeInt(0, 9)),
            TextSourceInput.inline(""),
            TextSourceInput.inline("R" + data.consumeInt(0, 9))),
        new HeaderFooterTextInput(
            TextSourceInput.inline(""),
            TextSourceInput.inline("P" + data.consumeInt(0, 9)),
            TextSourceInput.inline("")));
  }

  private static ExcelPrintLayout nextExcelPrintLayout(GridGrindFuzzData data) {
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

  private static ExcelPaneRegion nextProtocolPaneRegion(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) & 0x3) {
      case 0 -> ExcelPaneRegion.UPPER_LEFT;
      case 1 -> ExcelPaneRegion.UPPER_RIGHT;
      case 2 -> ExcelPaneRegion.LOWER_LEFT;
      default -> ExcelPaneRegion.LOWER_RIGHT;
    };
  }

  private static ExcelPaneRegion nextPaneRegion(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) & 0x3) {
      case 0 -> ExcelPaneRegion.UPPER_LEFT;
      case 1 -> ExcelPaneRegion.UPPER_RIGHT;
      case 2 -> ExcelPaneRegion.LOWER_LEFT;
      default -> ExcelPaneRegion.LOWER_RIGHT;
    };
  }

  private static SheetCopyPosition nextSheetCopyPosition(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new SheetCopyPosition.AppendAtEnd();
      default -> new SheetCopyPosition.AtIndex(data.consumeInt(0, 2));
    };
  }

  private static ExcelSheetCopyPosition nextExcelSheetCopyPosition(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new ExcelSheetCopyPosition.AppendAtEnd();
      default -> new ExcelSheetCopyPosition.AtIndex(data.consumeInt(0, 2));
    };
  }

  private static List<String> nextSelectedSheetNames(
      GridGrindFuzzData data, String primarySheet, String secondarySheet) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> List.of(primarySheet);
      case 1 -> List.of(secondarySheet);
      default -> List.of(primarySheet, secondarySheet);
    };
  }

  private static ExcelSheetVisibility nextProtocolSheetVisibility(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> ExcelSheetVisibility.VISIBLE;
      case 1 -> ExcelSheetVisibility.HIDDEN;
      default -> ExcelSheetVisibility.VERY_HIDDEN;
    };
  }

  private static ExcelSheetVisibility nextSheetVisibility(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> ExcelSheetVisibility.VISIBLE;
      case 1 -> ExcelSheetVisibility.HIDDEN;
      default -> ExcelSheetVisibility.VERY_HIDDEN;
    };
  }

  private static SheetProtectionSettings nextProtocolSheetProtectionSettings(
      GridGrindFuzzData data) {
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

  private static ExcelSheetProtectionSettings nextSheetProtectionSettings(GridGrindFuzzData data) {
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
      String primarySheet, String secondarySheet, GridGrindFuzzData data) throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-jazzer-workflow-");
    Path sourcePath = directory.resolve("source.xlsx");
    Path saveAsPath = directory.resolve("output.xlsx");

    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 ->
          new WorkflowStorage(
              new WorkbookPlan.WorkbookSource.New(),
              new WorkbookPlan.WorkbookPersistence.None(),
              directory);
      case 1 ->
          new WorkflowStorage(
              new WorkbookPlan.WorkbookSource.New(),
              new WorkbookPlan.WorkbookPersistence.SaveAs(saveAsPath.toString()),
              directory);
      case 2 -> {
        writeExistingWorkbook(sourcePath, primarySheet, secondarySheet, data);
        yield new WorkflowStorage(
            new WorkbookPlan.WorkbookSource.ExistingFile(sourcePath.toString()),
            new WorkbookPlan.WorkbookPersistence.None(),
            directory);
      }
      case 3 -> {
        writeExistingWorkbook(sourcePath, primarySheet, secondarySheet, data);
        yield new WorkflowStorage(
            new WorkbookPlan.WorkbookSource.ExistingFile(sourcePath.toString()),
            new WorkbookPlan.WorkbookPersistence.OverwriteSource(),
            directory);
      }
      default -> {
        writeExistingWorkbook(sourcePath, primarySheet, secondarySheet, data);
        yield new WorkflowStorage(
            new WorkbookPlan.WorkbookSource.ExistingFile(sourcePath.toString()),
            new WorkbookPlan.WorkbookPersistence.SaveAs(saveAsPath.toString()),
            directory);
      }
    };
  }

  private static void writeExistingWorkbook(
      Path sourcePath, String primarySheet, String secondarySheet, GridGrindFuzzData data)
      throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var primary = workbook.createSheet(primarySheet);
      primary.createRow(0).createCell(0).setCellValue("Month");
      primary.getRow(0).createCell(1).setCellValue("Plan");
      primary.getRow(0).createCell(2).setCellValue("Actual");
      primary.createRow(1).createCell(0).setCellValue("Jan");
      primary.getRow(1).createCell(1).setCellValue(2.0d);
      primary.getRow(1).createCell(2).setCellValue(4.0d);
      primary.getRow(1).createCell(3).setCellFormula("B2*2");
      primary.createRow(2).createCell(0).setCellValue("Feb");
      primary.getRow(2).createCell(1).setCellValue(3.0d);
      primary.getRow(2).createCell(2).setCellValue(6.0d);
      primary.createRow(3).createCell(0).setCellValue("Mar");
      primary.getRow(3).createCell(1).setCellValue(5.0d);
      primary.getRow(3).createCell(2).setCellValue(7.0d);
      primary.createRow(4).createCell(4).setCellValue("Queue");
      primary.getRow(4).createCell(5).setCellValue("Owner");
      primary.createRow(5).createCell(4).setCellValue("seed");
      primary.getRow(5).createCell(5).setCellValue("GridGrind");
      if (data.consumeBoolean()) {
        var secondary = workbook.createSheet(secondarySheet);
        secondary.createRow(0).createCell(0).setCellValue("Month");
        secondary.getRow(0).createCell(1).setCellValue("Plan");
        secondary.getRow(0).createCell(2).setCellValue("Actual");
        secondary.createRow(1).createCell(0).setCellValue("Jan");
        secondary.getRow(1).createCell(1).setCellValue(3.0d);
        secondary.getRow(1).createCell(2).setCellValue(9.0d);
        secondary.getRow(1).createCell(3).setCellFormula("B2*3");
        secondary.createRow(2).createCell(0).setCellValue("Feb");
        secondary.getRow(2).createCell(1).setCellValue(4.0d);
        secondary.getRow(2).createCell(2).setCellValue(8.0d);
        secondary.createRow(3).createCell(0).setCellValue("Mar");
        secondary.getRow(3).createCell(1).setCellValue(6.0d);
        secondary.getRow(3).createCell(2).setCellValue(10.0d);
      }
      Files.createDirectories(sourcePath.getParent());
      try (OutputStream outputStream = Files.newOutputStream(sourcePath)) {
        workbook.write(outputStream);
      }
    }
  }

  private static HyperlinkTarget nextHyperlinkTarget(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new HyperlinkTarget.Url("https://example.com/" + nextNamedRangeName(data, true));
      case 1 -> new HyperlinkTarget.Email(nextNamedRangeName(data, true) + "@example.com");
      case 2 -> new HyperlinkTarget.File("/tmp/" + nextNamedRangeName(data, true) + ".xlsx");
      default -> new HyperlinkTarget.Document("Budget!A1");
    };
  }

  private static ExcelHyperlink nextExcelHyperlink(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new ExcelHyperlink.Url("https://example.com/" + nextNamedRangeName(data, true));
      case 1 -> new ExcelHyperlink.Email(nextNamedRangeName(data, true) + "@example.com");
      case 2 -> new ExcelHyperlink.File("/tmp/" + nextNamedRangeName(data, true) + ".xlsx");
      default -> new ExcelHyperlink.Document("Budget!A1");
    };
  }

  private static CommentInput nextCommentInput(GridGrindFuzzData data) {
    return new CommentInput(
        TextSourceInput.inline("Note " + nextNamedRangeName(data, true)),
        "GridGrind",
        data.consumeBoolean());
  }

  private static PictureInput nextPictureInput(GridGrindFuzzData data) {
    return new PictureInput(
        DRAWING_PICTURE_NAME,
        nextPictureDataInput(),
        nextDrawingAnchorInput(data),
        data.consumeBoolean() ? TextSourceInput.inline("Queue preview") : null);
  }

  private static ChartInput nextChartInput(GridGrindFuzzData data) {
    DrawingAnchorInput.TwoCell anchor = nextDrawingAnchorInput(data);
    ChartInput.Title title = nextChartTitleInput(data);
    ChartInput.Legend legend = nextChartLegendInput(data);
    ExcelChartDisplayBlanksAs displayBlanksAs = nextChartDisplayBlanksAs(data);
    boolean plotOnlyVisibleCells = data.consumeBoolean();
    boolean varyColors = data.consumeBoolean();
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 ->
          new ChartInput.Bar(
              DRAWING_CHART_NAME,
              anchor,
              title,
              legend,
              displayBlanksAs,
              plotOnlyVisibleCells,
              varyColors,
              nextChartBarDirection(data),
              nextChartSeriesInputs(data, false));
      case 1 ->
          new ChartInput.Line(
              DRAWING_CHART_NAME,
              anchor,
              title,
              legend,
              displayBlanksAs,
              plotOnlyVisibleCells,
              varyColors,
              nextChartSeriesInputs(data, false));
      default ->
          new ChartInput.Pie(
              DRAWING_CHART_NAME,
              anchor,
              title,
              legend,
              displayBlanksAs,
              plotOnlyVisibleCells,
              varyColors,
              data.consumeBoolean() ? data.consumeInt(0, 360) : null,
              nextChartSeriesInputs(data, true));
    };
  }

  private static ShapeInput nextShapeInput(GridGrindFuzzData data) {
    if (data.consumeBoolean()) {
      return new ShapeInput(
          DRAWING_SHAPE_NAME,
          ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
          nextDrawingAnchorInput(data),
          data.consumeBoolean() ? "roundRect" : "rect",
          data.consumeBoolean() ? TextSourceInput.inline("Queue") : null);
    }
    return new ShapeInput(
        DRAWING_CONNECTOR_NAME,
        ExcelAuthoredDrawingShapeKind.CONNECTOR,
        nextDrawingAnchorInput(data),
        null,
        null);
  }

  private static EmbeddedObjectInput nextEmbeddedObjectInput(GridGrindFuzzData data) {
    return new EmbeddedObjectInput(
        DRAWING_EMBEDDED_OBJECT_NAME,
        "Ops payload",
        "ops-payload.txt",
        "open",
        BinarySourceInput.inlineBase64(
            Base64.getEncoder()
                .encodeToString(
                    ("GridGrind payload " + data.consumeInt(0, 9))
                        .getBytes(StandardCharsets.UTF_8))),
        nextPictureDataInput(),
        nextDrawingAnchorInput(data));
  }

  private static DrawingAnchorInput.TwoCell nextDrawingAnchorInput(GridGrindFuzzData data) {
    int firstColumn = data.consumeInt(0, 4);
    int firstRow = data.consumeInt(0, 8);
    int lastColumn = data.consumeInt(firstColumn + 1, firstColumn + 3);
    int lastRow = data.consumeInt(firstRow + 1, firstRow + 4);
    return new DrawingAnchorInput.TwoCell(
        new DrawingMarkerInput(firstColumn, firstRow, 0, 0),
        new DrawingMarkerInput(lastColumn, lastRow, 0, 0),
        nextDrawingAnchorBehavior(data));
  }

  private static PictureDataInput nextPictureDataInput() {
    return new PictureDataInput(
        ExcelPictureFormat.PNG, BinarySourceInput.inlineBase64(PNG_PIXEL_BASE64));
  }

  private static ExcelPictureDefinition nextExcelPictureDefinition(GridGrindFuzzData data) {
    return new ExcelPictureDefinition(
        DRAWING_PICTURE_NAME,
        new ExcelBinaryData(Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
        ExcelPictureFormat.PNG,
        nextExcelDrawingAnchor(data),
        data.consumeBoolean() ? "Queue preview" : null);
  }

  private static ExcelChartDefinition nextExcelChartDefinition(GridGrindFuzzData data) {
    ExcelDrawingAnchor.TwoCell anchor = nextExcelDrawingAnchor(data);
    ExcelChartDefinition.Title title = nextExcelChartTitle(data);
    ExcelChartDefinition.Legend legend = nextExcelChartLegend(data);
    ExcelChartDisplayBlanksAs displayBlanksAs = nextChartDisplayBlanksAs(data);
    boolean plotOnlyVisibleCells = data.consumeBoolean();
    boolean varyColors = data.consumeBoolean();
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 ->
          new ExcelChartDefinition.Bar(
              DRAWING_CHART_NAME,
              anchor,
              title,
              legend,
              displayBlanksAs,
              plotOnlyVisibleCells,
              varyColors,
              nextChartBarDirection(data),
              nextExcelChartSeries(data, false));
      case 1 ->
          new ExcelChartDefinition.Line(
              DRAWING_CHART_NAME,
              anchor,
              title,
              legend,
              displayBlanksAs,
              plotOnlyVisibleCells,
              varyColors,
              nextExcelChartSeries(data, false));
      default ->
          new ExcelChartDefinition.Pie(
              DRAWING_CHART_NAME,
              anchor,
              title,
              legend,
              displayBlanksAs,
              plotOnlyVisibleCells,
              varyColors,
              data.consumeBoolean() ? data.consumeInt(0, 360) : null,
              nextExcelChartSeries(data, true));
    };
  }

  private static ChartInput.Title nextChartTitleInput(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 -> new ChartInput.Title.None();
      case 1 -> new ChartInput.Title.Text(TextSourceInput.inline("Chart " + data.consumeInt(0, 9)));
      default -> new ChartInput.Title.Formula(data.consumeBoolean() ? "B1" : "C1");
    };
  }

  private static ExcelChartDefinition.Title nextExcelChartTitle(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 -> new ExcelChartDefinition.Title.None();
      case 1 -> new ExcelChartDefinition.Title.Text("Chart " + data.consumeInt(0, 9));
      default -> new ExcelChartDefinition.Title.Formula(data.consumeBoolean() ? "B1" : "C1");
    };
  }

  private static ChartInput.Legend nextChartLegendInput(GridGrindFuzzData data) {
    return data.consumeBoolean()
        ? new ChartInput.Legend.Hidden()
        : new ChartInput.Legend.Visible(nextChartLegendPosition(data));
  }

  private static ExcelChartDefinition.Legend nextExcelChartLegend(GridGrindFuzzData data) {
    return data.consumeBoolean()
        ? new ExcelChartDefinition.Legend.Hidden()
        : new ExcelChartDefinition.Legend.Visible(nextChartLegendPosition(data));
  }

  private static ExcelChartLegendPosition nextChartLegendPosition(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 5) {
      case 0 -> ExcelChartLegendPosition.BOTTOM;
      case 1 -> ExcelChartLegendPosition.LEFT;
      case 2 -> ExcelChartLegendPosition.RIGHT;
      case 3 -> ExcelChartLegendPosition.TOP;
      default -> ExcelChartLegendPosition.TOP_RIGHT;
    };
  }

  private static ExcelChartDisplayBlanksAs nextChartDisplayBlanksAs(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 -> ExcelChartDisplayBlanksAs.GAP;
      case 1 -> ExcelChartDisplayBlanksAs.SPAN;
      default -> ExcelChartDisplayBlanksAs.ZERO;
    };
  }

  private static ExcelChartBarDirection nextChartBarDirection(GridGrindFuzzData data) {
    return data.consumeBoolean() ? ExcelChartBarDirection.COLUMN : ExcelChartBarDirection.BAR;
  }

  private static List<ChartInput.Series> nextChartSeriesInputs(
      GridGrindFuzzData data, boolean pieChart) {
    List<ChartInput.Series> series = new ArrayList<>();
    series.add(nextChartSeriesInput(data, "B1", "B2:B4"));
    if (!pieChart && data.consumeBoolean()) {
      series.add(nextChartSeriesInput(data, "C1", "C2:C4"));
    }
    return List.copyOf(series);
  }

  private static List<ExcelChartDefinition.Series> nextExcelChartSeries(
      GridGrindFuzzData data, boolean pieChart) {
    List<ExcelChartDefinition.Series> series = new ArrayList<>();
    series.add(nextExcelChartSeries(data, "B1", "B2:B4"));
    if (!pieChart && data.consumeBoolean()) {
      series.add(nextExcelChartSeries(data, "C1", "C2:C4"));
    }
    return List.copyOf(series);
  }

  private static ChartInput.Series nextChartSeriesInput(
      GridGrindFuzzData data, String titleFormula, String valuesFormula) {
    return new ChartInput.Series(
        data.consumeBoolean()
            ? new ChartInput.Title.Formula(titleFormula)
            : new ChartInput.Title.Text(TextSourceInput.inline("Series " + data.consumeInt(0, 9))),
        new ChartInput.DataSource("A2:A4"),
        new ChartInput.DataSource(valuesFormula));
  }

  private static ExcelChartDefinition.Series nextExcelChartSeries(
      GridGrindFuzzData data, String titleFormula, String valuesFormula) {
    return new ExcelChartDefinition.Series(
        data.consumeBoolean()
            ? new ExcelChartDefinition.Title.Formula(titleFormula)
            : new ExcelChartDefinition.Title.Text("Series " + data.consumeInt(0, 9)),
        new ExcelChartDefinition.DataSource("A2:A4"),
        new ExcelChartDefinition.DataSource(valuesFormula));
  }

  private static ExcelShapeDefinition nextExcelShapeDefinition(GridGrindFuzzData data) {
    if (data.consumeBoolean()) {
      return new ExcelShapeDefinition(
          DRAWING_SHAPE_NAME,
          ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
          nextExcelDrawingAnchor(data),
          data.consumeBoolean() ? "roundRect" : "rect",
          data.consumeBoolean() ? "Queue" : null);
    }
    return new ExcelShapeDefinition(
        DRAWING_CONNECTOR_NAME,
        ExcelAuthoredDrawingShapeKind.CONNECTOR,
        nextExcelDrawingAnchor(data),
        null,
        null);
  }

  private static ExcelEmbeddedObjectDefinition nextExcelEmbeddedObjectDefinition(
      GridGrindFuzzData data) {
    return new ExcelEmbeddedObjectDefinition(
        DRAWING_EMBEDDED_OBJECT_NAME,
        "Ops payload",
        "ops-payload.txt",
        "open",
        new ExcelBinaryData(
            ("GridGrind payload " + data.consumeInt(0, 9)).getBytes(StandardCharsets.UTF_8)),
        ExcelPictureFormat.PNG,
        new ExcelBinaryData(Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
        nextExcelDrawingAnchor(data));
  }

  private static ExcelDrawingAnchor.TwoCell nextExcelDrawingAnchor(GridGrindFuzzData data) {
    DrawingAnchorInput.TwoCell anchor = nextDrawingAnchorInput(data);
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(
            anchor.from().columnIndex(),
            anchor.from().rowIndex(),
            anchor.from().dx(),
            anchor.from().dy()),
        new ExcelDrawingMarker(
            anchor.to().columnIndex(), anchor.to().rowIndex(), anchor.to().dx(), anchor.to().dy()),
        anchor.behavior());
  }

  private static ExcelDrawingAnchorBehavior nextDrawingAnchorBehavior(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) & 0x3) {
      case 0 -> ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE;
      case 1 -> ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE;
      default -> ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE;
    };
  }

  private static String nextDrawingObjectName(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 5) {
      case 0 -> DRAWING_PICTURE_NAME;
      case 1 -> DRAWING_CHART_NAME;
      case 2 -> DRAWING_SHAPE_NAME;
      case 3 -> DRAWING_CONNECTOR_NAME;
      default -> DRAWING_EMBEDDED_OBJECT_NAME;
    };
  }

  private static String nextDrawingBinaryObjectName(GridGrindFuzzData data) {
    return data.consumeBoolean() ? DRAWING_PICTURE_NAME : DRAWING_EMBEDDED_OBJECT_NAME;
  }

  private static DataValidationInput nextDataValidationInput(GridGrindFuzzData data) {
    return new DataValidationInput(
        data.consumeBoolean()
            ? new DataValidationRuleInput.ExplicitList(List.of("Queued", "Done"))
            : new DataValidationRuleInput.WholeNumber(
                ExcelComparisonOperator.GREATER_OR_EQUAL, "1", null),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean()
            ? new DataValidationPromptInput(
                TextSourceInput.inline("Status"),
                TextSourceInput.inline("Use an allowed value."),
                data.consumeBoolean())
            : null,
        data.consumeBoolean()
            ? new DataValidationErrorAlertInput(
                ExcelDataValidationErrorStyle.STOP,
                TextSourceInput.inline("Invalid"),
                TextSourceInput.inline("Use an allowed value."),
                data.consumeBoolean())
            : null);
  }

  private static ExcelDataValidationDefinition nextExcelDataValidationDefinition(
      GridGrindFuzzData data) {
    return new ExcelDataValidationDefinition(
        data.consumeBoolean()
            ? new ExcelDataValidationRule.ExplicitList(List.of("Queued", "Done"))
            : new ExcelDataValidationRule.WholeNumber(
                ExcelComparisonOperator.GREATER_OR_EQUAL, "1", null),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean()
            ? new ExcelDataValidationPrompt(
                "Status", "Use an allowed value.", data.consumeBoolean())
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
      GridGrindFuzzData data, boolean validRange) {
    return new ConditionalFormattingBlockInput(
        List.of(validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)),
        List.of(
            data.consumeBoolean()
                ? new ConditionalFormattingRuleInput.FormulaRule(
                    "A1>0", data.consumeBoolean(), nextDifferentialStyleInput(data))
                : new ConditionalFormattingRuleInput.CellValueRule(
                    ExcelComparisonOperator.GREATER_THAN,
                    "1",
                    null,
                    data.consumeBoolean(),
                    nextDifferentialStyleInput(data))));
  }

  private static ExcelConditionalFormattingBlockDefinition
      nextExcelConditionalFormattingBlockDefinition(GridGrindFuzzData data, boolean validRange) {
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

  private static DifferentialStyleInput nextDifferentialStyleInput(GridGrindFuzzData data) {
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
        numberFormat, bold, italic, null, fontColor, underline, strikeout, fillColor, null);
  }

  private static ExcelDifferentialStyle nextExcelDifferentialStyle(GridGrindFuzzData data) {
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
        numberFormat, bold, italic, null, fontColor, underline, strikeout, fillColor, null);
  }

  private static RangeSelector nextRangeSelector(
      GridGrindFuzzData data, String sheetName, boolean validRange) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new RangeSelector.AllOnSheet(sheetName);
      default ->
          new RangeSelector.ByRanges(
              sheetName,
              List.of(validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)));
    };
  }

  private static ExcelRangeSelection nextExcelRangeSelection(
      GridGrindFuzzData data, boolean validRange) {
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
    String base =
        sourceSheetName.length() <= 27 ? sourceSheetName : sourceSheetName.substring(0, 27);
    return base + "_B1";
  }

  private static TableInput nextTableInput(
      GridGrindFuzzData data, String sheetName, String tableName, boolean validRange) {
    return new TableInput(
        tableName,
        sheetName,
        validRange ? "A1:B3" : FuzzDataDecoders.nextNonBlankRange(data, false),
        data.consumeBoolean(),
        nextTableStyleInput(data));
  }

  private static ExcelTableDefinition nextExcelTableDefinition(
      GridGrindFuzzData data, String sheetName, String tableName, boolean validRange) {
    return new ExcelTableDefinition(
        tableName,
        sheetName,
        validRange ? "A1:B3" : FuzzDataDecoders.nextNonBlankRange(data, false),
        data.consumeBoolean(),
        nextExcelTableStyle(data));
  }

  private static TableStyleInput nextTableStyleInput(GridGrindFuzzData data) {
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

  private static ExcelTableStyle nextExcelTableStyle(GridGrindFuzzData data) {
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

  private static TableSelector nextTableSelector(
      GridGrindFuzzData data, String primarySheet, String secondarySheet) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new TableSelector.All();
      default ->
          new TableSelector.ByNames(
              List.of(
                  data.consumeBoolean()
                      ? nextTableName(data, true, primarySheet)
                      : nextTableName(data, true, secondarySheet)));
    };
  }

  private static PivotTableSelector nextPivotTableSelector(
      GridGrindFuzzData data,
      String primarySheet,
      String secondarySheet,
      String pivotTableName,
      boolean validName) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new PivotTableSelector.All();
      default ->
          data.consumeBoolean()
              ? new PivotTableSelector.ByNameOnSheet(
                  validName ? pivotTableName : nextPivotTableName(data, false), primarySheet)
              : new PivotTableSelector.ByNameOnSheet(
                  validName ? pivotTableName : nextPivotTableName(data, false), secondarySheet);
    };
  }

  private static PivotTableInput nextPivotTableInput(
      GridGrindFuzzData data,
      String targetSheet,
      String pivotTableName,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return new PivotTableInput(
        validName ? pivotTableName : nextPivotTableName(data, false),
        targetSheet,
        nextPivotTableSource(data, targetSheet, namedRangeName, tableName, validName, validRange),
        new PivotTableInput.Anchor(data.consumeBoolean() ? "F4" : "A6"),
        List.of("Month"),
        List.of(),
        List.of(),
        List.of(
            new PivotTableInput.DataField(
                "Actual", ExcelPivotDataConsolidateFunction.SUM, "Total Actual", null)));
  }

  private static PivotTableInput.Source nextPivotTableSource(
      GridGrindFuzzData data,
      String targetSheet,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 ->
          new PivotTableInput.Source.Range(
              targetSheet, validRange ? "A1:C4" : FuzzDataDecoders.nextNonBlankRange(data, false));
      case 1 ->
          new PivotTableInput.Source.NamedRange(
              validName ? namedRangeName : nextNamedRangeName(data, false));
      default ->
          new PivotTableInput.Source.Table(
              validName ? tableName : nextTableName(data, false, targetSheet));
    };
  }

  private static ExcelPivotTableDefinition nextExcelPivotTableDefinition(
      GridGrindFuzzData data,
      String targetSheet,
      String pivotTableName,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return new ExcelPivotTableDefinition(
        validName ? pivotTableName : nextPivotTableName(data, false),
        targetSheet,
        nextExcelPivotTableSource(
            data, targetSheet, namedRangeName, tableName, validName, validRange),
        new ExcelPivotTableDefinition.Anchor(data.consumeBoolean() ? "F4" : "A6"),
        List.of("Month"),
        List.of(),
        List.of(),
        List.of(
            new ExcelPivotTableDefinition.DataField(
                "Actual", ExcelPivotDataConsolidateFunction.SUM, "Total Actual", null)));
  }

  private static ExcelPivotTableDefinition.Source nextExcelPivotTableSource(
      GridGrindFuzzData data,
      String targetSheet,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 ->
          new ExcelPivotTableDefinition.Source.Range(
              targetSheet, validRange ? "A1:C4" : FuzzDataDecoders.nextNonBlankRange(data, false));
      case 1 ->
          new ExcelPivotTableDefinition.Source.NamedRange(
              validName ? namedRangeName : nextNamedRangeName(data, false));
      default ->
          new ExcelPivotTableDefinition.Source.Table(
              validName ? tableName : nextTableName(data, false, targetSheet));
    };
  }

  private static String nextTableName(GridGrindFuzzData data, boolean valid, String sheetName) {
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

  private static ExcelComment nextExcelComment(GridGrindFuzzData data) {
    return new ExcelComment(
        "Note " + nextNamedRangeName(data, true), "GridGrind", data.consumeBoolean());
  }

  private static String nextNamedRangeName(GridGrindFuzzData data, boolean valid) {
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

  private static String nextPivotTableName(GridGrindFuzzData data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");
    if (!valid) {
      return "";
    }
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> PIVOT_TABLE_NAME;
      case 1 -> "Budget Pivot";
      default -> "Pivot " + data.consumeInt(1, 9);
    };
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

  private record WorkflowStorage(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      Path cleanupRoot) {}
}
