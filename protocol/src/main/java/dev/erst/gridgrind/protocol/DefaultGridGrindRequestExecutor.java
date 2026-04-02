package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelCellSelection;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeSelection;
import dev.erst.gridgrind.excel.ExcelNamedRangeSelector;
import dev.erst.gridgrind.excel.ExcelNamedRangeSnapshot;
import dev.erst.gridgrind.excel.ExcelRangeSelection;
import dev.erst.gridgrind.excel.ExcelSheetSelection;
import dev.erst.gridgrind.excel.ExcelTableSelection;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Default request executor that applies the GridGrind workflow against the workbook core. */
public final class DefaultGridGrindRequestExecutor implements GridGrindRequestExecutor {
  private final WorkbookCommandExecutor commandExecutor;
  private final WorkbookReadExecutor readExecutor;
  private final WorkbookCloser workbookCloser;

  /** Creates the production request executor with the default workbook executors and closer. */
  public DefaultGridGrindRequestExecutor() {
    this(new WorkbookCommandExecutor(), new WorkbookReadExecutor(), ExcelWorkbook::close);
  }

  /** Constructor for testing, allowing injection of custom executors and closer. */
  DefaultGridGrindRequestExecutor(
      WorkbookCommandExecutor commandExecutor,
      WorkbookReadExecutor readExecutor,
      WorkbookCloser workbookCloser) {
    this.commandExecutor =
        Objects.requireNonNull(commandExecutor, "commandExecutor must not be null");
    this.readExecutor = Objects.requireNonNull(readExecutor, "readExecutor must not be null");
    this.workbookCloser = Objects.requireNonNull(workbookCloser, "workbookCloser must not be null");
  }

  /** Executes one complete GridGrind request. */
  @Override
  public GridGrindResponse execute(GridGrindRequest request) {
    GridGrindProtocolVersion protocolVersion =
        request == null ? GridGrindProtocolVersion.current() : request.protocolVersion();
    if (request == null) {
      return new GridGrindResponse.Failure(
          protocolVersion,
          GridGrindProblems.problem(
              GridGrindProblemCode.INVALID_REQUEST,
              "request must not be null",
              new GridGrindResponse.ProblemContext.ValidateRequest(null, null),
              (Throwable) null));
    }

    Optional<GridGrindResponse> validationError = validateRequest(protocolVersion, request);
    if (validationError.isPresent()) {
      return validationError.get();
    }

    ExcelWorkbook workbook;
    try {
      workbook = openWorkbook(request.source());
    } catch (Exception exception) {
      return new GridGrindResponse.Failure(
          protocolVersion,
          GridGrindProblems.fromException(
              exception,
              new GridGrindResponse.ProblemContext.OpenWorkbook(
                  reqSourceType(request), reqPersistenceType(request), reqSourcePath(request))));
    }

    return guardUnexpectedRuntime(
        protocolVersion,
        request,
        workbook,
        () -> executeWorkbookWorkflow(protocolVersion, request, workbook));
  }

  private GridGrindResponse executeWorkbookWorkflow(
      GridGrindProtocolVersion protocolVersion, GridGrindRequest request, ExcelWorkbook workbook) {
    WorkbookLocation workbookLocation =
        workbookLocationFor(request.source(), request.persistence());
    for (int operationIndex = 0; operationIndex < request.operations().size(); operationIndex++) {
      WorkbookOperation operation = request.operations().get(operationIndex);
      try {
        commandExecutor.apply(workbook, toCommand(operation));
      } catch (Exception exception) {
        return closeWorkbook(
            workbook,
            operationFailure(protocolVersion, request, operationIndex, operation, exception),
            request);
      }
    }

    List<WorkbookReadResult> reads = new ArrayList<>(request.reads().size());
    for (int readIndex = 0; readIndex < request.reads().size(); readIndex++) {
      WorkbookReadOperation read = request.reads().get(readIndex);
      try {
        WorkbookReadCommand command = toReadCommand(read);
        dev.erst.gridgrind.excel.WorkbookReadResult result =
            readExecutor.apply(workbook, workbookLocation, command).getFirst();
        reads.add(toReadResult(result));
      } catch (Exception exception) {
        return closeWorkbook(
            workbook, readFailure(protocolVersion, request, readIndex, read, exception), request);
      }
    }

    GridGrindResponse.PersistenceOutcome persistence;
    try {
      persistence = persistWorkbook(workbook, request.source(), request.persistence());
    } catch (Exception exception) {
      return closeWorkbook(
          workbook,
          new GridGrindResponse.Failure(
              protocolVersion,
              GridGrindProblems.fromException(
                  exception,
                  new GridGrindResponse.ProblemContext.PersistWorkbook(
                      reqSourceType(request),
                      reqPersistenceType(request),
                      reqSourcePath(request),
                      persistencePath(request.source(), request.persistence())))),
          request);
    }

    return closeWorkbook(
        workbook,
        new GridGrindResponse.Success(protocolVersion, persistence, List.copyOf(reads)),
        request);
  }

  /** Validates cross-field constraints that cannot be enforced at the record level. */
  private Optional<GridGrindResponse> validateRequest(
      GridGrindProtocolVersion protocolVersion, GridGrindRequest request) {
    return switch (request.persistence()) {
      case GridGrindRequest.WorkbookPersistence.OverwriteSource _ ->
          switch (request.source()) {
            case GridGrindRequest.WorkbookSource.New _ ->
                Optional.of(
                    new GridGrindResponse.Failure(
                        protocolVersion,
                        GridGrindProblems.problem(
                            GridGrindProblemCode.INVALID_REQUEST,
                            "OVERWRITE persistence requires an EXISTING source; "
                                + "a NEW workbook has no source file to overwrite",
                            new GridGrindResponse.ProblemContext.ValidateRequest(
                                reqSourceType(request), reqPersistenceType(request)),
                            (Throwable) null)));
            case GridGrindRequest.WorkbookSource.ExistingFile _ -> Optional.empty();
          };
      case GridGrindRequest.WorkbookPersistence.None _ -> Optional.empty();
      case GridGrindRequest.WorkbookPersistence.SaveAs _ -> Optional.empty();
    };
  }

  private ExcelWorkbook openWorkbook(GridGrindRequest.WorkbookSource source) throws IOException {
    return switch (source) {
      case GridGrindRequest.WorkbookSource.New _ -> ExcelWorkbook.create();
      case GridGrindRequest.WorkbookSource.ExistingFile existingFile ->
          ExcelWorkbook.open(Path.of(existingFile.path()));
    };
  }

  /** Converts one protocol operation into the matching workbook-core command. */
  static WorkbookCommand toCommand(WorkbookOperation operation) {
    return switch (operation) {
      case WorkbookOperation.EnsureSheet op -> new WorkbookCommand.CreateSheet(op.sheetName());
      case WorkbookOperation.RenameSheet op ->
          new WorkbookCommand.RenameSheet(op.sheetName(), op.newSheetName());
      case WorkbookOperation.DeleteSheet op -> new WorkbookCommand.DeleteSheet(op.sheetName());
      case WorkbookOperation.MoveSheet op ->
          new WorkbookCommand.MoveSheet(op.sheetName(), op.targetIndex());
      case WorkbookOperation.CopySheet op ->
          new WorkbookCommand.CopySheet(
              op.sourceSheetName(), op.newSheetName(), op.position().toExcelSheetCopyPosition());
      case WorkbookOperation.SetActiveSheet op ->
          new WorkbookCommand.SetActiveSheet(op.sheetName());
      case WorkbookOperation.SetSelectedSheets op ->
          new WorkbookCommand.SetSelectedSheets(op.sheetNames());
      case WorkbookOperation.SetSheetVisibility op ->
          new WorkbookCommand.SetSheetVisibility(op.sheetName(), op.visibility());
      case WorkbookOperation.SetSheetProtection op ->
          new WorkbookCommand.SetSheetProtection(op.sheetName(), op.protection());
      case WorkbookOperation.ClearSheetProtection op ->
          new WorkbookCommand.ClearSheetProtection(op.sheetName());
      case WorkbookOperation.MergeCells op ->
          new WorkbookCommand.MergeCells(op.sheetName(), op.range());
      case WorkbookOperation.UnmergeCells op ->
          new WorkbookCommand.UnmergeCells(op.sheetName(), op.range());
      case WorkbookOperation.SetColumnWidth op ->
          new WorkbookCommand.SetColumnWidth(
              op.sheetName(), op.firstColumnIndex(), op.lastColumnIndex(), op.widthCharacters());
      case WorkbookOperation.SetRowHeight op ->
          new WorkbookCommand.SetRowHeight(
              op.sheetName(), op.firstRowIndex(), op.lastRowIndex(), op.heightPoints());
      case WorkbookOperation.FreezePanes op ->
          new WorkbookCommand.FreezePanes(
              op.sheetName(), op.splitColumn(), op.splitRow(), op.leftmostColumn(), op.topRow());
      case WorkbookOperation.SetCell op ->
          new WorkbookCommand.SetCell(op.sheetName(), op.address(), op.value().toExcelCellValue());
      case WorkbookOperation.SetRange op ->
          new WorkbookCommand.SetRange(
              op.sheetName(),
              op.range(),
              op.rows().stream()
                  .map(row -> row.stream().map(CellInput::toExcelCellValue).toList())
                  .toList());
      case WorkbookOperation.ClearRange op ->
          new WorkbookCommand.ClearRange(op.sheetName(), op.range());
      case WorkbookOperation.SetHyperlink op ->
          new WorkbookCommand.SetHyperlink(
              op.sheetName(), op.address(), op.target().toExcelHyperlink());
      case WorkbookOperation.ClearHyperlink op ->
          new WorkbookCommand.ClearHyperlink(op.sheetName(), op.address());
      case WorkbookOperation.SetComment op ->
          new WorkbookCommand.SetComment(
              op.sheetName(), op.address(), op.comment().toExcelComment());
      case WorkbookOperation.ClearComment op ->
          new WorkbookCommand.ClearComment(op.sheetName(), op.address());
      case WorkbookOperation.ApplyStyle op ->
          new WorkbookCommand.ApplyStyle(op.sheetName(), op.range(), op.style().toExcelCellStyle());
      case WorkbookOperation.SetDataValidation op ->
          new WorkbookCommand.SetDataValidation(
              op.sheetName(), op.range(), op.validation().toExcelDataValidationDefinition());
      case WorkbookOperation.ClearDataValidations op ->
          new WorkbookCommand.ClearDataValidations(
              op.sheetName(), toExcelRangeSelection(op.selection()));
      case WorkbookOperation.SetConditionalFormatting op ->
          new WorkbookCommand.SetConditionalFormatting(
              op.sheetName(),
              op.conditionalFormatting().toExcelConditionalFormattingBlockDefinition());
      case WorkbookOperation.ClearConditionalFormatting op ->
          new WorkbookCommand.ClearConditionalFormatting(
              op.sheetName(), toExcelRangeSelection(op.selection()));
      case WorkbookOperation.SetAutofilter op ->
          new WorkbookCommand.SetAutofilter(op.sheetName(), op.range());
      case WorkbookOperation.ClearAutofilter op ->
          new WorkbookCommand.ClearAutofilter(op.sheetName());
      case WorkbookOperation.SetTable op ->
          new WorkbookCommand.SetTable(op.table().toExcelTableDefinition());
      case WorkbookOperation.DeleteTable op ->
          new WorkbookCommand.DeleteTable(op.name(), op.sheetName());
      case WorkbookOperation.SetNamedRange op ->
          new WorkbookCommand.SetNamedRange(
              new dev.erst.gridgrind.excel.ExcelNamedRangeDefinition(
                  op.name(),
                  op.scope().toExcelNamedRangeScope(),
                  op.target().toExcelNamedRangeTarget()));
      case WorkbookOperation.DeleteNamedRange op ->
          new WorkbookCommand.DeleteNamedRange(op.name(), op.scope().toExcelNamedRangeScope());
      case WorkbookOperation.AppendRow op ->
          new WorkbookCommand.AppendRow(
              op.sheetName(), op.values().stream().map(CellInput::toExcelCellValue).toList());
      case WorkbookOperation.AutoSizeColumns op ->
          new WorkbookCommand.AutoSizeColumns(op.sheetName());
      case WorkbookOperation.EvaluateFormulas _ -> new WorkbookCommand.EvaluateAllFormulas();
      case WorkbookOperation.ForceFormulaRecalculationOnOpen _ ->
          new WorkbookCommand.ForceFormulaRecalculationOnOpen();
    };
  }

  /** Converts one protocol read into the matching workbook-core read command. */
  static WorkbookReadCommand toReadCommand(WorkbookReadOperation read) {
    return switch (read) {
      case WorkbookReadOperation.GetWorkbookSummary op ->
          new WorkbookReadCommand.GetWorkbookSummary(op.requestId());
      case WorkbookReadOperation.GetNamedRanges op ->
          new WorkbookReadCommand.GetNamedRanges(
              op.requestId(), toExcelNamedRangeSelection(op.selection()));
      case WorkbookReadOperation.GetSheetSummary op ->
          new WorkbookReadCommand.GetSheetSummary(op.requestId(), op.sheetName());
      case WorkbookReadOperation.GetCells op ->
          new WorkbookReadCommand.GetCells(op.requestId(), op.sheetName(), op.addresses());
      case WorkbookReadOperation.GetWindow op ->
          new WorkbookReadCommand.GetWindow(
              op.requestId(), op.sheetName(), op.topLeftAddress(), op.rowCount(), op.columnCount());
      case WorkbookReadOperation.GetMergedRegions op ->
          new WorkbookReadCommand.GetMergedRegions(op.requestId(), op.sheetName());
      case WorkbookReadOperation.GetHyperlinks op ->
          new WorkbookReadCommand.GetHyperlinks(
              op.requestId(), op.sheetName(), toExcelCellSelection(op.selection()));
      case WorkbookReadOperation.GetComments op ->
          new WorkbookReadCommand.GetComments(
              op.requestId(), op.sheetName(), toExcelCellSelection(op.selection()));
      case WorkbookReadOperation.GetSheetLayout op ->
          new WorkbookReadCommand.GetSheetLayout(op.requestId(), op.sheetName());
      case WorkbookReadOperation.GetDataValidations op ->
          new WorkbookReadCommand.GetDataValidations(
              op.requestId(), op.sheetName(), toExcelRangeSelection(op.selection()));
      case WorkbookReadOperation.GetConditionalFormatting op ->
          new WorkbookReadCommand.GetConditionalFormatting(
              op.requestId(), op.sheetName(), toExcelRangeSelection(op.selection()));
      case WorkbookReadOperation.GetAutofilters op ->
          new WorkbookReadCommand.GetAutofilters(op.requestId(), op.sheetName());
      case WorkbookReadOperation.GetTables op ->
          new WorkbookReadCommand.GetTables(op.requestId(), toExcelTableSelection(op.selection()));
      case WorkbookReadOperation.GetFormulaSurface op ->
          new WorkbookReadCommand.GetFormulaSurface(
              op.requestId(), toExcelSheetSelection(op.selection()));
      case WorkbookReadOperation.GetSheetSchema op ->
          new WorkbookReadCommand.GetSheetSchema(
              op.requestId(), op.sheetName(), op.topLeftAddress(), op.rowCount(), op.columnCount());
      case WorkbookReadOperation.GetNamedRangeSurface op ->
          new WorkbookReadCommand.GetNamedRangeSurface(
              op.requestId(), toExcelNamedRangeSelection(op.selection()));
      case WorkbookReadOperation.AnalyzeFormulaHealth op ->
          new WorkbookReadCommand.AnalyzeFormulaHealth(
              op.requestId(), toExcelSheetSelection(op.selection()));
      case WorkbookReadOperation.AnalyzeDataValidationHealth op ->
          new WorkbookReadCommand.AnalyzeDataValidationHealth(
              op.requestId(), toExcelSheetSelection(op.selection()));
      case WorkbookReadOperation.AnalyzeConditionalFormattingHealth op ->
          new WorkbookReadCommand.AnalyzeConditionalFormattingHealth(
              op.requestId(), toExcelSheetSelection(op.selection()));
      case WorkbookReadOperation.AnalyzeAutofilterHealth op ->
          new WorkbookReadCommand.AnalyzeAutofilterHealth(
              op.requestId(), toExcelSheetSelection(op.selection()));
      case WorkbookReadOperation.AnalyzeTableHealth op ->
          new WorkbookReadCommand.AnalyzeTableHealth(
              op.requestId(), toExcelTableSelection(op.selection()));
      case WorkbookReadOperation.AnalyzeHyperlinkHealth op ->
          new WorkbookReadCommand.AnalyzeHyperlinkHealth(
              op.requestId(), toExcelSheetSelection(op.selection()));
      case WorkbookReadOperation.AnalyzeNamedRangeHealth op ->
          new WorkbookReadCommand.AnalyzeNamedRangeHealth(
              op.requestId(), toExcelNamedRangeSelection(op.selection()));
      case WorkbookReadOperation.AnalyzeWorkbookFindings op ->
          new WorkbookReadCommand.AnalyzeWorkbookFindings(op.requestId());
    };
  }

  GridGrindResponse.PersistenceOutcome persistWorkbook(
      ExcelWorkbook workbook,
      GridGrindRequest.WorkbookSource source,
      GridGrindRequest.WorkbookPersistence persistence)
      throws IOException {
    return switch (persistence) {
      case GridGrindRequest.WorkbookPersistence.None _ ->
          new GridGrindResponse.PersistenceOutcome.NotSaved();
      case GridGrindRequest.WorkbookPersistence.OverwriteSource _ -> {
        Path rawPath =
            switch (source) {
              case GridGrindRequest.WorkbookSource.ExistingFile existingFile ->
                  Path.of(existingFile.path());
              case GridGrindRequest.WorkbookSource.New _ ->
                  throw new IllegalArgumentException(
                      "OVERWRITE persistence requires an EXISTING source");
            };
        Path normalizedPath = rawPath.toAbsolutePath().normalize();
        workbook.save(normalizedPath);
        yield new GridGrindResponse.PersistenceOutcome.Overwritten(
            rawPath.toString(), normalizedPath.toString());
      }
      case GridGrindRequest.WorkbookPersistence.SaveAs saveAs -> {
        Path normalizedPath = Path.of(saveAs.path()).toAbsolutePath().normalize();
        workbook.save(normalizedPath);
        yield new GridGrindResponse.PersistenceOutcome.SavedAs(
            saveAs.path(), normalizedPath.toString());
      }
    };
  }

  /** Converts one workbook-core read result into the protocol response shape. */
  static WorkbookReadResult toReadResult(dev.erst.gridgrind.excel.WorkbookReadResult result) {
    return switch (result) {
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult workbookSummary ->
          new WorkbookReadResult.WorkbookSummaryResult(
              workbookSummary.requestId(), toWorkbookSummary(workbookSummary.workbook()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangesResult namedRanges ->
          new WorkbookReadResult.NamedRangesResult(
              namedRanges.requestId(),
              namedRanges.namedRanges().stream()
                  .map(DefaultGridGrindRequestExecutor::toNamedRangeReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummaryResult sheetSummary ->
          new WorkbookReadResult.SheetSummaryResult(
              sheetSummary.requestId(),
              new GridGrindResponse.SheetSummaryReport(
                  sheetSummary.sheet().sheetName(),
                  sheetSummary.sheet().visibility(),
                  toSheetProtectionReport(sheetSummary.sheet().protection()),
                  sheetSummary.sheet().physicalRowCount(),
                  sheetSummary.sheet().lastRowIndex(),
                  sheetSummary.sheet().lastColumnIndex()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.CellsResult cells ->
          new WorkbookReadResult.CellsResult(
              cells.requestId(),
              cells.sheetName(),
              cells.cells().stream().map(DefaultGridGrindRequestExecutor::toCellReport).toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.WindowResult window ->
          new WorkbookReadResult.WindowResult(window.requestId(), toWindowReport(window.window()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.MergedRegionsResult mergedRegions ->
          new WorkbookReadResult.MergedRegionsResult(
              mergedRegions.requestId(),
              mergedRegions.sheetName(),
              mergedRegions.mergedRegions().stream()
                  .map(region -> new GridGrindResponse.MergedRegionReport(region.range()))
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.HyperlinksResult hyperlinks ->
          new WorkbookReadResult.HyperlinksResult(
              hyperlinks.requestId(),
              hyperlinks.sheetName(),
              hyperlinks.hyperlinks().stream()
                  .map(DefaultGridGrindRequestExecutor::toCellHyperlinkReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.CommentsResult comments ->
          new WorkbookReadResult.CommentsResult(
              comments.requestId(),
              comments.sheetName(),
              comments.comments().stream()
                  .map(DefaultGridGrindRequestExecutor::toCellCommentReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayoutResult sheetLayout ->
          new WorkbookReadResult.SheetLayoutResult(
              sheetLayout.requestId(), toSheetLayoutReport(sheetLayout.layout()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.DataValidationsResult dataValidations ->
          new WorkbookReadResult.DataValidationsResult(
              dataValidations.requestId(),
              dataValidations.sheetName(),
              dataValidations.validations().stream()
                  .map(DefaultGridGrindRequestExecutor::toDataValidationEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingResult
              conditionalFormatting ->
          new WorkbookReadResult.ConditionalFormattingResult(
              conditionalFormatting.requestId(),
              conditionalFormatting.sheetName(),
              conditionalFormatting.conditionalFormattingBlocks().stream()
                  .map(ConditionalFormattingEntryReport::fromExcel)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.AutofiltersResult autofilters ->
          new WorkbookReadResult.AutofiltersResult(
              autofilters.requestId(),
              autofilters.sheetName(),
              autofilters.autofilters().stream()
                  .map(DefaultGridGrindRequestExecutor::toAutofilterEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.TablesResult tables ->
          new WorkbookReadResult.TablesResult(
              tables.requestId(),
              tables.tables().stream()
                  .map(DefaultGridGrindRequestExecutor::toTableEntryReport)
                  .toList());
      case dev.erst.gridgrind.excel.WorkbookReadResult.FormulaSurfaceResult formulaSurface ->
          new WorkbookReadResult.FormulaSurfaceResult(
              formulaSurface.requestId(), toFormulaSurfaceReport(formulaSurface.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetSchemaResult sheetSchema ->
          new WorkbookReadResult.SheetSchemaResult(
              sheetSchema.requestId(), toSheetSchemaReport(sheetSchema.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurfaceResult namedRangeSurface ->
          new WorkbookReadResult.NamedRangeSurfaceResult(
              namedRangeSurface.requestId(),
              toNamedRangeSurfaceReport(namedRangeSurface.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.FormulaHealthResult formulaHealth ->
          new WorkbookReadResult.FormulaHealthResult(
              formulaHealth.requestId(), toFormulaHealthReport(formulaHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.DataValidationHealthResult
              dataValidationHealth ->
          new WorkbookReadResult.DataValidationHealthResult(
              dataValidationHealth.requestId(),
              toDataValidationHealthReport(dataValidationHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingHealthResult
              conditionalFormattingHealth ->
          new WorkbookReadResult.ConditionalFormattingHealthResult(
              conditionalFormattingHealth.requestId(),
              toConditionalFormattingHealthReport(conditionalFormattingHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.AutofilterHealthResult autofilterHealth ->
          new WorkbookReadResult.AutofilterHealthResult(
              autofilterHealth.requestId(), toAutofilterHealthReport(autofilterHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.TableHealthResult tableHealth ->
          new WorkbookReadResult.TableHealthResult(
              tableHealth.requestId(), toTableHealthReport(tableHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.HyperlinkHealthResult hyperlinkHealth ->
          new WorkbookReadResult.HyperlinkHealthResult(
              hyperlinkHealth.requestId(), toHyperlinkHealthReport(hyperlinkHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeHealthResult namedRangeHealth ->
          new WorkbookReadResult.NamedRangeHealthResult(
              namedRangeHealth.requestId(), toNamedRangeHealthReport(namedRangeHealth.analysis()));
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookFindingsResult workbookFindings ->
          new WorkbookReadResult.WorkbookFindingsResult(
              workbookFindings.requestId(), toWorkbookFindingsReport(workbookFindings.analysis()));
    };
  }

  private static GridGrindResponse.WindowReport toWindowReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.Window window) {
    return new GridGrindResponse.WindowReport(
        window.sheetName(),
        window.topLeftAddress(),
        window.rowCount(),
        window.columnCount(),
        window.rows().stream()
            .map(
                row ->
                    new GridGrindResponse.WindowRowReport(
                        row.rowIndex(),
                        row.cells().stream()
                            .map(DefaultGridGrindRequestExecutor::toCellReport)
                            .toList()))
            .toList());
  }

  private static GridGrindResponse.CellHyperlinkReport toCellHyperlinkReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.CellHyperlink hyperlink) {
    return new GridGrindResponse.CellHyperlinkReport(
        hyperlink.address(), toHyperlinkTarget(hyperlink.hyperlink()));
  }

  private static GridGrindResponse.CellCommentReport toCellCommentReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.CellComment comment) {
    return new GridGrindResponse.CellCommentReport(
        comment.address(), toCommentReport(comment.comment()));
  }

  private static GridGrindResponse.SheetLayoutReport toSheetLayoutReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout layout) {
    return new GridGrindResponse.SheetLayoutReport(
        layout.sheetName(),
        toFreezePaneReport(layout.freezePanes()),
        layout.columns().stream()
            .map(
                column ->
                    new GridGrindResponse.ColumnLayoutReport(
                        column.columnIndex(), column.widthCharacters()))
            .toList(),
        layout.rows().stream()
            .map(row -> new GridGrindResponse.RowLayoutReport(row.rowIndex(), row.heightPoints()))
            .toList());
  }

  private static GridGrindResponse.FreezePaneReport toFreezePaneReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.FreezePane freezePane) {
    return switch (freezePane) {
      case dev.erst.gridgrind.excel.WorkbookReadResult.FreezePane.None _ ->
          new GridGrindResponse.FreezePaneReport.None();
      case dev.erst.gridgrind.excel.WorkbookReadResult.FreezePane.Frozen frozen ->
          new GridGrindResponse.FreezePaneReport.Frozen(
              frozen.splitColumn(), frozen.splitRow(), frozen.leftmostColumn(), frozen.topRow());
    };
  }

  private static GridGrindResponse.FormulaSurfaceReport toFormulaSurfaceReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.FormulaSurface analysis) {
    return new GridGrindResponse.FormulaSurfaceReport(
        analysis.totalFormulaCellCount(),
        analysis.sheets().stream()
            .map(
                sheet ->
                    new GridGrindResponse.SheetFormulaSurfaceReport(
                        sheet.sheetName(),
                        sheet.formulaCellCount(),
                        sheet.distinctFormulaCount(),
                        sheet.formulas().stream()
                            .map(
                                formula ->
                                    new GridGrindResponse.FormulaPatternReport(
                                        formula.formula(),
                                        formula.occurrenceCount(),
                                        formula.addresses()))
                            .toList()))
            .toList());
  }

  private static GridGrindResponse.SheetSchemaReport toSheetSchemaReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.SheetSchema analysis) {
    return new GridGrindResponse.SheetSchemaReport(
        analysis.sheetName(),
        analysis.topLeftAddress(),
        analysis.rowCount(),
        analysis.columnCount(),
        analysis.dataRowCount(),
        analysis.columns().stream()
            .map(
                column ->
                    new GridGrindResponse.SchemaColumnReport(
                        column.columnIndex(),
                        column.columnAddress(),
                        column.headerDisplayValue(),
                        column.populatedCellCount(),
                        column.blankCellCount(),
                        column.observedTypes().stream()
                            .map(
                                typeCount ->
                                    new GridGrindResponse.TypeCountReport(
                                        typeCount.type(), typeCount.count()))
                            .toList(),
                        column.dominantType()))
            .toList());
  }

  private static GridGrindResponse.NamedRangeSurfaceReport toNamedRangeSurfaceReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurface analysis) {
    return new GridGrindResponse.NamedRangeSurfaceReport(
        analysis.workbookScopedCount(),
        analysis.sheetScopedCount(),
        analysis.rangeBackedCount(),
        analysis.formulaBackedCount(),
        analysis.namedRanges().stream()
            .map(
                entry ->
                    new GridGrindResponse.NamedRangeSurfaceEntryReport(
                        entry.name(),
                        toNamedRangeScope(entry.scope()),
                        entry.refersToFormula(),
                        switch (entry.kind()) {
                          case RANGE -> GridGrindResponse.NamedRangeBackingKind.RANGE;
                          case FORMULA -> GridGrindResponse.NamedRangeBackingKind.FORMULA;
                        }))
            .toList());
  }

  private static GridGrindResponse.FormulaHealthReport toFormulaHealthReport(
      dev.erst.gridgrind.excel.WorkbookAnalysis.FormulaHealth analysis) {
    return new GridGrindResponse.FormulaHealthReport(
        analysis.checkedFormulaCellCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(DefaultGridGrindRequestExecutor::toAnalysisFindingReport)
            .toList());
  }

  private static DataValidationHealthReport toDataValidationHealthReport(
      dev.erst.gridgrind.excel.WorkbookAnalysis.DataValidationHealth analysis) {
    return new DataValidationHealthReport(
        analysis.checkedValidationCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(DefaultGridGrindRequestExecutor::toAnalysisFindingReport)
            .toList());
  }

  private static ConditionalFormattingHealthReport toConditionalFormattingHealthReport(
      dev.erst.gridgrind.excel.WorkbookAnalysis.ConditionalFormattingHealth analysis) {
    return new ConditionalFormattingHealthReport(
        analysis.checkedConditionalFormattingBlockCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(DefaultGridGrindRequestExecutor::toAnalysisFindingReport)
            .toList());
  }

  private static AutofilterHealthReport toAutofilterHealthReport(
      dev.erst.gridgrind.excel.WorkbookAnalysis.AutofilterHealth analysis) {
    return new AutofilterHealthReport(
        analysis.checkedAutofilterCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(DefaultGridGrindRequestExecutor::toAnalysisFindingReport)
            .toList());
  }

  private static TableHealthReport toTableHealthReport(
      dev.erst.gridgrind.excel.WorkbookAnalysis.TableHealth analysis) {
    return new TableHealthReport(
        analysis.checkedTableCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(DefaultGridGrindRequestExecutor::toAnalysisFindingReport)
            .toList());
  }

  private static GridGrindResponse.HyperlinkHealthReport toHyperlinkHealthReport(
      dev.erst.gridgrind.excel.WorkbookAnalysis.HyperlinkHealth analysis) {
    return new GridGrindResponse.HyperlinkHealthReport(
        analysis.checkedHyperlinkCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(DefaultGridGrindRequestExecutor::toAnalysisFindingReport)
            .toList());
  }

  private static GridGrindResponse.NamedRangeHealthReport toNamedRangeHealthReport(
      dev.erst.gridgrind.excel.WorkbookAnalysis.NamedRangeHealth analysis) {
    return new GridGrindResponse.NamedRangeHealthReport(
        analysis.checkedNamedRangeCount(),
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(DefaultGridGrindRequestExecutor::toAnalysisFindingReport)
            .toList());
  }

  private static GridGrindResponse.WorkbookFindingsReport toWorkbookFindingsReport(
      dev.erst.gridgrind.excel.WorkbookAnalysis.WorkbookFindings analysis) {
    return new GridGrindResponse.WorkbookFindingsReport(
        toAnalysisSummaryReport(analysis.summary()),
        analysis.findings().stream()
            .map(DefaultGridGrindRequestExecutor::toAnalysisFindingReport)
            .toList());
  }

  private static GridGrindResponse.AnalysisSummaryReport toAnalysisSummaryReport(
      dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary summary) {
    return new GridGrindResponse.AnalysisSummaryReport(
        summary.totalCount(), summary.errorCount(), summary.warningCount(), summary.infoCount());
  }

  private static GridGrindResponse.AnalysisFindingReport toAnalysisFindingReport(
      dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding finding) {
    return new GridGrindResponse.AnalysisFindingReport(
        finding.code(),
        finding.severity(),
        finding.title(),
        finding.message(),
        toAnalysisLocationReport(finding.location()),
        finding.evidence());
  }

  private static GridGrindResponse.AnalysisLocationReport toAnalysisLocationReport(
      dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation location) {
    return switch (location) {
      case dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Workbook _ ->
          new GridGrindResponse.AnalysisLocationReport.Workbook();
      case dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Sheet sheet ->
          new GridGrindResponse.AnalysisLocationReport.Sheet(sheet.sheetName());
      case dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Cell cell ->
          new GridGrindResponse.AnalysisLocationReport.Cell(cell.sheetName(), cell.address());
      case dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Range range ->
          new GridGrindResponse.AnalysisLocationReport.Range(range.sheetName(), range.range());
      case dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.NamedRange namedRange ->
          new GridGrindResponse.AnalysisLocationReport.NamedRange(
              namedRange.name(), toNamedRangeScope(namedRange.scope()));
    };
  }

  private static DataValidationEntryReport toDataValidationEntryReport(
      dev.erst.gridgrind.excel.ExcelDataValidationSnapshot snapshot) {
    return switch (snapshot) {
      case dev.erst.gridgrind.excel.ExcelDataValidationSnapshot.Supported supported ->
          new DataValidationEntryReport.Supported(
              supported.ranges(),
              DataValidationEntryReport.DataValidationDefinitionReport.fromExcel(
                  supported.validation()));
      case dev.erst.gridgrind.excel.ExcelDataValidationSnapshot.Unsupported unsupported ->
          new DataValidationEntryReport.Unsupported(
              unsupported.ranges(), unsupported.kind(), unsupported.detail());
    };
  }

  private static AutofilterEntryReport toAutofilterEntryReport(
      dev.erst.gridgrind.excel.ExcelAutofilterSnapshot snapshot) {
    return switch (snapshot) {
      case dev.erst.gridgrind.excel.ExcelAutofilterSnapshot.SheetOwned sheetOwned ->
          new AutofilterEntryReport.SheetOwned(sheetOwned.range());
      case dev.erst.gridgrind.excel.ExcelAutofilterSnapshot.TableOwned tableOwned ->
          new AutofilterEntryReport.TableOwned(tableOwned.range(), tableOwned.tableName());
    };
  }

  private static TableEntryReport toTableEntryReport(
      dev.erst.gridgrind.excel.ExcelTableSnapshot snapshot) {
    return new TableEntryReport(
        snapshot.name(),
        snapshot.sheetName(),
        snapshot.range(),
        snapshot.headerRowCount(),
        snapshot.totalsRowCount(),
        snapshot.columnNames(),
        toTableStyleReport(snapshot.style()),
        snapshot.hasAutofilter());
  }

  private static TableStyleReport toTableStyleReport(
      dev.erst.gridgrind.excel.ExcelTableStyleSnapshot snapshot) {
    return switch (snapshot) {
      case dev.erst.gridgrind.excel.ExcelTableStyleSnapshot.None _ -> new TableStyleReport.None();
      case dev.erst.gridgrind.excel.ExcelTableStyleSnapshot.Named named ->
          new TableStyleReport.Named(
              named.name(),
              named.showFirstColumn(),
              named.showLastColumn(),
              named.showRowStripes(),
              named.showColumnStripes());
    };
  }

  private static ExcelNamedRangeSelection toExcelNamedRangeSelection(
      NamedRangeSelection selection) {
    return switch (selection) {
      case NamedRangeSelection.All _ -> new ExcelNamedRangeSelection.All();
      case NamedRangeSelection.Selected selected ->
          new ExcelNamedRangeSelection.Selected(
              selected.selectors().stream()
                  .map(DefaultGridGrindRequestExecutor::toExcelNamedRangeSelector)
                  .toList());
    };
  }

  private static ExcelNamedRangeSelector toExcelNamedRangeSelector(NamedRangeSelector selector) {
    return switch (selector) {
      case NamedRangeSelector.ByName byName -> new ExcelNamedRangeSelector.ByName(byName.name());
      case NamedRangeSelector.WorkbookScope workbookScope ->
          new ExcelNamedRangeSelector.WorkbookScope(workbookScope.name());
      case NamedRangeSelector.SheetScope sheetScope ->
          new ExcelNamedRangeSelector.SheetScope(sheetScope.name(), sheetScope.sheetName());
    };
  }

  private static ExcelSheetSelection toExcelSheetSelection(SheetSelection selection) {
    return switch (selection) {
      case SheetSelection.All _ -> new ExcelSheetSelection.All();
      case SheetSelection.Selected selected ->
          new ExcelSheetSelection.Selected(selected.sheetNames());
    };
  }

  private static ExcelTableSelection toExcelTableSelection(TableSelection selection) {
    return switch (selection) {
      case TableSelection.All _ -> new ExcelTableSelection.All();
      case TableSelection.ByNames byNames -> new ExcelTableSelection.ByNames(byNames.names());
    };
  }

  private static ExcelRangeSelection toExcelRangeSelection(RangeSelection selection) {
    return switch (selection) {
      case RangeSelection.All _ -> new ExcelRangeSelection.All();
      case RangeSelection.Selected selected -> new ExcelRangeSelection.Selected(selected.ranges());
    };
  }

  private static ExcelCellSelection toExcelCellSelection(CellSelection selection) {
    return switch (selection) {
      case CellSelection.AllUsedCells _ -> new ExcelCellSelection.AllUsedCells();
      case CellSelection.Selected selected -> new ExcelCellSelection.Selected(selected.addresses());
    };
  }

  private static GridGrindResponse.CellReport toCellReport(ExcelCellSnapshot snapshot) {
    GridGrindResponse.CellStyleReport style =
        new GridGrindResponse.CellStyleReport(
            snapshot.style().numberFormat(),
            snapshot.style().bold(),
            snapshot.style().italic(),
            snapshot.style().wrapText(),
            snapshot.style().horizontalAlignment(),
            snapshot.style().verticalAlignment(),
            snapshot.style().fontName(),
            FontHeightReport.fromExcelFontHeight(snapshot.style().fontHeight()),
            snapshot.style().fontColor(),
            snapshot.style().underline(),
            snapshot.style().strikeout(),
            snapshot.style().fillColor(),
            snapshot.style().topBorderStyle(),
            snapshot.style().rightBorderStyle(),
            snapshot.style().bottomBorderStyle(),
            snapshot.style().leftBorderStyle());
    HyperlinkTarget hyperlink = toHyperlinkTarget(snapshot.metadata().hyperlink().orElse(null));
    GridGrindResponse.CommentReport comment =
        toCommentReport(snapshot.metadata().comment().orElse(null));

    return switch (snapshot) {
      case ExcelCellSnapshot.BlankSnapshot s ->
          new GridGrindResponse.CellReport.BlankReport(
              s.address(), s.declaredType(), s.displayValue(), style, hyperlink, comment);
      case ExcelCellSnapshot.TextSnapshot s ->
          new GridGrindResponse.CellReport.TextReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.stringValue());
      case ExcelCellSnapshot.NumberSnapshot s ->
          new GridGrindResponse.CellReport.NumberReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.numberValue());
      case ExcelCellSnapshot.BooleanSnapshot s ->
          new GridGrindResponse.CellReport.BooleanReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.booleanValue());
      case ExcelCellSnapshot.ErrorSnapshot s ->
          new GridGrindResponse.CellReport.ErrorReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.errorValue());
      case ExcelCellSnapshot.FormulaSnapshot s ->
          new GridGrindResponse.CellReport.FormulaReport(
              s.address(),
              s.declaredType(),
              s.displayValue(),
              style,
              hyperlink,
              comment,
              s.formula(),
              toCellReport(s.evaluation()));
    };
  }

  /** Converts workbook-core hyperlink metadata into the canonical protocol hyperlink shape. */
  static HyperlinkTarget toHyperlinkTarget(ExcelHyperlink hyperlink) {
    if (hyperlink == null) {
      return null;
    }
    return switch (hyperlink) {
      case ExcelHyperlink.Url url -> new HyperlinkTarget.Url(url.target());
      case ExcelHyperlink.Email email -> new HyperlinkTarget.Email(email.target());
      case ExcelHyperlink.File file -> new HyperlinkTarget.File(file.path());
      case ExcelHyperlink.Document document -> new HyperlinkTarget.Document(document.target());
    };
  }

  /** Converts workbook-core comment metadata into the protocol response shape. */
  static GridGrindResponse.CommentReport toCommentReport(ExcelComment comment) {
    if (comment == null) {
      return null;
    }
    return new GridGrindResponse.CommentReport(comment.text(), comment.author(), comment.visible());
  }

  /** Converts one workbook-core named-range snapshot into the protocol response shape. */
  static GridGrindResponse.NamedRangeReport toNamedRangeReport(ExcelNamedRangeSnapshot namedRange) {
    return switch (namedRange) {
      case ExcelNamedRangeSnapshot.RangeSnapshot rangeSnapshot ->
          new GridGrindResponse.NamedRangeReport.RangeReport(
              rangeSnapshot.name(),
              toNamedRangeScope(rangeSnapshot.scope()),
              rangeSnapshot.refersToFormula(),
              new NamedRangeTarget(
                  rangeSnapshot.target().sheetName(), rangeSnapshot.target().range()));
      case ExcelNamedRangeSnapshot.FormulaSnapshot formulaSnapshot ->
          new GridGrindResponse.NamedRangeReport.FormulaReport(
              formulaSnapshot.name(),
              toNamedRangeScope(formulaSnapshot.scope()),
              formulaSnapshot.refersToFormula());
    };
  }

  /** Converts the workbook-core named-range scope into the protocol scope variant. */
  static NamedRangeScope toNamedRangeScope(ExcelNamedRangeScope scope) {
    return switch (scope) {
      case ExcelNamedRangeScope.WorkbookScope _ -> new NamedRangeScope.Workbook();
      case ExcelNamedRangeScope.SheetScope sheetScope ->
          new NamedRangeScope.Sheet(sheetScope.sheetName());
    };
  }

  /** Converts workbook-core sheet-protection state into the protocol response variant. */
  static GridGrindResponse.SheetProtectionReport toSheetProtectionReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.SheetProtection protection) {
    return switch (protection) {
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetProtection.Unprotected _ ->
          new GridGrindResponse.SheetProtectionReport.Unprotected();
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetProtection.Protected protectedState ->
          new GridGrindResponse.SheetProtectionReport.Protected(protectedState.settings());
    };
  }

  /** Converts one workbook-core workbook summary into the protocol response shape. */
  static GridGrindResponse.WorkbookSummary toWorkbookSummary(
      dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary workbookSummary) {
    return switch (workbookSummary) {
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.Empty empty ->
          new GridGrindResponse.WorkbookSummary.Empty(
              empty.sheetCount(),
              empty.sheetNames(),
              empty.namedRangeCount(),
              empty.forceFormulaRecalculationOnOpen());
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.WithSheets withSheets ->
          new GridGrindResponse.WorkbookSummary.WithSheets(
              withSheets.sheetCount(),
              withSheets.sheetNames(),
              withSheets.activeSheetName(),
              withSheets.selectedSheetNames(),
              withSheets.namedRangeCount(),
              withSheets.forceFormulaRecalculationOnOpen());
    };
  }

  /** Delegates to GridGrindProblems to classify the exception into a problem code. */
  static GridGrindProblemCode problemCodeFor(Exception exception) {
    return GridGrindProblems.codeFor(exception);
  }

  /** Delegates to GridGrindProblems to extract a user-readable problem message. */
  static String messageFor(Exception exception) {
    return GridGrindProblems.messageFor(exception);
  }

  /**
   * Returns the workbook filesystem context implied by the source and persistence selection.
   *
   * <p>Package-private so tests can assert persisted-path resolution directly without going through
   * full request execution.
   */
  static WorkbookLocation workbookLocationFor(
      GridGrindRequest.WorkbookSource source, GridGrindRequest.WorkbookPersistence persistence) {
    String persistedPath = persistencePath(source, persistence);
    if (persistedPath != null) {
      return new WorkbookLocation.StoredWorkbook(Path.of(persistedPath));
    }
    return switch (source) {
      case GridGrindRequest.WorkbookSource.ExistingFile existingFile ->
          new WorkbookLocation.StoredWorkbook(
              Path.of(existingFile.path()).toAbsolutePath().normalize());
      case GridGrindRequest.WorkbookSource.New _ -> new WorkbookLocation.UnsavedWorkbook();
    };
  }

  private static String absolutePath(String path) {
    return Path.of(path).toAbsolutePath().normalize().toString();
  }

  private GridGrindResponse closeWorkbook(
      ExcelWorkbook workbook, GridGrindResponse response, GridGrindRequest request) {
    try {
      workbookCloser.close(workbook);
      return response;
    } catch (IOException exception) {
      return switch (response) {
        case GridGrindResponse.Failure failure ->
            new GridGrindResponse.Failure(
                failure.protocolVersion(),
                GridGrindProblems.appendCause(
                    failure.problem(),
                    GridGrindProblems.supplementalCause(
                        "EXECUTE_REQUEST",
                        exception,
                        "Workbook close failed after the primary problem")));
        case GridGrindResponse.Success _ ->
            new GridGrindResponse.Failure(
                response.protocolVersion(),
                GridGrindProblems.fromException(
                    exception,
                    new GridGrindResponse.ProblemContext.ExecuteRequest(
                        reqSourceType(request), reqPersistenceType(request))));
      };
    }
  }

  private GridGrindResponse guardUnexpectedRuntime(
      GridGrindProtocolVersion protocolVersion,
      GridGrindRequest request,
      ExcelWorkbook workbook,
      java.util.function.Supplier<GridGrindResponse> execution) {
    try {
      return execution.get();
    } catch (RuntimeException exception) {
      return closeWorkbook(
          workbook,
          new GridGrindResponse.Failure(
              protocolVersion,
              GridGrindProblems.fromException(
                  exception,
                  new GridGrindResponse.ProblemContext.ExecuteRequest(
                      reqSourceType(request), reqPersistenceType(request)))),
          request);
    }
  }

  private String reqSourceType(GridGrindRequest request) {
    return switch (request.source()) {
      case GridGrindRequest.WorkbookSource.New _ -> "NEW";
      case GridGrindRequest.WorkbookSource.ExistingFile _ -> "EXISTING";
    };
  }

  private String reqPersistenceType(GridGrindRequest request) {
    return switch (request.persistence()) {
      case GridGrindRequest.WorkbookPersistence.None _ -> "NONE";
      case GridGrindRequest.WorkbookPersistence.OverwriteSource _ -> "OVERWRITE";
      case GridGrindRequest.WorkbookPersistence.SaveAs _ -> "SAVE_AS";
    };
  }

  private String reqSourcePath(GridGrindRequest request) {
    return switch (request.source()) {
      case GridGrindRequest.WorkbookSource.ExistingFile existingFile ->
          absolutePath(existingFile.path());
      case GridGrindRequest.WorkbookSource.New _ -> null;
    };
  }

  static String persistencePath(
      GridGrindRequest.WorkbookSource source, GridGrindRequest.WorkbookPersistence persistence) {
    return switch (persistence) {
      case GridGrindRequest.WorkbookPersistence.SaveAs saveAs -> absolutePath(saveAs.path());
      case GridGrindRequest.WorkbookPersistence.OverwriteSource _ ->
          switch (source) {
            case GridGrindRequest.WorkbookSource.ExistingFile existingFile ->
                absolutePath(existingFile.path());
            case GridGrindRequest.WorkbookSource.New _ -> null;
          };
      case GridGrindRequest.WorkbookPersistence.None _ -> null;
    };
  }

  private GridGrindResponse operationFailure(
      GridGrindProtocolVersion protocolVersion,
      GridGrindRequest request,
      int operationIndex,
      WorkbookOperation operation,
      Exception exception) {
    return new GridGrindResponse.Failure(
        protocolVersion,
        GridGrindProblems.fromException(
            exception,
            new GridGrindResponse.ProblemContext.ApplyOperation(
                reqSourceType(request),
                reqPersistenceType(request),
                operationIndex,
                operation.operationType(),
                sheetNameFor(operation, exception),
                addressFor(operation, exception),
                rangeFor(operation, exception),
                formulaFor(operation, exception),
                namedRangeNameFor(operation, exception))));
  }

  private GridGrindResponse readFailure(
      GridGrindProtocolVersion protocolVersion,
      GridGrindRequest request,
      int readIndex,
      WorkbookReadOperation read,
      Exception exception) {
    return new GridGrindResponse.Failure(
        protocolVersion,
        GridGrindProblems.fromException(
            exception,
            new GridGrindResponse.ProblemContext.ExecuteRead(
                reqSourceType(request),
                reqPersistenceType(request),
                readIndex,
                readType(read),
                read.requestId(),
                sheetNameFor(read),
                addressFor(read, exception),
                GridGrindProblems.formulaFor(exception),
                namedRangeNameFor(read, exception))));
  }

  /** Returns the SCREAMING_SNAKE_CASE discriminator for one read operation. */
  static String readType(WorkbookReadOperation read) {
    return switch (read) {
      case WorkbookReadOperation.GetWorkbookSummary _ -> "GET_WORKBOOK_SUMMARY";
      case WorkbookReadOperation.GetNamedRanges _ -> "GET_NAMED_RANGES";
      case WorkbookReadOperation.GetSheetSummary _ -> "GET_SHEET_SUMMARY";
      case WorkbookReadOperation.GetCells _ -> "GET_CELLS";
      case WorkbookReadOperation.GetWindow _ -> "GET_WINDOW";
      case WorkbookReadOperation.GetMergedRegions _ -> "GET_MERGED_REGIONS";
      case WorkbookReadOperation.GetHyperlinks _ -> "GET_HYPERLINKS";
      case WorkbookReadOperation.GetComments _ -> "GET_COMMENTS";
      case WorkbookReadOperation.GetSheetLayout _ -> "GET_SHEET_LAYOUT";
      case WorkbookReadOperation.GetDataValidations _ -> "GET_DATA_VALIDATIONS";
      case WorkbookReadOperation.GetConditionalFormatting _ -> "GET_CONDITIONAL_FORMATTING";
      case WorkbookReadOperation.GetAutofilters _ -> "GET_AUTOFILTERS";
      case WorkbookReadOperation.GetTables _ -> "GET_TABLES";
      case WorkbookReadOperation.GetFormulaSurface _ -> "GET_FORMULA_SURFACE";
      case WorkbookReadOperation.GetSheetSchema _ -> "GET_SHEET_SCHEMA";
      case WorkbookReadOperation.GetNamedRangeSurface _ -> "GET_NAMED_RANGE_SURFACE";
      case WorkbookReadOperation.AnalyzeFormulaHealth _ -> "ANALYZE_FORMULA_HEALTH";
      case WorkbookReadOperation.AnalyzeDataValidationHealth _ -> "ANALYZE_DATA_VALIDATION_HEALTH";
      case WorkbookReadOperation.AnalyzeConditionalFormattingHealth _ ->
          "ANALYZE_CONDITIONAL_FORMATTING_HEALTH";
      case WorkbookReadOperation.AnalyzeAutofilterHealth _ -> "ANALYZE_AUTOFILTER_HEALTH";
      case WorkbookReadOperation.AnalyzeTableHealth _ -> "ANALYZE_TABLE_HEALTH";
      case WorkbookReadOperation.AnalyzeHyperlinkHealth _ -> "ANALYZE_HYPERLINK_HEALTH";
      case WorkbookReadOperation.AnalyzeNamedRangeHealth _ -> "ANALYZE_NAMED_RANGE_HEALTH";
      case WorkbookReadOperation.AnalyzeWorkbookFindings _ -> "ANALYZE_WORKBOOK_FINDINGS";
    };
  }

  /** Returns the sheet name associated with one read operation. */
  static String sheetNameFor(WorkbookReadOperation read) {
    return switch (read) {
      case WorkbookReadOperation.GetSheetSummary op -> op.sheetName();
      case WorkbookReadOperation.GetCells op -> op.sheetName();
      case WorkbookReadOperation.GetWindow op -> op.sheetName();
      case WorkbookReadOperation.GetMergedRegions op -> op.sheetName();
      case WorkbookReadOperation.GetHyperlinks op -> op.sheetName();
      case WorkbookReadOperation.GetComments op -> op.sheetName();
      case WorkbookReadOperation.GetSheetLayout op -> op.sheetName();
      case WorkbookReadOperation.GetDataValidations op -> op.sheetName();
      case WorkbookReadOperation.GetConditionalFormatting op -> op.sheetName();
      case WorkbookReadOperation.GetAutofilters op -> op.sheetName();
      case WorkbookReadOperation.GetSheetSchema op -> op.sheetName();
      case WorkbookReadOperation.GetFormulaSurface op -> singleSheetName(op.selection());
      case WorkbookReadOperation.AnalyzeFormulaHealth op -> singleSheetName(op.selection());
      case WorkbookReadOperation.AnalyzeDataValidationHealth op -> singleSheetName(op.selection());
      case WorkbookReadOperation.AnalyzeConditionalFormattingHealth op ->
          singleSheetName(op.selection());
      case WorkbookReadOperation.AnalyzeAutofilterHealth op -> singleSheetName(op.selection());
      case WorkbookReadOperation.AnalyzeHyperlinkHealth op -> singleSheetName(op.selection());
      case WorkbookReadOperation.GetWorkbookSummary _ -> null;
      case WorkbookReadOperation.GetNamedRanges _ -> null;
      case WorkbookReadOperation.GetTables _ -> null;
      case WorkbookReadOperation.GetNamedRangeSurface _ -> null;
      case WorkbookReadOperation.AnalyzeTableHealth _ -> null;
      case WorkbookReadOperation.AnalyzeNamedRangeHealth _ -> null;
      case WorkbookReadOperation.AnalyzeWorkbookFindings _ -> null;
    };
  }

  /** Returns the address associated with one read operation or exception. */
  static String addressFor(WorkbookReadOperation read, Exception exception) {
    String fromException = GridGrindProblems.addressFor(exception);
    if (fromException != null) {
      return fromException;
    }
    return switch (read) {
      case WorkbookReadOperation.GetWindow op -> op.topLeftAddress();
      case WorkbookReadOperation.GetSheetSchema op -> op.topLeftAddress();
      case WorkbookReadOperation.GetWorkbookSummary _ -> null;
      case WorkbookReadOperation.GetNamedRanges _ -> null;
      case WorkbookReadOperation.GetSheetSummary _ -> null;
      case WorkbookReadOperation.GetCells _ -> null;
      case WorkbookReadOperation.GetMergedRegions _ -> null;
      case WorkbookReadOperation.GetHyperlinks _ -> null;
      case WorkbookReadOperation.GetComments _ -> null;
      case WorkbookReadOperation.GetSheetLayout _ -> null;
      case WorkbookReadOperation.GetDataValidations _ -> null;
      case WorkbookReadOperation.GetConditionalFormatting _ -> null;
      case WorkbookReadOperation.GetAutofilters _ -> null;
      case WorkbookReadOperation.GetTables _ -> null;
      case WorkbookReadOperation.GetFormulaSurface _ -> null;
      case WorkbookReadOperation.GetNamedRangeSurface _ -> null;
      case WorkbookReadOperation.AnalyzeFormulaHealth _ -> null;
      case WorkbookReadOperation.AnalyzeDataValidationHealth _ -> null;
      case WorkbookReadOperation.AnalyzeConditionalFormattingHealth _ -> null;
      case WorkbookReadOperation.AnalyzeAutofilterHealth _ -> null;
      case WorkbookReadOperation.AnalyzeTableHealth _ -> null;
      case WorkbookReadOperation.AnalyzeHyperlinkHealth _ -> null;
      case WorkbookReadOperation.AnalyzeNamedRangeHealth _ -> null;
      case WorkbookReadOperation.AnalyzeWorkbookFindings _ -> null;
    };
  }

  /** Returns the named-range name associated with one read operation or exception. */
  static String namedRangeNameFor(WorkbookReadOperation read, Exception exception) {
    String fromException = GridGrindProblems.namedRangeNameFor(exception);
    if (fromException != null) {
      return fromException;
    }
    return switch (read) {
      case WorkbookReadOperation.GetWorkbookSummary _ -> null;
      case WorkbookReadOperation.GetNamedRanges _ -> null;
      case WorkbookReadOperation.GetSheetSummary _ -> null;
      case WorkbookReadOperation.GetCells _ -> null;
      case WorkbookReadOperation.GetWindow _ -> null;
      case WorkbookReadOperation.GetMergedRegions _ -> null;
      case WorkbookReadOperation.GetHyperlinks _ -> null;
      case WorkbookReadOperation.GetComments _ -> null;
      case WorkbookReadOperation.GetSheetLayout _ -> null;
      case WorkbookReadOperation.GetDataValidations _ -> null;
      case WorkbookReadOperation.GetConditionalFormatting _ -> null;
      case WorkbookReadOperation.GetAutofilters _ -> null;
      case WorkbookReadOperation.GetTables _ -> null;
      case WorkbookReadOperation.GetFormulaSurface _ -> null;
      case WorkbookReadOperation.GetSheetSchema _ -> null;
      case WorkbookReadOperation.GetNamedRangeSurface op -> singleNamedRangeName(op.selection());
      case WorkbookReadOperation.AnalyzeFormulaHealth _ -> null;
      case WorkbookReadOperation.AnalyzeDataValidationHealth _ -> null;
      case WorkbookReadOperation.AnalyzeConditionalFormattingHealth _ -> null;
      case WorkbookReadOperation.AnalyzeAutofilterHealth _ -> null;
      case WorkbookReadOperation.AnalyzeTableHealth _ -> null;
      case WorkbookReadOperation.AnalyzeHyperlinkHealth _ -> null;
      case WorkbookReadOperation.AnalyzeNamedRangeHealth op -> singleNamedRangeName(op.selection());
      case WorkbookReadOperation.AnalyzeWorkbookFindings _ -> null;
    };
  }

  private static String singleSheetName(SheetSelection selection) {
    return switch (selection) {
      case SheetSelection.All _ -> null;
      case SheetSelection.Selected selected ->
          selected.sheetNames().size() == 1 ? selected.sheetNames().getFirst() : null;
    };
  }

  private static String singleNamedRangeName(NamedRangeSelection selection) {
    return switch (selection) {
      case NamedRangeSelection.All _ -> null;
      case NamedRangeSelection.Selected selected ->
          selected.selectors().size() == 1
              ? switch (selected.selectors().getFirst()) {
                case NamedRangeSelector.ByName byName -> byName.name();
                case NamedRangeSelector.WorkbookScope workbookScope -> workbookScope.name();
                case NamedRangeSelector.SheetScope sheetScope -> sheetScope.name();
              }
              : null;
    };
  }

  /**
   * Returns the formula string for the failure context: checks the exception first, then falls back
   * to the write operation when it directly carries a formula.
   */
  static String formulaFor(WorkbookOperation operation, Exception exception) {
    String fromException = GridGrindProblems.formulaFor(exception);
    if (fromException != null) {
      return fromException;
    }
    return switch (operation) {
      case WorkbookOperation.SetCell op ->
          switch (op.value()) {
            case CellInput.Formula formula -> formula.formula();
            case CellInput.Blank _ -> null;
            case CellInput.Text _ -> null;
            case CellInput.Numeric _ -> null;
            case CellInput.BooleanValue _ -> null;
            case CellInput.Date _ -> null;
            case CellInput.DateTime _ -> null;
          };
      case WorkbookOperation.EnsureSheet _ -> null;
      case WorkbookOperation.RenameSheet _ -> null;
      case WorkbookOperation.DeleteSheet _ -> null;
      case WorkbookOperation.MoveSheet _ -> null;
      case WorkbookOperation.CopySheet _ -> null;
      case WorkbookOperation.SetActiveSheet _ -> null;
      case WorkbookOperation.SetSelectedSheets _ -> null;
      case WorkbookOperation.SetSheetVisibility _ -> null;
      case WorkbookOperation.SetSheetProtection _ -> null;
      case WorkbookOperation.ClearSheetProtection _ -> null;
      case WorkbookOperation.MergeCells _ -> null;
      case WorkbookOperation.UnmergeCells _ -> null;
      case WorkbookOperation.SetColumnWidth _ -> null;
      case WorkbookOperation.SetRowHeight _ -> null;
      case WorkbookOperation.FreezePanes _ -> null;
      case WorkbookOperation.SetRange _ -> null;
      case WorkbookOperation.ClearRange _ -> null;
      case WorkbookOperation.SetHyperlink _ -> null;
      case WorkbookOperation.ClearHyperlink _ -> null;
      case WorkbookOperation.SetComment _ -> null;
      case WorkbookOperation.ClearComment _ -> null;
      case WorkbookOperation.ApplyStyle _ -> null;
      case WorkbookOperation.SetDataValidation _ -> null;
      case WorkbookOperation.ClearDataValidations _ -> null;
      case WorkbookOperation.SetConditionalFormatting _ -> null;
      case WorkbookOperation.ClearConditionalFormatting _ -> null;
      case WorkbookOperation.SetAutofilter _ -> null;
      case WorkbookOperation.ClearAutofilter _ -> null;
      case WorkbookOperation.SetTable _ -> null;
      case WorkbookOperation.DeleteTable _ -> null;
      case WorkbookOperation.SetNamedRange _ -> null;
      case WorkbookOperation.DeleteNamedRange _ -> null;
      case WorkbookOperation.AppendRow _ -> null;
      case WorkbookOperation.AutoSizeColumns _ -> null;
      case WorkbookOperation.EvaluateFormulas _ -> null;
      case WorkbookOperation.ForceFormulaRecalculationOnOpen _ -> null;
    };
  }

  /** Returns the sheet name associated with one write operation or exception. */
  static String sheetNameFor(WorkbookOperation operation, Exception exception) {
    String fromOperation =
        switch (operation) {
          case WorkbookOperation.EnsureSheet op -> op.sheetName();
          case WorkbookOperation.RenameSheet op -> op.sheetName();
          case WorkbookOperation.DeleteSheet op -> op.sheetName();
          case WorkbookOperation.MoveSheet op -> op.sheetName();
          case WorkbookOperation.CopySheet op -> op.sourceSheetName();
          case WorkbookOperation.SetActiveSheet op -> op.sheetName();
          case WorkbookOperation.SetSelectedSheets _ -> null;
          case WorkbookOperation.SetSheetVisibility op -> op.sheetName();
          case WorkbookOperation.SetSheetProtection op -> op.sheetName();
          case WorkbookOperation.ClearSheetProtection op -> op.sheetName();
          case WorkbookOperation.MergeCells op -> op.sheetName();
          case WorkbookOperation.UnmergeCells op -> op.sheetName();
          case WorkbookOperation.SetColumnWidth op -> op.sheetName();
          case WorkbookOperation.SetRowHeight op -> op.sheetName();
          case WorkbookOperation.FreezePanes op -> op.sheetName();
          case WorkbookOperation.SetCell op -> op.sheetName();
          case WorkbookOperation.SetRange op -> op.sheetName();
          case WorkbookOperation.ClearRange op -> op.sheetName();
          case WorkbookOperation.SetHyperlink op -> op.sheetName();
          case WorkbookOperation.ClearHyperlink op -> op.sheetName();
          case WorkbookOperation.SetComment op -> op.sheetName();
          case WorkbookOperation.ClearComment op -> op.sheetName();
          case WorkbookOperation.ApplyStyle op -> op.sheetName();
          case WorkbookOperation.SetDataValidation op -> op.sheetName();
          case WorkbookOperation.ClearDataValidations op -> op.sheetName();
          case WorkbookOperation.SetConditionalFormatting op -> op.sheetName();
          case WorkbookOperation.ClearConditionalFormatting op -> op.sheetName();
          case WorkbookOperation.SetAutofilter op -> op.sheetName();
          case WorkbookOperation.ClearAutofilter op -> op.sheetName();
          case WorkbookOperation.SetTable op -> op.table().sheetName();
          case WorkbookOperation.DeleteTable op -> op.sheetName();
          case WorkbookOperation.SetNamedRange op -> op.target().sheetName();
          case WorkbookOperation.DeleteNamedRange op ->
              switch (op.scope()) {
                case NamedRangeScope.Workbook _ -> null;
                case NamedRangeScope.Sheet sheetScope -> sheetScope.sheetName();
              };
          case WorkbookOperation.AppendRow op -> op.sheetName();
          case WorkbookOperation.AutoSizeColumns op -> op.sheetName();
          case WorkbookOperation.EvaluateFormulas _ -> null;
          case WorkbookOperation.ForceFormulaRecalculationOnOpen _ -> null;
        };
    return fromOperation != null ? fromOperation : GridGrindProblems.sheetNameFor(exception);
  }

  /** Returns the cell address associated with one write operation or exception. */
  static String addressFor(WorkbookOperation operation, Exception exception) {
    String fromOperation =
        switch (operation) {
          case WorkbookOperation.SetCell op -> op.address();
          case WorkbookOperation.SetHyperlink op -> op.address();
          case WorkbookOperation.ClearHyperlink op -> op.address();
          case WorkbookOperation.SetComment op -> op.address();
          case WorkbookOperation.ClearComment op -> op.address();
          case WorkbookOperation.EnsureSheet _ -> null;
          case WorkbookOperation.RenameSheet _ -> null;
          case WorkbookOperation.DeleteSheet _ -> null;
          case WorkbookOperation.MoveSheet _ -> null;
          case WorkbookOperation.CopySheet _ -> null;
          case WorkbookOperation.SetActiveSheet _ -> null;
          case WorkbookOperation.SetSelectedSheets _ -> null;
          case WorkbookOperation.SetSheetVisibility _ -> null;
          case WorkbookOperation.SetSheetProtection _ -> null;
          case WorkbookOperation.ClearSheetProtection _ -> null;
          case WorkbookOperation.MergeCells _ -> null;
          case WorkbookOperation.UnmergeCells _ -> null;
          case WorkbookOperation.SetColumnWidth _ -> null;
          case WorkbookOperation.SetRowHeight _ -> null;
          case WorkbookOperation.FreezePanes _ -> null;
          case WorkbookOperation.SetRange _ -> null;
          case WorkbookOperation.ClearRange _ -> null;
          case WorkbookOperation.ApplyStyle _ -> null;
          case WorkbookOperation.SetDataValidation _ -> null;
          case WorkbookOperation.ClearDataValidations _ -> null;
          case WorkbookOperation.SetConditionalFormatting _ -> null;
          case WorkbookOperation.ClearConditionalFormatting _ -> null;
          case WorkbookOperation.SetAutofilter _ -> null;
          case WorkbookOperation.ClearAutofilter _ -> null;
          case WorkbookOperation.SetTable _ -> null;
          case WorkbookOperation.DeleteTable _ -> null;
          case WorkbookOperation.SetNamedRange _ -> null;
          case WorkbookOperation.DeleteNamedRange _ -> null;
          case WorkbookOperation.AppendRow _ -> null;
          case WorkbookOperation.AutoSizeColumns _ -> null;
          case WorkbookOperation.EvaluateFormulas _ -> null;
          case WorkbookOperation.ForceFormulaRecalculationOnOpen _ -> null;
        };
    return fromOperation != null ? fromOperation : GridGrindProblems.addressFor(exception);
  }

  /** Returns the range string associated with one write operation or exception. */
  static String rangeFor(WorkbookOperation operation, Exception exception) {
    String fromOperation =
        switch (operation) {
          case WorkbookOperation.SetRange op -> op.range();
          case WorkbookOperation.ClearRange op -> op.range();
          case WorkbookOperation.ApplyStyle op -> op.range();
          case WorkbookOperation.SetDataValidation op -> op.range();
          case WorkbookOperation.SetConditionalFormatting op ->
              op.conditionalFormatting().ranges().size() == 1
                  ? op.conditionalFormatting().ranges().getFirst()
                  : null;
          case WorkbookOperation.SetAutofilter op -> op.range();
          case WorkbookOperation.SetTable op -> op.table().range();
          case WorkbookOperation.MergeCells op -> op.range();
          case WorkbookOperation.UnmergeCells op -> op.range();
          case WorkbookOperation.SetNamedRange op -> op.target().range();
          case WorkbookOperation.EnsureSheet _ -> null;
          case WorkbookOperation.RenameSheet _ -> null;
          case WorkbookOperation.DeleteSheet _ -> null;
          case WorkbookOperation.MoveSheet _ -> null;
          case WorkbookOperation.CopySheet _ -> null;
          case WorkbookOperation.SetActiveSheet _ -> null;
          case WorkbookOperation.SetSelectedSheets _ -> null;
          case WorkbookOperation.SetSheetVisibility _ -> null;
          case WorkbookOperation.SetSheetProtection _ -> null;
          case WorkbookOperation.ClearSheetProtection _ -> null;
          case WorkbookOperation.SetColumnWidth _ -> null;
          case WorkbookOperation.SetRowHeight _ -> null;
          case WorkbookOperation.FreezePanes _ -> null;
          case WorkbookOperation.SetCell _ -> null;
          case WorkbookOperation.SetHyperlink _ -> null;
          case WorkbookOperation.ClearHyperlink _ -> null;
          case WorkbookOperation.SetComment _ -> null;
          case WorkbookOperation.ClearComment _ -> null;
          case WorkbookOperation.AppendRow _ -> null;
          case WorkbookOperation.ClearDataValidations _ -> null;
          case WorkbookOperation.ClearConditionalFormatting _ -> null;
          case WorkbookOperation.ClearAutofilter _ -> null;
          case WorkbookOperation.DeleteTable _ -> null;
          case WorkbookOperation.DeleteNamedRange _ -> null;
          case WorkbookOperation.AutoSizeColumns _ -> null;
          case WorkbookOperation.EvaluateFormulas _ -> null;
          case WorkbookOperation.ForceFormulaRecalculationOnOpen _ -> null;
        };
    return fromOperation != null ? fromOperation : GridGrindProblems.rangeFor(exception);
  }

  /** Returns the named-range identifier associated with one write operation or exception. */
  static String namedRangeNameFor(WorkbookOperation operation, Exception exception) {
    String fromOperation =
        switch (operation) {
          case WorkbookOperation.SetNamedRange op -> op.name();
          case WorkbookOperation.DeleteNamedRange op -> op.name();
          case WorkbookOperation.EnsureSheet _ -> null;
          case WorkbookOperation.RenameSheet _ -> null;
          case WorkbookOperation.DeleteSheet _ -> null;
          case WorkbookOperation.MoveSheet _ -> null;
          case WorkbookOperation.CopySheet _ -> null;
          case WorkbookOperation.SetActiveSheet _ -> null;
          case WorkbookOperation.SetSelectedSheets _ -> null;
          case WorkbookOperation.SetSheetVisibility _ -> null;
          case WorkbookOperation.SetSheetProtection _ -> null;
          case WorkbookOperation.ClearSheetProtection _ -> null;
          case WorkbookOperation.MergeCells _ -> null;
          case WorkbookOperation.UnmergeCells _ -> null;
          case WorkbookOperation.SetColumnWidth _ -> null;
          case WorkbookOperation.SetRowHeight _ -> null;
          case WorkbookOperation.FreezePanes _ -> null;
          case WorkbookOperation.SetCell _ -> null;
          case WorkbookOperation.SetRange _ -> null;
          case WorkbookOperation.ClearRange _ -> null;
          case WorkbookOperation.SetHyperlink _ -> null;
          case WorkbookOperation.ClearHyperlink _ -> null;
          case WorkbookOperation.SetComment _ -> null;
          case WorkbookOperation.ClearComment _ -> null;
          case WorkbookOperation.ApplyStyle _ -> null;
          case WorkbookOperation.SetDataValidation _ -> null;
          case WorkbookOperation.ClearDataValidations _ -> null;
          case WorkbookOperation.SetConditionalFormatting _ -> null;
          case WorkbookOperation.ClearConditionalFormatting _ -> null;
          case WorkbookOperation.SetAutofilter _ -> null;
          case WorkbookOperation.ClearAutofilter _ -> null;
          case WorkbookOperation.SetTable _ -> null;
          case WorkbookOperation.DeleteTable _ -> null;
          case WorkbookOperation.AppendRow _ -> null;
          case WorkbookOperation.AutoSizeColumns _ -> null;
          case WorkbookOperation.EvaluateFormulas _ -> null;
          case WorkbookOperation.ForceFormulaRecalculationOnOpen _ -> null;
        };
    return fromOperation != null ? fromOperation : GridGrindProblems.namedRangeNameFor(exception);
  }

  /** Functional interface for closing an ExcelWorkbook after request execution. */
  @FunctionalInterface
  interface WorkbookCloser {
    /** Closes the workbook, releasing any held resources. */
    void close(ExcelWorkbook workbook) throws IOException;
  }
}
