package dev.erst.gridgrind.protocol.exec;

import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.protocol.dto.*;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
import java.io.IOException;
import java.nio.file.Path;
import java.time.DateTimeException;
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
          problemFor(
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
              problemFor(
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
              op.sourceSheetName(), op.newSheetName(), toExcelSheetCopyPosition(op.position()));
      case WorkbookOperation.SetActiveSheet op ->
          new WorkbookCommand.SetActiveSheet(op.sheetName());
      case WorkbookOperation.SetSelectedSheets op ->
          new WorkbookCommand.SetSelectedSheets(op.sheetNames());
      case WorkbookOperation.SetSheetVisibility op ->
          new WorkbookCommand.SetSheetVisibility(
              op.sheetName(), toExcelSheetVisibility(op.visibility()));
      case WorkbookOperation.SetSheetProtection op ->
          new WorkbookCommand.SetSheetProtection(
              op.sheetName(), toExcelSheetProtectionSettings(op.protection()));
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
      case WorkbookOperation.SetSheetPane op ->
          new WorkbookCommand.SetSheetPane(op.sheetName(), toExcelSheetPane(op.pane()));
      case WorkbookOperation.SetSheetZoom op ->
          new WorkbookCommand.SetSheetZoom(op.sheetName(), op.zoomPercent());
      case WorkbookOperation.SetPrintLayout op ->
          new WorkbookCommand.SetPrintLayout(op.sheetName(), toExcelPrintLayout(op.printLayout()));
      case WorkbookOperation.ClearPrintLayout op ->
          new WorkbookCommand.ClearPrintLayout(op.sheetName());
      case WorkbookOperation.SetCell op ->
          new WorkbookCommand.SetCell(op.sheetName(), op.address(), toExcelCellValue(op.value()));
      case WorkbookOperation.SetRange op ->
          new WorkbookCommand.SetRange(
              op.sheetName(),
              op.range(),
              op.rows().stream()
                  .map(
                      row ->
                          row.stream()
                              .map(DefaultGridGrindRequestExecutor::toExcelCellValue)
                              .toList())
                  .toList());
      case WorkbookOperation.ClearRange op ->
          new WorkbookCommand.ClearRange(op.sheetName(), op.range());
      case WorkbookOperation.SetHyperlink op ->
          new WorkbookCommand.SetHyperlink(
              op.sheetName(), op.address(), toExcelHyperlink(op.target()));
      case WorkbookOperation.ClearHyperlink op ->
          new WorkbookCommand.ClearHyperlink(op.sheetName(), op.address());
      case WorkbookOperation.SetComment op ->
          new WorkbookCommand.SetComment(
              op.sheetName(), op.address(), toExcelComment(op.comment()));
      case WorkbookOperation.ClearComment op ->
          new WorkbookCommand.ClearComment(op.sheetName(), op.address());
      case WorkbookOperation.ApplyStyle op ->
          new WorkbookCommand.ApplyStyle(op.sheetName(), op.range(), toExcelCellStyle(op.style()));
      case WorkbookOperation.SetDataValidation op ->
          new WorkbookCommand.SetDataValidation(
              op.sheetName(), op.range(), toExcelDataValidationDefinition(op.validation()));
      case WorkbookOperation.ClearDataValidations op ->
          new WorkbookCommand.ClearDataValidations(
              op.sheetName(), toExcelRangeSelection(op.selection()));
      case WorkbookOperation.SetConditionalFormatting op ->
          new WorkbookCommand.SetConditionalFormatting(
              op.sheetName(), toExcelConditionalFormattingBlock(op.conditionalFormatting()));
      case WorkbookOperation.ClearConditionalFormatting op ->
          new WorkbookCommand.ClearConditionalFormatting(
              op.sheetName(), toExcelRangeSelection(op.selection()));
      case WorkbookOperation.SetAutofilter op ->
          new WorkbookCommand.SetAutofilter(op.sheetName(), op.range());
      case WorkbookOperation.ClearAutofilter op ->
          new WorkbookCommand.ClearAutofilter(op.sheetName());
      case WorkbookOperation.SetTable op ->
          new WorkbookCommand.SetTable(toExcelTableDefinition(op.table()));
      case WorkbookOperation.DeleteTable op ->
          new WorkbookCommand.DeleteTable(op.name(), op.sheetName());
      case WorkbookOperation.SetNamedRange op ->
          new WorkbookCommand.SetNamedRange(
              new dev.erst.gridgrind.excel.ExcelNamedRangeDefinition(
                  op.name(),
                  toExcelNamedRangeScope(op.scope()),
                  toExcelNamedRangeTarget(op.target())));
      case WorkbookOperation.DeleteNamedRange op ->
          new WorkbookCommand.DeleteNamedRange(op.name(), toExcelNamedRangeScope(op.scope()));
      case WorkbookOperation.AppendRow op ->
          new WorkbookCommand.AppendRow(
              op.sheetName(),
              op.values().stream().map(DefaultGridGrindRequestExecutor::toExcelCellValue).toList());
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
      case WorkbookReadOperation.GetPrintLayout op ->
          new WorkbookReadCommand.GetPrintLayout(op.requestId(), op.sheetName());
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
                  toSheetVisibility(sheetSummary.sheet().visibility()),
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
      case dev.erst.gridgrind.excel.WorkbookReadResult.PrintLayoutResult printLayout ->
          new WorkbookReadResult.PrintLayoutResult(
              printLayout.requestId(), toPrintLayoutReport(printLayout));
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
                  .map(DefaultGridGrindRequestExecutor::toConditionalFormattingEntryReport)
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
        toPaneReport(layout.pane()),
        layout.zoomPercent(),
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

  private static PrintLayoutReport toPrintLayoutReport(
      dev.erst.gridgrind.excel.WorkbookReadResult.PrintLayoutResult printLayout) {
    return new PrintLayoutReport(
        printLayout.sheetName(),
        toPrintAreaReport(printLayout.printLayout().printArea()),
        toPrintOrientation(printLayout.printLayout().orientation()),
        toPrintScalingReport(printLayout.printLayout().scaling()),
        toPrintTitleRowsReport(printLayout.printLayout().repeatingRows()),
        toPrintTitleColumnsReport(printLayout.printLayout().repeatingColumns()),
        toHeaderFooterTextReport(printLayout.printLayout().header()),
        toHeaderFooterTextReport(printLayout.printLayout().footer()));
  }

  private static PaneReport toPaneReport(dev.erst.gridgrind.excel.ExcelSheetPane pane) {
    return switch (pane) {
      case dev.erst.gridgrind.excel.ExcelSheetPane.None _ -> new PaneReport.None();
      case dev.erst.gridgrind.excel.ExcelSheetPane.Frozen frozen ->
          new PaneReport.Frozen(
              frozen.splitColumn(), frozen.splitRow(), frozen.leftmostColumn(), frozen.topRow());
      case dev.erst.gridgrind.excel.ExcelSheetPane.Split split ->
          new PaneReport.Split(
              split.xSplitPosition(),
              split.ySplitPosition(),
              split.leftmostColumn(),
              split.topRow(),
              toPaneRegion(split.activePane()));
    };
  }

  private static PrintAreaReport toPrintAreaReport(
      dev.erst.gridgrind.excel.ExcelPrintLayout.Area printArea) {
    return switch (printArea) {
      case dev.erst.gridgrind.excel.ExcelPrintLayout.Area.None _ -> new PrintAreaReport.None();
      case dev.erst.gridgrind.excel.ExcelPrintLayout.Area.Range range ->
          new PrintAreaReport.Range(range.range());
    };
  }

  private static PrintScalingReport toPrintScalingReport(
      dev.erst.gridgrind.excel.ExcelPrintLayout.Scaling scaling) {
    return switch (scaling) {
      case dev.erst.gridgrind.excel.ExcelPrintLayout.Scaling.Automatic _ ->
          new PrintScalingReport.Automatic();
      case dev.erst.gridgrind.excel.ExcelPrintLayout.Scaling.Fit fit ->
          new PrintScalingReport.Fit(fit.widthPages(), fit.heightPages());
    };
  }

  private static PrintTitleRowsReport toPrintTitleRowsReport(
      dev.erst.gridgrind.excel.ExcelPrintLayout.TitleRows repeatingRows) {
    return switch (repeatingRows) {
      case dev.erst.gridgrind.excel.ExcelPrintLayout.TitleRows.None _ ->
          new PrintTitleRowsReport.None();
      case dev.erst.gridgrind.excel.ExcelPrintLayout.TitleRows.Band band ->
          new PrintTitleRowsReport.Band(band.firstRowIndex(), band.lastRowIndex());
    };
  }

  private static PrintTitleColumnsReport toPrintTitleColumnsReport(
      dev.erst.gridgrind.excel.ExcelPrintLayout.TitleColumns repeatingColumns) {
    return switch (repeatingColumns) {
      case dev.erst.gridgrind.excel.ExcelPrintLayout.TitleColumns.None _ ->
          new PrintTitleColumnsReport.None();
      case dev.erst.gridgrind.excel.ExcelPrintLayout.TitleColumns.Band band ->
          new PrintTitleColumnsReport.Band(band.firstColumnIndex(), band.lastColumnIndex());
    };
  }

  private static HeaderFooterTextReport toHeaderFooterTextReport(
      dev.erst.gridgrind.excel.ExcelHeaderFooterText text) {
    return new HeaderFooterTextReport(text.left(), text.center(), text.right());
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
        toAnalysisFindingCode(finding.code()),
        toAnalysisSeverity(finding.severity()),
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

  static DataValidationEntryReport toDataValidationEntryReport(
      dev.erst.gridgrind.excel.ExcelDataValidationSnapshot snapshot) {
    return switch (snapshot) {
      case dev.erst.gridgrind.excel.ExcelDataValidationSnapshot.Supported supported ->
          new DataValidationEntryReport.Supported(
              supported.ranges(), toDataValidationDefinitionReport(supported.validation()));
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

  static dev.erst.gridgrind.excel.ExcelCellValue toExcelCellValue(CellInput value) {
    return switch (value) {
      case CellInput.Blank _ -> dev.erst.gridgrind.excel.ExcelCellValue.blank();
      case CellInput.Text text -> dev.erst.gridgrind.excel.ExcelCellValue.text(text.text());
      case CellInput.Numeric numeric ->
          dev.erst.gridgrind.excel.ExcelCellValue.number(numeric.number());
      case CellInput.BooleanValue booleanValue ->
          dev.erst.gridgrind.excel.ExcelCellValue.bool(booleanValue.bool());
      case CellInput.Date date -> dev.erst.gridgrind.excel.ExcelCellValue.date(date.date());
      case CellInput.DateTime dateTime ->
          dev.erst.gridgrind.excel.ExcelCellValue.dateTime(dateTime.dateTime());
      case CellInput.Formula formula ->
          dev.erst.gridgrind.excel.ExcelCellValue.formula(formula.formula());
    };
  }

  static dev.erst.gridgrind.excel.ExcelHyperlink toExcelHyperlink(HyperlinkTarget target) {
    return switch (target) {
      case HyperlinkTarget.Url url -> new dev.erst.gridgrind.excel.ExcelHyperlink.Url(url.target());
      case HyperlinkTarget.Email email ->
          new dev.erst.gridgrind.excel.ExcelHyperlink.Email(email.email());
      case HyperlinkTarget.File file ->
          new dev.erst.gridgrind.excel.ExcelHyperlink.File(file.path());
      case HyperlinkTarget.Document document ->
          new dev.erst.gridgrind.excel.ExcelHyperlink.Document(document.target());
    };
  }

  static ExcelComment toExcelComment(CommentInput comment) {
    return new ExcelComment(comment.text(), comment.author(), comment.visible());
  }

  static dev.erst.gridgrind.excel.ExcelCellStyle toExcelCellStyle(CellStyleInput style) {
    return new dev.erst.gridgrind.excel.ExcelCellStyle(
        style.numberFormat(),
        style.bold(),
        style.italic(),
        style.wrapText(),
        toExcelHorizontalAlignment(style.horizontalAlignment()),
        toExcelVerticalAlignment(style.verticalAlignment()),
        style.fontName(),
        toExcelFontHeight(style.fontHeight()),
        style.fontColor(),
        style.underline(),
        style.strikeout(),
        style.fillColor(),
        toExcelBorder(style.border()));
  }

  static dev.erst.gridgrind.excel.ExcelFontHeight toExcelFontHeight(FontHeightInput fontHeight) {
    if (fontHeight == null) {
      return null;
    }
    return switch (fontHeight) {
      case FontHeightInput.Points points ->
          dev.erst.gridgrind.excel.ExcelFontHeight.fromPoints(points.points());
      case FontHeightInput.Twips twips ->
          new dev.erst.gridgrind.excel.ExcelFontHeight(twips.twips());
    };
  }

  private static dev.erst.gridgrind.excel.ExcelHorizontalAlignment toExcelHorizontalAlignment(
      HorizontalAlignment alignment) {
    if (alignment == null) {
      return null;
    }
    return switch (alignment) {
      case GENERAL -> dev.erst.gridgrind.excel.ExcelHorizontalAlignment.GENERAL;
      case LEFT -> dev.erst.gridgrind.excel.ExcelHorizontalAlignment.LEFT;
      case CENTER -> dev.erst.gridgrind.excel.ExcelHorizontalAlignment.CENTER;
      case RIGHT -> dev.erst.gridgrind.excel.ExcelHorizontalAlignment.RIGHT;
      case FILL -> dev.erst.gridgrind.excel.ExcelHorizontalAlignment.FILL;
      case JUSTIFY -> dev.erst.gridgrind.excel.ExcelHorizontalAlignment.JUSTIFY;
      case CENTER_SELECTION -> dev.erst.gridgrind.excel.ExcelHorizontalAlignment.CENTER_SELECTION;
      case DISTRIBUTED -> dev.erst.gridgrind.excel.ExcelHorizontalAlignment.DISTRIBUTED;
    };
  }

  private static dev.erst.gridgrind.excel.ExcelVerticalAlignment toExcelVerticalAlignment(
      VerticalAlignment alignment) {
    if (alignment == null) {
      return null;
    }
    return switch (alignment) {
      case TOP -> dev.erst.gridgrind.excel.ExcelVerticalAlignment.TOP;
      case CENTER -> dev.erst.gridgrind.excel.ExcelVerticalAlignment.CENTER;
      case BOTTOM -> dev.erst.gridgrind.excel.ExcelVerticalAlignment.BOTTOM;
      case JUSTIFY -> dev.erst.gridgrind.excel.ExcelVerticalAlignment.JUSTIFY;
      case DISTRIBUTED -> dev.erst.gridgrind.excel.ExcelVerticalAlignment.DISTRIBUTED;
    };
  }

  static dev.erst.gridgrind.excel.ExcelBorder toExcelBorder(CellBorderInput border) {
    if (border == null) {
      return null;
    }
    return new dev.erst.gridgrind.excel.ExcelBorder(
        toExcelBorderSide(border.all()),
        toExcelBorderSide(border.top()),
        toExcelBorderSide(border.right()),
        toExcelBorderSide(border.bottom()),
        toExcelBorderSide(border.left()));
  }

  static dev.erst.gridgrind.excel.ExcelBorderSide toExcelBorderSide(CellBorderSideInput side) {
    return side == null
        ? null
        : new dev.erst.gridgrind.excel.ExcelBorderSide(toExcelBorderStyle(side.style()));
  }

  static dev.erst.gridgrind.excel.ExcelDataValidationDefinition toExcelDataValidationDefinition(
      DataValidationInput validation) {
    return new dev.erst.gridgrind.excel.ExcelDataValidationDefinition(
        toExcelDataValidationRule(validation.rule()),
        validation.allowBlank(),
        validation.suppressDropDownArrow(),
        toExcelDataValidationPrompt(validation.prompt()),
        toExcelDataValidationErrorAlert(validation.errorAlert()));
  }

  static dev.erst.gridgrind.excel.ExcelDataValidationRule toExcelDataValidationRule(
      DataValidationRuleInput rule) {
    return switch (rule) {
      case DataValidationRuleInput.ExplicitList explicitList ->
          new dev.erst.gridgrind.excel.ExcelDataValidationRule.ExplicitList(explicitList.values());
      case DataValidationRuleInput.FormulaList formulaList ->
          new dev.erst.gridgrind.excel.ExcelDataValidationRule.FormulaList(formulaList.formula());
      case DataValidationRuleInput.WholeNumber wholeNumber ->
          new dev.erst.gridgrind.excel.ExcelDataValidationRule.WholeNumber(
              toExcelComparisonOperator(wholeNumber.operator()),
              wholeNumber.formula1(),
              wholeNumber.formula2());
      case DataValidationRuleInput.DecimalNumber decimalNumber ->
          new dev.erst.gridgrind.excel.ExcelDataValidationRule.DecimalNumber(
              toExcelComparisonOperator(decimalNumber.operator()),
              decimalNumber.formula1(),
              decimalNumber.formula2());
      case DataValidationRuleInput.DateRule dateRule ->
          new dev.erst.gridgrind.excel.ExcelDataValidationRule.DateRule(
              toExcelComparisonOperator(dateRule.operator()),
              dateRule.formula1(),
              dateRule.formula2());
      case DataValidationRuleInput.TimeRule timeRule ->
          new dev.erst.gridgrind.excel.ExcelDataValidationRule.TimeRule(
              toExcelComparisonOperator(timeRule.operator()),
              timeRule.formula1(),
              timeRule.formula2());
      case DataValidationRuleInput.TextLength textLength ->
          new dev.erst.gridgrind.excel.ExcelDataValidationRule.TextLength(
              toExcelComparisonOperator(textLength.operator()),
              textLength.formula1(),
              textLength.formula2());
      case DataValidationRuleInput.CustomFormula customFormula ->
          new dev.erst.gridgrind.excel.ExcelDataValidationRule.CustomFormula(
              customFormula.formula());
    };
  }

  static dev.erst.gridgrind.excel.ExcelDataValidationPrompt toExcelDataValidationPrompt(
      DataValidationPromptInput prompt) {
    return prompt == null
        ? null
        : new dev.erst.gridgrind.excel.ExcelDataValidationPrompt(
            prompt.title(), prompt.text(), prompt.showPromptBox());
  }

  static dev.erst.gridgrind.excel.ExcelDataValidationErrorAlert toExcelDataValidationErrorAlert(
      DataValidationErrorAlertInput errorAlert) {
    return errorAlert == null
        ? null
        : new dev.erst.gridgrind.excel.ExcelDataValidationErrorAlert(
            toExcelDataValidationErrorStyle(errorAlert.style()),
            errorAlert.title(),
            errorAlert.text(),
            errorAlert.showErrorBox());
  }

  private static dev.erst.gridgrind.excel.ExcelDataValidationErrorStyle
      toExcelDataValidationErrorStyle(DataValidationErrorStyle style) {
    return switch (style) {
      case STOP -> dev.erst.gridgrind.excel.ExcelDataValidationErrorStyle.STOP;
      case WARNING -> dev.erst.gridgrind.excel.ExcelDataValidationErrorStyle.WARNING;
      case INFORMATION -> dev.erst.gridgrind.excel.ExcelDataValidationErrorStyle.INFORMATION;
    };
  }

  static dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition
      toExcelConditionalFormattingBlock(ConditionalFormattingBlockInput block) {
    return new dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition(
        block.ranges(),
        block.rules().stream()
            .map(DefaultGridGrindRequestExecutor::toExcelConditionalFormattingRule)
            .toList());
  }

  static dev.erst.gridgrind.excel.ExcelConditionalFormattingRule toExcelConditionalFormattingRule(
      ConditionalFormattingRuleInput rule) {
    return switch (rule) {
      case ConditionalFormattingRuleInput.FormulaRule formulaRule ->
          new dev.erst.gridgrind.excel.ExcelConditionalFormattingRule.FormulaRule(
              formulaRule.formula(),
              formulaRule.stopIfTrue(),
              toExcelDifferentialStyle(formulaRule.style()));
      case ConditionalFormattingRuleInput.CellValueRule cellValueRule ->
          new dev.erst.gridgrind.excel.ExcelConditionalFormattingRule.CellValueRule(
              toExcelComparisonOperator(cellValueRule.operator()),
              cellValueRule.formula1(),
              cellValueRule.formula2(),
              cellValueRule.stopIfTrue(),
              toExcelDifferentialStyle(cellValueRule.style()));
    };
  }

  static dev.erst.gridgrind.excel.ExcelDifferentialStyle toExcelDifferentialStyle(
      DifferentialStyleInput style) {
    return new dev.erst.gridgrind.excel.ExcelDifferentialStyle(
        style.numberFormat(),
        style.bold(),
        style.italic(),
        toExcelFontHeight(style.fontHeight()),
        style.fontColor(),
        style.underline(),
        style.strikeout(),
        style.fillColor(),
        toExcelDifferentialBorder(style.border()));
  }

  static dev.erst.gridgrind.excel.ExcelDifferentialBorder toExcelDifferentialBorder(
      DifferentialBorderInput border) {
    if (border == null) {
      return null;
    }
    return new dev.erst.gridgrind.excel.ExcelDifferentialBorder(
        toExcelDifferentialBorderSide(border.all()),
        toExcelDifferentialBorderSide(border.top()),
        toExcelDifferentialBorderSide(border.right()),
        toExcelDifferentialBorderSide(border.bottom()),
        toExcelDifferentialBorderSide(border.left()));
  }

  private static dev.erst.gridgrind.excel.ExcelDifferentialBorderSide toExcelDifferentialBorderSide(
      DifferentialBorderSideInput side) {
    return side == null
        ? null
        : new dev.erst.gridgrind.excel.ExcelDifferentialBorderSide(
            toExcelBorderStyle(side.style()), side.color());
  }

  private static dev.erst.gridgrind.excel.ExcelComparisonOperator toExcelComparisonOperator(
      ComparisonOperator operator) {
    return switch (operator) {
      case BETWEEN -> dev.erst.gridgrind.excel.ExcelComparisonOperator.BETWEEN;
      case NOT_BETWEEN -> dev.erst.gridgrind.excel.ExcelComparisonOperator.NOT_BETWEEN;
      case EQUAL -> dev.erst.gridgrind.excel.ExcelComparisonOperator.EQUAL;
      case NOT_EQUAL -> dev.erst.gridgrind.excel.ExcelComparisonOperator.NOT_EQUAL;
      case GREATER_THAN -> dev.erst.gridgrind.excel.ExcelComparisonOperator.GREATER_THAN;
      case LESS_THAN -> dev.erst.gridgrind.excel.ExcelComparisonOperator.LESS_THAN;
      case GREATER_OR_EQUAL -> dev.erst.gridgrind.excel.ExcelComparisonOperator.GREATER_OR_EQUAL;
      case LESS_OR_EQUAL -> dev.erst.gridgrind.excel.ExcelComparisonOperator.LESS_OR_EQUAL;
    };
  }

  static dev.erst.gridgrind.excel.ExcelTableDefinition toExcelTableDefinition(TableInput table) {
    return new dev.erst.gridgrind.excel.ExcelTableDefinition(
        table.name(),
        table.sheetName(),
        table.range(),
        table.showTotalsRow(),
        toExcelTableStyle(table.style()));
  }

  static dev.erst.gridgrind.excel.ExcelTableStyle toExcelTableStyle(TableStyleInput style) {
    return switch (style) {
      case TableStyleInput.None _ -> new dev.erst.gridgrind.excel.ExcelTableStyle.None();
      case TableStyleInput.Named named ->
          new dev.erst.gridgrind.excel.ExcelTableStyle.Named(
              named.name(),
              named.showFirstColumn(),
              named.showLastColumn(),
              named.showRowStripes(),
              named.showColumnStripes());
    };
  }

  static dev.erst.gridgrind.excel.ExcelNamedRangeScope toExcelNamedRangeScope(
      NamedRangeScope scope) {
    return switch (scope) {
      case NamedRangeScope.Workbook _ ->
          new dev.erst.gridgrind.excel.ExcelNamedRangeScope.WorkbookScope();
      case NamedRangeScope.Sheet sheet ->
          new dev.erst.gridgrind.excel.ExcelNamedRangeScope.SheetScope(sheet.sheetName());
    };
  }

  static dev.erst.gridgrind.excel.ExcelNamedRangeTarget toExcelNamedRangeTarget(
      NamedRangeTarget target) {
    return new dev.erst.gridgrind.excel.ExcelNamedRangeTarget(target.sheetName(), target.range());
  }

  static dev.erst.gridgrind.excel.ExcelSheetCopyPosition toExcelSheetCopyPosition(
      SheetCopyPosition position) {
    return switch (position) {
      case SheetCopyPosition.AppendAtEnd _ ->
          new dev.erst.gridgrind.excel.ExcelSheetCopyPosition.AppendAtEnd();
      case SheetCopyPosition.AtIndex atIndex ->
          new dev.erst.gridgrind.excel.ExcelSheetCopyPosition.AtIndex(atIndex.targetIndex());
    };
  }

  private static dev.erst.gridgrind.excel.ExcelSheetVisibility toExcelSheetVisibility(
      SheetVisibility visibility) {
    return switch (visibility) {
      case VISIBLE -> dev.erst.gridgrind.excel.ExcelSheetVisibility.VISIBLE;
      case HIDDEN -> dev.erst.gridgrind.excel.ExcelSheetVisibility.HIDDEN;
      case VERY_HIDDEN -> dev.erst.gridgrind.excel.ExcelSheetVisibility.VERY_HIDDEN;
    };
  }

  private static SheetVisibility toSheetVisibility(
      dev.erst.gridgrind.excel.ExcelSheetVisibility visibility) {
    return switch (visibility) {
      case VISIBLE -> SheetVisibility.VISIBLE;
      case HIDDEN -> SheetVisibility.HIDDEN;
      case VERY_HIDDEN -> SheetVisibility.VERY_HIDDEN;
    };
  }

  private static dev.erst.gridgrind.excel.ExcelSheetProtectionSettings
      toExcelSheetProtectionSettings(SheetProtectionSettings settings) {
    return new dev.erst.gridgrind.excel.ExcelSheetProtectionSettings(
        settings.autoFilterLocked(),
        settings.deleteColumnsLocked(),
        settings.deleteRowsLocked(),
        settings.formatCellsLocked(),
        settings.formatColumnsLocked(),
        settings.formatRowsLocked(),
        settings.insertColumnsLocked(),
        settings.insertHyperlinksLocked(),
        settings.insertRowsLocked(),
        settings.objectsLocked(),
        settings.pivotTablesLocked(),
        settings.scenariosLocked(),
        settings.selectLockedCellsLocked(),
        settings.selectUnlockedCellsLocked(),
        settings.sortLocked());
  }

  private static SheetProtectionSettings toSheetProtectionSettings(
      dev.erst.gridgrind.excel.ExcelSheetProtectionSettings settings) {
    return new SheetProtectionSettings(
        settings.autoFilterLocked(),
        settings.deleteColumnsLocked(),
        settings.deleteRowsLocked(),
        settings.formatCellsLocked(),
        settings.formatColumnsLocked(),
        settings.formatRowsLocked(),
        settings.insertColumnsLocked(),
        settings.insertHyperlinksLocked(),
        settings.insertRowsLocked(),
        settings.objectsLocked(),
        settings.pivotTablesLocked(),
        settings.scenariosLocked(),
        settings.selectLockedCellsLocked(),
        settings.selectUnlockedCellsLocked(),
        settings.sortLocked());
  }

  private static dev.erst.gridgrind.excel.ExcelPaneRegion toExcelPaneRegion(PaneRegion region) {
    return switch (region) {
      case UPPER_LEFT -> dev.erst.gridgrind.excel.ExcelPaneRegion.UPPER_LEFT;
      case UPPER_RIGHT -> dev.erst.gridgrind.excel.ExcelPaneRegion.UPPER_RIGHT;
      case LOWER_LEFT -> dev.erst.gridgrind.excel.ExcelPaneRegion.LOWER_LEFT;
      case LOWER_RIGHT -> dev.erst.gridgrind.excel.ExcelPaneRegion.LOWER_RIGHT;
    };
  }

  private static PaneRegion toPaneRegion(dev.erst.gridgrind.excel.ExcelPaneRegion region) {
    return switch (region) {
      case UPPER_LEFT -> PaneRegion.UPPER_LEFT;
      case UPPER_RIGHT -> PaneRegion.UPPER_RIGHT;
      case LOWER_LEFT -> PaneRegion.LOWER_LEFT;
      case LOWER_RIGHT -> PaneRegion.LOWER_RIGHT;
    };
  }

  private static dev.erst.gridgrind.excel.ExcelPrintOrientation toExcelPrintOrientation(
      PrintOrientation orientation) {
    return switch (orientation) {
      case PORTRAIT -> dev.erst.gridgrind.excel.ExcelPrintOrientation.PORTRAIT;
      case LANDSCAPE -> dev.erst.gridgrind.excel.ExcelPrintOrientation.LANDSCAPE;
    };
  }

  private static PrintOrientation toPrintOrientation(
      dev.erst.gridgrind.excel.ExcelPrintOrientation orientation) {
    return switch (orientation) {
      case PORTRAIT -> PrintOrientation.PORTRAIT;
      case LANDSCAPE -> PrintOrientation.LANDSCAPE;
    };
  }

  static FontHeightReport toFontHeightReport(dev.erst.gridgrind.excel.ExcelFontHeight fontHeight) {
    return fontHeight == null
        ? null
        : new FontHeightReport(fontHeight.twips(), fontHeight.points());
  }

  static HorizontalAlignment toHorizontalAlignment(
      dev.erst.gridgrind.excel.ExcelHorizontalAlignment alignment) {
    return switch (alignment) {
      case GENERAL -> HorizontalAlignment.GENERAL;
      case LEFT -> HorizontalAlignment.LEFT;
      case CENTER -> HorizontalAlignment.CENTER;
      case RIGHT -> HorizontalAlignment.RIGHT;
      case FILL -> HorizontalAlignment.FILL;
      case JUSTIFY -> HorizontalAlignment.JUSTIFY;
      case CENTER_SELECTION -> HorizontalAlignment.CENTER_SELECTION;
      case DISTRIBUTED -> HorizontalAlignment.DISTRIBUTED;
    };
  }

  static VerticalAlignment toVerticalAlignment(
      dev.erst.gridgrind.excel.ExcelVerticalAlignment alignment) {
    return switch (alignment) {
      case TOP -> VerticalAlignment.TOP;
      case CENTER -> VerticalAlignment.CENTER;
      case BOTTOM -> VerticalAlignment.BOTTOM;
      case JUSTIFY -> VerticalAlignment.JUSTIFY;
      case DISTRIBUTED -> VerticalAlignment.DISTRIBUTED;
    };
  }

  static BorderStyle toBorderStyle(dev.erst.gridgrind.excel.ExcelBorderStyle style) {
    return switch (style) {
      case NONE -> BorderStyle.NONE;
      case THIN -> BorderStyle.THIN;
      case MEDIUM -> BorderStyle.MEDIUM;
      case DASHED -> BorderStyle.DASHED;
      case DOTTED -> BorderStyle.DOTTED;
      case THICK -> BorderStyle.THICK;
      case DOUBLE -> BorderStyle.DOUBLE;
      case HAIR -> BorderStyle.HAIR;
      case MEDIUM_DASHED -> BorderStyle.MEDIUM_DASHED;
      case DASH_DOT -> BorderStyle.DASH_DOT;
      case MEDIUM_DASH_DOT -> BorderStyle.MEDIUM_DASH_DOT;
      case DASH_DOT_DOT -> BorderStyle.DASH_DOT_DOT;
      case MEDIUM_DASH_DOT_DOT -> BorderStyle.MEDIUM_DASH_DOT_DOT;
      case SLANTED_DASH_DOT -> BorderStyle.SLANTED_DASH_DOT;
    };
  }

  private static dev.erst.gridgrind.excel.ExcelBorderStyle toExcelBorderStyle(BorderStyle style) {
    return switch (style) {
      case NONE -> dev.erst.gridgrind.excel.ExcelBorderStyle.NONE;
      case THIN -> dev.erst.gridgrind.excel.ExcelBorderStyle.THIN;
      case MEDIUM -> dev.erst.gridgrind.excel.ExcelBorderStyle.MEDIUM;
      case DASHED -> dev.erst.gridgrind.excel.ExcelBorderStyle.DASHED;
      case DOTTED -> dev.erst.gridgrind.excel.ExcelBorderStyle.DOTTED;
      case THICK -> dev.erst.gridgrind.excel.ExcelBorderStyle.THICK;
      case DOUBLE -> dev.erst.gridgrind.excel.ExcelBorderStyle.DOUBLE;
      case HAIR -> dev.erst.gridgrind.excel.ExcelBorderStyle.HAIR;
      case MEDIUM_DASHED -> dev.erst.gridgrind.excel.ExcelBorderStyle.MEDIUM_DASHED;
      case DASH_DOT -> dev.erst.gridgrind.excel.ExcelBorderStyle.DASH_DOT;
      case MEDIUM_DASH_DOT -> dev.erst.gridgrind.excel.ExcelBorderStyle.MEDIUM_DASH_DOT;
      case DASH_DOT_DOT -> dev.erst.gridgrind.excel.ExcelBorderStyle.DASH_DOT_DOT;
      case MEDIUM_DASH_DOT_DOT -> dev.erst.gridgrind.excel.ExcelBorderStyle.MEDIUM_DASH_DOT_DOT;
      case SLANTED_DASH_DOT -> dev.erst.gridgrind.excel.ExcelBorderStyle.SLANTED_DASH_DOT;
    };
  }

  static DataValidationEntryReport.DataValidationDefinitionReport toDataValidationDefinitionReport(
      dev.erst.gridgrind.excel.ExcelDataValidationDefinition definition) {
    return new DataValidationEntryReport.DataValidationDefinitionReport(
        toDataValidationRuleInput(definition.rule()),
        definition.allowBlank(),
        definition.suppressDropDownArrow(),
        definition.prompt() == null
            ? null
            : new DataValidationPromptInput(
                definition.prompt().title(),
                definition.prompt().text(),
                definition.prompt().showPromptBox()),
        definition.errorAlert() == null
            ? null
            : new DataValidationErrorAlertInput(
                toDataValidationErrorStyle(definition.errorAlert().style()),
                definition.errorAlert().title(),
                definition.errorAlert().text(),
                definition.errorAlert().showErrorBox()));
  }

  private static DataValidationRuleInput toDataValidationRuleInput(
      dev.erst.gridgrind.excel.ExcelDataValidationRule rule) {
    return switch (rule) {
      case dev.erst.gridgrind.excel.ExcelDataValidationRule.ExplicitList explicitList ->
          new DataValidationRuleInput.ExplicitList(explicitList.values());
      case dev.erst.gridgrind.excel.ExcelDataValidationRule.FormulaList formulaList ->
          new DataValidationRuleInput.FormulaList(formulaList.formula());
      case dev.erst.gridgrind.excel.ExcelDataValidationRule.WholeNumber wholeNumber ->
          new DataValidationRuleInput.WholeNumber(
              toComparisonOperator(wholeNumber.operator()),
              wholeNumber.formula1(),
              wholeNumber.formula2());
      case dev.erst.gridgrind.excel.ExcelDataValidationRule.DecimalNumber decimalNumber ->
          new DataValidationRuleInput.DecimalNumber(
              toComparisonOperator(decimalNumber.operator()),
              decimalNumber.formula1(),
              decimalNumber.formula2());
      case dev.erst.gridgrind.excel.ExcelDataValidationRule.DateRule dateRule ->
          new DataValidationRuleInput.DateRule(
              toComparisonOperator(dateRule.operator()), dateRule.formula1(), dateRule.formula2());
      case dev.erst.gridgrind.excel.ExcelDataValidationRule.TimeRule timeRule ->
          new DataValidationRuleInput.TimeRule(
              toComparisonOperator(timeRule.operator()), timeRule.formula1(), timeRule.formula2());
      case dev.erst.gridgrind.excel.ExcelDataValidationRule.TextLength textLength ->
          new DataValidationRuleInput.TextLength(
              toComparisonOperator(textLength.operator()),
              textLength.formula1(),
              textLength.formula2());
      case dev.erst.gridgrind.excel.ExcelDataValidationRule.CustomFormula customFormula ->
          new DataValidationRuleInput.CustomFormula(customFormula.formula());
    };
  }

  private static DataValidationErrorStyle toDataValidationErrorStyle(
      dev.erst.gridgrind.excel.ExcelDataValidationErrorStyle style) {
    return switch (style) {
      case STOP -> DataValidationErrorStyle.STOP;
      case WARNING -> DataValidationErrorStyle.WARNING;
      case INFORMATION -> DataValidationErrorStyle.INFORMATION;
    };
  }

  static ConditionalFormattingEntryReport toConditionalFormattingEntryReport(
      dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockSnapshot block) {
    return new ConditionalFormattingEntryReport(
        block.ranges(),
        block.rules().stream()
            .map(DefaultGridGrindRequestExecutor::toConditionalFormattingRuleReport)
            .toList());
  }

  static ConditionalFormattingRuleReport toConditionalFormattingRuleReport(
      dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot rule) {
    return switch (rule) {
      case dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot.FormulaRule
              formulaRule ->
          new ConditionalFormattingRuleReport.FormulaRule(
              formulaRule.priority(),
              formulaRule.stopIfTrue(),
              formulaRule.formula(),
              toDifferentialStyleReport(formulaRule.style()));
      case dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot.CellValueRule
              cellValueRule ->
          new ConditionalFormattingRuleReport.CellValueRule(
              cellValueRule.priority(),
              cellValueRule.stopIfTrue(),
              toComparisonOperator(cellValueRule.operator()),
              cellValueRule.formula1(),
              cellValueRule.formula2(),
              toDifferentialStyleReport(cellValueRule.style()));
      case dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot.ColorScaleRule
              colorScaleRule ->
          new ConditionalFormattingRuleReport.ColorScaleRule(
              colorScaleRule.priority(),
              colorScaleRule.stopIfTrue(),
              colorScaleRule.thresholds().stream()
                  .map(DefaultGridGrindRequestExecutor::toConditionalFormattingThresholdReport)
                  .toList(),
              colorScaleRule.colors());
      case dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot.DataBarRule
              dataBarRule ->
          new ConditionalFormattingRuleReport.DataBarRule(
              dataBarRule.priority(),
              dataBarRule.stopIfTrue(),
              dataBarRule.color(),
              dataBarRule.iconOnly(),
              dataBarRule.leftToRight(),
              dataBarRule.widthMin(),
              dataBarRule.widthMax(),
              toConditionalFormattingThresholdReport(dataBarRule.minThreshold()),
              toConditionalFormattingThresholdReport(dataBarRule.maxThreshold()));
      case dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot.IconSetRule
              iconSetRule ->
          new ConditionalFormattingRuleReport.IconSetRule(
              iconSetRule.priority(),
              iconSetRule.stopIfTrue(),
              toConditionalFormattingIconSet(iconSetRule.iconSet()),
              iconSetRule.iconOnly(),
              iconSetRule.reversed(),
              iconSetRule.thresholds().stream()
                  .map(DefaultGridGrindRequestExecutor::toConditionalFormattingThresholdReport)
                  .toList());
      case dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot.UnsupportedRule
              unsupportedRule ->
          new ConditionalFormattingRuleReport.UnsupportedRule(
              unsupportedRule.priority(),
              unsupportedRule.stopIfTrue(),
              unsupportedRule.kind(),
              unsupportedRule.detail());
    };
  }

  static ConditionalFormattingThresholdReport toConditionalFormattingThresholdReport(
      dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdSnapshot threshold) {
    return new ConditionalFormattingThresholdReport(
        toConditionalFormattingThresholdType(threshold.type()),
        threshold.formula(),
        threshold.value());
  }

  private static ConditionalFormattingIconSet toConditionalFormattingIconSet(
      dev.erst.gridgrind.excel.ExcelConditionalFormattingIconSet iconSet) {
    return switch (iconSet) {
      case GYR_3_ARROW -> ConditionalFormattingIconSet.GYR_3_ARROW;
      case GREY_3_ARROWS -> ConditionalFormattingIconSet.GREY_3_ARROWS;
      case GYR_3_FLAGS -> ConditionalFormattingIconSet.GYR_3_FLAGS;
      case GYR_3_TRAFFIC_LIGHTS -> ConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS;
      case GYR_3_TRAFFIC_LIGHTS_BOX -> ConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS_BOX;
      case GYR_3_SHAPES -> ConditionalFormattingIconSet.GYR_3_SHAPES;
      case GYR_3_SYMBOLS_CIRCLE -> ConditionalFormattingIconSet.GYR_3_SYMBOLS_CIRCLE;
      case GYR_3_SYMBOLS -> ConditionalFormattingIconSet.GYR_3_SYMBOLS;
      case GYR_4_ARROWS -> ConditionalFormattingIconSet.GYR_4_ARROWS;
      case GREY_4_ARROWS -> ConditionalFormattingIconSet.GREY_4_ARROWS;
      case RB_4_TRAFFIC_LIGHTS -> ConditionalFormattingIconSet.RB_4_TRAFFIC_LIGHTS;
      case RATINGS_4 -> ConditionalFormattingIconSet.RATINGS_4;
      case GYRB_4_TRAFFIC_LIGHTS -> ConditionalFormattingIconSet.GYRB_4_TRAFFIC_LIGHTS;
      case GYYYR_5_ARROWS -> ConditionalFormattingIconSet.GYYYR_5_ARROWS;
      case GREY_5_ARROWS -> ConditionalFormattingIconSet.GREY_5_ARROWS;
      case RATINGS_5 -> ConditionalFormattingIconSet.RATINGS_5;
      case QUARTERS_5 -> ConditionalFormattingIconSet.QUARTERS_5;
    };
  }

  private static ConditionalFormattingThresholdType toConditionalFormattingThresholdType(
      dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdType type) {
    return switch (type) {
      case NUMBER -> ConditionalFormattingThresholdType.NUMBER;
      case MIN -> ConditionalFormattingThresholdType.MIN;
      case MAX -> ConditionalFormattingThresholdType.MAX;
      case PERCENT -> ConditionalFormattingThresholdType.PERCENT;
      case PERCENTILE -> ConditionalFormattingThresholdType.PERCENTILE;
      case UNALLOCATED -> ConditionalFormattingThresholdType.UNALLOCATED;
      case FORMULA -> ConditionalFormattingThresholdType.FORMULA;
    };
  }

  static DifferentialStyleReport toDifferentialStyleReport(
      dev.erst.gridgrind.excel.ExcelDifferentialStyleSnapshot style) {
    if (style == null) {
      return null;
    }
    return new DifferentialStyleReport(
        style.numberFormat(),
        style.bold(),
        style.italic(),
        toFontHeightReport(style.fontHeight()),
        style.fontColor(),
        style.underline(),
        style.strikeout(),
        style.fillColor(),
        toDifferentialBorderReport(style.border()),
        style.unsupportedFeatures().stream()
            .map(DefaultGridGrindRequestExecutor::toConditionalFormattingUnsupportedFeature)
            .toList());
  }

  static DifferentialBorderReport toDifferentialBorderReport(
      dev.erst.gridgrind.excel.ExcelDifferentialBorder border) {
    if (border == null) {
      return null;
    }
    return new DifferentialBorderReport(
        toDifferentialBorderSideReport(border.all()),
        toDifferentialBorderSideReport(border.top()),
        toDifferentialBorderSideReport(border.right()),
        toDifferentialBorderSideReport(border.bottom()),
        toDifferentialBorderSideReport(border.left()));
  }

  static DifferentialBorderSideReport toDifferentialBorderSideReport(
      dev.erst.gridgrind.excel.ExcelDifferentialBorderSide side) {
    return side == null
        ? null
        : new DifferentialBorderSideReport(toBorderStyle(side.style()), side.color());
  }

  private static ConditionalFormattingUnsupportedFeature toConditionalFormattingUnsupportedFeature(
      dev.erst.gridgrind.excel.ExcelConditionalFormattingUnsupportedFeature feature) {
    return switch (feature) {
      case STYLE_REFERENCE -> ConditionalFormattingUnsupportedFeature.STYLE_REFERENCE;
      case FONT_ATTRIBUTES -> ConditionalFormattingUnsupportedFeature.FONT_ATTRIBUTES;
      case FILL_PATTERN -> ConditionalFormattingUnsupportedFeature.FILL_PATTERN;
      case FILL_BACKGROUND_COLOR -> ConditionalFormattingUnsupportedFeature.FILL_BACKGROUND_COLOR;
      case BORDER_COMPLEXITY -> ConditionalFormattingUnsupportedFeature.BORDER_COMPLEXITY;
      case ALIGNMENT -> ConditionalFormattingUnsupportedFeature.ALIGNMENT;
      case PROTECTION -> ConditionalFormattingUnsupportedFeature.PROTECTION;
    };
  }

  private static ComparisonOperator toComparisonOperator(
      dev.erst.gridgrind.excel.ExcelComparisonOperator operator) {
    return switch (operator) {
      case BETWEEN -> ComparisonOperator.BETWEEN;
      case NOT_BETWEEN -> ComparisonOperator.NOT_BETWEEN;
      case EQUAL -> ComparisonOperator.EQUAL;
      case NOT_EQUAL -> ComparisonOperator.NOT_EQUAL;
      case GREATER_THAN -> ComparisonOperator.GREATER_THAN;
      case LESS_THAN -> ComparisonOperator.LESS_THAN;
      case GREATER_OR_EQUAL -> ComparisonOperator.GREATER_OR_EQUAL;
      case LESS_OR_EQUAL -> ComparisonOperator.LESS_OR_EQUAL;
    };
  }

  private static AnalysisFindingCode toAnalysisFindingCode(
      dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode code) {
    return AnalysisFindingCode.valueOf(code.name());
  }

  private static AnalysisSeverity toAnalysisSeverity(
      dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity severity) {
    return AnalysisSeverity.valueOf(severity.name());
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

  static ExcelTableSelection toExcelTableSelection(TableSelection selection) {
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

  private static dev.erst.gridgrind.excel.ExcelSheetPane toExcelSheetPane(PaneInput pane) {
    return switch (pane) {
      case PaneInput.None _ -> new dev.erst.gridgrind.excel.ExcelSheetPane.None();
      case PaneInput.Frozen frozen ->
          new dev.erst.gridgrind.excel.ExcelSheetPane.Frozen(
              frozen.splitColumn(), frozen.splitRow(), frozen.leftmostColumn(), frozen.topRow());
      case PaneInput.Split split ->
          new dev.erst.gridgrind.excel.ExcelSheetPane.Split(
              split.xSplitPosition(),
              split.ySplitPosition(),
              split.leftmostColumn(),
              split.topRow(),
              toExcelPaneRegion(split.activePane()));
    };
  }

  private static dev.erst.gridgrind.excel.ExcelPrintLayout toExcelPrintLayout(
      PrintLayoutInput printLayout) {
    return new dev.erst.gridgrind.excel.ExcelPrintLayout(
        toExcelPrintArea(printLayout.printArea()),
        toExcelPrintOrientation(printLayout.orientation()),
        toExcelPrintScaling(printLayout.scaling()),
        toExcelPrintTitleRows(printLayout.repeatingRows()),
        toExcelPrintTitleColumns(printLayout.repeatingColumns()),
        new dev.erst.gridgrind.excel.ExcelHeaderFooterText(
            printLayout.header().left(),
            printLayout.header().center(),
            printLayout.header().right()),
        new dev.erst.gridgrind.excel.ExcelHeaderFooterText(
            printLayout.footer().left(),
            printLayout.footer().center(),
            printLayout.footer().right()));
  }

  private static dev.erst.gridgrind.excel.ExcelPrintLayout.Area toExcelPrintArea(
      PrintAreaInput printArea) {
    return switch (printArea) {
      case PrintAreaInput.None _ -> new dev.erst.gridgrind.excel.ExcelPrintLayout.Area.None();
      case PrintAreaInput.Range range ->
          new dev.erst.gridgrind.excel.ExcelPrintLayout.Area.Range(range.range());
    };
  }

  private static dev.erst.gridgrind.excel.ExcelPrintLayout.Scaling toExcelPrintScaling(
      PrintScalingInput scaling) {
    return switch (scaling) {
      case PrintScalingInput.Automatic _ ->
          new dev.erst.gridgrind.excel.ExcelPrintLayout.Scaling.Automatic();
      case PrintScalingInput.Fit fit ->
          new dev.erst.gridgrind.excel.ExcelPrintLayout.Scaling.Fit(
              fit.widthPages(), fit.heightPages());
    };
  }

  private static dev.erst.gridgrind.excel.ExcelPrintLayout.TitleRows toExcelPrintTitleRows(
      PrintTitleRowsInput repeatingRows) {
    return switch (repeatingRows) {
      case PrintTitleRowsInput.None _ ->
          new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleRows.None();
      case PrintTitleRowsInput.Band band ->
          new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleRows.Band(
              band.firstRowIndex(), band.lastRowIndex());
    };
  }

  private static dev.erst.gridgrind.excel.ExcelPrintLayout.TitleColumns toExcelPrintTitleColumns(
      PrintTitleColumnsInput repeatingColumns) {
    return switch (repeatingColumns) {
      case PrintTitleColumnsInput.None _ ->
          new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleColumns.None();
      case PrintTitleColumnsInput.Band band ->
          new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleColumns.Band(
              band.firstColumnIndex(), band.lastColumnIndex());
    };
  }

  private static GridGrindResponse.CellReport toCellReport(ExcelCellSnapshot snapshot) {
    GridGrindResponse.CellStyleReport style =
        new GridGrindResponse.CellStyleReport(
            snapshot.style().numberFormat(),
            snapshot.style().bold(),
            snapshot.style().italic(),
            snapshot.style().wrapText(),
            toHorizontalAlignment(snapshot.style().horizontalAlignment()),
            toVerticalAlignment(snapshot.style().verticalAlignment()),
            snapshot.style().fontName(),
            toFontHeightReport(snapshot.style().fontHeight()),
            snapshot.style().fontColor(),
            snapshot.style().underline(),
            snapshot.style().strikeout(),
            snapshot.style().fillColor(),
            toBorderStyle(snapshot.style().topBorderStyle()),
            toBorderStyle(snapshot.style().rightBorderStyle()),
            toBorderStyle(snapshot.style().bottomBorderStyle()),
            toBorderStyle(snapshot.style().leftBorderStyle()));
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
          new GridGrindResponse.SheetProtectionReport.Protected(
              toSheetProtectionSettings(protectedState.settings()));
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

  /**
   * Classifies one request-execution exception into the corresponding public problem code.
   *
   * <p>Engine exceptions are owned here so protocol-level problem construction remains engine-blind
   * outside the single request-execution bridge.
   */
  static GridGrindProblemCode problemCodeFor(Exception exception) {
    return switch (exception) {
      case WorkbookNotFoundException _ -> GridGrindProblemCode.WORKBOOK_NOT_FOUND;
      case SheetNotFoundException _ -> GridGrindProblemCode.SHEET_NOT_FOUND;
      case NamedRangeNotFoundException _ -> GridGrindProblemCode.NAMED_RANGE_NOT_FOUND;
      case CellNotFoundException _ -> GridGrindProblemCode.CELL_NOT_FOUND;
      case InvalidCellAddressException _ -> GridGrindProblemCode.INVALID_CELL_ADDRESS;
      case InvalidRangeAddressException _ -> GridGrindProblemCode.INVALID_RANGE_ADDRESS;
      case UnsupportedFormulaException _ -> GridGrindProblemCode.UNSUPPORTED_FORMULA;
      case InvalidFormulaException _ -> GridGrindProblemCode.INVALID_FORMULA;
      case IOException _ -> GridGrindProblemCode.IO_ERROR;
      case DateTimeException _ -> GridGrindProblemCode.INVALID_REQUEST;
      default -> GridGrindProblems.codeFor(exception);
    };
  }

  /** Returns the public problem message for one execution exception. */
  static String messageFor(Exception exception) {
    return GridGrindProblems.messageFor(exception);
  }

  private static GridGrindResponse.Problem problemFor(
      Exception exception, GridGrindResponse.ProblemContext context) {
    Objects.requireNonNull(exception, "exception must not be null");
    Objects.requireNonNull(context, "context must not be null");
    return GridGrindProblems.problem(
        problemCodeFor(exception),
        messageFor(exception),
        enrichContext(context, exception),
        causesFor(exception, context.stage()));
  }

  private static List<GridGrindResponse.ProblemCause> causesFor(Exception exception, String stage) {
    return List.of(
        new GridGrindResponse.ProblemCause(
            problemCodeFor(exception), messageFor(exception), stage));
  }

  /**
   * Enriches one problem context with execution-specific exception fields while preserving payload
   * coordinates supplied by protocol parsing failures.
   */
  static GridGrindResponse.ProblemContext enrichContext(
      GridGrindResponse.ProblemContext context, Exception exception) {
    GridGrindResponse.ProblemContext protocolEnriched =
        GridGrindProblems.enrichContext(context, exception);
    return switch (protocolEnriched) {
      case GridGrindResponse.ProblemContext.ApplyOperation applyOperation ->
          enrichApplyOperationContext(applyOperation, exception);
      case GridGrindResponse.ProblemContext.ExecuteRead executeRead ->
          enrichExecuteReadContext(executeRead, exception);
      case GridGrindResponse.ProblemContext.ReadRequest _ -> protocolEnriched;
      case GridGrindResponse.ProblemContext.ParseArguments _ -> protocolEnriched;
      case GridGrindResponse.ProblemContext.ValidateRequest _ -> protocolEnriched;
      case GridGrindResponse.ProblemContext.OpenWorkbook _ -> protocolEnriched;
      case GridGrindResponse.ProblemContext.PersistWorkbook _ -> protocolEnriched;
      case GridGrindResponse.ProblemContext.ExecuteRequest _ -> protocolEnriched;
      case GridGrindResponse.ProblemContext.WriteResponse _ -> protocolEnriched;
    };
  }

  private static GridGrindResponse.ProblemContext.ApplyOperation enrichApplyOperationContext(
      GridGrindResponse.ProblemContext.ApplyOperation context, Exception exception) {
    if (exception instanceof FormulaException formulaException) {
      return context.withExceptionData(
          formulaException.sheetName(),
          formulaException.address(),
          null,
          formulaException.formula(),
          namedRangeNameFor(exception));
    }
    if (exception instanceof InvalidRangeAddressException invalidRangeAddressException) {
      return context.withExceptionData(
          null, null, invalidRangeAddressException.range(), null, namedRangeNameFor(exception));
    }
    if (exception instanceof NamedRangeNotFoundException) {
      return context.withExceptionData(null, null, null, null, namedRangeNameFor(exception));
    }
    return context;
  }

  private static GridGrindResponse.ProblemContext.ExecuteRead enrichExecuteReadContext(
      GridGrindResponse.ProblemContext.ExecuteRead context, Exception exception) {
    if (exception instanceof FormulaException formulaException) {
      return context.withExceptionData(
          formulaException.sheetName(),
          formulaException.address(),
          formulaException.formula(),
          namedRangeNameFor(exception));
    }
    if (exception instanceof NamedRangeNotFoundException) {
      return context.withExceptionData(null, null, null, namedRangeNameFor(exception));
    }
    return context;
  }

  static String formulaFor(Exception exception) {
    return exception instanceof FormulaException formulaException
        ? formulaException.formula()
        : null;
  }

  static String sheetNameFor(Exception exception) {
    return exception instanceof FormulaException formulaException
        ? formulaException.sheetName()
        : null;
  }

  static String addressFor(Exception exception) {
    return switch (exception) {
      case FormulaException formulaException -> formulaException.address();
      case CellNotFoundException cellNotFoundException -> cellNotFoundException.address();
      case InvalidCellAddressException invalidCellAddressException ->
          invalidCellAddressException.address();
      default -> null;
    };
  }

  static String rangeFor(Exception exception) {
    return exception instanceof InvalidRangeAddressException invalidRangeAddressException
        ? invalidRangeAddressException.range()
        : null;
  }

  static String namedRangeNameFor(Exception exception) {
    return exception instanceof NamedRangeNotFoundException namedRangeNotFoundException
        ? namedRangeNotFoundException.name()
        : null;
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
                problemFor(
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
              problemFor(
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
        problemFor(
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
        problemFor(
            exception,
            new GridGrindResponse.ProblemContext.ExecuteRead(
                reqSourceType(request),
                reqPersistenceType(request),
                readIndex,
                readType(read),
                read.requestId(),
                sheetNameFor(read),
                addressFor(read, exception),
                formulaFor(exception),
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
      case WorkbookReadOperation.GetPrintLayout _ -> "GET_PRINT_LAYOUT";
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
      case WorkbookReadOperation.GetPrintLayout op -> op.sheetName();
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
    String fromException = addressFor(exception);
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
      case WorkbookReadOperation.GetPrintLayout _ -> null;
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
    String fromException = namedRangeNameFor(exception);
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
      case WorkbookReadOperation.GetPrintLayout _ -> null;
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
    String fromException = formulaFor(exception);
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
      case WorkbookOperation.SetSheetPane _ -> null;
      case WorkbookOperation.SetSheetZoom _ -> null;
      case WorkbookOperation.SetPrintLayout _ -> null;
      case WorkbookOperation.ClearPrintLayout _ -> null;
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
          case WorkbookOperation.SetSheetPane op -> op.sheetName();
          case WorkbookOperation.SetSheetZoom op -> op.sheetName();
          case WorkbookOperation.SetPrintLayout op -> op.sheetName();
          case WorkbookOperation.ClearPrintLayout op -> op.sheetName();
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
    return fromOperation != null ? fromOperation : sheetNameFor(exception);
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
          case WorkbookOperation.SetSheetPane _ -> null;
          case WorkbookOperation.SetSheetZoom _ -> null;
          case WorkbookOperation.SetPrintLayout _ -> null;
          case WorkbookOperation.ClearPrintLayout _ -> null;
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
    return fromOperation != null ? fromOperation : addressFor(exception);
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
          case WorkbookOperation.SetSheetPane _ -> null;
          case WorkbookOperation.SetSheetZoom _ -> null;
          case WorkbookOperation.SetPrintLayout _ -> null;
          case WorkbookOperation.ClearPrintLayout _ -> null;
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
    return fromOperation != null ? fromOperation : rangeFor(exception);
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
          case WorkbookOperation.SetSheetPane _ -> null;
          case WorkbookOperation.SetSheetZoom _ -> null;
          case WorkbookOperation.SetPrintLayout _ -> null;
          case WorkbookOperation.ClearPrintLayout _ -> null;
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
    return fromOperation != null ? fromOperation : namedRangeNameFor(exception);
  }

  /** Functional interface for closing an ExcelWorkbook after request execution. */
  @FunctionalInterface
  interface WorkbookCloser {
    /** Closes the workbook, releasing any held resources. */
    void close(ExcelWorkbook workbook) throws IOException;
  }
}
