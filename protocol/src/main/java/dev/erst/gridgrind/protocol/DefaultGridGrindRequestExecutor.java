package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelCellSelection;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeSelection;
import dev.erst.gridgrind.excel.ExcelNamedRangeSelector;
import dev.erst.gridgrind.excel.ExcelNamedRangeSnapshot;
import dev.erst.gridgrind.excel.ExcelSheetSelection;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
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
                  reqSourceMode(request), reqPersistenceMode(request), reqSourcePath(request))));
    }

    return guardUnexpectedRuntime(
        protocolVersion,
        request,
        workbook,
        () -> executeWorkbookWorkflow(protocolVersion, request, workbook));
  }

  private GridGrindResponse executeWorkbookWorkflow(
      GridGrindProtocolVersion protocolVersion, GridGrindRequest request, ExcelWorkbook workbook) {
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
            readExecutor.apply(workbook, command).getFirst();
        reads.add(toReadResult(result));
      } catch (Exception exception) {
        return closeWorkbook(
            workbook, readFailure(protocolVersion, request, readIndex, read, exception), request);
      }
    }

    GridGrindResponse.PersistenceOutcome persistence;
    try {
      persistence = persistWorkbook(workbook, request.persistence(), reqSourcePath(request));
    } catch (Exception exception) {
      return closeWorkbook(
          workbook,
          new GridGrindResponse.Failure(
              protocolVersion,
              GridGrindProblems.fromException(
                  exception,
                  new GridGrindResponse.ProblemContext.PersistWorkbook(
                      reqSourceMode(request),
                      reqPersistenceMode(request),
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
                                reqSourceMode(request), reqPersistenceMode(request)),
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
      case WorkbookReadOperation.AnalyzeFormulaSurface op ->
          new WorkbookReadCommand.AnalyzeFormulaSurface(
              op.requestId(), toExcelSheetSelection(op.selection()));
      case WorkbookReadOperation.AnalyzeSheetSchema op ->
          new WorkbookReadCommand.AnalyzeSheetSchema(
              op.requestId(), op.sheetName(), op.topLeftAddress(), op.rowCount(), op.columnCount());
      case WorkbookReadOperation.AnalyzeNamedRangeSurface op ->
          new WorkbookReadCommand.AnalyzeNamedRangeSurface(
              op.requestId(), toExcelNamedRangeSelection(op.selection()));
    };
  }

  private GridGrindResponse.PersistenceOutcome persistWorkbook(
      ExcelWorkbook workbook, GridGrindRequest.WorkbookPersistence persistence, String sourcePath)
      throws IOException {
    return switch (persistence) {
      case GridGrindRequest.WorkbookPersistence.None _ ->
          new GridGrindResponse.PersistenceOutcome.NotSaved();
      case GridGrindRequest.WorkbookPersistence.OverwriteSource _ -> {
        Path path = Path.of(sourcePath);
        workbook.save(path);
        yield new GridGrindResponse.PersistenceOutcome.Saved(path.toAbsolutePath().toString());
      }
      case GridGrindRequest.WorkbookPersistence.SaveAs saveAs -> {
        Path path = Path.of(saveAs.path());
        workbook.save(path);
        yield new GridGrindResponse.PersistenceOutcome.Saved(path.toAbsolutePath().toString());
      }
    };
  }

  /** Converts one workbook-core read result into the protocol response shape. */
  static WorkbookReadResult toReadResult(dev.erst.gridgrind.excel.WorkbookReadResult result) {
    return switch (result) {
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult workbookSummary ->
          new WorkbookReadResult.WorkbookSummaryResult(
              workbookSummary.requestId(),
              new GridGrindResponse.WorkbookSummary(
                  workbookSummary.workbook().sheetCount(),
                  workbookSummary.workbook().sheetNames(),
                  workbookSummary.workbook().namedRangeCount(),
                  workbookSummary.workbook().forceFormulaRecalculationOnOpen()));
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
        hyperlink.address(), toHyperlinkReport(hyperlink.hyperlink()));
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
    GridGrindResponse.HyperlinkReport hyperlink =
        toHyperlinkReport(snapshot.metadata().hyperlink().orElse(null));
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

  /** Converts workbook-core hyperlink metadata into the protocol response shape. */
  static GridGrindResponse.HyperlinkReport toHyperlinkReport(ExcelHyperlink hyperlink) {
    if (hyperlink == null) {
      return null;
    }
    return new GridGrindResponse.HyperlinkReport(hyperlink.type(), hyperlink.target());
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

  /** Delegates to GridGrindProblems to classify the exception into a problem code. */
  static GridGrindProblemCode problemCodeFor(Exception exception) {
    return GridGrindProblems.codeFor(exception);
  }

  /** Delegates to GridGrindProblems to extract a user-readable problem message. */
  static String messageFor(Exception exception) {
    return GridGrindProblems.messageFor(exception);
  }

  private String absolutePath(String path) {
    return Path.of(path).toAbsolutePath().toString();
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
                        reqSourceMode(request), reqPersistenceMode(request))));
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
                      reqSourceMode(request), reqPersistenceMode(request)))),
          request);
    }
  }

  private String reqSourceMode(GridGrindRequest request) {
    return switch (request.source()) {
      case GridGrindRequest.WorkbookSource.New _ -> "NEW";
      case GridGrindRequest.WorkbookSource.ExistingFile _ -> "EXISTING";
    };
  }

  private String reqPersistenceMode(GridGrindRequest request) {
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

  String persistencePath(
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
                reqSourceMode(request),
                reqPersistenceMode(request),
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
                reqSourceMode(request),
                reqPersistenceMode(request),
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
      case WorkbookReadOperation.AnalyzeFormulaSurface _ -> "ANALYZE_FORMULA_SURFACE";
      case WorkbookReadOperation.AnalyzeSheetSchema _ -> "ANALYZE_SHEET_SCHEMA";
      case WorkbookReadOperation.AnalyzeNamedRangeSurface _ -> "ANALYZE_NAMED_RANGE_SURFACE";
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
      case WorkbookReadOperation.AnalyzeSheetSchema op -> op.sheetName();
      case WorkbookReadOperation.GetWorkbookSummary _ -> null;
      case WorkbookReadOperation.GetNamedRanges _ -> null;
      case WorkbookReadOperation.AnalyzeFormulaSurface _ -> null;
      case WorkbookReadOperation.AnalyzeNamedRangeSurface _ -> null;
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
      case WorkbookReadOperation.AnalyzeSheetSchema op -> op.topLeftAddress();
      case WorkbookReadOperation.GetWorkbookSummary _ -> null;
      case WorkbookReadOperation.GetNamedRanges _ -> null;
      case WorkbookReadOperation.GetSheetSummary _ -> null;
      case WorkbookReadOperation.GetCells _ -> null;
      case WorkbookReadOperation.GetMergedRegions _ -> null;
      case WorkbookReadOperation.GetHyperlinks _ -> null;
      case WorkbookReadOperation.GetComments _ -> null;
      case WorkbookReadOperation.GetSheetLayout _ -> null;
      case WorkbookReadOperation.AnalyzeFormulaSurface _ -> null;
      case WorkbookReadOperation.AnalyzeNamedRangeSurface _ -> null;
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
      case WorkbookReadOperation.AnalyzeFormulaSurface _ -> null;
      case WorkbookReadOperation.AnalyzeSheetSchema _ -> null;
      case WorkbookReadOperation.AnalyzeNamedRangeSurface _ -> null;
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
          case WorkbookOperation.MergeCells _ -> null;
          case WorkbookOperation.UnmergeCells _ -> null;
          case WorkbookOperation.SetColumnWidth _ -> null;
          case WorkbookOperation.SetRowHeight _ -> null;
          case WorkbookOperation.FreezePanes _ -> null;
          case WorkbookOperation.SetRange _ -> null;
          case WorkbookOperation.ClearRange _ -> null;
          case WorkbookOperation.ApplyStyle _ -> null;
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
          case WorkbookOperation.MergeCells op -> op.range();
          case WorkbookOperation.UnmergeCells op -> op.range();
          case WorkbookOperation.SetNamedRange op -> op.target().range();
          case WorkbookOperation.EnsureSheet _ -> null;
          case WorkbookOperation.RenameSheet _ -> null;
          case WorkbookOperation.DeleteSheet _ -> null;
          case WorkbookOperation.MoveSheet _ -> null;
          case WorkbookOperation.SetColumnWidth _ -> null;
          case WorkbookOperation.SetRowHeight _ -> null;
          case WorkbookOperation.FreezePanes _ -> null;
          case WorkbookOperation.SetCell _ -> null;
          case WorkbookOperation.SetHyperlink _ -> null;
          case WorkbookOperation.ClearHyperlink _ -> null;
          case WorkbookOperation.SetComment _ -> null;
          case WorkbookOperation.ClearComment _ -> null;
          case WorkbookOperation.AppendRow _ -> null;
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
