package dev.erst.gridgrind.engine.backend.streaming;

import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedRawFormula;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.engine.backend.xssf.SheetNotFoundException;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.engine.cells.CellErrors;
import dev.erst.gridgrind.engine.packaging.WorkbookFiles;
import dev.erst.gridgrind.engine.packaging.WorkbookNormalization;
import dev.erst.gridgrind.engine.poi.FormulaWrites;
import dev.erst.gridgrind.engine.styles.StyleAdapter;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Stateful low-memory append-oriented workbook authoring backed by POI's SXSSF writer. */
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_TEXT_CHARACTERS_PER_CELL
})
public final class StreamingWorkbookWriter implements AutoCloseable {
  private final SXSSFWorkbook workbook;
  private final StyleAdapter styleRegistry;
  private final Map<String, SXSSFSheet> sheets;
  private final Map<String, Integer> nextRowIndexes;

  /** Creates one empty streaming workbook session. */
  public StreamingWorkbookWriter() {
    this.workbook = new SXSSFWorkbook(new XSSFWorkbook(), 100, true, true);
    this.styleRegistry = new StyleAdapter(workbook.getXSSFWorkbook());
    this.sheets = new HashMap<>();
    this.nextRowIndexes = new HashMap<>();
  }

  /** Applies a materialized core mutation to the streaming session. */
  public void apply(ResolvedStep step) throws IOException {
    switch (step.action()) {
      case WorkbookMutationAction.EnsureSheet _ ->
          ensureSheet(((SheetSelector.ByName) step.target()).name().value());
      case ResolvedMutationAction.AppendRow action ->
          appendValues(
              requiredSheet(((SheetSelector.ByName) step.target()).name().value()),
              ((SheetSelector.ByName) step.target()).name().value(),
              action.values());
      default ->
          throw new IllegalArgumentException(
              "execution.mode.type=STREAMING_WRITE supports ENSURE_SHEET and APPEND_ROW only;"
                  + " unsupported mutation action type: "
                  + dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames
                      .mutationActionTypeName(step.action().authoredType()));
    }
  }

  private void ensureSheet(String name) throws IOException {
    if (sheets.containsKey(name)) return;
    SXSSFSheet created;
    try {
      created = workbook.createSheet(name);
    } catch (IllegalStateException exception) {
      // POI 6's createAndRegisterSXSSFSheet wraps scratch acquisition IOException in this type.
      // Restore that declared IO boundary without reclassifying unrelated native failures.
      if (exception.getCause() instanceof IOException failure) {
        throw new IOException(failure.getMessage(), exception);
      }
      throw exception;
    }
    sheets.put(name, created);
    nextRowIndexes.put(name, 0);
  }

  /** Marks the streamed workbook to recalculate formulas on the next Excel-compatible open. */
  public void markRecalculateOnOpen() {
    workbook.setForceFormulaRecalculation(true);
  }

  /** Materializes the current streaming workbook session to one `.xlsx` path. */
  public void save(Path workbookPath, WorkbookArtifactWriteDisposition writeDisposition)
      throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    Objects.requireNonNull(writeDisposition, "writeDisposition must not be null");
    Path absolutePath = workbookPath.toAbsolutePath().normalize();
    Files.createDirectories(absolutePath.getParent());
    WorkbookNormalization.normalizeDocumentProperties(workbook.getXSSFWorkbook());
    try (OutputStream outputStream =
        WorkbookFiles.newWorkbookOutputStream(absolutePath, writeDisposition)) {
      workbook.write(outputStream);
    }
    WorkbookNormalization.normalizeWorkbookPackage(absolutePath);
  }

  @Override
  public void close() throws IOException {
    workbook.close();
  }

  private SXSSFSheet requiredSheet(String sheetName) {
    SXSSFSheet sheet = sheets.get(sheetName);
    if (sheet == null) {
      throw new SheetNotFoundException(sheetName);
    }
    return sheet;
  }

  private void appendValues(
      SXSSFSheet sheet, String sheetName, java.util.List<ResolvedCellWrite> values) {
    values.forEach(value ->
        dev.erst.gridgrind.core.model.value.TemporalSemantics.requireWritable(
            value, dev.erst.gridgrind.core.model.value.ExcelDateSystem.WINDOWS_1900));
    int rowIndex = nextRowIndexes.compute(sheetName, (_, nextRowIndex) -> nextRowIndex + 1) - 1;
    SXSSFRow row = sheet.createRow(rowIndex);
    for (int columnIndex = 0; columnIndex < values.size(); columnIndex++) {
      writeCellValue(row.createCell(columnIndex), values.get(columnIndex));
    }
  }

  private void writeCellValue(SXSSFCell cell, ResolvedCellWrite value) {
    switch (value) {
      case CellValue.Blank _ -> cell.setBlank();
      case dev.erst.gridgrind.core.model.value.ResolvedCellText textValue -> {
        cell.setCellValue(textValue.value());
      }
      case ResolvedRichText richTextValue ->
          cell.setCellValue(StyleAdapter.toPoiRichText(workbook.getXSSFWorkbook(), richTextValue));
      case CellValueSpec.NumberValue numberValue -> cell.setCellValue(numberValue.value());
      case CellValue.BooleanValue booleanValue -> cell.setCellValue(booleanValue.value());
      case CellValueSpec.ErrorValue errorValue ->
          cell.setCellErrorValue(CellErrors.toPoiStoredFormulaError(errorValue.value()).getCode());
      case CellValueSpec.Date dateValue -> {
        cell.setCellValue(
            dev.erst.gridgrind.core.model.value.TemporalSemantics.serial(
                dateValue.value(),
                dev.erst.gridgrind.core.model.value.ExcelDateSystem.WINDOWS_1900));
        cell.setCellStyle(styleRegistry.localDateStyle(cell));
      }
      case CellValueSpec.DateTime dateTimeValue -> {
        cell.setCellValue(
            dev.erst.gridgrind.core.model.value.TemporalSemantics.serial(
                dateTimeValue.value(),
                dev.erst.gridgrind.core.model.value.ExcelDateSystem.WINDOWS_1900));
        cell.setCellStyle(styleRegistry.localDateTimeStyle(cell));
      }
      case ResolvedFormula formulaValue ->
          FormulaWrites.setAuthoredFormula(cell, formulaValue.value());
      case ResolvedRawFormula rawFormulaValue ->
          FormulaWrites.setOpaqueFormula(cell, rawFormulaValue.value());
    }
  }
}
