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
import dev.erst.gridgrind.engine.cells.CellTextLimits;
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
// Streaming authoring spans canonical scalar, temporal, formula and rich-text operands.
@SuppressWarnings("PMD.ExcessiveImports")
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
  public void apply(ResolvedStep step) {
    switch (step.action()) {
      case WorkbookMutationAction.EnsureSheet _ ->
          sheets.computeIfAbsent(
              ((SheetSelector.ByName) step.target()).name().value(),
              name -> {
                nextRowIndexes.put(name, 0);
                return workbook.createSheet(name);
              });
      case ResolvedMutationAction.AppendRow action ->
          appendValues(
              requiredSheet(((SheetSelector.ByName) step.target()).name().value()),
              ((SheetSelector.ByName) step.target()).name().value(),
              action.values());
      default ->
          throw new IllegalArgumentException(
              "execution.mode.type=STREAMING_WRITE supports ENSURE_SHEET and APPEND_ROW only; unsupported mutation action type: "
                  + dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames
                      .mutationActionTypeName(step.action().authoredType()));
    }
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
    int rowIndex = nextRowIndexes.compute(sheetName, (_, nextRowIndex) -> nextRowIndex + 1) - 1;
    SXSSFRow row = sheet.createRow(rowIndex);
    for (int columnIndex = 0; columnIndex < values.size(); columnIndex++) {
      writeCellValue(row.createCell(columnIndex), values.get(columnIndex));
    }
  }

  private void writeCellValue(SXSSFCell cell, ResolvedCellWrite value) {
    switch (value) {
      case CellValue.Blank _ -> cell.setBlank();
      case CellValue.Text textValue -> {
        CellTextLimits.requireSupportedLength(textValue.value(), "textValue.value"); // LIM-010
        cell.setCellValue(textValue.value());
      }
      case ResolvedRichText richTextValue ->
          cell.setCellValue(StyleAdapter.toPoiRichText(workbook.getXSSFWorkbook(), richTextValue));
      case CellValue.NumberValue numberValue -> cell.setCellValue(numberValue.value());
      case CellValue.BooleanValue booleanValue -> cell.setCellValue(booleanValue.value());
      case CellValue.ErrorValue errorValue ->
          cell.setCellErrorValue(CellErrors.toPoiStoredFormulaError(errorValue.value()).getCode());
      case CellValueSpec.Date dateValue -> {
        cell.setCellValue(dateValue.value());
        cell.setCellStyle(styleRegistry.localDateStyle(cell));
      }
      case CellValueSpec.DateTime dateTimeValue -> {
        cell.setCellValue(dateTimeValue.value());
        cell.setCellStyle(styleRegistry.localDateTimeStyle(cell));
      }
      case ResolvedFormula formulaValue ->
          FormulaWrites.setAuthoredFormula(cell, formulaValue.value());
      case ResolvedRawFormula rawFormulaValue ->
          FormulaWrites.setOpaqueFormula(cell, rawFormulaValue.value());
    }
  }
}
