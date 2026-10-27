package dev.erst.gridgrind.engine.cells;

import dev.erst.gridgrind.core.model.formula.ArrayFormulaReport;
import dev.erst.gridgrind.core.model.formula.ResolvedArrayFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedRawFormula;
import dev.erst.gridgrind.core.model.layout.WindowRowReport;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;
import dev.erst.gridgrind.core.plan.CellReadFacet;
import dev.erst.gridgrind.core.plan.CellReadProjection;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import dev.erst.gridgrind.engine.backend.xssf.CellNotFoundException;
import dev.erst.gridgrind.engine.backend.xssf.InvalidCellAddressException;
import dev.erst.gridgrind.engine.comments.CommentAdapter;
import dev.erst.gridgrind.engine.formulas.FormulaExceptions;
import dev.erst.gridgrind.engine.formulas.FormulaLimits;
import dev.erst.gridgrind.engine.formulas.FormulaRuntime;
import dev.erst.gridgrind.engine.links.LinkAdapter;
import dev.erst.gridgrind.engine.poi.CellBounds;
import dev.erst.gridgrind.engine.poi.FormulaWrites;
import dev.erst.gridgrind.engine.poi.SheetStructure;
import dev.erst.gridgrind.engine.styles.StyleAdapter;
import dev.erst.gridgrind.engine.tables.TableAdapter;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jspecify.annotations.Nullable;

/** Owns typed cell reads, projection, mutation, previews and array formula groups. */
@SuppressWarnings({"PMD.ExcessiveImports", "PMD.CommentRequired"})
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_TEXT_CHARACTERS_PER_CELL
})
public final class CellsAdapter {
  private final Sheet sheet;
  private final StyleAdapter styleRegistry;
  private final FormulaRuntime formulaRuntime;
  private final DataFormatter dataFormatter;

  public CellsAdapter(
      Sheet sheet,
      StyleAdapter styleRegistry,
      FormulaRuntime formulaRuntime,
      DataFormatter dataFormatter) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.styleRegistry = Objects.requireNonNull(styleRegistry, "styleRegistry must not be null");
    this.formulaRuntime = Objects.requireNonNull(formulaRuntime, "formulaRuntime must not be null");
    this.dataFormatter = Objects.requireNonNull(dataFormatter, "dataFormatter must not be null");
  }

  public String text(String address) {
    TextValidation.requireNonBlank(address, "address");
    return requiredCell(address).getStringCellValue();
  }

  public double number(String address) {
    TextValidation.requireNonBlank(address, "address");
    Cell cell = requiredCell(address);
    if (cell.getCellType() == CellType.FORMULA) {
      org.apache.poi.ss.usermodel.CellValue evaluatedCell;
      try {
        evaluatedCell = formulaRuntime.evaluate(cell);
      } catch (RuntimeException exception) {
        throw FormulaExceptions.wrap(
            formulaRuntime, sheet.getSheetName(), address, cell.getCellFormula(), exception);
      }
      if (evaluatedCell == null || evaluatedCell.getCellType() != CellType.NUMERIC) {
        throw new IllegalStateException("Cell does not evaluate to a numeric value: " + address);
      }
      return evaluatedCell.getNumberValue();
    }
    return cell.getNumericCellValue();
  }

  public boolean bool(String address) {
    TextValidation.requireNonBlank(address, "address");
    Cell cell = requiredCell(address);
    if (cell.getCellType() == CellType.FORMULA) {
      org.apache.poi.ss.usermodel.CellValue evaluatedCell;
      try {
        evaluatedCell = formulaRuntime.evaluate(cell);
      } catch (RuntimeException exception) {
        throw FormulaExceptions.wrap(
            formulaRuntime, sheet.getSheetName(), address, cell.getCellFormula(), exception);
      }
      if (evaluatedCell == null || evaluatedCell.getCellType() != CellType.BOOLEAN) {
        throw new IllegalStateException("Cell does not evaluate to a boolean value: " + address);
      }
      return evaluatedCell.getBooleanValue();
    }
    return cell.getBooleanCellValue();
  }

  public String formula(String address) {
    TextValidation.requireNonBlank(address, "address");
    Cell cell = requiredCell(address);
    if (cell.getCellType() != CellType.FORMULA) {
      throw new IllegalStateException("Cell does not contain a formula: " + address);
    }
    return cell.getCellFormula();
  }

  public CellReport snapshotCell(String address) {
    return snapshotCell(address, new CellReadProjection(java.util.List.of(CellReadFacet.values())));
  }

  public CellReport snapshotCell(String address, CellReadProjection projection) {
    TextValidation.requireNonBlank(address, "address");
    Objects.requireNonNull(projection, "projection must not be null");
    CellReference cellReference = parseCellReference(address);
    requireValidCellReference(address, cellReference);
    return snapshotCellOrBlank(address, cellReference.getRow(), cellReference.getCol(), projection);
  }

  public List<CellReport> snapshotCells(List<String> addresses) {
    return snapshotCells(
        addresses, new CellReadProjection(java.util.List.of(CellReadFacet.values())));
  }

  public List<CellReport> snapshotCells(List<String> addresses, CellReadProjection projection) {
    Objects.requireNonNull(addresses, "addresses must not be null");
    Objects.requireNonNull(projection, "projection must not be null");
    List<CellReport> cells = new ArrayList<>(addresses.size());
    for (String address : List.copyOf(addresses)) {
      TextValidation.requireNonBlank(address, "address");
      cells.add(snapshotCell(address, projection));
    }
    return List.copyOf(cells);
  }

  public List<WindowRowReport> preview(int maxRows, int maxColumns) {
    if (maxRows <= 0) {
      throw new IllegalArgumentException("maxRows must be greater than 0");
    }
    if (maxColumns <= 0) {
      throw new IllegalArgumentException("maxColumns must be greater than 0");
    }

    List<WindowRowReport> previewRows = new ArrayList<>();
    int lastRowIndex = sheet.getLastRowNum();
    int endingRowIndex = Math.min(lastRowIndex, maxRows - 1);
    for (int rowIndex = 0; rowIndex <= endingRowIndex; rowIndex++) {
      previewRows.add(new WindowRowReport(rowIndex, previewRow(rowIndex, maxColumns)));
    }
    return List.copyOf(previewRows);
  }

  public dev.erst.gridgrind.core.model.layout.WindowReport.Dense window(
      String topLeftAddress, int rowCount, int columnCount) {
    return window(
        topLeftAddress,
        rowCount,
        columnCount,
        new CellReadProjection(java.util.List.of(CellReadFacet.values())));
  }

  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  public dev.erst.gridgrind.core.model.layout.WindowReport.Dense window(
      String topLeftAddress, int rowCount, int columnCount, CellReadProjection projection) {
    TextValidation.requireNonBlank(topLeftAddress, "topLeftAddress");
    Objects.requireNonNull(projection, "projection must not be null");
    if (rowCount <= 0) {
      throw new IllegalArgumentException("rowCount must be greater than 0");
    }
    if (columnCount <= 0) {
      throw new IllegalArgumentException("columnCount must be greater than 0");
    }

    CellReference topLeft = parseCellReference(topLeftAddress);
    requireValidCellReference(topLeftAddress, topLeft);
    int lastRow = topLeft.getRow() + rowCount - 1;
    int lastCol = topLeft.getCol() + columnCount - 1;
    if (lastRow > SpreadsheetVersion.EXCEL2007.getLastRowIndex()
        || lastCol > SpreadsheetVersion.EXCEL2007.getLastColumnIndex()) {
      throw new IllegalArgumentException(
          "window extends beyond the Excel sheet boundary: topLeft="
              + topLeftAddress
              + " rowCount="
              + rowCount
              + " columnCount="
              + columnCount);
    }
    List<dev.erst.gridgrind.core.model.layout.WindowRowReport> rows = new ArrayList<>(rowCount);
    for (int rowOffset = 0; rowOffset < rowCount; rowOffset++) {
      int rowIndex = topLeft.getRow() + rowOffset;
      List<CellReport> cells = new ArrayList<>(columnCount);
      for (int columnOffset = 0; columnOffset < columnCount; columnOffset++) {
        int columnIndex = topLeft.getCol() + columnOffset;
        String address = new CellReference(rowIndex, columnIndex).formatAsString();
        cells.add(snapshotCellOrBlank(address, rowIndex, columnIndex, projection));
      }
      rows.add(
          new dev.erst.gridgrind.core.model.layout.WindowRowReport(rowIndex, List.copyOf(cells)));
    }
    return new dev.erst.gridgrind.core.model.layout.WindowReport.Dense(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheet.getSheetName()),
        new dev.erst.gridgrind.core.vocabulary.CellAddress(topLeftAddress),
        new dev.erst.gridgrind.core.model.layout.WindowDimensionsReport(rowCount, columnCount),
        List.copyOf(rows));
  }

  public List<ArrayFormulaReport> arrayFormulas() {
    Set<String> seenRanges = new LinkedHashSet<>();
    List<ArrayFormulaReport> formulas = new ArrayList<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (cell.getCellType() != CellType.FORMULA || !cell.isPartOfArrayFormulaGroup()) {
          continue;
        }
        var range = cell.getArrayFormulaRange();
        String rangeText = range.formatAsString();
        if (!seenRanges.add(rangeText)) {
          continue;
        }
        formulas.add(
            new ArrayFormulaReport(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheet.getSheetName()),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel(rangeText),
                new dev.erst.gridgrind.core.vocabulary.CellAddress(
                    new CellReference(range.getFirstRow(), range.getFirstColumn())
                        .formatAsString()),
                xssfSheet()
                    .getRow(range.getFirstRow())
                    .getCell(range.getFirstColumn())
                    .getCellFormula(),
                range.getNumberOfCells() == 1));
      }
    }
    return List.copyOf(formulas);
  }

  private Cell requiredCell(String address) {
    CellReference cellReference = parseCellReference(address);
    Row row = sheet.getRow(cellReference.getRow());
    if (row == null) {
      throw new CellNotFoundException(address);
    }

    Cell cell = row.getCell(cellReference.getCol());
    if (cell == null) {
      throw new CellNotFoundException(address);
    }

    return cell;
  }

  private List<CellReport> previewRow(int rowIndex, int maxColumns) {
    Row row = sheet.getRow(rowIndex);
    List<CellReport> cells = new ArrayList<>();
    if (row != null) {
      for (int columnIndex = 0; columnIndex < maxColumns; columnIndex++) {
        Cell cell = row.getCell(columnIndex);
        if (SheetStructure.shouldPreview(cell)) {
          cells.add(
              snapshot(
                  new CellReference(rowIndex, columnIndex).formatAsString(),
                  cell,
                  CellReadProjection.of(CellReadFacet.VALUE, CellReadFacet.FORMAT)));
        }
      }
    }
    return List.copyOf(cells);
  }

  private CellReport snapshotCellOrBlank(
      String address, int rowIndex, int columnIndex, CellReadProjection projection) {
    Row row = sheet.getRow(rowIndex);
    Cell cell = row == null ? null : row.getCell(columnIndex);
    if (cell == null) return blankReport(address, projection);
    return snapshot(address, cell, projection);
  }

  private CellReport snapshot(String address, Cell cell, CellReadProjection projection) {
    var coordinate = new dev.erst.gridgrind.core.vocabulary.CellAddress(address);
    CellType declared = cell.getCellType();
    String formula = declared == CellType.FORMULA ? cell.getCellFormula() : null;
    String display;
    try {
      display =
          !projection.includes(CellReadFacet.FORMAT)
              ? ""
              : declared == CellType.FORMULA
                  ? requiresFormulaEvaluation(projection)
                      ? formulaRuntime.displayValue(dataFormatter, cell)
                      : cell.getCellFormula()
                  : dataFormatter.formatCellValue(cell);
    } catch (RuntimeException failure) {
      throw FormulaExceptions.wrap(formulaRuntime, sheet.getSheetName(), address, formula, failure);
    }
    CellStyleReport effectiveStyle = styleRegistry.snapshot(cell);
    var format =
        projection.includes(CellReadFacet.FORMAT)
            ? java.util.Optional.of(display)
            : java.util.Optional.<String>empty();
    var style =
        projection.includes(CellReadFacet.STYLE)
            ? java.util.Optional.of(effectiveStyle)
            : java.util.Optional.<CellStyleReport>empty();
    var hyperlink =
        projection.includes(CellReadFacet.HYPERLINK)
            ? LinkAdapter.hyperlink(cell)
            : java.util.Optional.<dev.erst.gridgrind.core.model.value.HyperlinkTarget>empty();
    var comment =
        projection.includes(CellReadFacet.COMMENT)
            ? CommentAdapter.commentSnapshot(cell)
            : java.util.Optional.<dev.erst.gridgrind.core.model.drawing.CommentReport>empty();
    if (declared == CellType.FORMULA) {
      org.apache.poi.ss.usermodel.CellValue evaluated = null;
      if (requiresFormulaEvaluation(projection)) {
        try {
          evaluated = formulaRuntime.evaluate(cell);
        } catch (RuntimeException failure) {
          throw FormulaExceptions.wrap(
              formulaRuntime, sheet.getSheetName(), address, formula, failure);
        }
      }
      var value =
          projection.includes(CellReadFacet.VALUE)
              ? java.util.Optional.of(evaluatedValue(evaluated))
              : java.util.Optional.<dev.erst.gridgrind.core.model.value.CellValue>empty();
      var temporal =
          projection.includes(CellReadFacet.TEMPORAL)
                  && value.orElse(null)
                      instanceof dev.erst.gridgrind.core.model.value.CellValue.NumberValue number
              ? java.util.Optional.of(temporal(number.value(), effectiveStyle.numberFormat()))
              : java.util.Optional.<dev.erst.gridgrind.core.model.value.CellTemporalReport>empty();
      return new CellReport.FormulaReport(
          coordinate,
          format,
          style,
          hyperlink,
          comment,
          projection.includes(CellReadFacet.FORMULA)
              ? java.util.Optional.of(cell.getCellFormula())
              : java.util.Optional.empty(),
          value,
          java.util.Optional.empty(),
          temporal);
    }
    return switch (declared) {
      case STRING ->
          new CellReport.TextReport(
              coordinate,
              format,
              style,
              hyperlink,
              comment,
              projected(
                  projection,
                  new dev.erst.gridgrind.core.model.value.CellValue.Text(
                      cell.getStringCellValue())),
              projection.includes(CellReadFacet.RICH_TEXT_RUNS)
                  ? StyleAdapter.snapshot(
                      xssfSheet().getWorkbook(),
                      (XSSFRichTextString) cell.getRichStringCellValue(),
                      effectiveStyle.font())
                  : java.util.Optional.empty());
      case NUMERIC ->
          new CellReport.NumberReport(
              coordinate,
              format,
              style,
              hyperlink,
              comment,
              projected(
                  projection,
                  new dev.erst.gridgrind.core.model.value.CellValue.NumberValue(
                      cell.getNumericCellValue())),
              projection.includes(CellReadFacet.TEMPORAL)
                  ? java.util.Optional.of(
                      temporal(cell.getNumericCellValue(), effectiveStyle.numberFormat()))
                  : java.util.Optional.empty());
      case BOOLEAN ->
          new CellReport.BooleanReport(
              coordinate,
              format,
              style,
              hyperlink,
              comment,
              projected(
                  projection,
                  new dev.erst.gridgrind.core.model.value.CellValue.BooleanValue(
                      cell.getBooleanCellValue())));
      case ERROR ->
          new CellReport.ErrorReport(
              coordinate,
              format,
              style,
              hyperlink,
              comment,
              projected(
                  projection,
                  new dev.erst.gridgrind.core.model.value.CellValue.ErrorValue(
                      CellErrors.toReportedWireValue(
                          FormulaError.forInt(cell.getErrorCellValue())))));
      case BLANK, _NONE, FORMULA ->
          new CellReport.BlankReport(
              coordinate,
              format,
              style,
              hyperlink,
              comment,
              projected(projection, new dev.erst.gridgrind.core.model.value.CellValue.Blank()));
    };
  }

  private static boolean requiresFormulaEvaluation(CellReadProjection projection) {
    return projection.includes(CellReadFacet.VALUE) || projection.includes(CellReadFacet.FORMAT);
  }

  private static <T> java.util.Optional<T> projected(CellReadProjection projection, T value) {
    return projection.includes(CellReadFacet.VALUE)
        ? java.util.Optional.of(value)
        : java.util.Optional.empty();
  }

  private static dev.erst.gridgrind.core.model.value.CellValue evaluatedValue(
      org.apache.poi.ss.usermodel.@Nullable CellValue value) {
    if (value == null) return new dev.erst.gridgrind.core.model.value.CellValue.Blank();
    return switch (value.getCellType()) {
      case STRING -> new dev.erst.gridgrind.core.model.value.CellValue.Text(value.getStringValue());
      case NUMERIC ->
          new dev.erst.gridgrind.core.model.value.CellValue.NumberValue(value.getNumberValue());
      case BOOLEAN ->
          new dev.erst.gridgrind.core.model.value.CellValue.BooleanValue(value.getBooleanValue());
      case ERROR ->
          new dev.erst.gridgrind.core.model.value.CellValue.ErrorValue(
              CellErrors.toReportedWireValue(FormulaError.forInt(value.getErrorValue())));
      default -> new dev.erst.gridgrind.core.model.value.CellValue.Blank();
    };
  }

  private dev.erst.gridgrind.core.model.value.CellTemporalReport temporal(
      double number, String numberFormat) {
    var kind = TemporalFormats.observedKind(numberFormat);
    if (kind.isEmpty()) return dev.erst.gridgrind.core.model.value.CellTemporalReport.notDate();
    var value =
        org.apache.poi.ss.usermodel.DateUtil.getLocalDateTime(
            number, xssfSheet().getWorkbook().isDate1904());
    return switch (kind.orElseThrow()) {
      case DATE ->
          dev.erst.gridgrind.core.model.value.CellTemporalReport.temporal(
              dev.erst.gridgrind.core.model.value.CellTemporalKind.DATE,
              value.toLocalDate().toString());
      case TIME ->
          dev.erst.gridgrind.core.model.value.CellTemporalReport.temporal(
              dev.erst.gridgrind.core.model.value.CellTemporalKind.TIME,
              value.toLocalTime().toString());
      case DATE_TIME ->
          dev.erst.gridgrind.core.model.value.CellTemporalReport.temporal(
              dev.erst.gridgrind.core.model.value.CellTemporalKind.DATE_TIME, value.toString());
    };
  }

  private CellReport blankReport(String address, CellReadProjection projection) {
    return new CellReport.BlankReport(
        new dev.erst.gridgrind.core.vocabulary.CellAddress(address),
        projection.includes(CellReadFacet.FORMAT)
            ? java.util.Optional.of("")
            : java.util.Optional.empty(),
        projection.includes(CellReadFacet.STYLE)
            ? java.util.Optional.of(styleRegistry.defaultSnapshot())
            : java.util.Optional.empty(),
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        projected(projection, new dev.erst.gridgrind.core.model.value.CellValue.Blank()));
  }

  private XSSFSheet xssfSheet() {
    return (XSSFSheet) sheet;
  }

  private static CellReference parseCellReference(String address) {
    try {
      return new CellReference(address);
    } catch (IllegalArgumentException exception) {
      throw new InvalidCellAddressException(address, exception);
    }
  }

  public static void requireValidCellReference(String address, CellReference cellReference) {
    int row = cellReference.getRow();
    int col = cellReference.getCol();
    if (row < 0
        || col < 0
        || row > SpreadsheetVersion.EXCEL2007.getLastRowIndex()
        || col > SpreadsheetVersion.EXCEL2007.getLastColumnIndex()) {
      throw new InvalidCellAddressException(
          address, new IllegalArgumentException("not a valid A1-style cell address: " + address));
    }
  }

  public CellsAdapter setCell(String address, ResolvedCellWrite value) {
    TextValidation.requireNonBlank(address, "address");
    Objects.requireNonNull(value, "value must not be null");

    CellReference cellReference = parseCellReference(address);
    writeCellValue(cellReference.getRow(), cellReference.getCol(), value);
    syncTableHeaders(
        new CellBounds(
            cellReference.getRow(),
            cellReference.getRow(),
            cellReference.getCol(),
            cellReference.getCol()));
    return this;
  }

  public CellsAdapter setRange(String range, List<List<ResolvedCellWrite>> rows) {
    TextValidation.requireNonBlank(range, "range");
    Objects.requireNonNull(rows, "rows must not be null");
    if (rows.isEmpty()) {
      throw new IllegalArgumentException("rows must not be empty");
    }

    List<List<ResolvedCellWrite>> copiedRows = copyRows(rows);
    CellBounds excelRange = CellBounds.parse(range);
    if (copiedRows.size() != excelRange.rowCount()
        || copiedRows.getFirst().size() != excelRange.columnCount()) {
      throw new IllegalArgumentException(
          "range dimensions do not match provided values: "
              + range
              + " expects "
              + excelRange.rowCount()
              + "x"
              + excelRange.columnCount()
              + " but received "
              + copiedRows.size()
              + "x"
              + copiedRows.getFirst().size());
    }

    for (int rowOffset = 0; rowOffset < copiedRows.size(); rowOffset++) {
      List<ResolvedCellWrite> rowValues = copiedRows.get(rowOffset);
      for (int columnOffset = 0; columnOffset < rowValues.size(); columnOffset++) {
        writeCellValue(
            excelRange.firstRow() + rowOffset,
            excelRange.firstColumn() + columnOffset,
            rowValues.get(columnOffset));
      }
    }

    syncTableHeaders(excelRange);
    return this;
  }

  public CellsAdapter clearRange(String range) {
    TextValidation.requireNonBlank(range, "range");
    CellBounds excelRange = CellBounds.parse(range);
    for (int rowIndex = excelRange.firstRow(); rowIndex <= excelRange.lastRow(); rowIndex++) {
      Row row = sheet.getRow(rowIndex);
      if (row == null) {
        continue;
      }
      for (int columnIndex = excelRange.firstColumn();
          columnIndex <= excelRange.lastColumn();
          columnIndex++) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
          continue;
        }
        resetToDefaultStyle(cell);
        cell.removeHyperlink();
        CommentAdapter.clearCellComment(cell);
        cell.setBlank();
      }
    }
    syncTableHeaders(excelRange);
    return this;
  }

  public CellsAdapter setArrayFormula(String range, ResolvedArrayFormula formula) {
    TextValidation.requireNonBlank(range, "range");
    Objects.requireNonNull(formula, "formula must not be null");

    CellBounds excelRange = CellBounds.parse(range);
    CellRangeAddress poiRange =
        new CellRangeAddress(
            excelRange.firstRow(),
            excelRange.lastRow(),
            excelRange.firstColumn(),
            excelRange.lastColumn());
    try {
      FormulaLimits.requireSupportedFormula(
          new FormulaLimits.CellContext(
              xssfSheet().getWorkbook(), xssfSheet().getWorkbook().getSheetIndex(xssfSheet())),
          formula.formula());
      xssfSheet().setArrayFormula(formula.formula(), poiRange);
    } catch (RuntimeException exception) {
      throw FormulaExceptions.wrap(
          formulaRuntime,
          sheet.getSheetName(),
          new CellReference(excelRange.firstRow(), excelRange.firstColumn()).formatAsString(),
          formula.formula(),
          exception);
    }
    syncTableHeaders(excelRange);
    return this;
  }

  public CellsAdapter clearArrayFormula(String address) {
    TextValidation.requireNonBlank(address, "address");
    Cell cell = requiredArrayCell(address);
    if (!cell.isPartOfArrayFormulaGroup()) {
      throw new IllegalArgumentException("Cell " + address + " is not part of an array formula.");
    }
    CellRangeAddress removed = cell.getArrayFormulaRange();
    xssfSheet().removeArrayFormula(cell);
    syncTableHeaders(
        new CellBounds(
            removed.getFirstRow(),
            removed.getLastRow(),
            removed.getFirstColumn(),
            removed.getLastColumn()));
    return this;
  }

  public CellsAdapter applyStyle(String range, CellStylePatch style) {
    TextValidation.requireNonBlank(range, "range");
    Objects.requireNonNull(style, "style must not be null");

    CellBounds excelRange = CellBounds.parse(range);
    for (int rowIndex = excelRange.firstRow(); rowIndex <= excelRange.lastRow(); rowIndex++) {
      for (int columnIndex = excelRange.firstColumn();
          columnIndex <= excelRange.lastColumn();
          columnIndex++) {
        Cell cell = getOrCreateCell(rowIndex, columnIndex);
        cell.setCellStyle(styleRegistry.mergedStyle(cell, style));
      }
    }
    syncTableHeaders(excelRange);
    return this;
  }

  public CellsAdapter appendRow(ResolvedCellWrite... values) {
    Objects.requireNonNull(values, "values must not be null");
    for (ResolvedCellWrite value : values) {
      Objects.requireNonNull(value, "values must not contain nulls");
    }

    int rowIndex = nextAppendRowIndex();
    for (int columnIndex = 0; columnIndex < values.length; columnIndex++) {
      writeCellValue(rowIndex, columnIndex, values[columnIndex]);
    }
    if (values.length > 0) {
      syncTableHeaders(new CellBounds(rowIndex, rowIndex, 0, values.length - 1));
    }
    return this;
  }

  private void writeCellValue(int rowIndex, int columnIndex, ResolvedCellWrite value) {
    Cell cell = getOrCreateCell(rowIndex, columnIndex);

    switch (value) {
      case CellValue.Blank _ -> {
        clearExistingFormula(cell);
        cell.setBlank();
      }
      case CellValue.Text textValue -> {
        clearExistingFormula(cell);
        CellTextLimits.requireSupportedLength(textValue.value(), "textValue.value");
        cell.setCellValue(textValue.value());
      }
      case ResolvedRichText richTextValue -> {
        clearExistingFormula(cell);
        cell.setCellValue(StyleAdapter.toPoiRichText(xssfSheet().getWorkbook(), richTextValue));
      }
      case CellValue.NumberValue numberValue -> {
        clearExistingFormula(cell);
        cell.setCellValue(numberValue.value());
      }
      case CellValue.BooleanValue booleanValue -> {
        clearExistingFormula(cell);
        cell.setCellValue(booleanValue.value());
      }
      case CellValue.ErrorValue errorValue -> {
        byte storedError = CellErrors.toPoiStoredFormulaError(errorValue.value()).getCode();
        clearExistingFormula(cell);
        cell.setCellErrorValue(storedError);
      }
      case CellValueSpec.Date dateValue -> {
        clearExistingFormula(cell);
        cell.setCellValue(dateValue.value());
        cell.setCellStyle(styleRegistry.localDateStyle(cell));
      }
      case CellValueSpec.DateTime dateTimeValue -> {
        clearExistingFormula(cell);
        cell.setCellValue(dateTimeValue.value());
        cell.setCellStyle(styleRegistry.localDateTimeStyle(cell));
      }
      case ResolvedFormula formulaValue ->
          FormulaWrites.setAuthoredFormula(
              cell,
              formulaValue.value(),
              formulaRuntime,
              sheet.getSheetName(),
              new CellReference(rowIndex, columnIndex).formatAsString());
      case ResolvedRawFormula rawFormulaValue ->
          FormulaWrites.setOpaqueFormula(cell, rawFormulaValue.value());
    }
  }

  private static void clearExistingFormula(Cell cell) {
    if (cell.getCellType() == CellType.FORMULA) {
      cell.removeFormula();
    }
  }

  private void syncTableHeaders(CellBounds changedRange) {
    TableAdapter.syncAffectedHeaders(xssfSheet(), changedRange);
  }

  private Cell getOrCreateCell(int rowIndex, int columnIndex) {
    return getOrCreateRow(rowIndex)
        .getCell(columnIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
  }

  private Cell requiredArrayCell(String address) {
    CellReference cellReference = parseCellReference(address);
    requireValidCellReference(address, cellReference);
    Row row = sheet.getRow(cellReference.getRow());
    if (row == null) {
      throw new CellNotFoundException(address);
    }
    Cell cell = row.getCell(cellReference.getCol());
    if (cell == null) {
      throw new CellNotFoundException(address);
    }
    return cell;
  }

  public int nextAppendRowIndex() {
    int lastValueBearingRowIndex = -1;
    for (Row row : sheet) {
      if (rowHasValueBearingCell(row)) {
        lastValueBearingRowIndex = row.getRowNum();
      }
    }
    return lastValueBearingRowIndex + 1;
  }

  private static boolean rowHasValueBearingCell(Row row) {
    for (Cell cell : row) {
      if (cell.getCellType() != CellType.BLANK) {
        return true;
      }
    }
    return false;
  }

  private Row getOrCreateRow(int rowIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      row = sheet.createRow(rowIndex);
    }
    return row;
  }

  private void resetToDefaultStyle(Cell cell) {
    cell.setCellStyle(styleRegistry.defaultStyle());
  }

  private static List<List<ResolvedCellWrite>> copyRows(List<List<ResolvedCellWrite>> rows) {
    List<List<ResolvedCellWrite>> copiedRows = new ArrayList<>(rows.size());
    int expectedWidth = -1;
    for (List<ResolvedCellWrite> row : rows) {
      Objects.requireNonNull(row, "rows must not contain nulls");
      List<ResolvedCellWrite> copiedRow = List.copyOf(row);
      if (copiedRow.isEmpty()) {
        throw new IllegalArgumentException("rows must not contain empty rows");
      }
      if (expectedWidth < 0) {
        expectedWidth = copiedRow.size();
      } else if (copiedRow.size() != expectedWidth) {
        throw new IllegalArgumentException("rows must describe a rectangular matrix");
      }
      for (ResolvedCellWrite value : copiedRow) {
        Objects.requireNonNull(value, "rows must not contain null cell values");
      }
      copiedRows.add(copiedRow);
    }
    return List.copyOf(copiedRows);
  }
}
