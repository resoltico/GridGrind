package dev.erst.gridgrind.engine.cells;

import dev.erst.gridgrind.core.limits.Limit;
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
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.engine.backend.xssf.CellNotFoundException;
import dev.erst.gridgrind.engine.comments.CommentAdapter;
import dev.erst.gridgrind.engine.formulas.FormulaExceptions;
import dev.erst.gridgrind.engine.formulas.FormulaLimits;
import dev.erst.gridgrind.engine.formulas.FormulaRuntime;
import dev.erst.gridgrind.engine.links.LinkAdapter;
import dev.erst.gridgrind.engine.poi.FormulaWrites;
import dev.erst.gridgrind.engine.poi.SheetStructure;
import dev.erst.gridgrind.engine.styles.StyleAdapter;
import dev.erst.gridgrind.engine.tables.TableAdapter;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
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
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_TEXT_CHARACTERS_PER_CELL,
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_ROWS_PER_WORKSHEET,
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_COLUMNS_PER_WORKSHEET
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

  public String text(CellAddress address) {
    Objects.requireNonNull(address, "address must not be null");
    return requiredCell(address).getStringCellValue();
  }

  public double number(CellAddress address) {
    Objects.requireNonNull(address, "address must not be null");
    Cell cell = requiredCell(address);
    if (cell.getCellType() == CellType.FORMULA) {
      org.apache.poi.ss.usermodel.CellValue evaluatedCell;
      try {
        evaluatedCell = formulaRuntime.evaluate(cell);
      } catch (RuntimeException exception) {
        throw FormulaExceptions.wrap(
            formulaRuntime,
            sheet.getSheetName(),
            address.value(),
            cell.getCellFormula(),
            exception);
      }
      if (evaluatedCell == null || evaluatedCell.getCellType() != CellType.NUMERIC) {
        throw new IllegalStateException("Cell does not evaluate to a numeric value: " + address);
      }
      return evaluatedCell.getNumberValue();
    }
    return cell.getNumericCellValue();
  }

  public boolean bool(CellAddress address) {
    Objects.requireNonNull(address, "address must not be null");
    Cell cell = requiredCell(address);
    if (cell.getCellType() == CellType.FORMULA) {
      org.apache.poi.ss.usermodel.CellValue evaluatedCell;
      try {
        evaluatedCell = formulaRuntime.evaluate(cell);
      } catch (RuntimeException exception) {
        throw FormulaExceptions.wrap(
            formulaRuntime,
            sheet.getSheetName(),
            address.value(),
            cell.getCellFormula(),
            exception);
      }
      if (evaluatedCell == null || evaluatedCell.getCellType() != CellType.BOOLEAN) {
        throw new IllegalStateException("Cell does not evaluate to a boolean value: " + address);
      }
      return evaluatedCell.getBooleanValue();
    }
    return cell.getBooleanCellValue();
  }

  public String formula(CellAddress address) {
    Objects.requireNonNull(address, "address must not be null");
    Cell cell = requiredCell(address);
    if (cell.getCellType() != CellType.FORMULA) {
      throw new IllegalStateException("Cell does not contain a formula: " + address);
    }
    return cell.getCellFormula();
  }

  public CellReport snapshotCell(CellAddress address) {
    return snapshotCell(address, new CellReadProjection(java.util.List.of(CellReadFacet.values())));
  }

  public CellReport snapshotCell(CellAddress address, CellReadProjection projection) {
    Objects.requireNonNull(address, "address must not be null");
    Objects.requireNonNull(projection, "projection must not be null");
    return snapshotCellOrBlank(address, projection);
  }

  public List<CellReport> snapshotCells(List<CellAddress> addresses) {
    return snapshotCells(
        addresses, new CellReadProjection(java.util.List.of(CellReadFacet.values())));
  }

  public List<CellReport> snapshotCells(
      List<CellAddress> addresses, CellReadProjection projection) {
    Objects.requireNonNull(addresses, "addresses must not be null");
    Objects.requireNonNull(projection, "projection must not be null");
    List<CellReport> cells = new ArrayList<>(addresses.size());
    for (CellAddress address : List.copyOf(addresses)) {
      Objects.requireNonNull(address, "address must not be null");
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
      CellAddress topLeftAddress, int rowCount, int columnCount) {
    return window(
        topLeftAddress,
        rowCount,
        columnCount,
        new CellReadProjection(java.util.List.of(CellReadFacet.values())));
  }

  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  public dev.erst.gridgrind.core.model.layout.WindowReport.Dense window(
      CellAddress topLeftAddress, int rowCount, int columnCount, CellReadProjection projection) {
    Objects.requireNonNull(topLeftAddress, "topLeftAddress must not be null");
    Objects.requireNonNull(projection, "projection must not be null");
    if (rowCount <= 0) {
      throw new IllegalArgumentException("rowCount must be greater than 0");
    }
    if (columnCount <= 0) {
      throw new IllegalArgumentException("columnCount must be greater than 0");
    }

    CellAddress topLeft = topLeftAddress;
    long lastRow = (long) topLeft.rowIndex() + rowCount - 1;
    long lastCol = (long) topLeft.columnIndex() + columnCount - 1;
    if (lastRow > Limit.MAX_ROW_INDEX || lastCol > Limit.MAX_COLUMN_INDEX) {
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
      int rowIndex = topLeft.rowIndex() + rowOffset;
      List<CellReport> cells = new ArrayList<>(columnCount);
      for (int columnOffset = 0; columnOffset < columnCount; columnOffset++) {
        int columnIndex = topLeft.columnIndex() + columnOffset;
        CellAddress address = new CellAddress(rowIndex, columnIndex);
        cells.add(snapshotCellOrBlank(address, projection));
      }
      rows.add(
          new dev.erst.gridgrind.core.model.layout.WindowRowReport(rowIndex, List.copyOf(cells)));
    }
    return new dev.erst.gridgrind.core.model.layout.WindowReport.Dense(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheet.getSheetName()),
        topLeftAddress,
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
                new CellAddress(range.getFirstRow(), range.getFirstColumn()),
                xssfSheet()
                    .getRow(range.getFirstRow())
                    .getCell(range.getFirstColumn())
                    .getCellFormula(),
                range.getNumberOfCells() == 1));
      }
    }
    return List.copyOf(formulas);
  }

  private Cell requiredCell(CellAddress address) {
    Row row = sheet.getRow(address.rowIndex());
    if (row == null) {
      throw new CellNotFoundException(address.value());
    }

    Cell cell = row.getCell(address.columnIndex());
    if (cell == null) {
      throw new CellNotFoundException(address.value());
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
                  new CellAddress(rowIndex, columnIndex),
                  cell,
                  CellReadProjection.of(CellReadFacet.VALUE, CellReadFacet.FORMAT)));
        }
      }
    }
    return List.copyOf(cells);
  }

  private CellReport snapshotCellOrBlank(CellAddress address, CellReadProjection projection) {
    Row row = sheet.getRow(address.rowIndex());
    Cell cell = row == null ? null : row.getCell(address.columnIndex());
    if (cell == null) return blankReport(address, projection);
    return snapshot(address, cell, projection);
  }

  private CellReport snapshot(CellAddress address, Cell cell, CellReadProjection projection) {
    var coordinate = address;
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
      throw FormulaExceptions.wrap(
          formulaRuntime, sheet.getSheetName(), address.value(), formula, failure);
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
              formulaRuntime, sheet.getSheetName(), address.value(), formula, failure);
        }
      }
      var value =
          projection.includes(CellReadFacet.VALUE)
              ? java.util.Optional.of(evaluatedValue(evaluated))
              : java.util.Optional.<dev.erst.gridgrind.core.model.value.CellValue>empty();
      var temporal =
          projection.includes(CellReadFacet.TEMPORAL)
                  && evaluatedValue(evaluated)
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
    return projection.includes(CellReadFacet.VALUE)
        || projection.includes(CellReadFacet.FORMAT)
        || projection.includes(CellReadFacet.TEMPORAL);
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
    return TemporalFormats.observedKind(numberFormat)
        .<dev.erst.gridgrind.core.model.value.CellTemporalReport>map(kind ->
            dev.erst.gridgrind.core.model.value.TemporalSemantics.interpret(
                number, kind, dateSystem()))
        .orElseGet(dev.erst.gridgrind.core.model.value.CellTemporalReport.NotTemporal::new);
  }

  private dev.erst.gridgrind.core.model.value.ExcelDateSystem dateSystem() {
    return xssfSheet().getWorkbook().isDate1904()
        ? dev.erst.gridgrind.core.model.value.ExcelDateSystem.MAC_1904
        : dev.erst.gridgrind.core.model.value.ExcelDateSystem.WINDOWS_1900;
  }

  private CellReport blankReport(CellAddress address, CellReadProjection projection) {
    return new CellReport.BlankReport(
        address,
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

  public CellsAdapter setCell(CellAddress address, ResolvedCellWrite value) {
    Objects.requireNonNull(address, "address must not be null");
    Objects.requireNonNull(value, "value must not be null");

    dev.erst.gridgrind.core.model.value.TemporalSemantics.requireWritable(value, dateSystem());
    writeCellValue(address.rowIndex(), address.columnIndex(), value);
    syncTableHeaders(
        new RangeRef(
            address.rowIndex(), address.rowIndex(), address.columnIndex(), address.columnIndex()));
    return this;
  }

  public CellsAdapter setRange(RangeRef range, List<List<ResolvedCellWrite>> rows) {
    Objects.requireNonNull(range, "range must not be null");
    Objects.requireNonNull(rows, "rows must not be null");
    if (rows.isEmpty()) {
      throw new IllegalArgumentException("rows must not be empty");
    }

    List<List<ResolvedCellWrite>> copiedRows = copyRows(rows);
    copiedRows.forEach(row ->
        row.forEach(value ->
            dev.erst.gridgrind.core.model.value.TemporalSemantics.requireWritable(
                value, dateSystem())));
    RangeRef excelRange = range;
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

  public CellsAdapter clearRange(RangeRef range) {
    Objects.requireNonNull(range, "range must not be null");
    RangeRef excelRange = range;
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

  public CellsAdapter setArrayFormula(RangeRef range, ResolvedArrayFormula formula) {
    Objects.requireNonNull(range, "range must not be null");
    Objects.requireNonNull(formula, "formula must not be null");

    RangeRef excelRange = range;
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

  public CellsAdapter clearArrayFormula(CellAddress address) {
    Objects.requireNonNull(address, "address must not be null");
    Cell cell = requiredCell(address);
    if (!cell.isPartOfArrayFormulaGroup()) {
      throw new IllegalArgumentException("Cell " + address + " is not part of an array formula.");
    }
    CellRangeAddress removed = cell.getArrayFormulaRange();
    xssfSheet().removeArrayFormula(cell);
    syncTableHeaders(
        new RangeRef(
            removed.getFirstRow(),
            removed.getLastRow(),
            removed.getFirstColumn(),
            removed.getLastColumn()));
    return this;
  }

  public CellsAdapter applyStyle(RangeRef range, CellStylePatch style) {
    Objects.requireNonNull(range, "range must not be null");
    Objects.requireNonNull(style, "style must not be null");

    RangeRef excelRange = range;
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
      dev.erst.gridgrind.core.model.value.TemporalSemantics.requireWritable(value, dateSystem());
    }

    int rowIndex = nextAppendRowIndex();
    for (int columnIndex = 0; columnIndex < values.length; columnIndex++) {
      writeCellValue(rowIndex, columnIndex, values[columnIndex]);
    }
    if (values.length > 0) {
      syncTableHeaders(new RangeRef(rowIndex, rowIndex, 0, values.length - 1));
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
      case dev.erst.gridgrind.core.model.value.ResolvedCellText textValue -> {
        clearExistingFormula(cell);
        cell.setCellValue(textValue.value());
      }
      case ResolvedRichText richTextValue -> {
        clearExistingFormula(cell);
        cell.setCellValue(StyleAdapter.toPoiRichText(xssfSheet().getWorkbook(), richTextValue));
      }
      case CellValueSpec.NumberValue numberValue -> {
        clearExistingFormula(cell);
        cell.setCellValue(numberValue.value());
      }
      case CellValue.BooleanValue booleanValue -> {
        clearExistingFormula(cell);
        cell.setCellValue(booleanValue.value());
      }
      case CellValueSpec.ErrorValue errorValue -> {
        byte storedError = CellErrors.toPoiStoredFormulaError(errorValue.value()).getCode();
        clearExistingFormula(cell);
        cell.setCellErrorValue(storedError);
      }
      case CellValueSpec.Date dateValue -> {
        clearExistingFormula(cell);
        cell.setCellValue(
            dev.erst.gridgrind.core.model.value.TemporalSemantics.serial(
                dateValue.value(), dateSystem()));
        cell.setCellStyle(styleRegistry.localDateStyle(cell));
      }
      case CellValueSpec.DateTime dateTimeValue -> {
        clearExistingFormula(cell);
        cell.setCellValue(
            dev.erst.gridgrind.core.model.value.TemporalSemantics.serial(
                dateTimeValue.value(), dateSystem()));
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

  private void syncTableHeaders(RangeRef changedRange) {
    TableAdapter.syncAffectedHeaders(xssfSheet(), changedRange);
  }

  private Cell getOrCreateCell(int rowIndex, int columnIndex) {
    return getOrCreateRow(rowIndex)
        .getCell(columnIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
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
