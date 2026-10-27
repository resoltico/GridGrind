package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.formula.ArrayFormulaReport;
import dev.erst.gridgrind.core.model.layout.WindowRowReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.plan.CellReadFacet;
import dev.erst.gridgrind.core.plan.CellReadProjection;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jspecify.annotations.Nullable;

/** Typed cell reads, snapshots, windows, and previews for one sheet wrapper. */
final class ExcelSheetCellReadSupport {
  private final Sheet sheet;
  private final WorkbookStyleRegistry styleRegistry;
  private final ExcelFormulaRuntime formulaRuntime;
  private final DataFormatter dataFormatter;

  ExcelSheetCellReadSupport(
      Sheet sheet,
      WorkbookStyleRegistry styleRegistry,
      ExcelFormulaRuntime formulaRuntime,
      DataFormatter dataFormatter) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.styleRegistry = Objects.requireNonNull(styleRegistry, "styleRegistry must not be null");
    this.formulaRuntime = Objects.requireNonNull(formulaRuntime, "formulaRuntime must not be null");
    this.dataFormatter = Objects.requireNonNull(dataFormatter, "dataFormatter must not be null");
  }

  String text(String address) {
    ExcelSheet.requireNonBlank(address, "address");
    return requiredCell(address).getStringCellValue();
  }

  double number(String address) {
    ExcelSheet.requireNonBlank(address, "address");
    Cell cell = requiredCell(address);
    if (cell.getCellType() == CellType.FORMULA) {
      CellValue evaluatedCell;
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

  boolean bool(String address) {
    ExcelSheet.requireNonBlank(address, "address");
    Cell cell = requiredCell(address);
    if (cell.getCellType() == CellType.FORMULA) {
      CellValue evaluatedCell;
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

  String formula(String address) {
    ExcelSheet.requireNonBlank(address, "address");
    Cell cell = requiredCell(address);
    if (cell.getCellType() != CellType.FORMULA) {
      throw new IllegalStateException("Cell does not contain a formula: " + address);
    }
    return cell.getCellFormula();
  }

  CellReport snapshotCell(String address) {
    return snapshotCell(address, new CellReadProjection(java.util.List.of(CellReadFacet.values())));
  }

  CellReport snapshotCell(String address, CellReadProjection projection) {
    ExcelSheet.requireNonBlank(address, "address");
    Objects.requireNonNull(projection, "projection must not be null");
    CellReference cellReference = parseCellReference(address);
    requireValidCellReference(address, cellReference);
    return snapshotCellOrBlank(address, cellReference.getRow(), cellReference.getCol(), projection);
  }

  List<CellReport> snapshotCells(List<String> addresses) {
    return snapshotCells(
        addresses, new CellReadProjection(java.util.List.of(CellReadFacet.values())));
  }

  List<CellReport> snapshotCells(List<String> addresses, CellReadProjection projection) {
    Objects.requireNonNull(addresses, "addresses must not be null");
    Objects.requireNonNull(projection, "projection must not be null");
    List<CellReport> cells = new ArrayList<>(addresses.size());
    for (String address : List.copyOf(addresses)) {
      ExcelSheet.requireNonBlank(address, "address");
      cells.add(snapshotCell(address, projection));
    }
    return List.copyOf(cells);
  }

  List<WindowRowReport> preview(int maxRows, int maxColumns, int lastRowIndex) {
    if (maxRows <= 0) {
      throw new IllegalArgumentException("maxRows must be greater than 0");
    }
    if (maxColumns <= 0) {
      throw new IllegalArgumentException("maxColumns must be greater than 0");
    }

    List<WindowRowReport> previewRows = new ArrayList<>();
    int endingRowIndex = Math.min(lastRowIndex, maxRows - 1);
    for (int rowIndex = 0; rowIndex <= endingRowIndex; rowIndex++) {
      previewRows.add(new WindowRowReport(rowIndex, previewRow(rowIndex, maxColumns)));
    }
    return List.copyOf(previewRows);
  }

  dev.erst.gridgrind.core.model.layout.WindowReport.Dense window(
      String topLeftAddress, int rowCount, int columnCount) {
    return window(
        topLeftAddress,
        rowCount,
        columnCount,
        new CellReadProjection(java.util.List.of(CellReadFacet.values())));
  }

  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  dev.erst.gridgrind.core.model.layout.WindowReport.Dense window(
      String topLeftAddress, int rowCount, int columnCount, CellReadProjection projection) {
    ExcelSheet.requireNonBlank(topLeftAddress, "topLeftAddress");
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

  List<ArrayFormulaReport> arrayFormulas() {
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
        if (ExcelSheetStructureSupport.shouldPreview(cell)) {
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
            ? ExcelSheetHyperlinkSupport.hyperlink(cell)
            : java.util.Optional.<dev.erst.gridgrind.core.model.value.HyperlinkTarget>empty();
    var comment =
        projection.includes(CellReadFacet.COMMENT)
            ? ExcelSheetCommentSupport.commentSnapshot(cell)
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
                  ? ExcelRichTextSupport.snapshot(
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
                      ExcelCellErrorLiteralSupport.toReportedWireValue(
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
      @Nullable CellValue value) {
    if (value == null) return new dev.erst.gridgrind.core.model.value.CellValue.Blank();
    return switch (value.getCellType()) {
      case STRING -> new dev.erst.gridgrind.core.model.value.CellValue.Text(value.getStringValue());
      case NUMERIC ->
          new dev.erst.gridgrind.core.model.value.CellValue.NumberValue(value.getNumberValue());
      case BOOLEAN ->
          new dev.erst.gridgrind.core.model.value.CellValue.BooleanValue(value.getBooleanValue());
      case ERROR ->
          new dev.erst.gridgrind.core.model.value.CellValue.ErrorValue(
              ExcelCellErrorLiteralSupport.toReportedWireValue(
                  FormulaError.forInt(value.getErrorValue())));
      default -> new dev.erst.gridgrind.core.model.value.CellValue.Blank();
    };
  }

  private dev.erst.gridgrind.core.model.value.CellTemporalReport temporal(
      double number, String numberFormat) {
    var kind = ExcelTemporalFormatSupport.observedKind(numberFormat);
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

  private static void requireValidCellReference(String address, CellReference cellReference) {
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
}
