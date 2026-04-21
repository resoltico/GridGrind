package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Chart source-formula resolution and scalar decoding helpers. */
final class ExcelChartSourceSupport {
  private ExcelChartSourceSupport() {}

  static Name resolveDefinedNameReference(XSSFSheet contextSheet, String formula) {
    if (!formula.matches("^[A-Za-z_][A-Za-z0-9_.]*$")) {
      return null;
    }

    int contextSheetIndex = contextSheet.getWorkbook().getSheetIndex(contextSheet);
    Name workbookScopedMatch = null;
    for (Name name : contextSheet.getWorkbook().getAllNames()) {
      if (!formula.equals(name.getNameName())) {
        continue;
      }
      if (name.getSheetIndex() == contextSheetIndex) {
        return name;
      }
      if (name.getSheetIndex() < 0) {
        workbookScopedMatch = name;
      }
    }
    return workbookScopedMatch;
  }

  static String normalizeAreaFormulaForPoi(String formula) {
    return ExcelNamedRangeTargets.normalizeAreaFormulaForPoi(formula);
  }

  static ExcelDrawingController.CellScalar scalarFromFormula(Cell cell) {
    return switch (cell.getCachedFormulaResultType()) {
      case STRING ->
          new ExcelDrawingController.CellScalar(
              ExcelDrawingController.CellScalarKind.STRING, cell.getStringCellValue(), 0d);
      case NUMERIC ->
          new ExcelDrawingController.CellScalar(
              ExcelDrawingController.CellScalarKind.NUMERIC, null, cell.getNumericCellValue());
      case BOOLEAN ->
          new ExcelDrawingController.CellScalar(
              ExcelDrawingController.CellScalarKind.STRING,
              Boolean.toString(cell.getBooleanCellValue()),
              0d);
      case BLANK, _NONE ->
          new ExcelDrawingController.CellScalar(
              ExcelDrawingController.CellScalarKind.STRING, "", 0d);
      case ERROR ->
          throw new IllegalArgumentException("Chart source formulas must not cache error values");
      case FORMULA ->
          throw new IllegalArgumentException(
              "Chart source formulas must expose a cached scalar result");
    };
  }

  static String requiredDefinedNameFormula(Name definedName) {
    Objects.requireNonNull(definedName, "definedName must not be null");
    String refersToFormula = definedName.getRefersToFormula();
    if (refersToFormula == null || refersToFormula.isBlank()) {
      throw new IllegalArgumentException(
          "Defined name '"
              + definedName.getNameName()
              + "' does not resolve to one chart source area");
    }
    return refersToFormula;
  }

  static org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> toCategoryDataSource(
      XSSFSheet sheet, String formula) {
    ResolvedChartSource source = resolveChartSource(sheet, formula);
    return source.numeric()
        ? org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
            source.numericValues().toArray(Double[]::new), source.referenceFormula())
        : org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
            source.stringValues().toArray(String[]::new), source.referenceFormula());
  }

  static org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource<? extends Number>
      toValueDataSource(XSSFSheet sheet, String formula) {
    ResolvedChartSource source = resolveChartSource(sheet, formula);
    if (!source.numeric()) {
      throw new IllegalArgumentException(
          "Chart value source must resolve to numeric cells: " + formula);
    }
    return org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory.fromArray(
        source.numericValues().toArray(Double[]::new), source.referenceFormula());
  }

  static ResolvedChartSource resolveChartSource(XSSFSheet sheet, String formula) {
    String normalizedFormula = normalizeFormula(formula);
    ResolvedAreaReference resolved = resolveAreaReference(sheet, normalizedFormula);
    List<String> stringValues = new ArrayList<>();
    List<Double> numericValues = new ArrayList<>();
    boolean numeric = true;
    for (CellReference reference : resolved.areaReference().getAllReferencedCells()) {
      Cell cell =
          resolved.sheet().getRow(reference.getRow()) == null
              ? null
              : resolved.sheet().getRow(reference.getRow()).getCell(reference.getCol());
      ExcelDrawingController.CellScalar scalar = scalar(cell);
      if (scalar.kind() == ExcelDrawingController.CellScalarKind.NUMERIC) {
        stringValues.add(Double.toString(scalar.number()));
        numericValues.add(scalar.number());
      } else {
        numeric = false;
        stringValues.add(scalar.text());
      }
    }
    return new ResolvedChartSource(
        normalizedFormula,
        resolved.sheet(),
        resolved.areaReference(),
        numeric,
        stringValues,
        numericValues);
  }

  static ResolvedAreaReference resolveAreaReference(XSSFSheet contextSheet, String formula) {
    Name definedName = resolveDefinedNameReference(contextSheet, formula);
    if (definedName != null) {
      return resolveDefinedName(contextSheet, definedName);
    }

    AreaReference[] references =
        parseContiguousAreaReferences(
            normalizeAreaFormulaForPoi(formula),
            "Chart source formula must resolve to one contiguous area: " + formula);
    if (references.length != 1) {
      throw new IllegalArgumentException(
          "Chart source formula must resolve to one contiguous area: " + formula);
    }
    return new ResolvedAreaReference(
        requireAreaReferenceSheet(contextSheet, references[0], formula), references[0]);
  }

  static CellReference resolveSingleCellReference(
      XSSFSheet sheet, String formula, String detailLabel) {
    ResolvedAreaReference resolved = resolveAreaReference(sheet, formula);
    if (!resolved.areaReference().isSingleCell()) {
      throw new IllegalArgumentException(
          detailLabel + " must resolve to a single cell: " + formula);
    }
    CellReference firstCell = resolved.areaReference().getFirstCell();
    return new CellReference(
        resolved.sheet().getSheetName(), firstCell.getRow(), firstCell.getCol(), true, true);
  }

  static String scalarText(XSSFSheet sheet, CellReference reference) {
    Cell cell =
        sheet.getWorkbook().getSheet(reference.getSheetName()).getRow(reference.getRow()) == null
            ? null
            : sheet
                .getWorkbook()
                .getSheet(reference.getSheetName())
                .getRow(reference.getRow())
                .getCell(reference.getCol());
    ExcelDrawingController.CellScalar scalar = scalar(cell);
    return scalar.kind() == ExcelDrawingController.CellScalarKind.NUMERIC
        ? Double.toString(scalar.number())
        : scalar.text();
  }

  static String normalizeFormula(String formula) {
    String normalized = requireNonBlank(formula, "formula");
    return normalized.startsWith("=") ? normalized.substring(1) : normalized;
  }

  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  static String nullIfBlank(String value) {
    return value == null || value.isBlank() ? null : value;
  }

  private static ResolvedAreaReference resolveDefinedName(
      XSSFSheet contextSheet, Name definedName) {
    String refersToFormula = requiredDefinedNameFormula(definedName);
    AreaReference[] references =
        parseContiguousAreaReferences(
            normalizeAreaFormulaForPoi(refersToFormula),
            "Defined name '" + definedName.getNameName() + "' must resolve to one contiguous area");
    if (references.length != 1) {
      throw new IllegalArgumentException(
          "Defined name '" + definedName.getNameName() + "' must resolve to one contiguous area");
    }
    XSSFSheet resolvedSheet =
        references[0].getFirstCell().getSheetName() == null
            ? requireDefinedNameSheet(contextSheet, definedName)
            : requireSheet(
                contextSheet.getWorkbook(),
                references[0].getFirstCell().getSheetName(),
                definedName.getNameName());
    return new ResolvedAreaReference(resolvedSheet, references[0]);
  }

  private static XSSFSheet requireAreaReferenceSheet(
      XSSFSheet contextSheet, AreaReference reference, String formula) {
    String sheetName = reference.getFirstCell().getSheetName();
    return sheetName == null
        ? contextSheet
        : requireSheet(contextSheet.getWorkbook(), sheetName, formula);
  }

  private static XSSFSheet requireDefinedNameSheet(XSSFSheet contextSheet, Name definedName) {
    if (definedName.getSheetIndex() >= 0) {
      return contextSheet.getWorkbook().getSheetAt(definedName.getSheetIndex());
    }
    return contextSheet;
  }

  private static AreaReference[] parseContiguousAreaReferences(
      String formula, String failureMessage) {
    try {
      return AreaReference.generateContiguous(SpreadsheetVersion.EXCEL2007, formula);
    } catch (IllegalArgumentException | IllegalStateException exception) {
      throw new IllegalArgumentException(failureMessage, exception);
    }
  }

  private static XSSFSheet requireSheet(XSSFWorkbook workbook, String sheetName, String formula) {
    XSSFSheet sheet = workbook.getSheet(sheetName);
    if (sheet == null) {
      throw new IllegalArgumentException(
          "Chart source formula '" + formula + "' resolves to missing sheet '" + sheetName + "'");
    }
    return sheet;
  }

  private static ExcelDrawingController.CellScalar scalar(Cell cell) {
    if (cell == null) {
      return new ExcelDrawingController.CellScalar(
          ExcelDrawingController.CellScalarKind.STRING, "", 0d);
    }
    return switch (cell.getCellType()) {
      case STRING ->
          new ExcelDrawingController.CellScalar(
              ExcelDrawingController.CellScalarKind.STRING, cell.getStringCellValue(), 0d);
      case NUMERIC ->
          new ExcelDrawingController.CellScalar(
              ExcelDrawingController.CellScalarKind.NUMERIC, null, cell.getNumericCellValue());
      case BOOLEAN ->
          new ExcelDrawingController.CellScalar(
              ExcelDrawingController.CellScalarKind.STRING,
              Boolean.toString(cell.getBooleanCellValue()),
              0d);
      case BLANK, _NONE ->
          new ExcelDrawingController.CellScalar(
              ExcelDrawingController.CellScalarKind.STRING, "", 0d);
      case FORMULA -> scalarFromFormula(cell);
      case ERROR ->
          throw new IllegalArgumentException("Chart source cells must not contain error values");
    };
  }
}
