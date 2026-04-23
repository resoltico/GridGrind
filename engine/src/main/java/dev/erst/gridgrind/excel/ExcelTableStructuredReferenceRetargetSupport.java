package dev.erst.gridgrind.excel;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Retargets table structured-reference formulas after POI clone-sheet rewrites table names. */
final class ExcelTableStructuredReferenceRetargetSupport {
  private ExcelTableStructuredReferenceRetargetSupport() {}

  /**
   * Rewrites formula cells on one copied sheet from transient cloned table names to final names.
   */
  static void retargetFormulaCells(
      XSSFSheet sheet, List<String> transientTableNames, List<String> finalTableNames) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(transientTableNames, "transientTableNames must not be null");
    Objects.requireNonNull(finalTableNames, "finalTableNames must not be null");

    Map<String, String> tableNameMapping = tableNameMapping(transientTableNames, finalTableNames);
    if (tableNameMapping.isEmpty()) {
      return;
    }
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (cell.getCellType() != CellType.FORMULA) {
          continue;
        }
        String rewritten = retargetFormula(cell.getCellFormula(), tableNameMapping);
        if (!rewritten.equals(cell.getCellFormula())) {
          ExcelFormulaWriteSupport.setRewrittenFormula(
              cell, rewritten, "Table structured-reference retargeting");
        }
      }
    }
  }

  static String retargetFormula(String formula, Map<String, String> tableNameMapping) {
    Objects.requireNonNull(formula, "formula must not be null");
    Objects.requireNonNull(tableNameMapping, "tableNameMapping must not be null");
    String rewritten = formula;
    for (Map.Entry<String, String> entry : tableNameMapping.entrySet()) {
      rewritten = replaceTableName(rewritten, entry.getKey(), entry.getValue());
    }
    return rewritten;
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static Map<String, String> tableNameMapping(
      List<String> transientTableNames, List<String> finalTableNames) {
    int mappingCount = Math.min(transientTableNames.size(), finalTableNames.size());
    Map<String, String> mappings = new LinkedHashMap<>();
    for (int index = 0; index < mappingCount; index++) {
      String transientName = requireNonBlank(transientTableNames.get(index), "transientTableNames");
      String finalName = requireNonBlank(finalTableNames.get(index), "finalTableNames");
      if (!transientName.equals(finalName)) {
        mappings.put(transientName, finalName);
      }
    }
    return Map.copyOf(mappings);
  }

  private static String replaceTableName(String formula, String transientName, String finalName) {
    Pattern pattern =
        Pattern.compile("(?<![A-Za-z0-9_\\.])" + Pattern.quote(transientName) + "(?=\\[)");
    Matcher matcher = pattern.matcher(formula);
    return matcher.find() ? matcher.replaceAll(Matcher.quoteReplacement(finalName)) : formula;
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not contain null values");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not contain blank values");
    }
    return value;
  }
}
