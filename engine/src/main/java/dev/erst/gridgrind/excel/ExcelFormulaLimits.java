package dev.erst.gridgrind.excel;

import java.util.ArrayDeque;
import java.util.Deque;
import java.util.Objects;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Centralized authored-formula ceiling checks before POI persists a workbook formula token tree.
 */
final class ExcelFormulaLimits {
  static final int MAX_FORMULA_LENGTH = 8192; // LIM-013
  static final int MAX_NESTED_FUNCTION_LEVELS = 64; // LIM-014
  static final int MAX_FUNCTION_ARGUMENTS =
      SpreadsheetVersion.EXCEL2007.getMaxFunctionArgs(); // LIM-015

  private ExcelFormulaLimits() {}

  static void requireSupportedFormula(CellContext cellContext, String formula) {
    Objects.requireNonNull(cellContext, "cellContext must not be null");
    Objects.requireNonNull(formula, "formula must not be null");
    if (formula.length() > MAX_FORMULA_LENGTH) {
      throw new IllegalArgumentException(
          "formula must not exceed "
              + MAX_FORMULA_LENGTH
              + " characters (Excel formula length limit)");
    }

    FormulaShape shape = scanFormulaShape(formula);
    if (shape.maximumFunctionNesting() > MAX_NESTED_FUNCTION_LEVELS) {
      throw new IllegalArgumentException(
          "formula must not exceed "
              + MAX_NESTED_FUNCTION_LEVELS
              + " nested function levels (Excel formula nesting limit)");
    }
    if (shape.maximumFunctionArguments() > MAX_FUNCTION_ARGUMENTS) {
      throw new IllegalArgumentException(
          "formula must not exceed "
              + MAX_FUNCTION_ARGUMENTS
              + " function arguments (Excel function argument limit)");
    }

    Ptg[] tokens =
        FormulaParser.parse(
            formula,
            XSSFEvaluationWorkbook.create(xssfWorkbook(cellContext)),
            FormulaType.CELL,
            cellContext.sheetIndex());
    for (Ptg ignored : tokens) {
      // Force POI to parse the authored formula after GridGrind-owned limit checks run.
    }
  }

  private static XSSFWorkbook xssfWorkbook(CellContext cellContext) {
    if (cellContext.workbook() instanceof XSSFWorkbook) {
      return (XSSFWorkbook) cellContext.workbook();
    }
    if (cellContext.workbook() instanceof SXSSFWorkbook) {
      return ((SXSSFWorkbook) cellContext.workbook()).getXSSFWorkbook();
    }
    throw new IllegalArgumentException(
        "Formula limits require an XSSF workbook-backed cell context");
  }

  record CellContext(org.apache.poi.ss.usermodel.Workbook workbook, int sheetIndex) {
    CellContext {
      Objects.requireNonNull(workbook, "workbook must not be null");
      if (sheetIndex < 0) {
        throw new IllegalArgumentException("sheetIndex must not be negative");
      }
    }
  }

  private static FormulaShape scanFormulaShape(String formula) {
    Deque<FunctionFrame> functions = new ArrayDeque<>();
    boolean inString = false;
    int maximumFunctionNesting = 0;
    int maximumFunctionArguments = 0;
    int index = 0;
    while (index < formula.length()) {
      char current = formula.charAt(index);
      if (current == '"') {
        if (inString && index + 1 < formula.length() && formula.charAt(index + 1) == '"') {
          index += 2;
          continue;
        } else {
          inString = !inString;
        }
        index++;
        continue;
      }
      if (inString) {
        index++;
        continue;
      }
      if (current == '(' && looksLikeFunctionCall(formula, index)) {
        functions.addLast(new FunctionFrame());
        maximumFunctionNesting = Math.max(maximumFunctionNesting, functions.size());
        index++;
        continue;
      }
      if (current == ',' && !functions.isEmpty()) {
        functions.getLast().argumentSeparators++;
        functions.getLast().hasContent = true;
        index++;
        continue;
      }
      if (current == ')' && !functions.isEmpty()) {
        FunctionFrame completed = functions.removeLast();
        int argumentCount = completed.hasContent ? completed.argumentSeparators + 1 : 0;
        maximumFunctionArguments = Math.max(maximumFunctionArguments, argumentCount);
        index++;
        continue;
      }
      if (!functions.isEmpty() && !Character.isWhitespace(current)) {
        functions.getLast().hasContent = true;
      }
      index++;
    }
    return new FormulaShape(maximumFunctionNesting, maximumFunctionArguments);
  }

  private static boolean looksLikeFunctionCall(String formula, int openParenIndex) {
    int index = openParenIndex - 1;
    while (index >= 0 && Character.isWhitespace(formula.charAt(index))) {
      index--;
    }
    if (index < 0 || !isFunctionIdentifierCharacter(formula.charAt(index))) {
      return false;
    }
    while (index >= 0 && isFunctionIdentifierCharacter(formula.charAt(index))) {
      index--;
    }
    return true;
  }

  private static boolean isFunctionIdentifierCharacter(char value) {
    return Character.isLetterOrDigit(value) || value == '_' || value == '.';
  }

  /**
   * Tracks the current function scope while counting nesting depth and comma-separated arguments.
   */
  private static final class FunctionFrame {
    private boolean hasContent;
    private int argumentSeparators;
  }

  private record FormulaShape(int maximumFunctionNesting, int maximumFunctionArguments) {}
}
