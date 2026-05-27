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

  static FormulaShape scanFormulaShape(String formula) {
    return new FormulaShapeScanner(formula).scan();
  }

  static boolean looksLikeFunctionCall(String formula, int openParenIndex) {
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

  static boolean isFunctionIdentifierCharacter(char value) {
    return Character.isLetterOrDigit(value) || value == '_' || value == '.';
  }

  /**
   * Tracks the current function scope while counting nesting depth and comma-separated arguments.
   */
  private static final class FunctionFrame {
    private boolean hasContent;
    private int argumentSeparators;
  }

  /** Stateful scanner for authored formula nesting and argument-count limits. */
  private static final class FormulaShapeScanner {
    private final String formula;
    private final Deque<FunctionFrame> functions = new ArrayDeque<>();
    private boolean inString;
    private int maximumFunctionNesting;
    private int maximumFunctionArguments;
    private int index;

    private FormulaShapeScanner(String formula) {
      this.formula = formula;
    }

    private FormulaShape scan() {
      while (index < formula.length()) {
        if (consumeQuotedText()) {
          continue;
        }
        if (consumeStringBody()) {
          continue;
        }
        if (consumeFunctionOpen()) {
          continue;
        }
        if (consumeArgumentSeparator()) {
          continue;
        }
        if (consumeFunctionClose()) {
          continue;
        }
        markCurrentFunctionContent();
        index++;
      }
      return new FormulaShape(maximumFunctionNesting, maximumFunctionArguments);
    }

    private boolean consumeQuotedText() {
      if (currentChar() != '"') {
        return false;
      }
      if (inString && nextCharIs('"')) {
        index += 2;
      } else {
        inString = !inString;
        index++;
      }
      return true;
    }

    private boolean consumeStringBody() {
      if (!inString) {
        return false;
      }
      index++;
      return true;
    }

    private boolean consumeFunctionOpen() {
      if (currentChar() != '(' || !looksLikeFunctionCall(formula, index)) {
        return false;
      }
      functions.addLast(new FunctionFrame());
      maximumFunctionNesting = Math.max(maximumFunctionNesting, functions.size());
      index++;
      return true;
    }

    private boolean consumeArgumentSeparator() {
      if (currentChar() != ',' || functions.isEmpty()) {
        return false;
      }
      functions.getLast().argumentSeparators++;
      functions.getLast().hasContent = true;
      index++;
      return true;
    }

    private boolean consumeFunctionClose() {
      if (currentChar() != ')' || functions.isEmpty()) {
        return false;
      }
      FunctionFrame completed = functions.removeLast();
      int argumentCount = completed.hasContent ? completed.argumentSeparators + 1 : 0;
      maximumFunctionArguments = Math.max(maximumFunctionArguments, argumentCount);
      index++;
      return true;
    }

    private void markCurrentFunctionContent() {
      if (!functions.isEmpty() && !Character.isWhitespace(currentChar())) {
        functions.getLast().hasContent = true;
      }
    }

    private char currentChar() {
      return formula.charAt(index);
    }

    private boolean nextCharIs(char value) {
      return index + 1 < formula.length() && formula.charAt(index + 1) == value;
    }
  }

  record FormulaShape(int maximumFunctionNesting, int maximumFunctionArguments) {}
}
