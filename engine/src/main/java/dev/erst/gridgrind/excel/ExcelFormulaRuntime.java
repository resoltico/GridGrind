package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;
import java.util.regex.Pattern;
import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.AreaEval;
import org.apache.poi.ss.formula.eval.BlankEval;
import org.apache.poi.ss.formula.eval.BoolEval;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.MissingArgEval;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.RefEval;
import org.apache.poi.ss.formula.eval.StringEval;
import org.apache.poi.ss.formula.eval.ValueEval;
import org.apache.poi.ss.formula.functions.FreeRefFunction;
import org.apache.poi.ss.formula.udf.DefaultUDFFinder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Abstracts formula evaluation, caching, and environment setup behind a GridGrind-owned seam. */
interface ExcelFormulaRuntime extends AutoCloseable {
  Pattern ARGUMENT_PLACEHOLDER_PATTERN =
      Pattern.compile("\\bARG([1-9][0-9]*)\\b", Pattern.CASE_INSENSITIVE);
  String SCRATCH_ARGUMENT_NAME_PREFIX = "_GRIDGRIND_ARG_";

  /** Evaluates one formula cell and returns the computed value, or null when POI reports none. */
  CellValue evaluate(Cell cell);

  /** Evaluates one formula cell and stores its cached result while leaving the cell as FORMULA. */
  CellType evaluateFormulaCell(Cell cell);

  /** Clears all cached evaluator results owned by this workbook runtime. */
  void clearCachedResults();

  /** Formats one cell exactly as Excel would display it, evaluating formulas when needed. */
  String displayValue(DataFormatter formatter, Cell cell);

  /** Returns the evaluator configuration used for analysis and diagnostics. */
  ExcelFormulaRuntimeContext context();

  @Override
  default void close() throws IOException {
    // The default runtime owns no additional closeable resources.
  }

  /** Creates the default runtime backed only by the supplied Apache POI evaluator. */
  static ExcelFormulaRuntime poi(FormulaEvaluator formulaEvaluator) {
    return new PoiExcelFormulaRuntime(
        formulaEvaluator, ExcelFormulaEnvironment.defaults().runtimeContext(), List.of());
  }

  /** Creates the production runtime backed by Apache POI plus GridGrind formula environment. */
  @SuppressWarnings("PMD.CloseResource")
  static ExcelFormulaRuntime poi(XSSFWorkbook workbook, ExcelFormulaEnvironment environment)
      throws IOException {
    Objects.requireNonNull(workbook, "workbook must not be null");
    ExcelFormulaEnvironment configuredEnvironment =
        environment == null ? ExcelFormulaEnvironment.defaults() : environment;
    validateUdfTemplates(configuredEnvironment.udfToolpacks());
    registerUdfToolpacks(workbook, configuredEnvironment.udfToolpacks());

    List<ReferencedWorkbookHandle> referencedWorkbooks = new ArrayList<>();
    Map<String, FormulaEvaluator> referencedEvaluators = new ConcurrentHashMap<>();
    if (!configuredEnvironment.externalWorkbooks().isEmpty()) {
      try {
        for (ExcelFormulaExternalWorkbookBinding binding :
            configuredEnvironment.externalWorkbooks()) {
          ReferencedWorkbookHandle referencedWorkbook =
              openReferencedWorkbook(binding, configuredEnvironment);
          referencedWorkbooks.add(referencedWorkbook);
          workbook.linkExternalWorkbook(binding.workbookName(), referencedWorkbook.workbook());
        }
      } catch (IOException | RuntimeException exception) {
        closeReferencedWorkbooks(referencedWorkbooks);
        throw exception;
      }
    }

    FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
    formulaEvaluator.setIgnoreMissingWorkbooks(
        configuredEnvironment.missingWorkbookPolicy()
            == ExcelFormulaMissingWorkbookPolicy.USE_CACHED_VALUE);

    if (!referencedWorkbooks.isEmpty()) {
      for (ReferencedWorkbookHandle referencedWorkbook : referencedWorkbooks) {
        referencedEvaluators.put(
            referencedWorkbook.workbookName(), referencedWorkbook.formulaEvaluator());
      }
      formulaEvaluator.setupReferencedWorkbooks(referencedEvaluators);
    }

    return new PoiExcelFormulaRuntime(
        formulaEvaluator,
        configuredEnvironment.runtimeContext(),
        referencedWorkbooks.stream().map(ReferencedWorkbookHandle::workbook).toList());
  }

  private static void validateUdfTemplates(List<ExcelFormulaUdfToolpack> udfToolpacks) {
    for (ExcelFormulaUdfToolpack toolpack : udfToolpacks) {
      for (ExcelFormulaUdfFunction function : toolpack.functions()) {
        validateUdfTemplate(function);
      }
    }
  }

  private static String scratchArgumentName(int oneBasedIndex) {
    return SCRATCH_ARGUMENT_NAME_PREFIX + oneBasedIndex;
  }

  private static String toScratchFormula(String formulaTemplate) {
    java.util.regex.Matcher matcher = ARGUMENT_PLACEHOLDER_PATTERN.matcher(formulaTemplate);
    StringBuffer translated = new StringBuffer();
    while (matcher.find()) {
      matcher.appendReplacement(
          translated, scratchArgumentName(Integer.parseInt(matcher.group(1))));
    }
    matcher.appendTail(translated);
    return translated.toString();
  }

  private static void validateUdfTemplate(ExcelFormulaUdfFunction function) {
    try (XSSFWorkbook scratchWorkbook = new XSSFWorkbook()) {
      org.apache.poi.xssf.usermodel.XSSFSheet argsSheet = scratchWorkbook.createSheet("Args");
      org.apache.poi.xssf.usermodel.XSSFSheet calcSheet = scratchWorkbook.createSheet("Calc");
      int placeholderCount = function.maximumArgumentCount();
      for (int index = 0; index < placeholderCount; index++) {
        defineScratchArgument(scratchWorkbook, argsSheet, index);
      }
      calcSheet
          .createRow(0)
          .createCell(0)
          .setCellFormula(toScratchFormula(function.formulaTemplate()));
    } catch (IOException | RuntimeException exception) {
      throw new IllegalArgumentException(
          "Invalid formulaTemplate for user-defined function "
              + function.name()
              + ": "
              + function.formulaTemplate(),
          exception);
    }
  }

  private static void defineScratchArgument(
      XSSFWorkbook scratchWorkbook,
      org.apache.poi.xssf.usermodel.XSSFSheet argsSheet,
      int zeroBasedIndex) {
    org.apache.poi.ss.usermodel.Row row = argsSheet.getRow(zeroBasedIndex);
    if (row == null) {
      row = argsSheet.createRow(zeroBasedIndex);
    }
    row.createCell(0).setCellValue(zeroBasedIndex + 1.0d);
    org.apache.poi.ss.usermodel.Name name = scratchWorkbook.createName();
    name.setNameName(scratchArgumentName(zeroBasedIndex + 1));
    name.setRefersToFormula("Args!$A$" + (zeroBasedIndex + 1));
  }

  private static void registerUdfToolpacks(
      Workbook workbook, List<ExcelFormulaUdfToolpack> udfToolpacks) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    if (udfToolpacks.isEmpty()) {
      return;
    }
    List<String> names = new ArrayList<>();
    List<FreeRefFunction> functions = new ArrayList<>();
    for (ExcelFormulaUdfToolpack toolpack : udfToolpacks) {
      for (ExcelFormulaUdfFunction function : toolpack.functions()) {
        names.add(function.name());
        functions.add(new TemplateFormulaUdfFunction(function));
      }
    }
    workbook.addToolPack(
        new DefaultUDFFinder(
            names.toArray(String[]::new), functions.toArray(FreeRefFunction[]::new)));
  }

  private static ReferencedWorkbookHandle openReferencedWorkbook(
      ExcelFormulaExternalWorkbookBinding binding, ExcelFormulaEnvironment environment)
      throws IOException {
    Workbook referencedWorkbook =
        org.apache.poi.ss.usermodel.WorkbookFactory.create(binding.path().toFile());
    registerUdfToolpacks(referencedWorkbook, environment.udfToolpacks());
    FormulaEvaluator referencedEvaluator =
        referencedWorkbook.getCreationHelper().createFormulaEvaluator();
    referencedEvaluator.setIgnoreMissingWorkbooks(
        environment.missingWorkbookPolicy() == ExcelFormulaMissingWorkbookPolicy.USE_CACHED_VALUE);
    return new ReferencedWorkbookHandle(
        binding.workbookName(), referencedWorkbook, referencedEvaluator);
  }

  @SuppressWarnings("PMD.CloseResource")
  private static void closeReferencedWorkbooks(List<ReferencedWorkbookHandle> referencedWorkbooks)
      throws IOException {
    IOException failure = null;
    for (ReferencedWorkbookHandle referencedWorkbook : referencedWorkbooks) {
      try {
        referencedWorkbook.close();
      } catch (IOException exception) {
        if (failure == null) {
          failure = exception;
        } else {
          failure.addSuppressed(exception);
        }
      }
    }
    if (failure != null) {
      throw failure;
    }
  }

  @SuppressWarnings("PMD.CloseResource")
  private static void closeAll(List<Workbook> referencedWorkbooks) throws IOException {
    IOException failure = null;
    for (Workbook referencedWorkbook : referencedWorkbooks) {
      try {
        referencedWorkbook.close();
      } catch (IOException exception) {
        if (failure == null) {
          failure = exception;
        } else {
          failure.addSuppressed(exception);
        }
      }
    }
    if (failure != null) {
      throw failure;
    }
  }

  /** Owns one externally linked workbook plus the evaluator configured for that link. */
  record ReferencedWorkbookHandle(
      String workbookName, Workbook workbook, FormulaEvaluator formulaEvaluator)
      implements AutoCloseable {
    public ReferencedWorkbookHandle {
      Objects.requireNonNull(workbookName, "workbookName must not be null");
      Objects.requireNonNull(workbook, "workbook must not be null");
      Objects.requireNonNull(formulaEvaluator, "formulaEvaluator must not be null");
    }

    @Override
    public void close() throws IOException {
      workbook.close();
    }
  }

  /** Production runtime that delegates evaluation and display formatting to Apache POI. */
  record PoiExcelFormulaRuntime(
      FormulaEvaluator formulaEvaluator,
      ExcelFormulaRuntimeContext context,
      List<Workbook> referencedWorkbooks)
      implements ExcelFormulaRuntime {
    public PoiExcelFormulaRuntime {
      Objects.requireNonNull(formulaEvaluator, "formulaEvaluator must not be null");
      Objects.requireNonNull(context, "context must not be null");
      referencedWorkbooks = List.copyOf(referencedWorkbooks);
    }

    @Override
    public CellValue evaluate(Cell cell) {
      Objects.requireNonNull(cell, "cell must not be null");
      formulaEvaluator.clearAllCachedResultValues();
      return formulaEvaluator.evaluate(cell);
    }

    @Override
    public CellType evaluateFormulaCell(Cell cell) {
      Objects.requireNonNull(cell, "cell must not be null");
      return formulaEvaluator.evaluateFormulaCell(cell);
    }

    @Override
    public void clearCachedResults() {
      formulaEvaluator.clearAllCachedResultValues();
    }

    @Override
    public String displayValue(DataFormatter formatter, Cell cell) {
      Objects.requireNonNull(formatter, "formatter must not be null");
      Objects.requireNonNull(cell, "cell must not be null");
      formulaEvaluator.clearAllCachedResultValues();
      return formatter.formatCellValue(cell, formulaEvaluator);
    }

    @Override
    public void close() throws IOException {
      closeAll(referencedWorkbooks);
    }
  }

  /** Safe template-backed UDF adapter built on top of a tiny scratch workbook. */
  final class TemplateFormulaUdfFunction implements FreeRefFunction {
    private final ExcelFormulaUdfFunction function;

    TemplateFormulaUdfFunction(ExcelFormulaUdfFunction function) {
      this.function = Objects.requireNonNull(function, "function must not be null");
    }

    @Override
    public ValueEval evaluate(ValueEval[] args, OperationEvaluationContext ec) {
      if (args.length < function.minimumArgumentCount()
          || args.length > function.maximumArgumentCount()) {
        return ErrorEval.VALUE_INVALID;
      }
      try (XSSFWorkbook scratchWorkbook = new XSSFWorkbook()) {
        org.apache.poi.xssf.usermodel.XSSFSheet argsSheet = scratchWorkbook.createSheet("Args");
        org.apache.poi.xssf.usermodel.XSSFSheet calcSheet = scratchWorkbook.createSheet("Calc");
        int nextStartRow = 0;
        for (int index = 0; index < args.length; index++) {
          org.apache.poi.ss.util.CellRangeAddress address =
              writeArgument(argsSheet, nextStartRow, args[index]);
          org.apache.poi.ss.usermodel.Name name = scratchWorkbook.createName();
          name.setNameName(scratchArgumentName(index + 1));
          name.setRefersToFormula("Args!" + absoluteAreaReference(address));
          nextStartRow = address.getLastRow() + 2;
        }
        Cell formulaCell = calcSheet.createRow(0).createCell(0);
        formulaCell.setCellFormula(toScratchFormula(function.formulaTemplate()));
        return toValueEval(
            scratchWorkbook.getCreationHelper().createFormulaEvaluator().evaluate(formulaCell));
      } catch (RuntimeException | IOException exception) {
        return ErrorEval.VALUE_INVALID;
      }
    }

    private static org.apache.poi.ss.util.CellRangeAddress writeArgument(
        org.apache.poi.xssf.usermodel.XSSFSheet argsSheet, int startRow, ValueEval argument) {
      if (argument instanceof RefEval refEval) {
        if (refEval.getNumberOfSheets() != 1) {
          throw new IllegalArgumentException(
              "3D references are not supported in template-backed UDF arguments");
        }
        writeScalarValue(
            cell(argsSheet, startRow, 0), refEval.getInnerValueEval(refEval.getFirstSheetIndex()));
        return new org.apache.poi.ss.util.CellRangeAddress(startRow, startRow, 0, 0);
      }
      if (argument instanceof AreaEval areaEval) {
        for (int rowOffset = 0; rowOffset < areaEval.getHeight(); rowOffset++) {
          for (int columnOffset = 0; columnOffset < areaEval.getWidth(); columnOffset++) {
            writeScalarValue(
                cell(argsSheet, startRow + rowOffset, columnOffset),
                areaEval.getRelativeValue(rowOffset, columnOffset));
          }
        }
        return new org.apache.poi.ss.util.CellRangeAddress(
            startRow, startRow + areaEval.getHeight() - 1, 0, areaEval.getWidth() - 1);
      }
      writeScalarValue(cell(argsSheet, startRow, 0), argument);
      return new org.apache.poi.ss.util.CellRangeAddress(startRow, startRow, 0, 0);
    }

    private static Cell cell(
        org.apache.poi.xssf.usermodel.XSSFSheet sheet, int rowIndex, int columnIndex) {
      org.apache.poi.ss.usermodel.Row row = sheet.getRow(rowIndex);
      if (row == null) {
        row = sheet.createRow(rowIndex);
      }
      Cell cell = row.getCell(columnIndex);
      return cell == null ? row.createCell(columnIndex) : cell;
    }

    private static void writeScalarValue(Cell cell, ValueEval valueEval) {
      switch (valueEval) {
        case BlankEval _ -> cell.setBlank();
        case MissingArgEval _ -> cell.setBlank();
        case NumberEval numberEval -> cell.setCellValue(numberEval.getNumberValue());
        case StringEval stringEval -> cell.setCellValue(stringEval.getStringValue());
        case BoolEval boolEval -> cell.setCellValue(boolEval.getBooleanValue());
        case ErrorEval errorEval -> cell.setCellErrorValue((byte) errorEval.getErrorCode());
        case RefEval _ ->
            throw new IllegalArgumentException("Nested RefEval values are not supported");
        case AreaEval _ ->
            throw new IllegalArgumentException("Nested AreaEval values are not supported");
        case null, default ->
            throw new IllegalArgumentException("Unsupported ValueEval: " + valueEval);
      }
    }

    private static String absoluteAreaReference(org.apache.poi.ss.util.CellRangeAddress address) {
      return absoluteCellReference(address.getFirstRow(), address.getFirstColumn())
          + ":"
          + absoluteCellReference(address.getLastRow(), address.getLastColumn());
    }

    private static String absoluteCellReference(int rowIndex, int columnIndex) {
      return new org.apache.poi.ss.util.CellReference(rowIndex, columnIndex).formatAsString(true);
    }

    private static ValueEval toValueEval(CellValue value) {
      if (value == null) {
        return BlankEval.instance;
      }
      return switch (value.getCellType()) {
        case NUMERIC -> new NumberEval(value.getNumberValue());
        case STRING -> new StringEval(value.getStringValue());
        case BOOLEAN -> BoolEval.valueOf(value.getBooleanValue());
        case ERROR -> ErrorEval.valueOf(value.getErrorValue());
        case BLANK, _NONE -> BlankEval.instance;
        case FORMULA -> throw new IllegalStateException("Scratch evaluator returned FORMULA");
      };
    }
  }
}
