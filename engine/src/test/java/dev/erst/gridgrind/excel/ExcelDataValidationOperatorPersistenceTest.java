package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationController;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import java.util.zip.ZipFile;
import javax.xml.parsers.DocumentBuilderFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies persisted OOXML comparison operators for every modeled validation family. */
class ExcelDataValidationOperatorPersistenceTest {
  @TempDir Path temporaryDirectory;

  @Test
  void persistsAndReadsEveryComparisonOperatorAcrossEveryComparisonFamily() throws IOException {
    Path workbookPath = temporaryDirectory.resolve("validation-operators.xlsx");
    List<ExcelComparisonOperator> operators = List.of(ExcelComparisonOperator.values());

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Validations");
      ExcelDataValidationController controller = new ExcelDataValidationController();
      int validationIndex = 0;
      for (Family family : Family.values()) {
        for (ExcelComparisonOperator operator : operators) {
          String column = CellReference.convertNumToColString(validationIndex);
          validationIndex++;
          controller.setDataValidation(
              sheet, column + "1:" + column + "3", definition(family, operator));
        }
      }
      try (var output = Files.newOutputStream(workbookPath)) {
        workbook.write(output);
      }
    }

    assertPersistedWorksheetOperators(workbookPath, operators);

    try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      XSSFSheet sheet = workbook.getSheet("Validations");
      ExcelDataValidationController controller = new ExcelDataValidationController();
      List<DataValidationEntryReport> snapshots = controller.dataValidations(sheet);
      assertEquals(Family.values().length * operators.size(), snapshots.size());

      int validationIndex = 0;
      for (Family family : Family.values()) {
        for (ExcelComparisonOperator operator : operators) {
          var persisted =
              sheet.getCTWorksheet().getDataValidations().getDataValidationArray(validationIndex);
          assertEquals(
              xmlOperator(operator),
              persisted.getOperator().toString(),
              family + " " + operator + " must persist its exact OOXML operator");

          DataValidationEntryReport.Supported snapshot =
              assertInstanceOf(
                  DataValidationEntryReport.Supported.class, snapshots.get(validationIndex));
          assertEquals(
              operator,
              operatorOf(snapshot.validation().rule()),
              family + " " + operator + " must read back without inversion");
          assertEquals(
              operator == ExcelComparisonOperator.BETWEEN
                  || operator == ExcelComparisonOperator.NOT_BETWEEN,
              formula2Of(snapshot.validation().rule()).isPresent(),
              family + " " + operator + " must preserve comparison-formula arity");
          validationIndex++;
        }
      }
    }
  }

  private static void assertPersistedWorksheetOperators(
      Path workbookPath, List<ExcelComparisonOperator> operators) throws IOException {
    try (ZipFile archive = new ZipFile(workbookPath.toFile())) {
      var worksheet = archive.getEntry("xl/worksheets/sheet1.xml");
      org.junit.jupiter.api.Assertions.assertNotNull(worksheet);
      DocumentBuilderFactory factory = DocumentBuilderFactory.newDefaultInstance();
      factory.setNamespaceAware(true);
      var document = factory.newDocumentBuilder().parse(archive.getInputStream(worksheet));
      var validations = document.getElementsByTagNameNS("*", "dataValidation");
      assertEquals(Family.values().length * operators.size(), validations.getLength());
      int validationIndex = 0;
      for (Family family : Family.values()) {
        for (ExcelComparisonOperator operator : operators) {
          assertEquals(
              xmlOperator(operator),
              ((org.w3c.dom.Element) validations.item(validationIndex)).getAttribute("operator"),
              family + " " + operator + " must persist its exact worksheet XML attribute");
          validationIndex++;
        }
      }
    } catch (javax.xml.parsers.ParserConfigurationException | org.xml.sax.SAXException exception) {
      throw new IOException("Could not parse persisted worksheet XML", exception);
    }
  }

  private static ResolvedDataValidation definition(
      Family family, ExcelComparisonOperator operator) {
    Optional<String> formula2 =
        switch (operator) {
          case BETWEEN, NOT_BETWEEN -> Optional.of(family.secondFormula());
          default -> Optional.empty();
        };
    DataValidationRuleSpec rule =
        switch (family) {
          case WHOLE_NUMBER ->
              new DataValidationRuleSpec.WholeNumber(
                  operator,
                  ExcelComparisonFormulaSupport.normalizeFormula("10", "formula1"),
                  ExcelComparisonFormulaSupport.normalizeOptionalFormula(formula2, "formula2"));
          case DECIMAL ->
              new DataValidationRuleSpec.DecimalNumber(
                  operator,
                  ExcelComparisonFormulaSupport.normalizeFormula("10.5", "formula1"),
                  ExcelComparisonFormulaSupport.normalizeOptionalFormula(formula2, "formula2"));
          case DATE ->
              new DataValidationRuleSpec.DateRule(
                  operator,
                  ExcelComparisonFormulaSupport.normalizeFormula("DATE(2026,1,1)", "formula1"),
                  ExcelComparisonFormulaSupport.normalizeOptionalFormula(formula2, "formula2"));
          case TIME ->
              new DataValidationRuleSpec.TimeRule(
                  operator,
                  ExcelComparisonFormulaSupport.normalizeFormula("TIME(9,0,0)", "formula1"),
                  ExcelComparisonFormulaSupport.normalizeOptionalFormula(formula2, "formula2"));
          case TEXT_LENGTH ->
              new DataValidationRuleSpec.TextLength(
                  operator,
                  ExcelComparisonFormulaSupport.normalizeFormula("5", "formula1"),
                  ExcelComparisonFormulaSupport.normalizeOptionalFormula(formula2, "formula2"));
        };
    return new ResolvedDataValidation(rule, false, false, Optional.empty(), Optional.empty());
  }

  private static ExcelComparisonOperator operatorOf(DataValidationRuleSpec rule) {
    return switch (rule) {
      case DataValidationRuleSpec.WholeNumber wholeNumber -> wholeNumber.operator();
      case DataValidationRuleSpec.DecimalNumber decimalNumber -> decimalNumber.operator();
      case DataValidationRuleSpec.DateRule dateRule -> dateRule.operator();
      case DataValidationRuleSpec.TimeRule timeRule -> timeRule.operator();
      case DataValidationRuleSpec.TextLength textLength -> textLength.operator();
      case DataValidationRuleSpec.ExplicitList _,
          DataValidationRuleSpec.FormulaList _,
          DataValidationRuleSpec.CustomFormula _ ->
          throw new AssertionError("Expected one comparison validation rule");
    };
  }

  private static Optional<String> formula2Of(DataValidationRuleSpec rule) {
    return switch (rule) {
      case DataValidationRuleSpec.WholeNumber wholeNumber -> wholeNumber.formula2();
      case DataValidationRuleSpec.DecimalNumber decimalNumber -> decimalNumber.formula2();
      case DataValidationRuleSpec.DateRule dateRule -> dateRule.formula2();
      case DataValidationRuleSpec.TimeRule timeRule -> timeRule.formula2();
      case DataValidationRuleSpec.TextLength textLength -> textLength.formula2();
      case DataValidationRuleSpec.ExplicitList _,
          DataValidationRuleSpec.FormulaList _,
          DataValidationRuleSpec.CustomFormula _ ->
          throw new AssertionError("Expected one comparison validation rule");
    };
  }

  private static String xmlOperator(ExcelComparisonOperator operator) {
    return switch (operator) {
      case BETWEEN -> "between";
      case NOT_BETWEEN -> "notBetween";
      case EQUAL -> "equal";
      case NOT_EQUAL -> "notEqual";
      case GREATER_THAN -> "greaterThan";
      case LESS_THAN -> "lessThan";
      case GREATER_OR_EQUAL -> "greaterThanOrEqual";
      case LESS_OR_EQUAL -> "lessThanOrEqual";
    };
  }

  /** Comparison validation families that share the persisted operator matrix. */
  private enum Family {
    /** Whole-number constraints. */
    WHOLE_NUMBER("20"),
    /** Decimal-number constraints. */
    DECIMAL("20.5"),
    /** Date constraints. */
    DATE("DATE(2026,12,31)"),
    /** Time constraints. */
    TIME("TIME(17,0,0)"),
    /** Text-length constraints. */
    TEXT_LENGTH("10");

    private final String secondFormula;

    Family(String secondFormula) {
      this.secondFormula = secondFormula;
    }

    String secondFormula() {
      return secondFormula;
    }
  }
}
