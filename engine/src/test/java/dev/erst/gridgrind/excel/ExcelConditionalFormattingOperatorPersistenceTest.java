package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import java.util.zip.ZipFile;
import javax.xml.parsers.DocumentBuilderFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Guards the conditional-formatting operator bridge independently from data validation. */
class ExcelConditionalFormattingOperatorPersistenceTest {
  @TempDir Path temporaryDirectory;

  @Test
  void persistsEveryConditionalFormattingComparisonOperatorWithItsOwnPoiBridge() throws Exception {
    Path workbookPath = temporaryDirectory.resolve("conditional-formatting-operators.xlsx");
    List<ExcelComparisonOperator> operators = List.of(ExcelComparisonOperator.values());
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Rules");
      List<ExcelConditionalFormattingRule> rules =
          operators.stream()
              .<ExcelConditionalFormattingRule>map(
                  operator ->
                      new ExcelConditionalFormattingRule.CellValueRule(
                          operator,
                          "10",
                          switch (operator) {
                            case BETWEEN, NOT_BETWEEN -> Optional.of("20");
                            default -> Optional.empty();
                          },
                          false,
                          Optional.empty()))
              .toList();
      new ExcelConditionalFormattingController()
          .setConditionalFormatting(
              sheet, new ExcelConditionalFormattingBlockDefinition(List.of("A1:A8"), rules));
      try (var output = Files.newOutputStream(workbookPath)) {
        workbook.write(output);
      }
    }

    try (ZipFile archive = new ZipFile(workbookPath.toFile())) {
      var worksheet = archive.getEntry("xl/worksheets/sheet1.xml");
      org.junit.jupiter.api.Assertions.assertNotNull(worksheet);
      DocumentBuilderFactory factory = DocumentBuilderFactory.newDefaultInstance();
      factory.setNamespaceAware(true);
      var document = factory.newDocumentBuilder().parse(archive.getInputStream(worksheet));
      var rules = document.getElementsByTagNameNS("*", "cfRule");
      assertEquals(operators.size(), rules.getLength());
      for (int index = 0; index < operators.size(); index++) {
        assertEquals(
            xmlOperator(operators.get(index)),
            ((org.w3c.dom.Element) rules.item(index)).getAttribute("operator"));
      }
    }
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
}
