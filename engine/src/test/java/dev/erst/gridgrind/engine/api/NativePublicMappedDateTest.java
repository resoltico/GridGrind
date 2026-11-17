package dev.erst.gridgrind.engine.api;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.grant.*;
import dev.erst.gridgrind.core.model.customxml.*;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.value.*;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.problem.*;
import dev.erst.gridgrind.core.result.*;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.*;
import java.nio.file.*;
import java.util.*;
import java.util.zip.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/**
 * Real public import observes the current workbook calendar and locates mapped XML date refusals.
 */
@Tag("native-public")
class NativePublicMappedDateTest {
  @TempDir Path root;

  @ParameterizedTest
  @CsvSource({
    "false,1903-12-31,true,1461",
    "true,1903-12-31,false,0",
    "false,2026-03-08Z,false,0",
    "true,2026-02-30,false,0",
    "false,2026-03-08,true,46089",
    "true,2026-03-08,true,44627"
  })
  void mappedDateUsesActualCalendarOrRefusesWithKnownXPath(
      boolean mac, String date, boolean valid, double expected) throws Exception {
    Path source = fixture(mac);
    String sheetName;
    try (var workbook = new XSSFWorkbook(Files.newInputStream(source))) {
      sheetName = workbook.getSheetAt(0).getSheetName();
    }
    var steps =
        List.<WorkbookStep>of(
            new MutationStep(
                new StepId("import"),
                new WorkbookSelector.Current(),
                new StructuredMutationAction.ImportCustomXmlMapping(
                    new CustomXmlImportSpec(
                        new CustomXmlMappingLocator(Optional.of(1L), Optional.of("CORSO_mapping")),
                        TextSourceSpec.inline(
                            "<CORSO><NOME>new ordinary field</NOME><DOCENTE>"
                                + date
                                + "</DOCENTE></CORSO>")))),
            new InspectionStep(
                new StepId("date"),
                new CellSelector.ByAddress(new SheetName(sheetName), new CellAddress("B1")),
                new SheetIntrospectionQuery.GetCells()));
    var request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile(source.toString()),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            steps);
    var grant =
        new ExecutionGrant(
            1024 * 1024,
            List.of(new ExecutionGrant.ReadAuthority.File(source)),
            List.of("IMPORT_CUSTOM_XML_MAPPING", "GET_CELLS"),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.None(),
            List.of(),
            AcceptancePolicy.minimum());
    var result =
        GridGrindEngine.requestExecutor()
            .execute(request, new GridGrindRequestInputs(root, root.resolve("scratch"), grant));
    if (valid) {
      var success = assertInstanceOf(WorkbookResult.Success.class, result, result.toString());
      var cells =
          assertInstanceOf(
              SheetInspectionResult.CellsResult.class, success.inspections().getFirst());
      assertEquals(
          new CellValue.NumberValue(expected), cells.cells().getFirst().value().orElseThrow());
    } else {
      var failure = assertInstanceOf(WorkbookResult.Failure.class, result, result.toString());
      assertEquals(
          GridGrindProblemCode.INVALID_REQUEST,
          failure.problem().code(),
          failure.problem().toString());
      var context = assertInstanceOf(ProblemContext.ExecuteStep.class, failure.problem().context());
      var location =
          assertInstanceOf(
              ProblemContextWorkbookSurfaces.ProblemLocation.MappedXmlField.class,
              context.location());
      assertEquals("/CORSO/DOCENTE", location.xpath());
      assertEquals(0, location.mappingRowIndex());
    }
    Path scratch = root.resolve("scratch");
    if (Files.exists(scratch))
      try (var files = Files.walk(scratch)) {
        assertTrue(files.noneMatch(Files::isRegularFile));
      }
  }

  private Path fixture(boolean mac) throws java.io.IOException {
    Path output = root.resolve("source.xlsx");
    try (var input =
            Objects.requireNonNull(
                getClass()
                    .getResourceAsStream(
                        "/dev/erst/gridgrind/engine/backend/xssf/CustomXMLMappings.xlsx"));
        var zip = new ZipInputStream(input);
        var target = new ZipOutputStream(Files.newOutputStream(output))) {
      ZipEntry entry = zip.getNextEntry();
      while (entry != null) {
        copyFixtureEntry(zip, target, entry, mac);
        entry = zip.getNextEntry();
      }
    }
    return output;
  }

  private static void copyFixtureEntry(
      ZipInputStream zip, ZipOutputStream target, ZipEntry entry, boolean mac)
      throws java.io.IOException {
    byte[] bytes = zip.readAllBytes();
    String name = entry.getName();
    if ("xl/workbook.xml".equals(name)) {
      String xml = new String(bytes, java.nio.charset.StandardCharsets.UTF_8);
      if (xml.contains("date1904="))
        xml = xml.replaceAll("date1904=\"[^\"]*\"", "date1904=\"" + (mac ? "1" : "0") + "\"");
      else if (xml.contains("<workbookPr/>"))
        xml = xml.replace("<workbookPr/>", "<workbookPr date1904=\"" + (mac ? "1" : "0") + "\"/>");
      else xml = xml.replace("<workbookPr ", "<workbookPr date1904=\"" + (mac ? "1" : "0") + "\" ");
      assertTrue(xml.contains("date1904="));
      bytes = xml.getBytes(java.nio.charset.StandardCharsets.UTF_8);
    }
    if ("xl/tables/tableSingleCells1.xml".equals(name)) {
      String xml = new String(bytes, java.nio.charset.StandardCharsets.UTF_8);
      assertTrue(xml.contains("xpath=\"/CORSO/DOCENTE\" xmlDataType=\"string\""));
      bytes =
          xml.replace(
                  "xpath=\"/CORSO/DOCENTE\" xmlDataType=\"string\"",
                  "xpath=\"/CORSO/DOCENTE\" xmlDataType=\"date\"")
              .getBytes(java.nio.charset.StandardCharsets.UTF_8);
    }
    if ("xl/xmlMaps.xml".equals(name))
      bytes =
          new String(bytes, java.nio.charset.StandardCharsets.UTF_8)
              .replace("name=\"DOCENTE\" type=\"xsd:string\"", "name=\"DOCENTE\" type=\"xsd:date\"")
              .getBytes(java.nio.charset.StandardCharsets.UTF_8);
    target.putNextEntry(new ZipEntry(name));
    target.write(bytes);
    target.closeEntry();
  }
}
