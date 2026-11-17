package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.source.SecretReference;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import dev.erst.gridgrind.json.codec.GridGrindJsonOutput;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies native observations and their response JSON contract. */
class ExampleExecutionFixturesTest {
  @TempDir Path tempDir;

  @Test
  void storedNonfiniteNumericObservationsRemainVisible() throws IOException {
    Path source = tempDir.resolve("observed-number.xlsx");
    try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
      var row = workbook.createSheet("Data").createRow(0);
      var observed = row.createCell(0);
      observed.setCellValue(1.0);
      observed.getCTCell().setV("NaN");
      row.createCell(1).setCellValue(1.0);
      var formula = row.createCell(2);
      formula.setCellFormula("A1");
      formula.getCTCell().setT(org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellType.N);
      formula.getCTCell().setV("NaN");
      try (OutputStream output = Files.newOutputStream(source)) {
        workbook.write(output);
      }
    }
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
            {"protocolVersion":1,"source":{"type":"EXISTING","path":"observed-number.xlsx"},
             "persistence":{"type":"NONE"},"steps":[{"stepId":"numbers",
             "target":{"type":"CELL_BY_ADDRESSES","sheetName":"Data","addresses":["A1","B1","C1"]},
             "query":{"type":"GET_CELLS"}}]}
            """);
    var executed =
        new DefaultGridGrindRequestExecutor(
                NativeInvocationContext.enterOwned(
                    dev.erst.gridgrind.core.semantics.ExecutionConventions.STANDARD))
            .execute(request, bindingsForExample(tempDir, request));
    var result =
        assertInstanceOf(WorkbookResult.Success.class, executed, () -> executed.toString());
    var cells =
        assertInstanceOf(
                dev.erst.gridgrind.core.plan.SheetInspectionResult.CellsResult.class,
                result.inspections().getFirst())
            .cells();
    var observed =
        assertInstanceOf(
                dev.erst.gridgrind.core.model.value.CellReport.NumberReport.class, cells.getFirst())
            .value()
            .orElseThrow();
    assertTrue(Double.isNaN(observed.value()));
    assertEquals(
        1.0,
        assertInstanceOf(
                dev.erst.gridgrind.core.model.value.CellReport.NumberReport.class, cells.get(1))
            .value()
            .orElseThrow()
            .value());
    var formulaValue =
        assertInstanceOf(
                dev.erst.gridgrind.core.model.value.CellReport.FormulaReport.class, cells.get(2))
            .value()
            .orElseThrow();
    assertTrue(
        Double.isNaN(
            assertInstanceOf(
                    dev.erst.gridgrind.core.model.value.CellValue.NumberValue.class, formulaValue)
                .value()));
    assertEquals(
        result,
        GridGrindJson.readWorkbookResult(GridGrindJsonOutput.writeWorkbookResultBytes(result)));
  }

  @Test
  void observedTableLabelsMayBeEmptyWithoutBecomingAuthoredIdentifiers() throws IOException {
    Path source = tempDir.resolve("observed-table-name.xlsx");
    try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
      var sheet = workbook.createSheet("Data");
      sheet.createRow(0).createCell(0).setCellValue("Item");
      sheet.createRow(1).createCell(0).setCellValue("Observed");
      var table =
          sheet.createTable(
              new org.apache.poi.ss.util.AreaReference(
                  "A1:A2", org.apache.poi.ss.SpreadsheetVersion.EXCEL2007));
      table.setName("ValidTable");
      table.setDisplayName("ValidTable");
      table.getCTTable().setName("");
      try (OutputStream output = Files.newOutputStream(source)) {
        workbook.write(output);
      }
    }
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
            {"protocolVersion":1,"source":{"type":"EXISTING","path":"observed-table-name.xlsx"},
            "persistence":{"type":"NONE"},"steps":[{"stepId":"tables","target":{"type":"TABLE_ALL"},
            "query":{"type":"GET_TABLES"}}]}
            """);
    var result =
        assertInstanceOf(
            WorkbookResult.Success.class,
            new DefaultGridGrindRequestExecutor(
                    NativeInvocationContext.enterOwned(
                        dev.erst.gridgrind.core.semantics.ExecutionConventions.STANDARD))
                .execute(request, bindingsForExample(tempDir, request)));
    var tables =
        assertInstanceOf(
            dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.TablesResult.class,
            result.inspections().getFirst());
    assertEquals("", tables.tables().getFirst().name().value());
    assertEquals(
        result,
        GridGrindJson.readWorkbookResult(GridGrindJsonOutput.writeWorkbookResultBytes(result)));
  }

  @Test
  void observedDefinedNameLabelsDoNotAcquireAuthoringRestrictions() throws IOException {
    Path source = tempDir.resolve("observed-defined-name.xlsx");
    try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
      workbook.createSheet("Data").createRow(0).createCell(0).setCellValue("Observed");
      var name = workbook.createName();
      name.setNameName("ValidName");
      name.setRefersToFormula("Data!$A$1");
      try (OutputStream output = Files.newOutputStream(source)) {
        workbook.write(output);
      }
    }
    try (var archive = java.nio.file.FileSystems.newFileSystem(source)) {
      Path workbookXml = archive.getPath("/xl/workbook.xml");
      String xml = Files.readString(workbookXml);
      assertTrue(xml.contains("name=\"ValidName\""));
      Files.writeString(workbookXml, xml.replace("name=\"ValidName\"", "name=\"1invalid\""));
    }
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
            {"protocolVersion":1,"source":{"type":"EXISTING","path":"observed-defined-name.xlsx"},
            "persistence":{"type":"NONE"},"steps":[{"stepId":"names","target":{"type":"NAMED_RANGE_ALL"},
            "query":{"type":"GET_NAMED_RANGES"}}]}
            """);
    var result =
        assertInstanceOf(
            WorkbookResult.Success.class,
            new DefaultGridGrindRequestExecutor(
                    NativeInvocationContext.enterOwned(
                        dev.erst.gridgrind.core.semantics.ExecutionConventions.STANDARD))
                .execute(request, bindingsForExample(tempDir, request)));
    var names =
        assertInstanceOf(
            dev.erst.gridgrind.core.plan.WorkbookInspectionResult.NamedRangesResult.class,
            result.inspections().getFirst());
    var observed =
        assertInstanceOf(
            dev.erst.gridgrind.core.model.formula.NamedRangeReport.RangeReport.class,
            names.namedRanges().getFirst());
    assertEquals("1invalid", observed.name().value());
    assertEquals("Data!$A$1", observed.refersToFormula());
    assertEquals(
        result,
        GridGrindJson.readWorkbookResult(GridGrindJsonOutput.writeWorkbookResultBytes(result)));
  }

  @Test
  void observedMalformedValidationRangesRemainVisible() throws IOException {
    Path source = tempDir.resolve("observed-range.xlsx");
    try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
      var sheet = workbook.createSheet("Data");
      var validation = sheet.getCTWorksheet().addNewDataValidations().addNewDataValidation();
      var node = (org.w3c.dom.Element) validation.getDomNode();
      node.setAttribute("type", "future-validation");
      node.setAttribute("sqref", "#REF!");
      try (OutputStream output = Files.newOutputStream(source)) {
        workbook.write(output);
      }
    }
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
            {"protocolVersion":1,"source":{"type":"EXISTING","path":"observed-range.xlsx"},
            "persistence":{"type":"NONE"},"steps":[{"stepId":"ranges",
            "target":{"type":"RANGE_ALL_ON_SHEET","sheetName":"Data"},"query":{"type":"GET_DATA_VALIDATIONS"}}]}
            """);
    var result =
        assertInstanceOf(
            WorkbookResult.Success.class,
            new DefaultGridGrindRequestExecutor(
                    NativeInvocationContext.enterOwned(
                        dev.erst.gridgrind.core.semantics.ExecutionConventions.STANDARD))
                .execute(request, bindingsForExample(tempDir, request)));
    var validations =
        assertInstanceOf(
            dev.erst.gridgrind.core.plan.SheetInspectionResult.DataValidationsResult.class,
            result.inspections().getFirst());
    var unsupported =
        assertInstanceOf(
            dev.erst.gridgrind.core.model.validation.DataValidationEntryReport.Unsupported.class,
            validations.validations().getFirst());
    assertEquals("#REF!", unsupported.ranges().getFirst().value());
    assertEquals("UNKNOWN", unsupported.kind());
    assertEquals(
        result,
        GridGrindJson.readWorkbookResult(GridGrindJsonOutput.writeWorkbookResultBytes(result)));
  }

  @Test
  void observedWorksheetLabelsDoNotAcquireAuthoringNameRestrictions() throws IOException {
    Path source = tempDir.resolve("observed-label.xlsx");
    try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
      var sheet = workbook.createSheet("Valid");
      sheet.createRow(0).createCell(0).setCellValue("Item");
      sheet.createRow(1).createCell(0).setCellValue("Observed");
      var table =
          sheet.createTable(
              new org.apache.poi.ss.util.AreaReference(
                  "A1:A2", org.apache.poi.ss.SpreadsheetVersion.EXCEL2007));
      table.setName("ObservedTable");
      table.setDisplayName("ObservedTable");
      workbook.getCTWorkbook().getSheets().getSheetArray(0).setName("invalid[]");
      try (OutputStream output = Files.newOutputStream(source)) {
        workbook.write(output);
      }
    }
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
            {"protocolVersion":1,"source":{"type":"EXISTING","path":"observed-label.xlsx"},"persistence":{"type":"NONE"},"steps":[{"stepId":"summary","target":{"type":"WORKBOOK_CURRENT"},"query":{"type":"GET_WORKBOOK_SUMMARY"}},{"stepId":"tables","target":{"type":"TABLE_ALL"},"query":{"type":"GET_TABLES"}}]}
            """);
    WorkbookResult.Success result =
        assertInstanceOf(
            WorkbookResult.Success.class,
            new DefaultGridGrindRequestExecutor(
                    NativeInvocationContext.enterOwned(
                        dev.erst.gridgrind.core.semantics.ExecutionConventions.STANDARD))
                .execute(request, bindingsForExample(tempDir, request)));
    var summaryResult =
        assertInstanceOf(
            dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookSummaryResult.class,
            result.inspections().getFirst());
    var summary =
        assertInstanceOf(
            dev.erst.gridgrind.core.result.WorkbookSummary.WithSheets.class,
            summaryResult.workbook());
    assertEquals(
        java.util.List.of("invalid[]"),
        summary.sheetNames().stream()
            .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::value)
            .toList());
    assertEquals("invalid[]", summary.activeSheetName().value());
    var tables =
        assertInstanceOf(
            dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.TablesResult.class,
            result.inspections().get(1));
    assertEquals("ObservedTable", tables.tables().getFirst().name().value());
    assertEquals("invalid[]", tables.tables().getFirst().sheetName().value());
    var reopened =
        GridGrindJson.readWorkbookResult(GridGrindJsonOutput.writeWorkbookResultBytes(result));
    assertEquals(result, reopened);
  }

  private static ExecutionInputBindings bindingsForExample(Path workspace, WorkbookPlan request) {
    if (request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile
        && existingFile
            .security()
            .flatMap(dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec::passwordRef)
            .equals(java.util.Optional.of(new SecretReference("source-open-password")))) {
      return ExecutionInputBindingsFixtureSupport.bindings(
          workspace,
          request,
          Map.of(new SecretReference("source-open-password"), "GridGrind-2026"));
    }
    return ExecutionInputBindingsFixtureSupport.bindings(workspace, request);
  }
}
