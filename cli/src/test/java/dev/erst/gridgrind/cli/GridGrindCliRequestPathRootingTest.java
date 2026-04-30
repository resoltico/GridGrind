package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponsePersistence;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.query.InspectionResult;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Black-box CLI regressions for request-file-relative execution paths and payloads. */
class GridGrindCliRequestPathRootingTest extends GridGrindCliTestSupport {
  @Test
  void requestFileRootsRelativePersistencePathsInItsOwnDirectory() throws IOException {
    Path requestDirectory = Files.createTempDirectory("gridgrind-cli-request-root-");
    Path requestPath = requestDirectory.resolve("relative-save.json");
    Files.writeString(
        requestPath,
        requestJson(
            "{ \"type\": \"NEW\" }",
            "{ \"type\": \"SAVE_AS\", \"path\": \"result.xlsx\" }",
            """
            [
              {
                "stepId": "ensure-budget",
                "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
                "action": { "type": "ENSURE_SHEET" }
              }
            ]
            """));
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--request", requestPath.toString()},
                InputStream.nullInputStream(),
                stdout);

    GridGrindResponse.Success response =
        assertInstanceOf(
            GridGrindResponse.Success.class, GridGrindJson.readResponse(stdout.toByteArray()));
    GridGrindResponsePersistence.PersistenceOutcome.SavedAs persistence =
        assertInstanceOf(
            GridGrindResponsePersistence.PersistenceOutcome.SavedAs.class, response.persistence());

    assertEquals(0, exitCode);
    assertEquals(requestDirectory.resolve("result.xlsx").toString(), persistence.executionPath());
    assertTrue(Files.exists(requestDirectory.resolve("result.xlsx")));
  }

  @Test
  void requestFileRootsRelativeExistingWorkbookPathsInItsOwnDirectory() throws IOException {
    Path requestDirectory = Files.createTempDirectory("gridgrind-cli-existing-root-");
    Path workbookPath = requestDirectory.resolve("input.xlsx");
    writeSingleTextWorkbook(workbookPath, "Ops", "A1", "Quarterly Budget");
    Path requestPath = requestDirectory.resolve("existing-source.json");
    Files.writeString(
        requestPath,
        requestJson(
            "{ \"type\": \"EXISTING\", \"path\": \"input.xlsx\" }",
            "{ \"type\": \"NONE\" }",
            """
            [
              {
                "stepId": "cells",
                "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Ops", "address": "A1" },
                "query": { "type": "GET_CELLS" }
              }
            ]
            """));
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--request", requestPath.toString()},
                InputStream.nullInputStream(),
                stdout);

    GridGrindResponse.Success response =
        assertInstanceOf(
            GridGrindResponse.Success.class, GridGrindJson.readResponse(stdout.toByteArray()));
    InspectionResult.CellsResult cells =
        assertInstanceOf(InspectionResult.CellsResult.class, response.inspections().getFirst());

    assertEquals(0, exitCode);
    assertEquals(
        "Quarterly Budget",
        assertInstanceOf(
                dev.erst.gridgrind.contract.dto.CellReport.TextReport.class,
                cells.cells().getFirst())
            .stringValue());
  }

  @Test
  void requestFileRootsRelativeSourceBackedPayloadsInItsOwnDirectory() throws IOException {
    Path requestDirectory = Files.createTempDirectory("gridgrind-cli-source-backed-root-");
    Path requestPath = requestDirectory.resolve("source-backed.json");
    Files.writeString(
        requestDirectory.resolve("title.txt"), "Quarterly Budget", StandardCharsets.UTF_8);
    Files.writeString(
        requestPath,
        requestJson(
            "{ \"type\": \"NEW\" }",
            "{ \"type\": \"NONE\" }",
            """
            [
              {
                "stepId": "ensure-budget",
                "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
                "action": { "type": "ENSURE_SHEET" }
              },
              {
                "stepId": "set-title",
                "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                "action": {
                  "type": "SET_CELL",
                  "value": {
                    "type": "TEXT",
                    "source": { "type": "UTF8_FILE", "path": "title.txt" }
                  }
                }
              },
              {
                "stepId": "read-title",
                "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                "query": { "type": "GET_CELLS" }
              }
            ]
            """));
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--request", requestPath.toString()},
                InputStream.nullInputStream(),
                stdout);

    GridGrindResponse.Success response =
        assertInstanceOf(
            GridGrindResponse.Success.class, GridGrindJson.readResponse(stdout.toByteArray()));
    InspectionResult.CellsResult cells =
        assertInstanceOf(InspectionResult.CellsResult.class, response.inspections().getFirst());

    assertEquals(0, exitCode);
    assertEquals(
        "Quarterly Budget",
        assertInstanceOf(
                dev.erst.gridgrind.contract.dto.CellReport.TextReport.class,
                cells.cells().getFirst())
            .stringValue());
  }

  @Test
  void requestFileRootsRelativeFormulaEnvironmentBindingsInItsOwnDirectory() throws IOException {
    ExternalFormulaScenario scenario = createExternalFormulaScenario();
    Path requestPath = scenario.directory().resolve("external-formula-request.json");
    Files.writeString(
        requestPath,
        requestJson(
            "{ \"type\": \"EXISTING\", \"path\": \"external-formula.xlsx\" }",
            "{ \"type\": \"NONE\" }",
            evaluateAllExecutionJson(),
            """
            {
              "externalWorkbooks": [
                { "workbookName": "referenced.xlsx", "path": "refs/referenced.xlsx" }
              ],
              "missingWorkbookPolicy": "ERROR",
              "udfToolpacks": []
            }
            """,
            """
            [
              {
                "stepId": "cells",
                "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Ops", "address": "B1" },
                "query": { "type": "GET_CELLS" }
              }
            ]
            """));
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--request", requestPath.toString()},
                InputStream.nullInputStream(),
                stdout);

    GridGrindResponse.Success response =
        assertInstanceOf(
            GridGrindResponse.Success.class, GridGrindJson.readResponse(stdout.toByteArray()));
    InspectionResult.CellsResult cells =
        assertInstanceOf(InspectionResult.CellsResult.class, response.inspections().getFirst());
    dev.erst.gridgrind.contract.dto.CellReport.FormulaReport formula =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.CellReport.FormulaReport.class,
            cells.cells().getFirst());

    assertEquals(0, exitCode);
    assertEquals(
        7.5d,
        assertInstanceOf(
                dev.erst.gridgrind.contract.dto.CellReport.NumberReport.class, formula.evaluation())
            .numberValue());
  }

  private static void writeSingleTextWorkbook(
      Path workbookPath, String sheetName, String address, String value) throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet(sheetName);
      org.apache.poi.ss.util.CellReference reference =
          new org.apache.poi.ss.util.CellReference(address);
      sheet.createRow(reference.getRow()).createCell(reference.getCol()).setCellValue(value);
      try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }
  }

  private static ExternalFormulaScenario createExternalFormulaScenario() throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-cli-external-formula-");
    Path referencesDirectory = Files.createDirectory(directory.resolve("refs"));
    Path referencedWorkbookPath = referencesDirectory.resolve("referenced.xlsx");
    Path workbookPath = directory.resolve("external-formula.xlsx");

    try (XSSFWorkbook referencedWorkbook = new XSSFWorkbook()) {
      referencedWorkbook.createSheet("Rates").createRow(0).createCell(0).setCellValue(7.5d);
      try (OutputStream outputStream = Files.newOutputStream(referencedWorkbookPath)) {
        referencedWorkbook.write(outputStream);
      }
    }

    try (Workbook referencedWorkbook = WorkbookFactory.create(referencedWorkbookPath.toFile());
        XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.linkExternalWorkbook("referenced.xlsx", referencedWorkbook);
      workbook.setCellFormulaValidation(false);
      XSSFSheet sheet = workbook.createSheet("Ops");
      sheet.createRow(0).createCell(0).setCellValue("External");
      sheet.getRow(0).createCell(1).setCellFormula("[referenced.xlsx]Rates!$A$1");
      var evaluator = workbook.getCreationHelper().createFormulaEvaluator();
      evaluator.setupReferencedWorkbooks(
          Map.of(
              "referenced.xlsx", referencedWorkbook.getCreationHelper().createFormulaEvaluator()));
      evaluator.evaluateFormulaCell(sheet.getRow(0).getCell(1));
      try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    return new ExternalFormulaScenario(directory);
  }

  private record ExternalFormulaScenario(Path directory) {}
}
