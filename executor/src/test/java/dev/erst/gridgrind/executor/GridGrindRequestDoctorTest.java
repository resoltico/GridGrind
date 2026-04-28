package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Tests for non-executing request linting through the executor-owned doctor surface. */
class GridGrindRequestDoctorTest {
  @Test
  void diagnosesCleanRequestsAndSummarizesStandardInputRequirements() throws IOException {
    WorkbookPlan cleanRequest =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "NEW" },
              "steps": [
                {
                  "stepId": "ensure-budget",
                  "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
                  "action": { "type": "ENSURE_SHEET" }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));
    WorkbookPlan standardInputRequest =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "NEW" },
              "steps": [
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
                    "value": { "type": "TEXT", "source": { "type": "STANDARD_INPUT" } }
                  }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    RequestDoctorReport cleanReport = new GridGrindRequestDoctor().diagnose(cleanRequest);
    RequestDoctorReport standardInputReport =
        new GridGrindRequestDoctor().diagnose(standardInputRequest);

    assertTrue(cleanReport.valid());
    assertEquals(AnalysisSeverity.INFO, cleanReport.severity());
    assertEquals("NEW", cleanReport.summary().orElseThrow().sourceType());
    assertEquals("NONE", cleanReport.summary().orElseThrow().persistenceType());
    assertEquals(1, cleanReport.summary().orElseThrow().stepCount());
    assertEquals(1, cleanReport.summary().orElseThrow().mutationStepCount());
    assertFalse(cleanReport.summary().orElseThrow().requiresStandardInputBinding());
    assertTrue(standardInputReport.valid());
    assertTrue(standardInputReport.summary().orElseThrow().requiresStandardInputBinding());
  }

  @Test
  void diagnosesWarningsAndBlockingValidationFailures() throws IOException {
    WorkbookPlan warningRequest =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "NEW" },
              "steps": [
                {
                  "stepId": "ensure-budget",
                  "target": { "type": "SHEET_BY_NAME", "name": "Budget Sheet" },
                  "action": { "type": "ENSURE_SHEET" }
                },
                {
                  "stepId": "set-formula",
                  "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget Sheet", "address": "B2" },
                  "action": {
                    "type": "SET_CELL",
                    "value": {
                      "type": "FORMULA",
                      "source": { "type": "INLINE", "text": "Budget Sheet!A1" }
                    }
                  }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));
    WorkbookPlan invalidRequest =
        new WorkbookPlan(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.OverwriteSource(),
            java.util.List.of());

    RequestDoctorReport warningReport = new GridGrindRequestDoctor().diagnose(warningRequest);
    RequestDoctorReport invalidReport = new GridGrindRequestDoctor().diagnose(invalidRequest);

    assertTrue(warningReport.valid());
    assertEquals(AnalysisSeverity.WARNING, warningReport.severity());
    assertEquals(1, warningReport.warnings().size());
    assertTrue(warningReport.warnings().getFirst().message().contains("single quotes"));
    assertFalse(invalidReport.valid());
    assertEquals(AnalysisSeverity.ERROR, invalidReport.severity());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST, invalidReport.problem().orElseThrow().code());
    assertTrue(
        invalidReport.problem().orElseThrow().message().contains("OVERWRITE persistence requires"));
  }

  @Test
  void diagnoseUsesDelegatingConstructorsAndCountsStepKindsWithoutBindings() throws IOException {
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "NEW" },
              "steps": [
                {
                  "stepId": "ensure-budget",
                  "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
                  "action": { "type": "ENSURE_SHEET" }
                },
                {
                  "stepId": "assert-budget",
                  "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                  "assertion": {
                    "type": "EXPECT_CELL_VALUE",
                    "expectedValue": { "type": "BLANK" }
                  }
                },
                {
                  "stepId": "read-budget",
                  "target": { "type": "WORKBOOK_CURRENT" },
                  "query": { "type": "GET_WORKBOOK_SUMMARY" }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    RequestDoctorReport report =
        new GridGrindRequestDoctor(new ExecutionValidationSupport()).diagnose(request);

    assertTrue(report.valid());
    assertEquals(3, report.summary().orElseThrow().stepCount());
    assertEquals(1, report.summary().orElseThrow().mutationStepCount());
    assertEquals(1, report.summary().orElseThrow().assertionStepCount());
    assertEquals(1, report.summary().orElseThrow().inspectionStepCount());
  }

  @Test
  void productionConstructorsInstantiateConcreteWorkbookSupport() {
    assertTrue(
        new GridGrindRequestDoctor()
            .diagnose(null)
            .problem()
            .orElseThrow()
            .message()
            .contains("request"));
    assertTrue(
        new GridGrindRequestDoctor(new ExecutionValidationSupport())
            .diagnose(null)
            .problem()
            .orElseThrow()
            .message()
            .contains("request"));
    assertTrue(
        new GridGrindRequestDoctor(
                new ExecutionValidationSupport(),
                new ExecutionWorkbookSupport(Files::createTempFile))
            .diagnose(null)
            .problem()
            .orElseThrow()
            .message()
            .contains("request"));
    assertThrows(
        NullPointerException.class,
        () ->
            new GridGrindRequestDoctor()
                .diagnose(
                    new WorkbookPlan(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        java.util.List.of()),
                    null));
  }

  @Test
  void diagnoseWithBindingsReturnsCleanReportWhenRelativeInputsResolve() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-doctor-bindings-");
    Files.writeString(workingDirectory.resolve("title.txt"), "Quarterly Budget");
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "NEW" },
              "steps": [
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
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    RequestDoctorReport report =
        new GridGrindRequestDoctor()
            .diagnose(request, new ExecutionInputBindings(workingDirectory));

    assertTrue(report.valid());
    assertEquals(AnalysisSeverity.INFO, report.severity());
    assertEquals(2, report.summary().orElseThrow().stepCount());
  }

  @Test
  void diagnoseWithBindingsRecordsInputKindWhenStandardInputBindingIsMissing() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-doctor-stdin-");
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "NEW" },
              "steps": [
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
                      "source": { "type": "STANDARD_INPUT" }
                    }
                  }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    RequestDoctorReport report =
        new GridGrindRequestDoctor()
            .diagnose(request, new ExecutionInputBindings(workingDirectory));
    dev.erst.gridgrind.contract.dto.ProblemContext.ResolveInputs context =
        (dev.erst.gridgrind.contract.dto.ProblemContext.ResolveInputs)
            report.problem().orElseThrow().context();

    assertFalse(report.valid());
    assertEquals(java.util.Optional.of("cell text"), context.inputKind());
    assertEquals(java.util.Optional.empty(), context.inputPath());
  }

  @Test
  void diagnoseWithBindingsRejectsBlankResolvedCellText() throws IOException {
    WorkbookPlan invalidRequest =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "NEW" },
              "steps": [
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
                      "source": { "type": "INLINE", "text": "" }
                    }
                  }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    RequestDoctorReport report =
        new GridGrindRequestDoctor()
            .diagnose(invalidRequest, ExecutionInputBindings.processDefault());

    assertFalse(report.valid());
    assertEquals(AnalysisSeverity.ERROR, report.severity());
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, report.problem().orElseThrow().code());
    assertEquals("RESOLVE_INPUTS", report.problem().orElseThrow().context().stage());
    assertEquals("cell text must not be blank", report.problem().orElseThrow().message());
  }

  @Test
  void diagnoseWithBindingsCapturesInputSourceContextWhenFilesAreMissing() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-doctor-missing-");
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "NEW" },
              "steps": [
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
                      "source": { "type": "UTF8_FILE", "path": "missing.txt" }
                    }
                  }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    RequestDoctorReport report =
        new GridGrindRequestDoctor()
            .diagnose(request, new ExecutionInputBindings(workingDirectory));

    assertFalse(report.valid());
    assertEquals(AnalysisSeverity.ERROR, report.severity());
    assertEquals(
        GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND, report.problem().orElseThrow().code());
    var context =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.ProblemContext.ResolveInputs.class,
            report.problem().orElseThrow().context());
    assertEquals("RESOLVE_INPUTS", context.stage());
    assertEquals(java.util.Optional.of("cell text"), context.inputKind());
    assertEquals(
        java.util.Optional.of(workingDirectory.resolve("missing.txt").toString()),
        context.inputPath());
  }

  @Test
  void diagnoseWithBindingsPreflightsExistingWorkbookSources() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-doctor-workbook-");
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "EXISTING", "path": "missing-workbook.xlsx" },
              "steps": [
                {
                  "stepId": "summary",
                  "target": { "type": "WORKBOOK_CURRENT" },
                  "query": { "type": "GET_WORKBOOK_SUMMARY" }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    RequestDoctorReport report =
        new GridGrindRequestDoctor()
            .diagnose(request, new ExecutionInputBindings(workingDirectory));

    assertFalse(report.valid());
    assertEquals(AnalysisSeverity.ERROR, report.severity());
    assertEquals(GridGrindProblemCode.WORKBOOK_NOT_FOUND, report.problem().orElseThrow().code());
    var context =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.ProblemContext.OpenWorkbook.class,
            report.problem().orElseThrow().context());
    assertEquals("OPEN_WORKBOOK", context.stage());
    assertEquals(
        java.util.Optional.of(workingDirectory.resolve("missing-workbook.xlsx").toString()),
        context.sourceWorkbookPath());
  }

  @Test
  void diagnoseWithBindingsAcceptsExistingWorkbookSourcesWhenOpenSucceeds() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-doctor-existing-success-");
    Path workbookPath = workingDirectory.resolve("existing.xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.save(workbookPath);
    }
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "EXISTING", "path": "existing.xlsx" },
              "steps": [
                {
                  "stepId": "summary",
                  "target": { "type": "WORKBOOK_CURRENT" },
                  "query": { "type": "GET_WORKBOOK_SUMMARY" }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    RequestDoctorReport report =
        new GridGrindRequestDoctor()
            .diagnose(request, new ExecutionInputBindings(workingDirectory));

    assertTrue(report.valid());
    assertEquals(AnalysisSeverity.INFO, report.severity());
    assertEquals("EXISTING", report.summary().orElseThrow().sourceType());
  }

  @Test
  void diagnosesNullRequestsAsInvalid() {
    RequestDoctorReport report = new GridGrindRequestDoctor().diagnose(null);

    assertFalse(report.valid());
    assertEquals(AnalysisSeverity.ERROR, report.severity());
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, report.problem().orElseThrow().code());
    assertEquals("request must not be null", report.problem().orElseThrow().message());
  }
}
