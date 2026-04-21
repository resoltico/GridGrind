package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.AnalysisSeverity;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
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
                  "target": { "type": "BY_NAME", "name": "Budget" },
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
                  "target": { "type": "BY_NAME", "name": "Budget" },
                  "action": { "type": "ENSURE_SHEET" }
                },
                {
                  "stepId": "set-title",
                  "target": { "type": "BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
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
    assertEquals("NEW", cleanReport.summary().sourceType());
    assertEquals("NONE", cleanReport.summary().persistenceType());
    assertEquals(1, cleanReport.summary().stepCount());
    assertEquals(1, cleanReport.summary().mutationStepCount());
    assertFalse(cleanReport.summary().requiresStandardInputBinding());
    assertTrue(standardInputReport.valid());
    assertTrue(standardInputReport.summary().requiresStandardInputBinding());
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
                  "target": { "type": "BY_NAME", "name": "Budget Sheet" },
                  "action": { "type": "ENSURE_SHEET" }
                },
                {
                  "stepId": "set-formula",
                  "target": { "type": "BY_ADDRESS", "sheetName": "Budget Sheet", "address": "B2" },
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
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, invalidReport.problem().code());
    assertTrue(invalidReport.problem().message().contains("OVERWRITE persistence requires"));
  }

  @Test
  void diagnosesNullRequestsAsInvalid() {
    RequestDoctorReport report = new GridGrindRequestDoctor().diagnose(null);

    assertFalse(report.valid());
    assertEquals(AnalysisSeverity.ERROR, report.severity());
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, report.problem().code());
    assertEquals("request must not be null", report.problem().message());
  }
}
