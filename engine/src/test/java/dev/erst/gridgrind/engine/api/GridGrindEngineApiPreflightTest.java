package dev.erst.gridgrind.engine.api;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.RequestAnalysis;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UncheckedIOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.concurrent.atomic.AtomicReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies production API preflight behavior across tolerant diagnostics and execution modes. */
class GridGrindEngineApiPreflightTest {
  @TempDir Path temporaryDirectory;

  @Test
  void productionDoctorUsesTheSameStaticCoreForTolerantAnalyses() {
    GridGrindRequestDoctor doctor = GridGrindEngine.requestDoctor();
    ProblemContextRequestSurfaces.RequestInput requestInput =
        ProblemContextRequestSurfaces.RequestInput.standardInput();

    RequestDoctorReport clean =
        doctor.diagnose(
            GridGrindJson.analyzeRequest(minimalRequest().getBytes(StandardCharsets.UTF_8)),
            requestInput);
    RequestDoctorReport cleanWithInputs =
        doctor.diagnose(
            GridGrindJson.analyzeRequest(minimalRequest().getBytes(StandardCharsets.UTF_8)),
            requestInput,
            inputs());
    RequestDoctorReport staticallyInvalid =
        doctor.diagnose(
            GridGrindJson.analyzeRequest(
                staticTargetMismatchRequest().getBytes(StandardCharsets.UTF_8)),
            requestInput);
    RequestDoctorReport structurallyInvalid =
        doctor.diagnose(
            GridGrindJson.analyzeRequest(
                "{\"protocolVersion\":\"V2\"}".getBytes(StandardCharsets.UTF_8)),
            requestInput,
            inputs());

    assertTrue(clean.valid());
    assertTrue(cleanWithInputs.valid());
    assertFalse(staticallyInvalid.valid());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST, staticallyInvalid.problems().getFirst().code());
    assertFalse(structurallyInvalid.valid());
    assertTrue(structurallyInvalid.summary().isEmpty());
  }

  @Test
  void productionExecutorReportsWorkbookOpenFailuresForFullAndEventReads() throws IOException {
    GridGrindRequestExecutor executor = GridGrindEngine.requestExecutor();
    Path sourcePath = temporaryDirectory.resolve("source.xlsx");
    createWorkbook(sourcePath);
    AtomicReference<Boolean> fullSourceDeleted = new AtomicReference<>(false);
    WorkbookResult.Failure fullFailure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            executor.execute(
                GridGrindJson.readRequest(
                    existingSourceRequest("FULL_XSSF").getBytes(StandardCharsets.UTF_8)),
                inputs(),
                deleteSourceAfterPreflight(sourcePath, fullSourceDeleted)));

    createWorkbook(sourcePath);
    AtomicReference<Boolean> eventReadSourceDeleted = new AtomicReference<>(false);
    WorkbookResult.Failure eventReadFailure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            executor.execute(
                GridGrindJson.readRequest(
                    existingSourceRequest("EVENT_READ").getBytes(StandardCharsets.UTF_8)),
                inputs(),
                deleteSourceAfterPreflight(sourcePath, eventReadSourceDeleted)));

    assertEquals(GridGrindProblemCode.WORKBOOK_NOT_FOUND, fullFailure.problem().code());
    assertEquals(GridGrindProblemCode.WORKBOOK_NOT_FOUND, eventReadFailure.problem().code());
    assertTrue(fullSourceDeleted.get());
    assertTrue(eventReadSourceDeleted.get());
  }

  @Test
  void productionExecutorUsesTheDoctorOrderedPrimaryPreflightProblem() {
    RequestAnalysis analysis =
        GridGrindJson.analyzeRequest(preflightFailureRequest().getBytes(StandardCharsets.UTF_8));
    ProblemContextRequestSurfaces.RequestInput requestInput =
        ProblemContextRequestSurfaces.RequestInput.standardInput();
    GridGrindRequestDoctor doctor = GridGrindEngine.requestDoctor();
    RequestDoctorReport doctorReport = doctor.diagnose(analysis, requestInput, inputs());
    WorkbookResult.Failure executionFailure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            GridGrindEngine.requestExecutor().execute(analysis, inputs()));

    assertFalse(doctorReport.valid());
    assertEquals(doctorReport.primaryProblem().orElseThrow(), executionFailure.problem());
    assertEquals(
        "steps[0].action.rows.cells[0][0].source.path",
        assertInstanceOf(
                dev.erst.gridgrind.contract.dto.ProblemContext.ResolveInputs.class,
                executionFailure.problem().context())
            .json()
            .orElseThrow()
            .jsonPathValue()
            .orElseThrow());
  }

  private GridGrindRequestInputs inputs() {
    return new GridGrindRequestInputs(temporaryDirectory, temporaryDirectory.resolve("temp-root"));
  }

  private static String minimalRequest() {
    return """
        {
          "protocolVersion": "V2",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": []
        }
        """;
  }

  private static String staticTargetMismatchRequest() {
    return """
        {
          "protocolVersion": "V2",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": [
            {
              "stepId": "set-value",
              "target": { "type": "WORKBOOK_CURRENT" },
              "action": { "type": "SET_CELL", "value": { "type": "NUMBER", "number": 1.0 } }
            }
          ]
        }
        """;
  }

  private static String existingSourceRequest(String mode) {
    return """
        {
          "protocolVersion": "V2",
          "source": { "type": "EXISTING", "path": "source.xlsx" },
          "persistence": { "type": "NONE" },
          "execution": {
            "mode": { "type": "%s" },
            "journal": { "level": "VERBOSE" }
          },
          "steps": []
        }
        """
        .formatted(mode);
  }

  private static String preflightFailureRequest() {
    return """
        {
          "protocolVersion": "V2",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": [
            {
              "stepId": "set-values",
              "target": { "type": "RANGE_BY_RANGE", "sheetName": "Budget", "range": "A1:B1" },
              "action": {
                "type": "SET_RANGE",
                "rows": {
                  "type": "TYPED",
                  "cells": [[
                    { "type": "TEXT", "source": { "type": "UTF8_FILE", "path": "missing-first.txt" } },
                    { "type": "TEXT", "source": { "type": "UTF8_FILE", "path": "missing-second.txt" } }
                  ]]
                }
              }
            }
          ]
        }
        """;
  }

  private static GridGrindJournalSink deleteSourceAfterPreflight(
      Path sourcePath, AtomicReference<Boolean> sourceDeleted) {
    return event -> {
      if ("RESOLVE_INPUTS".equals(event.category()) && "succeeded".equals(event.detail())) {
        try {
          Files.delete(sourcePath);
          sourceDeleted.set(true);
        } catch (IOException exception) {
          throw new UncheckedIOException(exception);
        }
      }
    };
  }

  private static void createWorkbook(Path path) throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook();
        OutputStream output = Files.newOutputStream(path)) {
      workbook.createSheet("Budget");
      workbook.write(output);
    }
  }
}
