package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused coverage for typed problem-context variants and their optional accessors. */
class ProblemContextCoverageTest {
  @Test
  void foundationalVariantsExposeOptionalFactsWithoutNullPadding() {
    ProblemContextRequestSurfaces.RequestShape unknownShape =
        ProblemContextRequestSurfaces.RequestShape.unknown();
    ProblemContextRequestSurfaces.RequestInput standardInput =
        ProblemContextRequestSurfaces.RequestInput.standardInput();
    ProblemContextRequestSurfaces.ResponseOutput standardOutput =
        ProblemContextRequestSurfaces.ResponseOutput.standardOutput();
    ProblemContextRequestSurfaces.JsonLocation unavailable =
        ProblemContextRequestSurfaces.JsonLocation.unavailable();
    ProblemContextRequestSurfaces.JsonLocation lineColumn =
        ProblemContextRequestSurfaces.JsonLocation.lineColumn(4, 12);
    ProblemContextRequestSurfaces.CliArgument unknownArgument =
        ProblemContextRequestSurfaces.CliArgument.unknown();
    ProblemContextWorkbookSurfaces.InputReference unknownInput =
        ProblemContextWorkbookSurfaces.InputReference.unknown();
    ProblemContextWorkbookSurfaces.InputReference kindInput =
        ProblemContextWorkbookSurfaces.InputReference.kind("cell text");
    ProblemContextWorkbookSurfaces.WorkbookReference newWorkbook =
        ProblemContextWorkbookSurfaces.WorkbookReference.newWorkbook();
    ProblemContextWorkbookSurfaces.PersistenceReference overwrite =
        ProblemContextWorkbookSurfaces.PersistenceReference.overwriteSource("/tmp/source.xlsx");
    ProblemContextWorkbookSurfaces.PersistenceReference saveAs =
        ProblemContextWorkbookSurfaces.PersistenceReference.saveAs("/tmp/output.xlsx");

    assertEquals(Optional.empty(), unknownShape.known());
    assertEquals(Optional.empty(), unknownShape.sourceTypeValue());
    assertEquals(Optional.empty(), unknownShape.persistenceTypeValue());
    assertEquals(Optional.empty(), standardInput.requestPathValue());
    assertEquals(Optional.empty(), standardOutput.responsePathValue());
    assertEquals(Optional.empty(), unavailable.jsonPathValue());
    assertEquals(Optional.empty(), unavailable.jsonLineValue());
    assertEquals(Optional.empty(), unavailable.jsonColumnValue());
    assertEquals(Optional.empty(), lineColumn.jsonPathValue());
    assertEquals(Optional.of(4), lineColumn.jsonLineValue());
    assertEquals(Optional.of(12), lineColumn.jsonColumnValue());
    assertEquals(Optional.empty(), unknownArgument.argumentValue());
    assertEquals(Optional.empty(), unknownInput.inputKindValue());
    assertEquals(Optional.empty(), unknownInput.inputPathValue());
    assertEquals(Optional.of("cell text"), kindInput.inputKindValue());
    assertEquals(Optional.empty(), kindInput.inputPathValue());
    assertEquals(Optional.empty(), newWorkbook.sourceWorkbookPathValue());
    assertEquals(Optional.of("/tmp/source.xlsx"), overwrite.sourceWorkbookPathValue());
    assertEquals(Optional.empty(), overwrite.persistencePathValue());
    assertEquals(Optional.empty(), saveAs.sourceWorkbookPathValue());
    assertEquals(Optional.of("/tmp/output.xlsx"), saveAs.persistencePathValue());
  }

  @Test
  void stageContextsExposeTypedRequestInputLocationAndOutputFacts() {
    ProblemContextRequestSurfaces.RequestShape requestShape =
        ProblemContextRequestSurfaces.RequestShape.known("EXISTING", "OVERWRITE");
    ProblemContext.ReadRequest readRequest =
        new ProblemContext.ReadRequest(
                ProblemContextRequestSurfaces.RequestInput.standardInput(), null)
            .withJson(ProblemContextRequestSurfaces.JsonLocation.lineColumn(7, 3));
    ProblemContext.ValidateRequest validateRequest =
        new ProblemContext.ValidateRequest(requestShape);
    ProblemContext.ResolveInputs resolveInputs =
        new ProblemContext.ResolveInputs(
            requestShape, ProblemContextWorkbookSurfaces.InputReference.kind("comment"));
    ProblemContext.OpenWorkbook openWorkbook =
        new ProblemContext.OpenWorkbook(
            requestShape, ProblemContextWorkbookSurfaces.WorkbookReference.newWorkbook());
    ProblemContext.ExecuteCalculation.Preflight preflight =
        new ProblemContext.ExecuteCalculation.Preflight(
            requestShape, ProblemContextWorkbookSurfaces.ProblemLocation.range("Budget", "A1:B2"));
    ProblemContext.ExecuteCalculation.Execution execution =
        new ProblemContext.ExecuteCalculation.Execution(
            requestShape,
            ProblemContextWorkbookSurfaces.ProblemLocation.formulaCell(
                "Budget", "B4", "SUM(B2:B3)"));
    ProblemContext.ExecuteStep executeStep =
        new ProblemContext.ExecuteStep(
            requestShape,
            new ProblemContextWorkbookSurfaces.StepReference(
                2, "step-2", "ASSERTION", "EXPECT_CELL_VALUE"),
            ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("Budget", "BudgetTotal"));
    ProblemContext.PersistWorkbook persistOverwrite =
        new ProblemContext.PersistWorkbook(
            requestShape,
            ProblemContextWorkbookSurfaces.PersistenceReference.overwriteSource(
                "/tmp/source.xlsx"));
    ProblemContext.PersistWorkbook persistSaveAs =
        new ProblemContext.PersistWorkbook(
            requestShape,
            ProblemContextWorkbookSurfaces.PersistenceReference.saveAs("/tmp/out.xlsx"));
    ProblemContext.ExecuteRequest executeRequest = new ProblemContext.ExecuteRequest(requestShape);
    ProblemContext.WriteResponse writeResponse =
        new ProblemContext.WriteResponse(
            ProblemContextRequestSurfaces.ResponseOutput.responseFile("/tmp/response.json"));

    assertEquals(Optional.empty(), readRequest.requestPath());
    assertEquals(Optional.empty(), readRequest.jsonPath());
    assertEquals(Optional.of(7), readRequest.jsonLine());
    assertEquals(Optional.of(3), readRequest.jsonColumn());
    assertEquals(Optional.of("EXISTING"), validateRequest.sourceType());
    assertEquals(Optional.of("OVERWRITE"), validateRequest.persistenceType());
    assertEquals(Optional.of("EXISTING"), resolveInputs.sourceType());
    assertEquals(Optional.of("OVERWRITE"), resolveInputs.persistenceType());
    assertEquals(Optional.of("comment"), resolveInputs.inputKind());
    assertEquals(Optional.empty(), resolveInputs.inputPath());
    assertEquals(Optional.of("EXISTING"), openWorkbook.sourceType());
    assertEquals(Optional.of("OVERWRITE"), openWorkbook.persistenceType());
    assertEquals(Optional.empty(), openWorkbook.sourceWorkbookPath());
    assertEquals(Optional.of("EXISTING"), preflight.sourceType());
    assertEquals(Optional.of("OVERWRITE"), preflight.persistenceType());
    assertEquals(Optional.of("Budget"), preflight.sheetName());
    assertEquals(Optional.of("A1:B2"), preflight.range());
    assertEquals(Optional.of("Budget"), execution.sheetName());
    assertEquals(Optional.of("B4"), execution.address());
    assertEquals(Optional.of("SUM(B2:B3)"), execution.formula());
    assertEquals(Optional.of("EXISTING"), executeStep.sourceType());
    assertEquals(Optional.of("OVERWRITE"), executeStep.persistenceType());
    assertEquals("ASSERTION", executeStep.stepKind());
    assertEquals("EXPECT_CELL_VALUE", executeStep.stepType());
    assertEquals(Optional.of("Budget"), executeStep.sheetName());
    assertEquals(Optional.of("BudgetTotal"), executeStep.namedRangeName());
    assertEquals(
        Optional.of("BudgetTotal"),
        new ProblemContext.ExecuteCalculation.Execution(
                requestShape,
                ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal"))
            .namedRangeName());
    assertEquals(Optional.of("EXISTING"), persistOverwrite.sourceType());
    assertEquals(Optional.of("OVERWRITE"), persistOverwrite.persistenceType());
    assertEquals(Optional.of("/tmp/source.xlsx"), persistOverwrite.sourceWorkbookPath());
    assertEquals(Optional.empty(), persistOverwrite.persistencePath());
    assertEquals(Optional.of("EXISTING"), persistSaveAs.sourceType());
    assertEquals(Optional.of("OVERWRITE"), persistSaveAs.persistenceType());
    assertEquals(Optional.empty(), persistSaveAs.sourceWorkbookPath());
    assertEquals(Optional.of("/tmp/out.xlsx"), persistSaveAs.persistencePath());
    assertEquals(Optional.of("EXISTING"), executeRequest.sourceType());
    assertEquals(Optional.of("OVERWRITE"), executeRequest.persistenceType());
    assertEquals(Optional.of("/tmp/response.json"), writeResponse.responsePath());
  }

  @Test
  void supportingVariantsValidateNegativeAndAbsentBranches() {
    assertEquals(
        Optional.empty(),
        ProblemContextWorkbookSurfaces.ProblemLocation.address("B4").sheetNameValue());
    assertEquals(
        Optional.of("B4"),
        ProblemContextWorkbookSurfaces.ProblemLocation.address("B4").addressValue());
    assertEquals(
        Optional.empty(),
        ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2").sheetNameValue());
    assertEquals(
        Optional.of("A1:B2"),
        ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2").rangeValue());
    assertEquals(
        Optional.empty(),
        ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal").sheetNameValue());
    assertEquals(
        Optional.of("BudgetTotal"),
        ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal")
            .namedRangeNameValue());
    assertEquals(
        Optional.of("BudgetTotal"),
        ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("Budget", "BudgetTotal")
            .namedRangeNameValue());
    assertEquals(
        Optional.empty(), ProblemContextWorkbookSurfaces.ProblemLocation.unknown().formulaValue());
    assertEquals(
        "stepIndex must be greater than or equal to 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ProblemContextWorkbookSurfaces.StepReference(
                        -1, "step", "MUTATION", "SET_CELL"))
            .getMessage());
    assertEquals(
        "jsonLine must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> ProblemContextRequestSurfaces.JsonLocation.lineColumn(0, 1))
            .getMessage());
    assertEquals(
        "jsonColumn must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> ProblemContextRequestSurfaces.JsonLocation.lineColumn(1, 0))
            .getMessage());
    assertEquals(
        "jsonLine must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> ProblemContextRequestSurfaces.JsonLocation.located("steps[0]", 0, 1))
            .getMessage());
    assertEquals(
        "jsonColumn must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> ProblemContextRequestSurfaces.JsonLocation.located("steps[0]", 1, 0))
            .getMessage());
  }
}
