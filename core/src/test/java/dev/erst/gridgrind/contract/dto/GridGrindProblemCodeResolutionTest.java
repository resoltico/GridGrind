package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

/** Covers cause-specific public resolution text for problem codes. */
class GridGrindProblemCodeResolutionTest {
  @Test
  void resolutionForUsesCauseSpecificGuidanceAcrossAssertionAndIoContexts() {
    ProblemContext.ExecuteStep assertionContext =
        new ProblemContext.ExecuteStep(
            ProblemContextRequestSurfaces.RequestShape.known("NEW", "NONE"),
            new ProblemContextWorkbookSurfaces.StepReference(
                0, "assert-total", "ASSERTION", "EXPECT_CELL_VALUE"),
            ProblemContextWorkbookSurfaces.ProblemLocation.cell("Budget", "B2"));
    ProblemContext.PersistWorkbook saveAsContext =
        new ProblemContext.PersistWorkbook(
            ProblemContextRequestSurfaces.RequestShape.known("EXISTING", "SAVE_AS"),
            ProblemContextWorkbookSurfaces.PersistenceReference.saveAs("/tmp/out.xlsx"));
    ProblemContext.PersistWorkbook overwriteSourceContext =
        new ProblemContext.PersistWorkbook(
            ProblemContextRequestSurfaces.RequestShape.known("EXISTING", "OVERWRITE"),
            ProblemContextWorkbookSurfaces.PersistenceReference.overwrite("/tmp/source.xlsx"));
    ProblemContext.OpenWorkbook openWorkbookContext =
        new ProblemContext.OpenWorkbook(
            ProblemContextRequestSurfaces.RequestShape.known("EXISTING", "NONE"),
            ProblemContextWorkbookSurfaces.WorkbookReference.existingFile("/tmp/source.xlsx"));
    ProblemContext.OpenWorkbook newWorkbookContext =
        new ProblemContext.OpenWorkbook(
            ProblemContextRequestSurfaces.RequestShape.known("NEW", "NONE"),
            ProblemContextWorkbookSurfaces.WorkbookReference.newWorkbook());
    ProblemContext.WriteResponse writeResponseContext =
        new ProblemContext.WriteResponse(
            ProblemContextRequestSurfaces.ResponseOutput.responseFile("/tmp/report.json"));
    ProblemContext.WriteResponse stdoutResponseContext =
        new ProblemContext.WriteResponse(
            ProblemContextRequestSurfaces.ResponseOutput.standardOutput());

    assertEquals(
        "Inspect problem.assertionFailure observations, then adjust the failing assertion or preceding workbook mutations and retry.",
        GridGrindProblemCode.ASSERTION_FAILED.resolutionFor("assertion failed", assertionContext));
    assertEquals(
        "Choose a new SAVE_AS destination path, remove the conflicting file, or set SAVE_AS.ifExists=REPLACE before retrying.",
        GridGrindProblemCode.IO_ERROR.resolutionFor(
            "Could not write workbook: already exists", saveAsContext));
    assertEquals(
        "Check the SAVE_AS destination path, parent directory permissions, free disk space, and file locks before retrying.",
        GridGrindProblemCode.IO_ERROR.resolutionFor("disk full", saveAsContext));
    assertEquals(
        "Check the destination workbook path, permissions, free disk space, and file locks before retrying.",
        GridGrindProblemCode.IO_ERROR.resolutionFor("access denied", overwriteSourceContext));
    assertEquals(
        "Check the source workbook path, permissions, and file locks before retrying.",
        GridGrindProblemCode.IO_ERROR.resolutionFor("access denied", openWorkbookContext));
    assertEquals(
        GridGrindProblemCode.IO_ERROR.resolution(),
        GridGrindProblemCode.IO_ERROR.resolutionFor("access denied", newWorkbookContext));
    assertEquals(
        "Choose a new --response path or remove the conflicting file, then retry.",
        GridGrindProblemCode.IO_ERROR.resolutionFor(
            "Could not write response: already exists", writeResponseContext));
    assertEquals(
        "Check the --response destination path, parent directory permissions, free disk space, and file locks before retrying.",
        GridGrindProblemCode.IO_ERROR.resolutionFor("access denied", writeResponseContext));
    assertEquals(
        GridGrindProblemCode.IO_ERROR.resolution(),
        GridGrindProblemCode.IO_ERROR.resolutionFor("access denied", stdoutResponseContext));
  }

  @Test
  void resolutionForCoversContextualInvalidArgumentBranches() {
    ProblemContext.ParseArguments requestArgumentContext =
        new ProblemContext.ParseArguments(
            ProblemContextRequestSurfaces.CliArgument.named("--request"));
    ProblemContext.ParseArguments lookupArgumentContext =
        new ProblemContext.ParseArguments(
            ProblemContextRequestSurfaces.CliArgument.named("--lookup"));
    ProblemContext.ParseArguments queryArgumentContext =
        new ProblemContext.ParseArguments(
            ProblemContextRequestSurfaces.CliArgument.named("--query"));

    assertEquals(
        "Requests that bind STANDARD_INPUT-authored payloads must arrive by --request so standard input stays available for payload bytes.",
        GridGrindProblemCode.INVALID_ARGUMENTS.resolutionFor(
            "Requests with STANDARD_INPUT-authored payloads must be provided via --request.",
            requestArgumentContext));
    assertEquals(
        "Use --print-recipe-catalog first when you need the stable recipe ids, requestFileName, advisory, and requiredWorkspacePaths, or use --print-recipe-keyword-match when you know the goal but not the id.",
        GridGrindProblemCode.INVALID_ARGUMENTS.resolutionFor(
            "Unknown recipe: workbook_health", lookupArgumentContext));
    assertEquals(
        "Use a natural-language query that leaves at least one searchable non-stop-word term after normalization.",
        GridGrindProblemCode.INVALID_ARGUMENTS.resolutionFor(
            "Invalid keyword query: only stop words remain after normalization.",
            queryArgumentContext));
  }
}
