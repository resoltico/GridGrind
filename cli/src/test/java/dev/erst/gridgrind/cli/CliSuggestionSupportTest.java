package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests that recovery commands are derived from request facts rather than canned search text. */
class CliSuggestionSupportTest {
  @Test
  void lookupSuggestionsNormalizeWhitespaceAndEscapeShellQuotedSearches() {
    assertEquals(
        Optional.of("gridgrind --print-protocol-catalog --search \"workbook plan\""),
        CliSuggestionSupport.protocolCatalogSearchCommandForLookupId("  workbook plan  "));
    assertEquals(
        Optional.empty(), CliSuggestionSupport.protocolCatalogSearchCommandForLookupId("  "));
    assertEquals(
        "gridgrind --print-protocol-catalog --search \"a\\\\b\\\"c\"",
        CliSuggestionSupport.protocolCatalogSearchCommand("a\\b\"c"));
  }

  @Test
  void problemSuggestionsPreferThePreciseJsonPathThenFallBackToAQuotedFieldName() {
    GridGrindProblemDetail.Problem fromPath =
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST_SHAPE,
            "Field 'ignored' is invalid",
            new ProblemContext.ReadRequest(
                RequestInput.standardInput(),
                JsonLocation.pathOnly("steps[12].formula_environment.type")));
    GridGrindProblemDetail.Problem fromMessage =
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST_SHAPE,
            "Field 'persistence' is invalid",
            new ProblemContext.ReadRequest(
                RequestInput.standardInput(), JsonLocation.unavailable()));
    GridGrindProblemDetail.Problem fromNonTypePath =
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST_SHAPE,
            "Field 'ignored' is invalid",
            new ProblemContext.BindRequest(
                RequestInput.standardInput(),
                JsonLocation.pathOnly("steps[0].action.zoomPercent")));
    GridGrindProblemDetail.Problem withoutContextualFact =
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST,
            "Invalid request",
            new ProblemContext.ValidateRequest(RequestShape.unknown()));
    GridGrindProblemDetail.Problem withoutQuotedFallback =
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST_SHAPE,
            "Invalid request",
            new ProblemContext.ReadRequest(
                RequestInput.standardInput(), JsonLocation.unavailable()));

    assertEquals(
        Optional.of(
            "gridgrind --print-protocol-catalog --search \"steps formula environment type\""),
        CliSuggestionSupport.protocolCatalogSearchCommandForProblem(fromPath));
    assertEquals(
        Optional.of("gridgrind --print-protocol-catalog --search \"persistence\""),
        CliSuggestionSupport.protocolCatalogSearchCommandForProblem(fromMessage));
    assertEquals(
        Optional.of("gridgrind --print-protocol-catalog --search \"steps action zoomPercent\""),
        CliSuggestionSupport.protocolCatalogSearchCommandForProblem(fromNonTypePath));
    assertEquals(
        Optional.empty(),
        CliSuggestionSupport.protocolCatalogSearchCommandForProblem(withoutContextualFact));
    assertEquals(
        Optional.empty(),
        CliSuggestionSupport.protocolCatalogSearchCommandForProblem(withoutQuotedFallback));
  }
}
