package dev.erst.gridgrind.core.problem;

import static org.junit.jupiter.api.Assertions.*;

import java.util.Optional;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/** Owns the distinction between invalid native values and valid unsupported preservation. */
class WorkbookContentIssueTest {
  @ParameterizedTest
  @CsvSource({
    "XML_SOURCE_REFUSED,INVALID_WORKBOOK_CONTENT",
    "INVALID_NATIVE_VALUE,INVALID_WORKBOOK_CONTENT",
    "UNSUPPORTED_PRESERVATION,UNSUPPORTED_WORKBOOK_CONTENT",
    "UNSUPPORTED_DEPENDENCY,UNSUPPORTED_WORKBOOK_CONTENT"
  })
  void declaresCallerCorrectableContentRefusals(String reason, String code) {
    var issue = WorkbookContentIssue.valueOf(reason);
    assertEquals(GridGrindProblemCode.valueOf(code), issue.problemCode());
    assertFalse(issue.message().isBlank());
  }

  @Test
  void locationRejectsInventedPositionsAndRetainsOnlyKnownCoordinates() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookContentLocation(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookContentLocation(
                Optional.empty(),
                Optional.of("/xl/workbook.xml"),
                Optional.empty(),
                Optional.of(0),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookContentLocation(
                Optional.empty(),
                Optional.of(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    var location =
        new WorkbookContentLocation(
            Optional.empty(),
            Optional.of("/xl/workbook.xml"),
            Optional.empty(),
            Optional.of(1),
            Optional.of(31));
    assertEquals(Optional.of(31), location.column());
    assertTrue(location.sheetName().isEmpty());
    assertTrue(location.fieldName().isEmpty());
    var nativeLocation = new ProblemContextWorkbookSurfaces.ProblemLocation.NativeContent(location);
    assertEquals(
        nativeLocation,
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.unknown(), nativeLocation));
  }
}
