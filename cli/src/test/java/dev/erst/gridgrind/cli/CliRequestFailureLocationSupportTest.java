package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.contract.json.InvalidJsonException;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/**
 * Verifies that late payload failures retain their location instead of falling back to unavailable.
 */
class CliRequestFailureLocationSupportTest {
  @Test
  void projectsEveryNormalizedPayloadLocationVariant() {
    assertEquals(
        JsonLocation.located("steps[0].action.zoomPercent", 8, 17),
        CliRequestFailureLocationSupport.locationFor(
            failure(Optional.of("steps[0].action.zoomPercent"), Optional.of(8), Optional.of(17))));
    assertEquals(
        JsonLocation.pathOnly("steps[0].action.zoomPercent"),
        CliRequestFailureLocationSupport.locationFor(
            failure(
                Optional.of("steps[0].action.zoomPercent"), Optional.empty(), Optional.empty())));
    assertEquals(
        JsonLocation.lineColumn(8, 17),
        CliRequestFailureLocationSupport.locationFor(
            failure(Optional.empty(), Optional.of(8), Optional.of(17))));
    assertEquals(
        JsonLocation.unavailable(),
        CliRequestFailureLocationSupport.locationFor(
            failure(Optional.empty(), Optional.empty(), Optional.empty())));
  }

  @Test
  void findsPayloadFailuresNestedInAnotherException() {
    assertEquals(
        JsonLocation.pathOnly("steps[0].action.zoomPercent"),
        CliRequestFailureLocationSupport.locationFor(
            new IllegalArgumentException(
                "wrapper",
                failure(
                    Optional.of("steps[0].action.zoomPercent"),
                    Optional.empty(),
                    Optional.empty()))));
  }

  @Test
  void leavesNonPayloadFailuresUnlocated() {
    assertEquals(
        JsonLocation.unavailable(),
        CliRequestFailureLocationSupport.locationFor(new IllegalArgumentException("unrelated")));
  }

  private static InvalidJsonException failure(
      Optional<String> jsonPath, Optional<Integer> jsonLine, Optional<Integer> jsonColumn) {
    return new InvalidJsonException("Invalid value", jsonPath, jsonLine, jsonColumn, null);
  }
}
