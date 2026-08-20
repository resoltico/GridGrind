package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Locks exact IEEE-backed numeric intake before Jackson binds request values. */
class RequestIeeeNumberExactnessTest {
  @Test
  void acceptsOrdinaryDecimalTokensThatRoundTripThroughDouble() {
    for (String token : List.of("0.1", "0.5", "1e0", "1.0", "100.00")) {
      assertDoesNotThrow(() -> GridGrindJson.readRequest(request(token)));
    }
  }

  @Test
  void rejectsNonFiniteAndLossyDoubleTokensAtTheirAuthoredField() {
    for (String token : List.of("1e400", "12345678901234567")) {
      byte[] bytes = request(token).getBytes(StandardCharsets.UTF_8);
      RequestAnalysis analysis = GridGrindJson.analyzeRequest(bytes);

      RequestNumberNotRepresentable problem =
          assertInstanceOf(
              RequestNumberNotRepresentable.class,
              analysis.structuralProblems().stream()
                  .filter(
                      candidate ->
                          candidate
                              .jsonPath()
                              .equals(java.util.Optional.of("steps[0].action.value.number")))
                  .findFirst()
                  .orElseThrow());
      assertEquals(token, problem.token());
      assertEquals(
          bytesAsText(bytes).indexOf("\"number\": " + token), problem.byteOffset().orElseThrow());
      NumberNotRepresentableException exception =
          assertThrows(
              NumberNotRepresentableException.class, () -> GridGrindJson.readRequest(bytes));
      assertEquals(problem.message(), exception.getMessage());
      assertEquals(Optional.of("steps[0].action.value.number"), exception.jsonPath());
    }
  }

  @Test
  void ordersRepresentabilityProblemsByTheirDedicatedPublicCode() {
    RequestNumberNotRepresentable number = new RequestNumberNotRepresentable("number", "1e400", 1L);
    RequestInvalidJson laterProblem =
        new RequestInvalidJson("invalid", Optional.of("number"), Optional.of(1L));

    assertEquals(
        List.of(laterProblem, number),
        RequestStructuralProblemOrder.order(List.of(laterProblem, number)));
  }

  @Test
  void rejectsLossyFloatTokensWithoutRejectingExactFloatOrIntegralCreatorValues() {
    List<RequestStructuralProblem> problems = new ArrayList<>();

    RequestNonContainerNodeValidator.validate(
        new RequestJsonNumber(0, "0.1"), Float.class, "float", 0, problems);
    RequestNonContainerNodeValidator.validate(
        new RequestJsonNumber(1, "1.5"), float.class, "exactFloat", 1, problems);
    RequestNonContainerNodeValidator.validate(
        new RequestJsonNumber(2, "1.5"), Double.class, "boxedDouble", 2, problems);
    RequestNonContainerNodeValidator.validate(
        new RequestJsonNumber(3, "42"), int.class, "integer", 3, problems);

    RequestNumberNotRepresentable problem =
        assertInstanceOf(RequestNumberNotRepresentable.class, problems.getFirst());
    assertEquals("float", problem.path());
    assertEquals("0.1", problem.token());
    assertEquals(1, problems.size());
  }

  private static String request(String number) {
    return """
        {
          "protocolVersion": "V2",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": [
            {
              "stepId": "set-number",
              "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
              "action": { "type": "SET_CELL", "value": { "type": "NUMBER", "number": %s } }
            }
          ]
        }
        """
        .formatted(number);
  }

  private static String bytesAsText(byte[] bytes) {
    return new String(bytes, StandardCharsets.UTF_8);
  }
}
