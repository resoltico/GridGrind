package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Covers binding-failure normalization independently from full request analysis. */
class RequestBindingFailureTest {
  @Test
  void normalizesADirectConstructorFailureAtTheFragmentOwner() {
    RequestBindingFailure failure =
        RequestBindingFailure.from(
            new IllegalArgumentException("direct failure"),
            new RequestJsonString(0, "value"),
            "source.path",
            RequestDiagnosticRedactor.empty());

    assertEquals("direct failure", failure.exception().getMessage());
    assertEquals("source.path", failure.jsonPath());
    assertTrue(failure.byteOffset().isEmpty());
  }

  @Test
  void translatesAJacksonBindingFailureWithoutDroppingTheFragment() {
    List<RequestBindingFailure> bindingFailures = new ArrayList<>();

    assertTrue(
        RequestFragmentBinder.bindNode(
                new RequestJsonString(0, "not-an-integer"),
                Integer.class,
                "count",
                List.of(),
                bindingFailures,
                RequestDiagnosticRedactor.empty())
            .isEmpty());
    assertEquals(1, bindingFailures.size());
    assertEquals("count", bindingFailures.getFirst().jsonPath());
  }

  @Test
  void retainsACompletePlanInvariantAtItsExistingPath() {
    InvalidRequestException exception =
        new InvalidRequestException(
            new DuplicateStepId("duplicate", "steps[1].stepId"),
            java.util.Optional.of("steps[1].stepId"),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            null);

    RequestBindingFailure failure =
        RequestBindingFailure.fromCompletePlan(exception, RequestDiagnosticRedactor.empty());

    assertEquals("steps[1].stepId", failure.jsonPath());
    assertEquals(exception.getMessage(), failure.exception().getMessage());
    assertTrue(failure.byteOffset().isEmpty());
  }

  @Test
  void retainsACompletePlanShapeProblemAtItsExistingPath() {
    InvalidRequestShapeException exception =
        new InvalidRequestShapeException(
            new MessageShape("invalid complete plan", Optional.of("steps")),
            Optional.of("steps"),
            Optional.empty(),
            Optional.empty(),
            null);

    RequestBindingFailure failure =
        RequestBindingFailure.fromCompletePlan(exception, RequestDiagnosticRedactor.empty());

    assertEquals("steps", failure.jsonPath());
    assertEquals(exception.getMessage(), failure.exception().getMessage());
    assertTrue(failure.byteOffset().isEmpty());
  }

  @Test
  void normalizesAnUnstructuredCompletePlanFailureAtTheRequestRoot() {
    RequestBindingFailure failure =
        RequestBindingFailure.fromCompletePlan(
            new IllegalArgumentException("complete plan failure"),
            RequestDiagnosticRedactor.empty());

    assertEquals("request", failure.jsonPath());
    assertEquals("complete plan failure", failure.exception().getMessage());
  }

  @Test
  void rejectsInvalidBindingFailureCoordinates() {
    IllegalArgumentException exception = new IllegalArgumentException("failure");

    assertThrows(
        IllegalArgumentException.class,
        () -> new RequestBindingFailure(exception, "", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new RequestBindingFailure(exception, "path", Optional.of(-1L)));
  }
}
