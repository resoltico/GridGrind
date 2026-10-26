package dev.erst.gridgrind.json.intake;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.InvalidFormulaInputException;
import dev.erst.gridgrind.core.model.formula.InvalidRawFormulaTextException;
import dev.erst.gridgrind.core.problem.DuplicateStepId;
import dev.erst.gridgrind.core.problem.FormulaRequestException;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.problem.InvalidRequestShapeException;
import dev.erst.gridgrind.core.problem.MessageInvariant;
import dev.erst.gridgrind.core.problem.MessageShape;
import dev.erst.gridgrind.core.problem.RequestBindingFailure;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Covers binding-failure normalization independently from full request analysis. */
class RequestBindingFailureTest {
  @Test
  void locatesADirectConstructorFailureAtTheFragmentOwner() {
    var failure =
        RequestConstructorValidation.problem(
            new IllegalArgumentException("direct failure"), "source.path", false);
    assertEquals("direct failure", failure.exception().getMessage());
    assertEquals("source.path", failure.jsonPath());
    assertTrue(failure.byteOffset().isEmpty());
  }

  @Test
  void retainsACompletePlanInvariantAtItsExistingPath() {
    InvalidRequestException exception =
        new InvalidRequestException(
            new DuplicateStepId(new StepId("duplicate"), "steps[1].stepId"),
            java.util.Optional.of("steps[1].stepId"),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            null);

    RequestBindingFailure failure = RequestConstructorValidation.problem(exception, "", false);

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

    RequestBindingFailure failure = RequestConstructorValidation.problem(exception, "", false);

    assertEquals("steps", failure.jsonPath());
    assertEquals(exception.getMessage(), failure.exception().getMessage());
    assertTrue(failure.byteOffset().isEmpty());
  }

  @Test
  void normalizesAnUnstructuredCompletePlanFailureAtTheRequestRoot() {
    RequestBindingFailure failure =
        RequestConstructorValidation.problem(
            new IllegalArgumentException("complete plan failure"), "", false);

    assertEquals("request", failure.jsonPath());
    assertEquals("complete plan failure", failure.exception().getMessage());
  }

  @Test
  void repeatedFieldNamesRemainDistinctRelativeSegments() {
    var invalid =
        new InvalidRequestException(
            new MessageInvariant("invalid value", Optional.of("value")),
            Optional.of("value"),
            Optional.empty(),
            Optional.empty(),
            null);
    assertEquals(
        "steps[0].action.value.value",
        RequestConstructorValidation.problem(invalid, "steps[0].action.value", false).jsonPath());
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

  @Test
  void retainsFormulaCodesAndOwnedRelativeFieldPaths() {
    var normal =
        RequestConstructorValidation.problem(
            new InvalidFormulaInputException("formula must not begin with =", "value.text"),
            "steps[0].action.value",
            false);
    var raw =
        RequestConstructorValidation.problem(
            new InvalidRawFormulaTextException("forbidden XML character", "value.text"),
            "steps[1].action.value",
            false);
    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA,
        assertInstanceOf(FormulaRequestException.class, normal.exception()).problemCode());
    assertEquals("steps[0].action.value.value.text", normal.jsonPath());
    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA_TEXT,
        assertInstanceOf(FormulaRequestException.class, raw.exception()).problemCode());
    assertEquals("steps[1].action.value.value.text", raw.jsonPath());
  }

  @Test
  void rebasesTypedFailuresAtOneQualifiedNestedFieldWithoutChangingGenericFailures() {
    InvalidRequestException invariant =
        new InvalidRequestException(
            new MessageInvariant("name must not be blank", Optional.of("name")),
            Optional.of("name"),
            Optional.empty(),
            Optional.empty(),
            null);
    FormulaRequestException formula =
        new FormulaRequestException(
            GridGrindProblemCode.INVALID_FORMULA,
            "formula must not begin with =",
            Optional.of("formula"),
            Optional.empty(),
            Optional.empty(),
            null);

    assertEquals(
        "steps[1].action.name",
        new RequestBindingFailure(invariant, "name", Optional.empty())
            .rebasedAt("steps[1].action.name", Optional.of(12L))
            .jsonPath());
    assertEquals(
        "steps[1].action.formula",
        new RequestBindingFailure(formula, "formula", Optional.empty())
            .rebasedAt("steps[1].action.formula", Optional.empty())
            .jsonPath());
    IllegalArgumentException generic = new IllegalArgumentException("generic");
    assertEquals(
        generic,
        new RequestBindingFailure(generic, "value", Optional.empty())
            .rebasedAt("steps[1].action.value", Optional.empty())
            .exception());
  }
}
