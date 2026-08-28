package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.InvalidFormulaInputException;
import dev.erst.gridgrind.contract.dto.InvalidRawFormulaTextException;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.exc.MismatchedInputException;

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

  @Test
  void preservesFormulaSpecificFailuresAcrossFragmentAndCompletePlanNormalization() {
    RequestBindingFailure normal =
        RequestBindingFailure.from(
            new InvalidFormulaInputException("formula must not begin with ="),
            new RequestJsonString(0, "formula"),
            "steps[0].action.value",
            RequestDiagnosticRedactor.empty());
    RequestBindingFailure raw =
        RequestBindingFailure.fromCompletePlan(
            new InvalidRawFormulaTextException("raw formula contains a forbidden XML character"),
            RequestDiagnosticRedactor.empty());
    FormulaRequestException wrappedFormula =
        new FormulaRequestException(
            GridGrindProblemCode.INVALID_FORMULA,
            "formula must not begin with =",
            Optional.of("source.text"),
            Optional.empty(),
            Optional.empty(),
            null);
    RequestBindingFailure rebased =
        RequestBindingFailure.from(
            new IllegalArgumentException(wrappedFormula),
            new RequestJsonObject(0, List.of()),
            "steps[0].action.value",
            RequestDiagnosticRedactor.empty());
    MismatchedInputException jacksonCause =
        MismatchedInputException.from(null, Object.class, "formula binding failure");
    FormulaRequestException jacksonWrappedFormula =
        new FormulaRequestException(
            GridGrindProblemCode.INVALID_FORMULA,
            "formula must not begin with =",
            Optional.of("source.text"),
            Optional.empty(),
            Optional.empty(),
            jacksonCause);
    RequestBindingFailure rebasedWithJacksonLocation =
        RequestBindingFailure.from(
            new IllegalArgumentException(jacksonWrappedFormula),
            new RequestJsonObject(0, List.of()),
            "steps[0].action.value",
            RequestDiagnosticRedactor.empty());

    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA,
        assertInstanceOf(FormulaRequestException.class, normal.exception()).problemCode());
    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA_TEXT,
        assertInstanceOf(FormulaRequestException.class, raw.exception()).problemCode());
    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA,
        assertInstanceOf(FormulaRequestException.class, rebased.exception()).problemCode());
    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA,
        assertInstanceOf(FormulaRequestException.class, rebasedWithJacksonLocation.exception())
            .problemCode());
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

  @Test
  void rebasesUniqueNestedFormulaAndShapeFailuresAtTheExactAuthoredField() {
    RequestJsonObject fragment =
        new RequestJsonObject(
            0,
            List.of(
                new RequestJsonMember("formula", 1, new RequestJsonString(11, "SUM(A1)")),
                new RequestJsonMember("nested", 22, new RequestJsonArray(31, List.of()))));
    FormulaRequestException formula =
        new FormulaRequestException(
            GridGrindProblemCode.INVALID_FORMULA,
            "formula must not begin with =",
            Optional.of("formula"),
            Optional.empty(),
            Optional.empty(),
            null);
    InvalidRequestShapeException shape =
        new InvalidRequestShapeException(
            new MessageShape("formula must be a string", Optional.of("formula")),
            Optional.of("formula"),
            Optional.empty(),
            Optional.empty(),
            null);

    RequestBindingFailure formulaFailure =
        RequestBindingFailure.from(
            formula, fragment, "steps[2].action", RequestDiagnosticRedactor.empty());
    RequestBindingFailure inputFormulaFailure =
        RequestBindingFailure.from(
            new InvalidFormulaInputException("formula must not begin with ="),
            fragment,
            "steps[2].action",
            RequestDiagnosticRedactor.empty());
    RequestBindingFailure shapeFailure =
        RequestBindingFailure.from(
            shape, fragment, "steps[2].action", RequestDiagnosticRedactor.empty());

    assertEquals("steps[2].action.formula", formulaFailure.jsonPath());
    assertInstanceOf(FormulaRequestException.class, formulaFailure.exception());
    assertEquals("steps[2].action.formula", inputFormulaFailure.jsonPath());
    assertInstanceOf(FormulaRequestException.class, inputFormulaFailure.exception());
    assertEquals("steps[2].action.formula", shapeFailure.jsonPath());
    assertInstanceOf(InvalidRequestShapeException.class, shapeFailure.exception());
  }

  @Test
  void resolvesOneInvariantFieldNestedInsideAnArray() {
    RequestJsonNode fragment =
        new RequestJsonArray(
            0,
            List.of(
                new RequestJsonObject(
                    1,
                    List.of(new RequestJsonMember("name", 2, new RequestJsonString(9, "value"))))));

    assertEquals(
        Optional.of("[0].name"),
        RequestBindingPathSupport.directChildPathForInvariant(
            fragment, new IllegalArgumentException("name must not be blank")));
  }
}
