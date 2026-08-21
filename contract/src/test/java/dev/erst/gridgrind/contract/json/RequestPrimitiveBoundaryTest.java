package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.catalog.GridGrindProtocolContractSupport;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Proves that every primitive request field is checked on the raw syntax tree before binding. */
class RequestPrimitiveBoundaryTest {
  @Test
  void everyPrimitiveRequestFieldRejectsExplicitNullBeforeTypedBinding() {
    List<GridGrindProtocolContractSupport.RequestPrimitiveField> fields =
        GridGrindProtocolContractSupport.requestPrimitiveFields();

    assertFalse(fields.isEmpty());
    for (GridGrindProtocolContractSupport.RequestPrimitiveField field : fields) {
      List<RequestStructuralProblem> problems = validateExplicitNull(field);
      String jsonPath = "request." + field.fieldName();

      assertTrue(
          problems.stream()
              .anyMatch(
                  problem ->
                      problem instanceof RequestExplicitNullField
                          && problem.jsonPath().equals(java.util.Optional.of(jsonPath))),
          () -> "Primitive request field accepted explicit null: " + field);
    }
  }

  @Test
  void everyRequiredPrimitiveRequestFieldRejectsOmissionBeforeTypedBinding() {
    List<GridGrindProtocolContractSupport.RequestPrimitiveField> requiredFields =
        GridGrindProtocolContractSupport.requestPrimitiveFields().stream()
            .filter(GridGrindProtocolContractSupport.RequestPrimitiveField::required)
            .toList();

    assertFalse(requiredFields.isEmpty());
    for (GridGrindProtocolContractSupport.RequestPrimitiveField field : requiredFields) {
      List<RequestStructuralProblem> problems = validateOmission(field);
      String jsonPath = "request." + field.fieldName();

      assertTrue(
          problems.stream()
              .anyMatch(
                  problem ->
                      problem instanceof RequestMissingRequiredField
                          && problem.jsonPath().equals(java.util.Optional.of(jsonPath))),
          () -> "Required primitive request field silently defaulted: " + field);
    }
  }

  @Test
  void everyOptionalPrimitiveRequestFieldPublishesItsEffectiveDefault() {
    List<GridGrindProtocolContractSupport.RequestPrimitiveField> optionalFields =
        GridGrindProtocolContractSupport.requestPrimitiveFields().stream()
            .filter(field -> !field.required())
            .toList();

    assertFalse(optionalFields.isEmpty());
    assertTrue(
        optionalFields.stream().allMatch(field -> field.defaultBoolean().isPresent()),
        () -> "Optional primitive request field has no declared default: " + optionalFields);
  }

  private static List<RequestStructuralProblem> validateExplicitNull(
      GridGrindProtocolContractSupport.RequestPrimitiveField field) {
    List<RequestStructuralProblem> problems = new ArrayList<>();
    RequestRecordValidator.validate(
        new RequestJsonObject(
            0, List.of(new RequestJsonMember(field.fieldName(), 0, new RequestJsonNull(1)))),
        field.recordType(),
        "request",
        0,
        problems);
    return problems;
  }

  private static List<RequestStructuralProblem> validateOmission(
      GridGrindProtocolContractSupport.RequestPrimitiveField field) {
    List<RequestStructuralProblem> problems = new ArrayList<>();
    RequestRecordValidator.validate(
        new RequestJsonObject(0, List.of()), field.recordType(), "request", 0, problems);
    return problems;
  }
}
