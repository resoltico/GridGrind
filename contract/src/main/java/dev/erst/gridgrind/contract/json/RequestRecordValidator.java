package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.catalog.GridGrindProtocolContractSupport;
import java.lang.reflect.RecordComponent;
import java.util.List;

/** Validates one record creator contract, including all visible component values. */
final class RequestRecordValidator {
  private RequestRecordValidator() {}

  static void validate(
      RequestJsonNode node,
      Class<? extends Record> recordType,
      String jsonPath,
      long diagnosticByteOffset,
      List<RequestStructuralProblem> problems) {
    validate(
        node,
        recordType,
        jsonPath,
        GridGrindProtocolContractSupport.effectiveObjectContract(recordType),
        diagnosticByteOffset,
        problems);
  }

  static void validate(
      RequestJsonNode node,
      Class<? extends Record> recordType,
      String jsonPath,
      GridGrindProtocolContractSupport.EffectiveObjectContract contract,
      long diagnosticByteOffset,
      List<RequestStructuralProblem> problems) {
    if (!(node instanceof RequestJsonObject object)) {
      problems.add(new RequestMalformedScalar(jsonPath, "a JSON object", diagnosticByteOffset));
      return;
    }
    validate(
        RequestObjectMembers.index(object),
        recordType,
        jsonPath,
        contract,
        diagnosticByteOffset,
        problems);
  }

  static void validate(
      RequestObjectMembers.Index object,
      Class<? extends Record> recordType,
      String jsonPath,
      GridGrindProtocolContractSupport.EffectiveObjectContract contract,
      long diagnosticByteOffset,
      List<RequestStructuralProblem> problems) {
    RequestObjectMembers.collectFields(
        object, jsonPath, contract.requiredFields(), contract.optionalFields(), problems);
    validateRecordComponents(object, recordType, jsonPath, problems);
  }

  static void validateVariant(
      RequestObjectMembers.Index object,
      Class<? extends Record> recordType,
      String jsonPath,
      GridGrindProtocolContractSupport.EffectiveObjectContract contract,
      List<RequestStructuralProblem> problems) {
    RequestObjectMembers.collectMissingRequiredFields(
        object, jsonPath, contract.requiredFields(), problems);
    validateRecordComponents(object, recordType, jsonPath, problems);
  }

  private static void validateRecordComponents(
      RequestObjectMembers.Index object,
      Class<? extends Record> recordType,
      String jsonPath,
      List<RequestStructuralProblem> problems) {
    List<RecordComponent> components = RequestObjectMembers.visibleRecordComponents(recordType);
    for (RecordComponent component : components) {
      String fieldName = GridGrindProtocolContractSupport.wireFieldName(component);
      for (RequestJsonMember member : object.membersNamed(fieldName)) {
        RequestNodeValidator.validateNode(
            member.value(),
            component.getGenericType(),
            RequestObjectMembers.childPath(jsonPath, fieldName),
            member.nameByteOffset(),
            problems);
      }
    }
  }
}
