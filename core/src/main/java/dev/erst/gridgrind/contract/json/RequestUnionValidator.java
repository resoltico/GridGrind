package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.catalog.GridGrindProtocolContractSupport;
import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadataSupport;
import dev.erst.gridgrind.contract.selector.SelectorJsonSupport;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Optional;
import java.util.Set;

/** Validates sealed and selector discriminators before validating the selected record shape. */
final class RequestUnionValidator {
  private RequestUnionValidator() {}

  static void validateUnion(
      RequestJsonNode node,
      Class<?> unionType,
      String jsonPath,
      long diagnosticByteOffset,
      List<RequestStructuralProblem> problems) {
    String discriminator =
        GridGrindProtocolContractSupport.discriminatorField(unionType)
            .orElseThrow(
                () ->
                    new IllegalStateException(
                        "Sealed request union must declare a discriminator: "
                            + unionType.getName()));
    validateDiscriminatedObject(
        node,
        jsonPath,
        diagnosticByteOffset,
        discriminator,
        Optional.of(unionType),
        typeId -> ProtocolTypeMetadataSupport.subtypeForTypeId(unionType, typeId),
        ProtocolTypeMetadataSupport.subtypesForType(unionType),
        problems);
  }

  static void validateSelector(
      RequestJsonNode node,
      String jsonPath,
      long diagnosticByteOffset,
      List<RequestStructuralProblem> problems) {
    validateDiscriminatedObject(
        node,
        jsonPath,
        diagnosticByteOffset,
        "type",
        Optional.empty(),
        typeId -> SelectorJsonSupport.typeFor(typeId).map(type -> type.asSubclass(Record.class)),
        SelectorJsonSupport.selectorTypes(),
        problems);
  }

  private static void validateDiscriminatedObject(
      RequestJsonNode node,
      String jsonPath,
      long diagnosticByteOffset,
      String discriminator,
      Optional<Class<?>> similarityRoot,
      java.util.function.Function<String, Optional<Class<? extends Record>>> subtypeForTypeId,
      List<Class<? extends Record>> possibleSubtypes,
      List<RequestStructuralProblem> problems) {
    if (!(node instanceof RequestJsonObject object)) {
      problems.add(new RequestMalformedScalar(jsonPath, "a JSON object", diagnosticByteOffset));
      return;
    }
    RequestObjectMembers.Index members = RequestObjectMembers.index(object);
    String typePath =
        RequestObjectMembers.childPath(
            jsonPath,
            GridGrindProtocolContractSupport.discriminatorContract(discriminator)
                .requiredFields()
                .getFirst());
    List<RequestJsonMember> typeMembers = members.membersNamed(discriminator);
    if (typeMembers.isEmpty()) {
      problems.add(new RequestMissingTypeDiscriminator(typePath));
      collectFieldsUnknownToEveryVariant(
          members, jsonPath, discriminator, possibleSubtypes, problems);
      return;
    }
    Set<Class<? extends Record>> viableSubtypes = new LinkedHashSet<>();
    for (RequestJsonMember typeMember : typeMembers) {
      validateDiscriminatorMember(typeMember, typePath, similarityRoot, subtypeForTypeId, problems)
          .ifPresent(viableSubtypes::add);
    }
    if (viableSubtypes.isEmpty()) {
      collectFieldsUnknownToEveryVariant(
          members, jsonPath, discriminator, possibleSubtypes, problems);
      return;
    }
    collectUnknownFieldsAcrossViableVariants(
        members, jsonPath, discriminator, viableSubtypes, problems);
    for (Class<? extends Record> recordType : viableSubtypes) {
      collectVariantProblems(members, recordType, jsonPath, discriminator, problems);
    }
  }

  private static void collectVariantProblems(
      RequestObjectMembers.Index object,
      Class<? extends Record> recordType,
      String jsonPath,
      String discriminator,
      List<RequestStructuralProblem> problems) {
    List<RequestStructuralProblem> variantProblems = new java.util.ArrayList<>();
    RequestRecordValidator.validateVariant(
        object,
        recordType,
        jsonPath,
        GridGrindProtocolContractSupport.effectiveObjectContract(recordType, discriminator),
        variantProblems);
    variantProblems.stream().filter(problem -> !problems.contains(problem)).forEach(problems::add);
  }

  private static void collectFieldsUnknownToEveryVariant(
      RequestObjectMembers.Index object,
      String jsonPath,
      String discriminator,
      List<Class<? extends Record>> possibleSubtypes,
      List<RequestStructuralProblem> problems) {
    RequestObjectMembers.collectUnknownFields(
        object, jsonPath, allowedFields(possibleSubtypes, discriminator), problems);
  }

  private static void collectUnknownFieldsAcrossViableVariants(
      RequestObjectMembers.Index object,
      String jsonPath,
      String discriminator,
      Set<Class<? extends Record>> viableSubtypes,
      List<RequestStructuralProblem> problems) {
    RequestObjectMembers.collectUnknownFields(
        object, jsonPath, fieldsAllowedByEveryVariant(viableSubtypes, discriminator), problems);
  }

  private static Set<String> allowedFields(
      Iterable<Class<? extends Record>> subtypes, String discriminator) {
    Set<String> allowedFields = new LinkedHashSet<>();
    for (Class<? extends Record> subtype : subtypes) {
      allowedFields.addAll(
          GridGrindProtocolContractSupport.effectiveObjectContract(subtype, discriminator)
              .fields());
    }
    return Set.copyOf(allowedFields);
  }

  private static Set<String> fieldsAllowedByEveryVariant(
      Set<Class<? extends Record>> subtypes, String discriminator) {
    java.util.Iterator<Class<? extends Record>> iterator = subtypes.iterator();
    Set<String> fields =
        new LinkedHashSet<>(
            GridGrindProtocolContractSupport.effectiveObjectContract(iterator.next(), discriminator)
                .fields());
    while (iterator.hasNext()) {
      fields.retainAll(
          GridGrindProtocolContractSupport.effectiveObjectContract(iterator.next(), discriminator)
              .fields());
    }
    return Set.copyOf(fields);
  }

  private static Optional<Class<? extends Record>> validateDiscriminatorMember(
      RequestJsonMember typeMember,
      String typePath,
      Optional<Class<?>> similarityRoot,
      java.util.function.Function<String, Optional<Class<? extends Record>>> subtypeForTypeId,
      List<RequestStructuralProblem> problems) {
    if (typeMember.value() instanceof RequestJsonNull) {
      problems.add(new RequestExplicitNullField(typePath, typeMember.nameByteOffset()));
      return Optional.empty();
    }
    if (!(typeMember.value() instanceof RequestJsonString typeValue)) {
      problems.add(
          new RequestMalformedScalar(
              typePath, "a JSON string type id", typeMember.nameByteOffset()));
      return Optional.empty();
    }
    Optional<Class<? extends Record>> subtype = subtypeForTypeId.apply(typeValue.value());
    if (subtype.isPresent()) {
      return subtype;
    }
    problems.add(
        new RequestUnknownTypeDiscriminator(
            typePath,
            typeValue.value(),
            similarityRoot
                .map(
                    root ->
                        GridGrindJsonSubtypeProblemSupport.similarTypeIds(root, typeValue.value()))
                .orElse(List.of()),
            GridGrindJsonSubtypeProblemSupport.specificGuidance(
                typePath, typeValue.value(), similarityRoot),
            typeMember.nameByteOffset()));
    return Optional.empty();
  }
}
