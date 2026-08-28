package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.node.ArrayNode;
import tools.jackson.databind.node.JsonNodeFactory;
import tools.jackson.databind.node.ObjectNode;
import tools.jackson.databind.node.StringNode;

/** Synthesizes executable step templates from the published protocol catalog metadata. */
final class CatalogStepTemplateSupport {
  private static final JsonNodeFactory JSON = JsonNodeFactory.instance;

  private CatalogStepTemplateSupport() {}

  static Catalog attach(Catalog catalog) {
    Objects.requireNonNull(catalog, "catalog must not be null");
    return new Catalog(
        catalog.protocolVersion(),
        catalog.discriminatorField(),
        catalog.requestType(),
        catalog.sourceTypes(),
        catalog.persistenceTypes(),
        catalog.stepTypes(),
        enrich(catalog, catalog.mutationActionTypes(), MutationStep.class, "action"),
        enrich(catalog, catalog.assertionTypes(), AssertionStep.class, "assertion"),
        enrich(catalog, catalog.inspectionQueryTypes(), InspectionStep.class, "query"),
        catalog.nestedTypes(),
        catalog.plainTypes(),
        catalog.notes());
  }

  private static List<TypeEntry> enrich(
      Catalog catalog, List<TypeEntry> entries, Class<?> stepType, String bodyField) {
    String stepKind =
        GridGrindProtocolTypeNames.workbookStepTypeName(
            stepType.asSubclass(dev.erst.gridgrind.contract.step.WorkbookStep.class));
    return entries.stream()
        .map(
            entry ->
                new TypeEntry(
                    entry.id(),
                    entry.summary(),
                    entry.fields(),
                    entry.targetSelectors(),
                    entry.targetSelectorRule(),
                    entry.noteRefs(),
                    entry.preconditions(),
                    Optional.of(buildTemplate(catalog, entry, stepKind, bodyField))))
        .toList();
  }

  private static ProtocolStepTemplate buildTemplate(
      Catalog catalog, TypeEntry entry, String stepKind, String bodyField) {
    List<String> notes = new ArrayList<>();
    ObjectNode step = JSON.objectNode();
    step.put("stepId", "step-" + entry.id().toLowerCase(java.util.Locale.ROOT).replace('_', '-'));
    step.set(
        bodyField,
        typeTemplate(catalog, entry, TypePlacement.topLevel(), new LinkedHashSet<>(), notes));
    step.set("target", targetTemplate(catalog, entry, notes));
    return new ProtocolStepTemplate(stepKind, bodyField, step, List.copyOf(notes));
  }

  private static JsonNode targetTemplate(Catalog catalog, TypeEntry entry, List<String> notes) {
    Objects.requireNonNull(catalog, "catalog must not be null");
    Objects.requireNonNull(entry, "entry must not be null");
    Objects.requireNonNull(notes, "notes must not be null");
    List<String> targetIds =
        entry.targetSelectors().stream()
            .flatMap(selector -> selector.typeIds().stream())
            .distinct()
            .toList();
    if (!targetIds.isEmpty()) {
      TypeEntry targetEntry =
          targetIds.stream()
              .map(targetId -> findTypeEntryById(catalog, targetId))
              .flatMap(Optional::stream)
              .min(
                  java.util.Comparator.comparingInt(
                      target -> CatalogStepTemplateDefaults.selectorPreference(target.id())))
              .orElseThrow(
                  () ->
                      new IllegalStateException(
                          "Catalog target selector ids did not resolve for " + entry.id()));
      if (targetIds.size() > 1) {
        notes.add(
            "Replace the placeholder target when needed; "
                + entry.id()
                + " accepts "
                + String.join(", ", targetIds)
                + ".");
      }
      return typeTemplate(
          catalog, targetEntry, TypePlacement.topLevel(), new LinkedHashSet<>(), notes);
    }
    String targetRule =
        entry
            .targetSelectorRule()
            .orElse("replace the placeholder target with one that matches this operation");
    notes.add("Target rule: " + targetRule);
    return typeTemplate(
        catalog,
        findTypeEntryById(catalog, "WORKBOOK_CURRENT")
            .orElseThrow(
                () ->
                    new IllegalStateException(
                        "Catalog is missing WORKBOOK_CURRENT for dynamic target placeholder")),
        TypePlacement.topLevel(),
        new LinkedHashSet<>(),
        notes);
  }

  private static JsonNode typeTemplate(
      Catalog catalog,
      TypeEntry entry,
      TypePlacement typePlacement,
      Set<String> recursionGuard,
      List<String> notes) {
    Objects.requireNonNull(catalog, "catalog must not be null");
    Objects.requireNonNull(entry, "entry must not be null");
    Objects.requireNonNull(typePlacement, "typePlacement must not be null");
    Objects.requireNonNull(recursionGuard, "recursionGuard must not be null");
    Objects.requireNonNull(notes, "notes must not be null");
    if (!recursionGuard.add(entry.id())) {
      notes.add(
          "Replace the recursive placeholder for "
              + entry.id()
              + " with one concrete non-recursive variant.");
      return recursivePlaceholder(entry, typePlacement);
    }
    ObjectNode object = JSON.objectNode();
    typePlacement.applyDiscriminator(object, entry.id());
    for (FieldEntry field : entry.fields()) {
      if (field.requirement() != FieldRequirement.REQUIRED) {
        continue;
      }
      object.set(field.name(), fieldValue(catalog, field, recursionGuard, notes));
    }
    applyTypeSpecificDefaults(catalog, entry.id(), object, recursionGuard, notes);
    recursionGuard.remove(entry.id());
    return object;
  }

  private static JsonNode fieldValue(
      Catalog catalog, FieldEntry field, Set<String> recursionGuard, List<String> notes) {
    return switch (field.shape()) {
      case FieldShape.Scalar scalar -> scalarValue(field, scalar.scalarType());
      case FieldShape.ListShape listShape -> {
        ArrayNode array = JSON.arrayNode();
        array.add(listElementValue(catalog, field, listShape, recursionGuard, notes));
        yield array;
      }
      default -> shapeValue(catalog, field.name(), field.shape(), recursionGuard, notes);
    };
  }

  private static JsonNode listElementValue(
      Catalog catalog,
      FieldEntry field,
      FieldShape.ListShape listShape,
      Set<String> recursionGuard,
      List<String> notes) {
    Objects.requireNonNull(catalog, "catalog must not be null");
    Objects.requireNonNull(field, "field must not be null");
    Objects.requireNonNull(listShape, "listShape must not be null");
    Objects.requireNonNull(recursionGuard, "recursionGuard must not be null");
    Objects.requireNonNull(notes, "notes must not be null");
    if (listShape.elementShape() instanceof FieldShape.Scalar scalar
        && !field.enumValues().isEmpty()) {
      return scalarValue(
          new FieldEntry(field.name(), FieldRequirement.REQUIRED, scalar, field.enumValues()),
          scalar.scalarType());
    }
    return shapeValue(catalog, field.name(), listShape.elementShape(), recursionGuard, notes);
  }

  private static JsonNode shapeValue(
      Catalog catalog,
      String fieldName,
      FieldShape shape,
      Set<String> recursionGuard,
      List<String> notes) {
    return switch (shape) {
      case FieldShape.Scalar scalar ->
          scalarValue(
              new FieldEntry(fieldName, FieldRequirement.REQUIRED, shape, List.of()),
              scalar.scalarType());
      case FieldShape.ListShape listShape -> {
        ArrayNode array = JSON.arrayNode();
        array.add(shapeValue(catalog, fieldName, listShape.elementShape(), recursionGuard, notes));
        yield array;
      }
      case FieldShape.TopLevelTypeSetRef topLevel ->
          topLevelTemplate(
              catalog,
              chooseTopLevel(catalog, topLevel.typeSet(), recursionGuard),
              recursionGuard,
              notes);
      case FieldShape.NestedTypeGroupRef nested ->
          nestedGroupTemplate(catalog, nested.group(), recursionGuard, notes);
      case FieldShape.NestedTypeGroupUnionRef union ->
          nestedUnionTemplate(catalog, union.groups(), recursionGuard, notes);
      case FieldShape.PlainTypeGroupRef plain ->
          plainGroupTemplate(catalog, plain.group(), recursionGuard, notes);
    };
  }

  private static JsonNode topLevelTemplate(
      Catalog catalog, TypeEntry entry, Set<String> recursionGuard, List<String> notes) {
    return typeTemplate(catalog, entry, TypePlacement.topLevel(), recursionGuard, notes);
  }

  static JsonNode nestedGroupTemplate(
      Catalog catalog, String groupName, Set<String> recursionGuard, List<String> notes) {
    NestedTypeGroup group = nestedGroup(catalog, groupName);
    return typeTemplate(
        catalog,
        chooseEntry(group.types(), recursionGuard, groupName),
        TypePlacement.nested(group.discriminatorField()),
        recursionGuard,
        notes);
  }

  private static JsonNode nestedUnionTemplate(
      Catalog catalog, List<String> groupNames, Set<String> recursionGuard, List<String> notes) {
    for (String groupName : groupNames) {
      NestedTypeGroup group = nestedGroup(catalog, groupName);
      List<TypeEntry> candidates =
          group.types().stream()
              .filter(candidate -> !recursionGuard.contains(candidate.id()))
              .toList();
      if (!candidates.isEmpty()) {
        return typeTemplate(
            catalog,
            chooseEntry(candidates, recursionGuard, groupName),
            TypePlacement.nested(group.discriminatorField()),
            recursionGuard,
            notes);
      }
    }
    NestedTypeGroup fallbackGroup = nestedGroup(catalog, groupNames.getFirst());
    return typeTemplate(
        catalog,
        chooseEntry(fallbackGroup.types(), recursionGuard, fallbackGroup.group()),
        TypePlacement.nested(fallbackGroup.discriminatorField()),
        recursionGuard,
        notes);
  }

  private static JsonNode plainGroupTemplate(
      Catalog catalog, String groupName, Set<String> recursionGuard, List<String> notes) {
    return typeTemplate(
        catalog,
        plainGroup(catalog, groupName).type(),
        TypePlacement.plain(),
        recursionGuard,
        notes);
  }

  private static JsonNode scalarValue(FieldEntry field, ScalarType scalarType) {
    if (!field.enumValues().isEmpty()) {
      return StringNode.valueOf(field.enumValues().getFirst());
    }
    return switch (scalarType) {
      case STRING ->
          StringNode.valueOf(CatalogStepTemplateDefaults.stringPlaceholder(field.name()));
      case NUMBER -> JSON.numberNode(CatalogStepTemplateDefaults.numberPlaceholder(field.name()));
      case BOOLEAN ->
          JSON.booleanNode(CatalogStepTemplateDefaults.booleanPlaceholder(field.name()));
    };
  }

  private static TypeEntry chooseTopLevel(
      Catalog catalog, String typeSet, Set<String> recursionGuard) {
    return chooseEntry(topLevelEntries(catalog, typeSet), recursionGuard, typeSet);
  }

  private static TypeEntry chooseEntry(
      List<TypeEntry> entries, Set<String> recursionGuard, String groupName) {
    return entries.stream()
        .sorted(
            java.util.Comparator.comparingInt(
                candidate ->
                    CatalogStepTemplateDefaults.entryPreference(groupName, candidate.id())))
        .filter(candidate -> !recursionGuard.contains(candidate.id()))
        .findFirst()
        .orElseGet(
            () ->
                entries.stream()
                    .sorted(
                        java.util.Comparator.comparingInt(
                            candidate ->
                                CatalogStepTemplateDefaults.entryPreference(
                                    groupName, candidate.id())))
                    .findFirst()
                    .orElseThrow(
                        () ->
                            new IllegalStateException(
                                "Catalog group " + groupName + " did not expose any entries")));
  }

  private static List<TypeEntry> topLevelEntries(Catalog catalog, String typeSet) {
    if ("sourceTypes".equals(typeSet)) {
      return catalog.sourceTypes();
    }
    if ("persistenceTypes".equals(typeSet)) {
      return catalog.persistenceTypes();
    }
    if ("stepTypes".equals(typeSet)) {
      return catalog.stepTypes();
    }
    if ("mutationActionTypes".equals(typeSet)) {
      return catalog.mutationActionTypes();
    }
    if ("assertionTypes".equals(typeSet)) {
      return catalog.assertionTypes();
    }
    if ("inspectionQueryTypes".equals(typeSet)) {
      return catalog.inspectionQueryTypes();
    }
    throw new IllegalArgumentException("Unsupported top-level type set: " + typeSet);
  }

  private static NestedTypeGroup nestedGroup(Catalog catalog, String group) {
    return catalog.nestedTypes().stream()
        .filter(candidate -> candidate.group().equals(group))
        .findFirst()
        .orElseThrow(() -> new IllegalArgumentException("Unknown nested type group: " + group));
  }

  private static PlainTypeGroup plainGroup(Catalog catalog, String group) {
    return catalog.plainTypes().stream()
        .filter(candidate -> candidate.group().equals(group))
        .findFirst()
        .orElseThrow(() -> new IllegalArgumentException("Unknown plain type group: " + group));
  }

  private static Optional<TypeEntry> findTypeEntryById(Catalog catalog, String typeId) {
    return allEntries(catalog).stream().filter(entry -> entry.id().equals(typeId)).findFirst();
  }

  private static List<TypeEntry> allEntries(Catalog catalog) {
    List<TypeEntry> entries = new ArrayList<>();
    entries.add(catalog.requestType());
    entries.addAll(catalog.sourceTypes());
    entries.addAll(catalog.persistenceTypes());
    entries.addAll(catalog.stepTypes());
    entries.addAll(catalog.mutationActionTypes());
    entries.addAll(catalog.assertionTypes());
    entries.addAll(catalog.inspectionQueryTypes());
    catalog.nestedTypes().forEach(group -> entries.addAll(group.types()));
    catalog.plainTypes().forEach(group -> entries.add(group.type()));
    return List.copyOf(entries);
  }

  private static JsonNode recursivePlaceholder(TypeEntry entry, TypePlacement typePlacement) {
    ObjectNode placeholder = JSON.objectNode();
    typePlacement.applyDiscriminator(placeholder, entry.id());
    return placeholder;
  }

  private static void applyTypeSpecificDefaults(
      Catalog catalog,
      String typeId,
      ObjectNode object,
      Set<String> recursionGuard,
      List<String> notes) {
    Objects.requireNonNull(catalog, "catalog must not be null");
    Objects.requireNonNull(typeId, "typeId must not be null");
    Objects.requireNonNull(object, "object must not be null");
    Objects.requireNonNull(recursionGuard, "recursionGuard must not be null");
    Objects.requireNonNull(notes, "notes must not be null");
    CatalogStepTemplateDefaults.applyTypeSpecificDefaults(
        catalog, typeId, object, recursionGuard, notes);
  }

  static JsonNode typeTemplateById(
      Catalog catalog, String typeId, Set<String> recursionGuard, List<String> notes) {
    return typeTemplate(
        catalog,
        findTypeEntryById(catalog, typeId)
            .orElseThrow(() -> new IllegalStateException("Catalog is missing type " + typeId)),
        TypePlacement.topLevel(),
        recursionGuard,
        notes);
  }

  private record TypePlacement(Optional<String> discriminatorField) {
    private TypePlacement {
      Objects.requireNonNull(discriminatorField, "discriminatorField must not be null");
    }

    static TypePlacement topLevel() {
      return new TypePlacement(Optional.of("type"));
    }

    static TypePlacement nested(String discriminatorField) {
      return new TypePlacement(Optional.of(discriminatorField));
    }

    static TypePlacement plain() {
      return new TypePlacement(Optional.empty());
    }

    void applyDiscriminator(ObjectNode object, String typeId) {
      Objects.requireNonNull(object, "object must not be null");
      Objects.requireNonNull(typeId, "typeId must not be null");
      discriminatorField.ifPresent(field -> object.put(field, typeId));
    }
  }
}
