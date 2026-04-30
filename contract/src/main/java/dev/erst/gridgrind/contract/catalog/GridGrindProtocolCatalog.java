package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.catalog.gather.CatalogDuplicateFailures;
import dev.erst.gridgrind.contract.catalog.gather.CatalogFieldMetadataSupport;
import dev.erst.gridgrind.contract.catalog.gather.CatalogGatherers;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.contract.step.WorkbookStepTargeting;
import java.lang.reflect.RecordComponent;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.function.Function;

/**
 * Publishes the machine-readable GridGrind protocol surface used by CLI discovery commands.
 *
 * <p>This registry intentionally spans the full public contract surface, so PMD's import-count
 * heuristic is not a useful coupling signal here.
 */
public final class GridGrindProtocolCatalog {
  private static final String DISCRIMINATOR_FIELD = "type";
  private static final WorkbookPlan REQUEST_TEMPLATE =
      new WorkbookPlan(
          GridGrindProtocolVersion.current(),
          Optional.empty(),
          new WorkbookPlan.WorkbookSource.New(),
          new WorkbookPlan.WorkbookPersistence.None(),
          dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
          dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
          List.of());
  private static final TypeEntry REQUEST_TYPE =
      typeEntry(
          WorkbookPlan.class,
          "WorkbookPlan",
          "Complete GridGrind plan for workbook source, execution policy, formula environment,"
              + " ordered mutation, assertion, and inspection steps, and persistence.",
          List.of(
              "protocolVersion",
              "planId",
              "persistence",
              "execution",
              "formulaEnvironment",
              "steps"));
  private static final CliSurface CLI_SURFACE = GridGrindProtocolCatalogCliSurface.CLI_SURFACE;
  private static final List<CatalogTypeDescriptor> STEP_TYPES =
      GridGrindProtocolCatalogTypeDescriptors.STEP_TYPES;
  private static final List<CatalogTypeDescriptor> SOURCE_TYPES =
      GridGrindProtocolCatalogTypeDescriptors.SOURCE_TYPES;
  private static final List<CatalogTypeDescriptor> PERSISTENCE_TYPES =
      GridGrindProtocolCatalogTypeDescriptors.PERSISTENCE_TYPES;
  private static final List<CatalogTypeDescriptor> MUTATION_ACTION_TYPES =
      GridGrindProtocolCatalogTypeDescriptors.MUTATION_ACTION_TYPES;
  private static final List<CatalogTypeDescriptor> ASSERTION_TYPES =
      GridGrindProtocolCatalogTypeDescriptors.ASSERTION_TYPES;
  private static final List<CatalogTypeDescriptor> INSPECTION_QUERY_TYPES =
      GridGrindProtocolCatalogTypeDescriptors.INSPECTION_QUERY_TYPES;
  private static final List<CatalogNestedTypeDescriptor> NESTED_TYPE_GROUPS =
      GridGrindProtocolCatalogFieldGroupSupport.NESTED_TYPE_GROUPS;
  private static final List<CatalogPlainTypeDescriptor> PLAIN_TYPE_DESCRIPTORS =
      GridGrindProtocolCatalogFieldGroupSupport.PLAIN_TYPE_DESCRIPTORS;

  private static final Catalog CATALOG = buildCatalog();

  private GridGrindProtocolCatalog() {}

  /** Returns the minimal successful request emitted by the CLI template command. */
  public static WorkbookPlan requestTemplate() {
    return REQUEST_TEMPLATE;
  }

  /** Returns the machine-readable protocol catalog emitted by the CLI discovery command. */
  public static Catalog catalog() {
    return CATALOG;
  }

  /** Returns the ordered built-in example set exposed by the CLI and catalog. */
  public static List<GridGrindShippedExamples.ShippedExample> shippedExamples() {
    return GridGrindShippedExamples.examples();
  }

  /** Returns one built-in example by its stable discovery id, or empty when unknown. */
  public static Optional<GridGrindShippedExamples.ShippedExample> exampleFor(String id) {
    return GridGrindShippedExamples.find(id);
  }

  /**
   * Returns the single catalog entry matching the given lookup token.
   *
   * <p>Unqualified lookups succeed only when exactly one entry with that id exists across the full
   * catalog surface. When an id repeats across groups, callers must qualify it as {@code
   * <group>:<id>}; ambiguous or unknown lookups return empty.
   */
  public static Optional<TypeEntry> entryFor(String idOrQualifiedId) {
    return GridGrindProtocolCatalogLookupSupport.entryFor(CATALOG, idOrQualifiedId);
  }

  /**
   * Returns the single catalog item matching the given lookup token.
   *
   * <p>Lookups may resolve either one concrete type entry such as {@code SET_CELL} or one nested /
   * plain type-group descriptor such as {@code cellInputTypes} or {@code plainTypes:
   * chartInputType}. Ambiguous or unknown lookups return empty.
   */
  public static Optional<Object> lookupValueFor(String idOrQualifiedId) {
    return GridGrindProtocolCatalogLookupSupport.lookupValueFor(CATALOG, idOrQualifiedId);
  }

  /**
   * Returns every catalog match for the given lookup token as stable qualified ids.
   *
   * <p>Unqualified duplicate ids expand to every matching {@code <group>:<id>} so callers can
   * surface deterministic disambiguation guidance.
   */
  public static List<String> matchingEntryIds(String idOrQualifiedId) {
    return GridGrindProtocolCatalogLookupSupport.matchingEntryIds(CATALOG, idOrQualifiedId);
  }

  /**
   * Returns every catalog match for the given lookup token as stable qualified ids.
   *
   * <p>This superset includes nested and plain type-group descriptors so CLI lookup can expose the
   * exact authoring groups operators need when the protocol catalog advertises them.
   */
  public static List<String> matchingLookupIds(String idOrQualifiedId) {
    return GridGrindProtocolCatalogLookupSupport.matchingLookupIds(CATALOG, idOrQualifiedId);
  }

  /** Performs case-insensitive discovery across ids, qualified ids, groups, and summaries. */
  public static CatalogSearchResult searchCatalog(String query) {
    return GridGrindProtocolCatalogLookupSupport.search(CATALOG, query);
  }

  private static Catalog buildCatalog() {
    validateFieldShapeGroupMappings();
    validateCoverage(WorkbookPlan.WorkbookSource.class, SOURCE_TYPES);
    validateCoverage(WorkbookPlan.WorkbookPersistence.class, PERSISTENCE_TYPES);
    validateCoverage(WorkbookStep.class, STEP_TYPES);
    validateCoverage(MutationAction.class, MUTATION_ACTION_TYPES);
    validateCoverage(Assertion.class, ASSERTION_TYPES);
    validateCoverage(InspectionQuery.class, INSPECTION_QUERY_TYPES);
    for (CatalogNestedTypeDescriptor nestedTypeGroup : NESTED_TYPE_GROUPS) {
      validateCoverage(nestedTypeGroup.sealedType(), nestedTypeGroup.typeDescriptors());
    }
    return new Catalog(
        GridGrindProtocolVersion.current(),
        DISCRIMINATOR_FIELD,
        REQUEST_TYPE,
        CLI_SURFACE,
        GridGrindShippedExamples.catalogEntries(),
        publicEntries(SOURCE_TYPES),
        publicEntries(PERSISTENCE_TYPES),
        publicEntries(STEP_TYPES),
        publicEntries(MUTATION_ACTION_TYPES),
        publicEntries(ASSERTION_TYPES),
        publicEntries(INSPECTION_QUERY_TYPES),
        NESTED_TYPE_GROUPS.stream().map(GridGrindProtocolCatalog::publicGroup).toList(),
        PLAIN_TYPE_DESCRIPTORS.stream().map(GridGrindProtocolCatalog::publicPlainGroup).toList());
  }

  private static NestedTypeGroup publicGroup(CatalogNestedTypeDescriptor descriptor) {
    return new NestedTypeGroup(
        descriptor.group(),
        descriptor.discriminatorField(),
        publicEntries(descriptor.typeDescriptors()));
  }

  private static PlainTypeGroup publicPlainGroup(CatalogPlainTypeDescriptor descriptor) {
    return new PlainTypeGroup(descriptor.group(), descriptor.typeEntry());
  }

  private static List<TypeEntry> publicEntries(List<CatalogTypeDescriptor> descriptors) {
    return descriptors.stream().map(CatalogTypeDescriptor::typeEntry).toList();
  }

  static CatalogNestedTypeDescriptor nestedTypeGroup(
      String group, Class<?> sealedType, List<CatalogTypeDescriptor> typeDescriptors) {
    return new CatalogNestedTypeDescriptor(
        group, discriminatorFieldFor(sealedType), sealedType, typeDescriptors);
  }

  static CatalogPlainTypeDescriptor plainTypeDescriptor(
      String group,
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> optionalFields) {
    return new CatalogPlainTypeDescriptor(
        group, recordType, typeEntry(recordType, id, summary, optionalFields));
  }

  static CatalogTypeDescriptor descriptor(
      Class<? extends Record> recordType, String id, String summary, String... optionalFields) {
    return new CatalogTypeDescriptor(
        recordType, typeEntry(recordType, id, summary, List.of(optionalFields)));
  }

  private static TypeEntry typeEntry(
      Class<? extends Record> recordType, String id, String summary, List<String> optionalFields) {
    Optional<WorkbookStepTargeting.TargetSurface> targetSurface =
        TypeEntryTargetingSupport.optionalTargetSurfaceFor(recordType);
    return new TypeEntry(
        canonicalTypeId(recordType, id),
        summary,
        fieldEntries(recordType, optionalFields),
        TypeEntryTargetingSupport.targetSelectorEntries(targetSurface),
        targetSurface.flatMap(WorkbookStepTargeting.TargetSurface::rule).orElse(null));
  }

  @SuppressWarnings("unchecked")
  private static String canonicalTypeId(Class<? extends Record> recordType, String suppliedId) {
    if (WorkbookPlan.WorkbookSource.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.workbookSourceTypeName(
              (Class<? extends WorkbookPlan.WorkbookSource>) recordType),
          recordType);
    }
    if (WorkbookPlan.WorkbookPersistence.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.workbookPersistenceTypeName(
              (Class<? extends WorkbookPlan.WorkbookPersistence>) recordType),
          recordType);
    }
    if (WorkbookStep.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.workbookStepTypeName(
              (Class<? extends WorkbookStep>) recordType),
          recordType);
    }
    if (MutationAction.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.mutationActionTypeName(
              (Class<? extends MutationAction>) recordType),
          recordType);
    }
    if (Assertion.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.assertionTypeName((Class<? extends Assertion>) recordType),
          recordType);
    }
    if (InspectionQuery.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.inspectionQueryTypeName(
              (Class<? extends InspectionQuery>) recordType),
          recordType);
    }
    return suppliedId;
  }

  static String requireMatchingCatalogId(
      String suppliedId, String canonicalId, Class<? extends Record> recordType) {
    if (!suppliedId.equals(canonicalId)) {
      throw new IllegalStateException(
          "Catalog type id mismatch for "
              + recordType.getName()
              + ": supplied="
              + suppliedId
              + ", canonical="
              + canonicalId);
    }
    return canonicalId;
  }

  private static List<FieldEntry> fieldEntries(
      Class<? extends Record> recordType, List<String> optionalFields) {
    requiredFields(recordType, optionalFields);
    Set<String> optionalFieldSet = Set.copyOf(optionalFields);
    return Arrays.stream(recordType.getRecordComponents())
        .filter(GridGrindProtocolCatalog::isCatalogVisible)
        .gather(CatalogGatherers.expandFieldsWithMetadata(optionalFieldSet))
        .toList();
  }

  private static void validateFieldShapeGroupMappings() {
    Set<Class<?>> descriptorNestedTypes =
        NESTED_TYPE_GROUPS.stream()
            .map(CatalogNestedTypeDescriptor::sealedType)
            .collect(java.util.stream.Collectors.toSet());
    Set<Class<?>> descriptorPlainTypes =
        PLAIN_TYPE_DESCRIPTORS.stream()
            .map(CatalogPlainTypeDescriptor::recordType)
            .collect(java.util.stream.Collectors.toSet());

    for (CatalogNestedTypeDescriptor descriptor : NESTED_TYPE_GROUPS) {
      CatalogFieldMetadataSupport.validateNestedTypeGroupMapping(
          descriptor.sealedType(), descriptor.group());
    }
    for (CatalogPlainTypeDescriptor descriptor : PLAIN_TYPE_DESCRIPTORS) {
      CatalogFieldMetadataSupport.validatePlainTypeGroupMapping(
          descriptor.recordType(), descriptor.group());
    }

    // Reverse check: every registered type must appear in a descriptor.
    validateReverseGroupMappings(descriptorNestedTypes, descriptorPlainTypes);
  }

  /**
   * Validates that every type registered in the field-shape maps appears in one of the provided
   * descriptor sets. Exposed as package-private so tests can exercise the failure paths with
   * synthetic descriptor sets that are intentionally incomplete.
   */
  static void validateReverseGroupMappings(
      Set<Class<?>> descriptorNestedTypes, Set<Class<?>> descriptorPlainTypes) {
    for (Class<?> registeredType : CatalogFieldMetadataSupport.registeredNestedTypes()) {
      if (!descriptorNestedTypes.contains(registeredType)) {
        throw new IllegalStateException(
            "Field-shape nested group map contains type with no catalog descriptor: "
                + registeredType.getName());
      }
    }
    for (Class<?> registeredType : CatalogFieldMetadataSupport.registeredPlainTypes()) {
      if (!descriptorPlainTypes.contains(registeredType)) {
        throw new IllegalStateException(
            "Field-shape plain group map contains type with no catalog descriptor: "
                + registeredType.getName());
      }
    }
  }

  /** Returns the required record fields after removing the explicitly optional ones. */
  static List<String> requiredFields(
      Class<? extends Record> recordType, List<String> optionalFields) {
    List<String> recordFields = recordFields(recordType);
    for (String optionalField : optionalFields) {
      if (!recordFields.contains(optionalField)) {
        throw new IllegalStateException(
            "Catalog optional field '%s' does not exist on %s"
                .formatted(optionalField, recordType.getName()));
      }
    }
    return recordFields.stream().filter(field -> !optionalFields.contains(field)).toList();
  }

  private static List<String> recordFields(Class<? extends Record> recordType) {
    return Arrays.stream(recordType.getRecordComponents())
        .filter(GridGrindProtocolCatalog::isCatalogVisible)
        .map(RecordComponent::getName)
        .toList();
  }

  private static boolean isCatalogVisible(RecordComponent component) {
    return !component.isAnnotationPresent(CatalogIgnored.class)
        && !component.getAccessor().isAnnotationPresent(CatalogIgnored.class)
        && !component.getAccessor().isAnnotationPresent(JsonIgnore.class);
  }

  private static void validateCoverage(
      Class<?> sealedType, List<CatalogTypeDescriptor> descriptors) {
    validateCoverage(
        sealedType,
        toOrderedMap(
            descriptors,
            CatalogTypeDescriptor::recordType,
            descriptor -> descriptor.typeEntry().id(),
            "catalog descriptor"));
  }

  /** Validates that a tagged union and the catalog expose the same ordered discriminator ids. */
  static void validateCoverage(Class<?> sealedType, Map<Class<?>, String> catalogIds) {
    if (sealedType.equals(WorkbookStep.class)) {
      Set<Class<?>> permitted =
          Arrays.stream(sealedType.getPermittedSubclasses())
              .collect(java.util.stream.Collectors.toCollection(java.util.LinkedHashSet::new));
      if (!permitted.equals(catalogIds.keySet())) {
        throw new IllegalStateException(
            "Catalog coverage mismatch for "
                + sealedType.getName()
                + ": permitted="
                + permitted
                + ", catalog="
                + catalogIds.keySet());
      }
      return;
    }
    discriminatorFieldFor(sealedType);
    JsonSubTypes jsonSubTypes = sealedType.getAnnotation(JsonSubTypes.class);
    if (jsonSubTypes == null) {
      throw new IllegalStateException(
          "Catalog coverage requires @JsonSubTypes on " + sealedType.getName());
    }

    Map<Class<?>, String> annotationIds =
        toOrderedMap(
            Arrays.asList(jsonSubTypes.value()),
            JsonSubTypes.Type::value,
            JsonSubTypes.Type::name,
            "annotation subtype");

    for (Class<?> recordType : catalogIds.keySet()) {
      if (!recordType.isRecord()) {
        throw new IllegalStateException(
            "Catalog entry %s does not target a record type".formatted(recordType));
      }
    }

    if (!annotationIds.keySet().equals(catalogIds.keySet())) {
      throw new IllegalStateException(
          "Catalog coverage mismatch for "
              + sealedType.getName()
              + ": annotated="
              + annotationIds.keySet()
              + ", catalog="
              + catalogIds.keySet());
    }

    for (Map.Entry<Class<?>, String> annotationEntry : annotationIds.entrySet()) {
      String catalogId = catalogIds.get(annotationEntry.getKey());
      if (!annotationEntry.getValue().equals(catalogId)) {
        throw new IllegalStateException(
            "Catalog id mismatch for "
                + annotationEntry.getKey().getName()
                + ": annotation="
                + annotationEntry.getValue()
                + ", catalog="
                + catalogId);
      }
    }
  }

  private static String discriminatorFieldFor(Class<?> sealedType) {
    JsonTypeInfo jsonTypeInfo = sealedType.getAnnotation(JsonTypeInfo.class);
    if (jsonTypeInfo == null || jsonTypeInfo.property().isBlank()) {
      throw new IllegalStateException(
          "Catalog coverage requires %s to declare a non-blank @JsonTypeInfo property"
              .formatted(sealedType.getName()));
    }
    return jsonTypeInfo.property();
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  static <T, K, V> Map<K, V> toOrderedMap(
      List<T> items, Function<T, K> keyFn, Function<T, V> valueFn, String label) {
    Map<K, V> result = new LinkedHashMap<>();
    for (T item : items) {
      K key = keyFn.apply(item);
      V value = valueFn.apply(item);
      if (result.containsKey(key)) {
        throw CatalogDuplicateFailures.duplicateEntryFailure(label, result.get(key), value);
      }
      result.put(key, value);
    }
    return result;
  }
}
