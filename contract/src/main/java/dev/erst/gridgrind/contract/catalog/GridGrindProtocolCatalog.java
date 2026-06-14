package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import java.util.Optional;

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
      CatalogTypeEntryFactory.typeEntry(
          WorkbookPlan.class,
          "WorkbookPlan",
          "Complete GridGrind plan for workbook source, execution policy, formula environment,"
              + " ordered mutation, assertion, and inspection steps, and persistence.",
          List.of("planId"));
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
    GridGrindProtocolCatalogCoverageValidator.validateFieldShapeGroupMappings(
        NESTED_TYPE_GROUPS, PLAIN_TYPE_DESCRIPTORS);
    for (CatalogTopLevelTypeDescriptorGroup group :
        GridGrindProtocolCatalogTypeDescriptors.TOP_LEVEL_GROUPS) {
      GridGrindProtocolCatalogCoverageValidator.validateCoverage(
          group.sealedType(), group.typeDescriptors());
    }
    for (CatalogNestedTypeDescriptor nestedTypeGroup : NESTED_TYPE_GROUPS) {
      GridGrindProtocolCatalogCoverageValidator.validateCoverage(
          nestedTypeGroup.sealedType(), nestedTypeGroup.typeDescriptors());
    }
    return CatalogStepTemplateSupport.attach(
        new Catalog(
            GridGrindProtocolVersion.current(),
            DISCRIMINATOR_FIELD,
            REQUEST_TYPE,
            publicEntries(SOURCE_TYPES),
            publicEntries(PERSISTENCE_TYPES),
            publicEntries(STEP_TYPES),
            publicEntries(MUTATION_ACTION_TYPES),
            publicEntries(ASSERTION_TYPES),
            publicEntries(INSPECTION_QUERY_TYPES),
            NESTED_TYPE_GROUPS.stream().map(GridGrindProtocolCatalog::publicGroup).toList(),
            PLAIN_TYPE_DESCRIPTORS.stream()
                .map(GridGrindProtocolCatalog::publicPlainGroup)
                .toList()));
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
}
