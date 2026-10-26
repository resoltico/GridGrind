package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.plan.GridGrindProtocolVersion;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
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
          dev.erst.gridgrind.core.plan.ExecutionPolicySpec.defaults(),
          dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec.empty(),
          List.of());
  private static final TypeEntry REQUEST_TYPE =
      CatalogTypeEntryFactory.typeEntry(
          WorkbookPlan.class,
          "WorkbookPlan",
          "Complete GridGrind plan for workbook source, execution policy, formula environment,"
              + " ordered mutation, assertion, and inspection steps, and persistence.");
  private static final List<CatalogTypeDescriptor> STEP_TYPES = CatalogTypes.STEP_TYPES;
  private static final List<CatalogTypeDescriptor> SOURCE_TYPES = CatalogTypes.SOURCE_TYPES;
  private static final List<CatalogTypeDescriptor> PERSISTENCE_TYPES =
      CatalogTypes.PERSISTENCE_TYPES;
  private static final List<CatalogTypeDescriptor> MUTATION_ACTION_TYPES =
      CatalogTypes.MUTATION_ACTION_TYPES;
  private static final List<CatalogTypeDescriptor> ASSERTION_TYPES = CatalogTypes.ASSERTION_TYPES;
  private static final List<CatalogTypeDescriptor> INSPECTION_QUERY_TYPES =
      CatalogTypes.INSPECTION_QUERY_TYPES;
  private static final List<CatalogNestedTypeDescriptor> NESTED_TYPE_GROUPS =
      CatalogTypes.NESTED_TYPE_GROUPS;
  private static final List<CatalogPlainTypeDescriptor> PLAIN_TYPE_DESCRIPTORS =
      CatalogTypes.PLAIN_TYPE_DESCRIPTORS;

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
   * plain type-group descriptor such as {@code cellValueSpecTypes} or {@code plainTypes:
   * chartSpecType}. Ambiguous or unknown lookups return empty.
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

  /** Resolves the shared notes referenced by one previously resolved lookup value. */
  public static List<CatalogNote> referencedNotesForLookupValue(Object lookupValue) {
    return CatalogNoteResolutionSupport.referencedNotes(CATALOG, lookupValue);
  }

  private static Catalog buildCatalog() {
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
                .toList(),
            GridGrindProtocolCatalogNotes.notes()));
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
