package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Objects;

/** JSON-serializable top-level protocol catalog. */
public record Catalog(
    ProtocolVersion protocolVersion,
    String discriminatorField,
    TypeEntry requestType,
    List<TypeEntry> sourceTypes,
    List<TypeEntry> persistenceTypes,
    List<TypeEntry> stepTypes,
    List<TypeEntry> mutationActionTypes,
    List<TypeEntry> assertionTypes,
    List<TypeEntry> inspectionQueryTypes,
    List<NestedTypeGroup> nestedTypes,
    List<PlainTypeGroup> plainTypes,
    List<CatalogNote> notes) {
  public Catalog {
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    discriminatorField = TextValidation.requireNonBlank(discriminatorField, "discriminatorField");
    Objects.requireNonNull(requestType, "requestType must not be null");
    sourceTypes = CatalogRecordValidation.copyEntries(sourceTypes, "sourceTypes");
    persistenceTypes = CatalogRecordValidation.copyEntries(persistenceTypes, "persistenceTypes");
    stepTypes = CatalogRecordValidation.copyEntries(stepTypes, "stepTypes");
    mutationActionTypes =
        CatalogRecordValidation.copyEntries(mutationActionTypes, "mutationActionTypes");
    assertionTypes = CatalogRecordValidation.copyEntries(assertionTypes, "assertionTypes");
    inspectionQueryTypes =
        CatalogRecordValidation.copyEntries(inspectionQueryTypes, "inspectionQueryTypes");
    nestedTypes = CatalogRecordValidation.copyGroups(nestedTypes, "nestedTypes");
    plainTypes = CatalogRecordValidation.copyPlainGroups(plainTypes, "plainTypes");
    notes = Objects.requireNonNullElseGet(notes, List::of);
    notes = CatalogRecordValidation.copyNotes(notes, "notes");
    CatalogNoteResolutionSupport.validateCatalogNoteRefs(
        List.of(
            List.of(requestType),
            sourceTypes,
            persistenceTypes,
            stepTypes,
            mutationActionTypes,
            assertionTypes,
            inspectionQueryTypes),
        nestedTypes,
        plainTypes,
        notes);
  }

  /**
   * Returns the ordered list of top-level type categories addressable by their published bare
   * names.
   */
  public List<TopLevelTypeGroup> topLevelGroups() {
    return List.of(
        new TopLevelTypeGroup("workbookSourceTypes", sourceTypes),
        new TopLevelTypeGroup("workbookPersistenceTypes", persistenceTypes),
        new TopLevelTypeGroup("workbookStepTypes", stepTypes),
        new TopLevelTypeGroup("mutationActionTypes", mutationActionTypes),
        new TopLevelTypeGroup("assertionTypes", assertionTypes),
        new TopLevelTypeGroup("inspectionQueryTypes", inspectionQueryTypes));
  }

  /** Returns the canonical operational and format limits published with this catalog. */
  public List<dev.erst.gridgrind.core.limits.Limit> limits() {
    return dev.erst.gridgrind.core.limits.Limit.all();
  }

  /** Returns one published shared note by its stable id, or empty when unknown. */
  public java.util.Optional<CatalogNote> note(String id) {
    Objects.requireNonNull(id, "id must not be null");
    return notes.stream().filter(note -> note.id().equals(id)).findFirst();
  }
}
