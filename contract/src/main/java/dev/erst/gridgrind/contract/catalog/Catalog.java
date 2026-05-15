package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.List;
import java.util.Objects;

/** JSON-serializable top-level catalog emitted by {@code --print-protocol-catalog}. */
public record Catalog(
    GridGrindProtocolVersion protocolVersion,
    String discriminatorField,
    TypeEntry requestType,
    List<TypeEntry> sourceTypes,
    List<TypeEntry> persistenceTypes,
    List<TypeEntry> stepTypes,
    List<TypeEntry> mutationActionTypes,
    List<TypeEntry> assertionTypes,
    List<TypeEntry> inspectionQueryTypes,
    List<NestedTypeGroup> nestedTypes,
    List<PlainTypeGroup> plainTypes) {
  public Catalog {
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    discriminatorField =
        CatalogRecordValidation.requireNonBlank(discriminatorField, "discriminatorField");
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
  }

  /**
   * Returns the ordered list of top-level type categories addressable by bare name via --lookup.
   */
  public List<TopLevelTypeGroup> topLevelGroups() {
    return List.of(
        new TopLevelTypeGroup("sourceTypes", sourceTypes),
        new TopLevelTypeGroup("persistenceTypes", persistenceTypes),
        new TopLevelTypeGroup("stepTypes", stepTypes),
        new TopLevelTypeGroup("mutationActionTypes", mutationActionTypes),
        new TopLevelTypeGroup("assertionTypes", assertionTypes),
        new TopLevelTypeGroup("inspectionQueryTypes", inspectionQueryTypes));
  }
}
