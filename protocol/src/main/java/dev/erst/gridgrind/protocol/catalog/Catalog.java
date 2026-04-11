package dev.erst.gridgrind.protocol.catalog;

import dev.erst.gridgrind.protocol.dto.GridGrindProtocolVersion;
import java.util.List;
import java.util.Objects;

/** JSON-serializable top-level catalog emitted by {@code --print-protocol-catalog}. */
public record Catalog(
    GridGrindProtocolVersion protocolVersion,
    String discriminatorField,
    TypeEntry requestType,
    List<TypeEntry> sourceTypes,
    List<TypeEntry> persistenceTypes,
    List<TypeEntry> operationTypes,
    List<TypeEntry> readTypes,
    List<NestedTypeGroup> nestedTypes,
    List<PlainTypeGroup> plainTypes) {
  public Catalog {
    protocolVersion =
        protocolVersion == null ? GridGrindProtocolVersion.current() : protocolVersion;
    discriminatorField =
        CatalogRecordValidation.requireNonBlank(discriminatorField, "discriminatorField");
    Objects.requireNonNull(requestType, "requestType must not be null");
    sourceTypes = CatalogRecordValidation.copyEntries(sourceTypes, "sourceTypes");
    persistenceTypes = CatalogRecordValidation.copyEntries(persistenceTypes, "persistenceTypes");
    operationTypes = CatalogRecordValidation.copyEntries(operationTypes, "operationTypes");
    readTypes = CatalogRecordValidation.copyEntries(readTypes, "readTypes");
    nestedTypes = CatalogRecordValidation.copyGroups(nestedTypes, "nestedTypes");
    plainTypes = CatalogRecordValidation.copyPlainGroups(plainTypes, "plainTypes");
  }
}
