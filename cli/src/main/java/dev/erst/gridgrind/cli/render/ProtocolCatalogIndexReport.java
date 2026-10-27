package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Objects;

/** Compact first-contact index for the catalog command default output. */
public record ProtocolCatalogIndexReport(
    ProtocolVersion protocolVersion,
    String discriminatorField,
    String requestTypeId,
    List<ProtocolCatalogGroupIndex> topLevelGroups,
    List<ProtocolCatalogGroupIndex> nestedTypeGroups,
    List<ProtocolCatalogGroupIndex> plainTypeGroups,
    List<ProtocolCatalogFieldMetadataKey> fieldMetadataKeys,
    List<ProtocolCatalogLookupNamespace> lookupNamespaces,
    List<dev.erst.gridgrind.core.limits.Limit> limits) {
  public ProtocolCatalogIndexReport {
    limits = List.copyOf(limits);
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    discriminatorField = TextValidation.requireNonBlank(discriminatorField, "discriminatorField");
    requestTypeId = TextValidation.requireNonBlank(requestTypeId, "requestTypeId");
    Objects.requireNonNull(topLevelGroups, "topLevelGroups must not be null");
    Objects.requireNonNull(nestedTypeGroups, "nestedTypeGroups must not be null");
    Objects.requireNonNull(plainTypeGroups, "plainTypeGroups must not be null");
    Objects.requireNonNull(fieldMetadataKeys, "fieldMetadataKeys must not be null");
    Objects.requireNonNull(lookupNamespaces, "lookupNamespaces must not be null");
    topLevelGroups =
        topLevelGroups.stream()
            .map(group -> Objects.requireNonNull(group, "topLevelGroups must not contain nulls"))
            .toList();
    nestedTypeGroups =
        nestedTypeGroups.stream()
            .map(group -> Objects.requireNonNull(group, "nestedTypeGroups must not contain nulls"))
            .toList();
    plainTypeGroups =
        plainTypeGroups.stream()
            .map(group -> Objects.requireNonNull(group, "plainTypeGroups must not contain nulls"))
            .toList();
    fieldMetadataKeys =
        fieldMetadataKeys.stream()
            .map(
                fieldMetadataKey ->
                    Objects.requireNonNull(
                        fieldMetadataKey, "fieldMetadataKeys must not contain nulls"))
            .toList();
    lookupNamespaces =
        lookupNamespaces.stream()
            .map(
                namespace ->
                    Objects.requireNonNull(namespace, "lookupNamespaces must not contain nulls"))
            .toList();
  }
}
