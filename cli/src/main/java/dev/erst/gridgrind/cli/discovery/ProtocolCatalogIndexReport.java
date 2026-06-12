package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.List;
import java.util.Objects;

/** Compact first-contact index for `--print-protocol-catalog` default output. */
public record ProtocolCatalogIndexReport(
    GridGrindProtocolVersion protocolVersion,
    String discriminatorField,
    String requestTypeId,
    List<ProtocolCatalogGroupIndex> topLevelGroups,
    List<ProtocolCatalogGroupIndex> nestedTypeGroups,
    List<ProtocolCatalogGroupIndex> plainTypeGroups,
    List<ProtocolCatalogLookupNamespace> lookupNamespaces) {
  public ProtocolCatalogIndexReport {
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    discriminatorField =
        CliDiscoveryValidation.requireNonBlank(discriminatorField, "discriminatorField");
    requestTypeId = CliDiscoveryValidation.requireNonBlank(requestTypeId, "requestTypeId");
    Objects.requireNonNull(topLevelGroups, "topLevelGroups must not be null");
    Objects.requireNonNull(nestedTypeGroups, "nestedTypeGroups must not be null");
    Objects.requireNonNull(plainTypeGroups, "plainTypeGroups must not be null");
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
    lookupNamespaces =
        lookupNamespaces.stream()
            .map(
                namespace ->
                    Objects.requireNonNull(namespace, "lookupNamespaces must not contain nulls"))
            .toList();
  }
}
