package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.core.discovery.DiscoveryValidation;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;

/** Structured version payload for packaged product discovery. */
public record CliVersionReport(
    ProtocolVersion protocolVersion,
    String version,
    String description,
    String documentRef,
    String containerImageRef) {
  public CliVersionReport {
    protocolVersion = DiscoveryValidation.requireProtocolVersion(protocolVersion);
    version = DiscoveryValidation.requireNonBlank(version, "version");
    description = DiscoveryValidation.requireNonBlank(description, "description");
    documentRef = DiscoveryValidation.requireNonBlank(documentRef, "documentRef");
    containerImageRef = DiscoveryValidation.requireNonBlank(containerImageRef, "containerImageRef");
  }
}
