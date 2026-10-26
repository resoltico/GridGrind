package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;

/** Structured version payload for packaged product discovery. */
public record CliVersionReport(
    ProtocolVersion protocolVersion,
    String version,
    String description,
    String documentRef,
    String containerImageRef) {
  public CliVersionReport {
    protocolVersion = CliDiscoveryValidation.requireProtocolVersion(protocolVersion);
    version = CliDiscoveryValidation.requireNonBlank(version, "version");
    description = CliDiscoveryValidation.requireNonBlank(description, "description");
    documentRef = CliDiscoveryValidation.requireNonBlank(documentRef, "documentRef");
    containerImageRef =
        CliDiscoveryValidation.requireNonBlank(containerImageRef, "containerImageRef");
  }
}
