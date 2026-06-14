package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;

/** Structured version payload for packaged product discovery. */
public record CliVersionReport(
    GridGrindProtocolVersion protocolVersion,
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
