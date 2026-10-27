package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.core.discovery.DiscoveryValidation;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** Structured version payload for packaged product discovery. */
public record CliVersionReport(
    ProtocolVersion protocolVersion,
    String version,
    String description,
    String documentRef,
    String containerImageRef) {
  public CliVersionReport {
    protocolVersion = DiscoveryValidation.requireProtocolVersion(protocolVersion);
    version = TextValidation.requireNonBlank(version, "version");
    description = TextValidation.requireNonBlank(description, "description");
    documentRef = TextValidation.requireNonBlank(documentRef, "documentRef");
    containerImageRef = TextValidation.requireNonBlank(containerImageRef, "containerImageRef");
  }
}
