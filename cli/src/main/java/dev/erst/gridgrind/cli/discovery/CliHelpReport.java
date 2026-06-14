package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.cli.CliSurface;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.Objects;

/** Structured help payload for one CLI help surface. */
public record CliHelpReport(
    GridGrindProtocolVersion protocolVersion,
    String topic,
    String version,
    String description,
    String documentRef,
    String containerImageRef,
    CliSurface surface) {
  public CliHelpReport {
    protocolVersion = CliDiscoveryValidation.requireProtocolVersion(protocolVersion);
    topic = CliDiscoveryValidation.requireNonBlank(topic, "topic");
    version = CliDiscoveryValidation.requireNonBlank(version, "version");
    description = CliDiscoveryValidation.requireNonBlank(description, "description");
    documentRef = CliDiscoveryValidation.requireNonBlank(documentRef, "documentRef");
    containerImageRef =
        CliDiscoveryValidation.requireNonBlank(containerImageRef, "containerImageRef");
    Objects.requireNonNull(surface, "surface must not be null");
  }
}
