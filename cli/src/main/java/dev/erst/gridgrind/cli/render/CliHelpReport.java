package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.core.discovery.DiscoveryValidation;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.util.Objects;

/** Structured help payload for one CLI help surface. */
public record CliHelpReport(
    ProtocolVersion protocolVersion,
    String topic,
    String version,
    String description,
    String documentRef,
    String containerImageRef,
    HelpSurfaceReport surface) {
  public CliHelpReport {
    protocolVersion = DiscoveryValidation.requireProtocolVersion(protocolVersion);
    topic = DiscoveryValidation.requireNonBlank(topic, "topic");
    version = DiscoveryValidation.requireNonBlank(version, "version");
    description = DiscoveryValidation.requireNonBlank(description, "description");
    documentRef = DiscoveryValidation.requireNonBlank(documentRef, "documentRef");
    containerImageRef = DiscoveryValidation.requireNonBlank(containerImageRef, "containerImageRef");
    Objects.requireNonNull(surface, "surface must not be null");
  }
}
