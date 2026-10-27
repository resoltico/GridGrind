package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.core.discovery.DiscoveryValidation;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
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
    topic = TextValidation.requireNonBlank(topic, "topic");
    version = TextValidation.requireNonBlank(version, "version");
    description = TextValidation.requireNonBlank(description, "description");
    documentRef = TextValidation.requireNonBlank(documentRef, "documentRef");
    containerImageRef = TextValidation.requireNonBlank(containerImageRef, "containerImageRef");
    Objects.requireNonNull(surface, "surface must not be null");
  }
}
