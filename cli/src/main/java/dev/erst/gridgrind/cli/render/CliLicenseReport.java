package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.core.discovery.DiscoveryValidation;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** Structured license payload for CLI discovery and distribution introspection. */
public record CliLicenseReport(
    ProtocolVersion protocolVersion, String version, String licenseText) {
  public CliLicenseReport {
    protocolVersion = DiscoveryValidation.requireProtocolVersion(protocolVersion);
    version = TextValidation.requireNonBlank(version, "version");
    licenseText = TextValidation.requireNonBlank(licenseText, "licenseText");
  }
}
