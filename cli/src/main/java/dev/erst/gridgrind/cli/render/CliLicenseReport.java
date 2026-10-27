package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.core.discovery.DiscoveryValidation;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;

/** Structured license payload for CLI discovery and distribution introspection. */
public record CliLicenseReport(
    ProtocolVersion protocolVersion, String version, String licenseText) {
  public CliLicenseReport {
    protocolVersion = DiscoveryValidation.requireProtocolVersion(protocolVersion);
    version = DiscoveryValidation.requireNonBlank(version, "version");
    licenseText = DiscoveryValidation.requireNonBlank(licenseText, "licenseText");
  }
}
