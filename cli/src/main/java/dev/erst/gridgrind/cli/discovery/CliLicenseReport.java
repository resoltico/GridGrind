package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;

/** Structured license payload for CLI discovery and distribution introspection. */
public record CliLicenseReport(
    ProtocolVersion protocolVersion, String version, String licenseText) {
  public CliLicenseReport {
    protocolVersion = CliDiscoveryValidation.requireProtocolVersion(protocolVersion);
    version = CliDiscoveryValidation.requireNonBlank(version, "version");
    licenseText = CliDiscoveryValidation.requireNonBlank(licenseText, "licenseText");
  }
}
