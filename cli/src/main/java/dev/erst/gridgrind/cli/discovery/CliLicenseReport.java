package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;

/** Structured license payload for CLI discovery and distribution introspection. */
public record CliLicenseReport(
    GridGrindProtocolVersion protocolVersion, String version, String licenseText) {
  public CliLicenseReport {
    protocolVersion = CliDiscoveryValidation.requireProtocolVersion(protocolVersion);
    version = CliDiscoveryValidation.requireNonBlank(version, "version");
    licenseText = CliDiscoveryValidation.requireNonBlank(licenseText, "licenseText");
  }
}
