package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.catalog.GridGrindContractText;

/** Owns the CLI/help surface published by the command-line transport. */
final class GridGrindProtocolCatalogCliSurface {
  static final CliSurface CLI_SURFACE =
      new CliSurface(
          GridGrindCliSurfaceSynopsisSections.usage(),
          GridGrindCliSurfaceSynopsisSections.workflows(),
          GridGrindCliSurfaceSynopsisSections.execution(),
          GridGrindCliSurfaceSynopsisSections.limits(),
          GridGrindCliSurfaceRequestSections.request(),
          GridGrindCliSurfaceRequestSections.fileWorkflow(),
          GridGrindCliSurfaceGuidanceSections.coordinateSystems(),
          GridGrindCliSurfaceGuidanceSections.minimalValidRequest(),
          GridGrindCliSurfaceGuidanceSections.stdinExample(),
          GridGrindCliSurfaceGuidanceSections.dockerExample(),
          GridGrindCliSurfaceGuidanceSections.discovery(),
          GridGrindCliSurfaceGuidanceSections.docs(),
          GridGrindCliSurfaceRequestSections.flags(),
          GridGrindContractText.standardInputRequiresRequestMessage());

  private GridGrindProtocolCatalogCliSurface() {}
}
