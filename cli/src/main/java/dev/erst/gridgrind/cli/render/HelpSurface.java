package dev.erst.gridgrind.cli.render;

/** Owns the CLI/help surface published by the command-line transport. */
final class HelpSurface {
  static final HelpSurfaceReport CLI_SURFACE =
      new HelpSurfaceReport(
          HelpSynopsis.usage(),
          HelpSynopsis.workflows(),
          HelpSynopsis.execution(),
          HelpSynopsis.limits(),
          HelpProtocol.request(),
          HelpProtocol.fileWorkflow(),
          HelpGuidance.coordinateSystems(),
          HelpGuidance.minimalValidRequest(),
          HelpGuidance.stdinExample(),
          HelpGuidance.dockerExample(),
          HelpGuidance.discovery(),
          HelpGuidance.docs(),
          HelpProtocol.flags(),
          PathGuidance.standardInputRequiresRequestMessage());

  private HelpSurface() {}
}
