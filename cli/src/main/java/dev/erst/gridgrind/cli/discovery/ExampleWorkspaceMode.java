package dev.erst.gridgrind.cli.discovery;

/** Portability contract for one printed request emitted by the CLI. */
public enum ExampleWorkspaceMode {
  /** The printed request runs from a blank working directory without extra copied assets. */
  SELF_CONTAINED,

  /** The printed request expects copied asset paths beside the request file before execution. */
  REQUIRES_EXAMPLE_ASSETS
}
