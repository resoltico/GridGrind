package dev.erst.gridgrind.cli.discovery;

/** Portability contract for one built-in example request emitted by the CLI. */
public enum ExampleWorkspaceMode {
  /** The printed request runs from a blank working directory without extra copied assets. */
  SELF_CONTAINED,

  /**
   * The printed request expects copied `examples/` assets beside the request file before execution.
   */
  REQUIRES_EXAMPLE_ASSETS
}
