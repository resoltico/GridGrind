package dev.erst.gridgrind.cli.discovery;

/** Structured workspace preparation advice published with every built-in recipe. */
public enum RecipeAdvisory {
  /** The printed request runs from a blank working directory without copied assets. */
  SELF_CONTAINED,

  /** The printed request requires the published workspace-relative assets before execution. */
  REQUIRES_EXAMPLE_ASSETS
}
