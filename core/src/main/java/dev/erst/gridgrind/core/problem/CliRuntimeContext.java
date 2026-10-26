package dev.erst.gridgrind.core.problem;

import dev.erst.gridgrind.core.metadata.Variant;

/** Context for a CLI runtime failure before workbook execution begins. */
@Variant(
    id = "CLI_RUNTIME",
    summary = "Context for a CLI runtime failure before workbook execution begins.")
public record CliRuntimeContext() implements ProblemContext {
  @Override
  public String stage() {
    return "CLI_RUNTIME";
  }
}
