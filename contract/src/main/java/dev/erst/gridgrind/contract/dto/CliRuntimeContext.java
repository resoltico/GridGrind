package dev.erst.gridgrind.contract.dto;

/** Context for a CLI runtime failure before workbook execution begins. */
public record CliRuntimeContext() implements ProblemContext {
  @Override
  public String stage() {
    return "CLI_RUNTIME";
  }
}
