package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

/** Focused coverage for execution-parser unknown-argument handling. */
class CliExecutionCommandParserTest {
  @Test
  void unknownArgumentExceptionRejectsRemovedTaskPlanFlagAsUnknown() {
    CliArgumentsException failure =
        CliExecutionCommandParser.unknownArgumentException("--print-task-plan");

    assertEquals("--print-task-plan", failure.argument());
    assertEquals("Unknown argument: --print-task-plan", failure.getMessage());
  }
}
