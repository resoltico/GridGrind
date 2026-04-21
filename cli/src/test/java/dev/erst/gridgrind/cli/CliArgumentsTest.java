package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

/** Focused parser tests for doctor-request specific argument behavior. */
class CliArgumentsTest {
  @Test
  void printGoalPlanParsesIntoItsDedicatedCommand() {
    CliCommand.PrintGoalPlan command =
        assertInstanceOf(
            CliCommand.PrintGoalPlan.class,
            CliArguments.parse(
                new String[] {"--print-goal-plan", "monthly sales dashboard with charts"}));

    assertEquals("monthly sales dashboard with charts", command.goal());
  }

  @Test
  void doctorRequestIgnoresUnknownTrailingArguments() {
    CliCommand.DoctorRequest command =
        assertInstanceOf(
            CliCommand.DoctorRequest.class,
            CliArguments.parse(new String[] {"--doctor-request", "--goal", "budget report"}));

    assertNull(command.requestPath());
  }

  @Test
  void doctorRequestRejectsResponsePaths() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {"--doctor-request", "--response", "doctor-report.json"}));

    assertEquals("--response", exception.argument());
    assertEquals("--response is not supported with --doctor-request", exception.getMessage());
  }
}
