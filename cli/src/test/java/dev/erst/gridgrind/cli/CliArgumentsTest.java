package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

/** Focused parser tests for CLI command exclusivity and argument validation. */
class CliArgumentsTest {
  @Test
  void printProtocolCatalogSearchParsesIntoDedicatedCommand() {
    CliCommand.PrintProtocolCatalog command =
        assertInstanceOf(
            CliCommand.PrintProtocolCatalog.class,
            CliArguments.parse(new String[] {"--print-protocol-catalog", "--search", "sheet"}));

    assertNull(command.operationFilter());
    assertEquals("sheet", command.searchQuery());
    assertNull(command.responsePath());
  }

  @Test
  void printProtocolCatalogOperationParsesIntoDedicatedCommand() {
    CliCommand.PrintProtocolCatalog command =
        assertInstanceOf(
            CliCommand.PrintProtocolCatalog.class,
            CliArguments.parse(
                new String[] {"--print-protocol-catalog", "--operation", "SET_CELL"}));

    assertEquals("SET_CELL", command.operationFilter());
    assertNull(command.searchQuery());
    assertNull(command.responsePath());
  }

  @Test
  void printProtocolCatalogParsesResponsePath() {
    CliCommand.PrintProtocolCatalog command =
        assertInstanceOf(
            CliCommand.PrintProtocolCatalog.class,
            CliArguments.parse(
                new String[] {
                  "--print-protocol-catalog", "--search", "sheet", "--response", "catalog.json"
                }));

    assertEquals("sheet", command.searchQuery());
    assertEquals("catalog.json", command.responsePath().toString());
  }

  @Test
  void printProtocolCatalogRejectsDuplicateSearchArguments() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {
                      "--print-protocol-catalog", "--search", "sheet", "--search", "layout"
                    }));

    assertEquals("--search", exception.argument());
    assertEquals("Duplicate argument: --search", exception.getMessage());
  }

  @Test
  void printProtocolCatalogRejectsDuplicateOperationArguments() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {
                      "--print-protocol-catalog",
                      "--operation",
                      "SET_CELL",
                      "--operation",
                      "GET_CELL"
                    }));

    assertEquals("--operation", exception.argument());
    assertEquals("Duplicate argument: --operation", exception.getMessage());
  }

  @Test
  void printProtocolCatalogRejectsMissingSearchValue() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-protocol-catalog", "--search"}));

    assertEquals("--search", exception.argument());
    assertEquals("Missing value for --search", exception.getMessage());
  }

  @Test
  void printProtocolCatalogRejectsOperationAndSearchTogether() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {
                      "--print-protocol-catalog", "--operation", "SET_CELL", "--search", "cell"
                    }));

    assertEquals("--search", exception.argument());
    assertEquals(
        "--print-protocol-catalog does not allow both --operation and --search",
        exception.getMessage());
  }

  @Test
  void printProtocolCatalogRejectsUnexpectedTrailingArguments() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {"--print-protocol-catalog", "--search", "sheet", "--version"}));

    assertEquals("--version", exception.argument());
    assertEquals("Unknown argument: --version", exception.getMessage());
  }

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
  void bareHelpAliasParsesIntoHelpCommand() {
    CliCommand.Help command =
        assertInstanceOf(CliCommand.Help.class, CliArguments.parse(new String[] {"help"}));

    assertNull(command.responsePath());
  }

  @Test
  void helpAcceptsResponsePath() {
    CliCommand.Help command =
        assertInstanceOf(
            CliCommand.Help.class,
            CliArguments.parse(new String[] {"--help", "--response", "help.txt"}));

    assertEquals("help.txt", command.responsePath().toString());
  }

  @Test
  void doctorRequestRejectsUnknownTrailingArguments() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--doctor-request", "--goal", "budget report"}));

    assertEquals("--goal", exception.argument());
    assertEquals("Unknown argument: --goal", exception.getMessage());
  }

  @Test
  void doctorRequestParsesResponsePaths() {
    CliCommand.DoctorRequest command =
        assertInstanceOf(
            CliCommand.DoctorRequest.class,
            CliArguments.parse(
                new String[] {"--doctor-request", "--response", "doctor-report.json"}));

    assertNull(command.requestPath());
    assertEquals("doctor-report.json", command.responsePath().toString());
  }

  @Test
  void printTaskCatalogParsesTaskFilterAndResponsePathTogether() {
    CliCommand.PrintTaskCatalog command =
        assertInstanceOf(
            CliCommand.PrintTaskCatalog.class,
            CliArguments.parse(
                new String[] {
                  "--print-task-catalog", "--task", "DASHBOARD", "--response", "task.json"
                }));

    assertEquals("DASHBOARD", command.taskFilter());
    assertEquals("task.json", command.responsePath().toString());
  }

  @Test
  void printTaskCatalogRejectsDuplicateTaskArguments() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {
                      "--print-task-catalog", "--task", "DASHBOARD", "--task", "AUDIT"
                    }));

    assertEquals("--task", exception.argument());
    assertEquals("Duplicate argument: --task", exception.getMessage());
  }

  @Test
  void printTaskCatalogRejectsDuplicateResponseArguments() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {
                      "--print-task-catalog",
                      "--task",
                      "DASHBOARD",
                      "--response",
                      "a.json",
                      "--response",
                      "b.json"
                    }));

    assertEquals("--response", exception.argument());
    assertEquals("Duplicate argument: --response", exception.getMessage());
  }

  @Test
  void printProtocolCatalogRejectsDuplicateResponseArguments() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {
                      "--print-protocol-catalog",
                      "--search",
                      "sheet",
                      "--response",
                      "a.json",
                      "--response",
                      "b.json"
                    }));

    assertEquals("--response", exception.argument());
    assertEquals("Duplicate argument: --response", exception.getMessage());
  }

  @Test
  void immediateCommandsRejectDuplicateTrailingResponseArguments() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {"--help", "--response", "a.txt", "--response", "b.txt"}));

    assertEquals("--response", exception.argument());
    assertEquals("Duplicate argument: --response", exception.getMessage());
  }

  @Test
  void immediateCommandsRejectTrailingExecutionFlags() {
    CliArgumentsException versionFailure =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--version", "--request", "ignored.json"}));
    assertEquals("--request", versionFailure.argument());
    assertEquals("Unknown argument: --request", versionFailure.getMessage());

    CliArgumentsException taskFailure =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {"--print-task-plan", "DASHBOARD", "--request", "ignored.json"}));
    assertEquals("--request", taskFailure.argument());
    assertEquals("Unknown argument: --request", taskFailure.getMessage());
  }

  @Test
  void immediateCommandsMustAppearFirstWhenPresent() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--request", "req.json", "--version"}));

    assertEquals("--version", exception.argument());
    assertEquals("Unknown argument: --version", exception.getMessage());
  }
}
