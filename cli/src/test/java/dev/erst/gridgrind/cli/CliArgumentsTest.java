package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

/** Focused parser tests for CLI command exclusivity and argument validation. */
class CliArgumentsTest {
  @Test
  void printProtocolCatalogSearchParsesIntoDedicatedCommand() {
    CliCommand.PrintProtocolCatalogSearch command =
        assertInstanceOf(
            CliCommand.PrintProtocolCatalogSearch.class,
            CliArguments.parse(new String[] {"--print-protocol-catalog", "--search", "sheet"}));

    assertEquals("sheet", command.searchQuery());
    org.junit.jupiter.api.Assertions.assertNull(command.responsePath());
  }

  @Test
  void printProtocolCatalogOperationParsesIntoDedicatedCommand() {
    CliCommand.PrintProtocolCatalogLookup command =
        assertInstanceOf(
            CliCommand.PrintProtocolCatalogLookup.class,
            CliArguments.parse(
                new String[] {"--print-protocol-catalog", "--operation", "SET_CELL"}));

    assertEquals("SET_CELL", command.operationFilter());
    org.junit.jupiter.api.Assertions.assertNull(command.responsePath());
  }

  @Test
  void printProtocolCatalogParsesResponsePath() {
    CliCommand.PrintProtocolCatalogSearch command =
        assertInstanceOf(
            CliCommand.PrintProtocolCatalogSearch.class,
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
  void printProtocolCatalogRejectsBlankSearchAndOperationValues() {
    CliArgumentsException blankSearch =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-protocol-catalog", "--search", ""}));
    assertEquals("--search", blankSearch.argument());
    assertEquals("search query must not be blank", blankSearch.getMessage());

    CliArgumentsException blankOperation =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-protocol-catalog", "--operation", ""}));
    assertEquals("--operation", blankOperation.argument());
    assertEquals("protocol catalog lookup id must not be blank", blankOperation.getMessage());
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

    org.junit.jupiter.api.Assertions.assertNull(command.responsePath());
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
  void immediateCommandsAcceptResponsePathBeforeTheCommandFlag() {
    CliCommand.Help helpCommand =
        assertInstanceOf(
            CliCommand.Help.class,
            CliArguments.parse(new String[] {"--response", "help.txt", "--help"}));
    assertEquals("help.txt", helpCommand.responsePath().toString());

    CliCommand.Version versionCommand =
        assertInstanceOf(
            CliCommand.Version.class,
            CliArguments.parse(new String[] {"--response", "version.json", "--version"}));
    assertEquals("version.json", versionCommand.responsePath().toString());

    CliCommand.PrintProtocolCatalogSearch searchCommand =
        assertInstanceOf(
            CliCommand.PrintProtocolCatalogSearch.class,
            CliArguments.parse(
                new String[] {
                  "--response", "catalog.json", "--print-protocol-catalog", "--search", "sheet"
                }));
    assertEquals("catalog.json", searchCommand.responsePath().toString());
    assertEquals("sheet", searchCommand.searchQuery());
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

    org.junit.jupiter.api.Assertions.assertNull(command.requestPath());
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
  void printTaskAndImmediateCatalogCommandsRejectBlankValues() {
    CliArgumentsException blankTask =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-task-catalog", "--task", ""}));
    assertEquals("--task", blankTask.argument());
    assertEquals("task id must not be blank", blankTask.getMessage());

    CliArgumentsException blankTaskPlan =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-task-plan", ""}));
    assertEquals("--print-task-plan", blankTaskPlan.argument());
    assertEquals("task id must not be blank", blankTaskPlan.getMessage());

    CliArgumentsException blankExample =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-example", ""}));
    assertEquals("--print-example", blankExample.argument());
    assertEquals("example id must not be blank", blankExample.getMessage());
  }

  @Test
  void dependentFlagsExplainTheirRequiredParentCommand() {
    CliArgumentsException taskException =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--task", "BUDGET"}));
    assertEquals("--task", taskException.argument());
    assertEquals(
        "--task requires --print-task-catalog and one task id value", taskException.getMessage());

    CliArgumentsException operationException =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--operation", "SET_CELL"}));
    assertEquals("--operation", operationException.argument());
    assertEquals(
        "--operation requires --print-protocol-catalog and one lookup id value",
        operationException.getMessage());

    CliArgumentsException searchException =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--search", "layout"}));
    assertEquals("--search", searchException.argument());
    assertEquals(
        "--search requires --print-protocol-catalog and one search text value",
        searchException.getMessage());
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
