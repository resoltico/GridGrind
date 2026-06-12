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
    assertEquals(java.util.Optional.empty(), command.responsePath());
  }

  @Test
  void printProtocolCatalogLookupParsesIntoDedicatedCommand() {
    CliCommand.PrintProtocolCatalogLookup command =
        assertInstanceOf(
            CliCommand.PrintProtocolCatalogLookup.class,
            CliArguments.parse(new String[] {"--print-protocol-catalog", "--lookup", "SET_CELL"}));

    assertEquals("SET_CELL", command.lookupId());
    assertEquals(java.util.Optional.empty(), command.responsePath());
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
    assertEquals(
        java.util.Optional.of(java.nio.file.Path.of("catalog.json")), command.responsePath());
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
  void printProtocolCatalogRejectsDuplicateLookupArguments() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {
                      "--print-protocol-catalog", "--lookup", "SET_CELL", "--lookup", "GET_CELL"
                    }));

    assertEquals("--lookup", exception.argument());
    assertEquals("Duplicate argument: --lookup", exception.getMessage());
  }

  @Test
  void printProtocolCatalogRejectsMissingAndBlankValues() {
    CliArgumentsException missingSearch =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-protocol-catalog", "--search"}));
    assertEquals("--search", missingSearch.argument());
    assertEquals("Missing value for --search", missingSearch.getMessage());

    CliArgumentsException blankSearch =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-protocol-catalog", "--search", ""}));
    assertEquals("--search", blankSearch.argument());
    assertEquals("search query must not be blank", blankSearch.getMessage());

    CliArgumentsException blankLookup =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-protocol-catalog", "--lookup", ""}));
    assertEquals("--lookup", blankLookup.argument());
    assertEquals("protocol catalog lookup id must not be blank", blankLookup.getMessage());
  }

  @Test
  void printProtocolCatalogRejectsLookupAndSearchTogether() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {
                      "--print-protocol-catalog", "--lookup", "SET_CELL", "--search", "cell"
                    }));

    assertEquals("--search", exception.argument());
    assertEquals(
        "--print-protocol-catalog does not allow both --lookup and --search",
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
    assertEquals(
        "Only one primary command may be used per invocation; --print-protocol-catalog cannot"
            + " be combined with --version",
        exception.getMessage());
  }

  @Test
  void immediateCommandsRejectCompetingPrimaryCommandsWithExplicitConflictMessage() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {
                      "--print-example", "--lookup", "BUDGET", "--print-task-catalog"
                    }));

    assertEquals("--print-task-catalog", exception.argument());
    assertEquals(
        "Only one primary command may be used per invocation; --print-example cannot be"
            + " combined with --print-task-catalog",
        exception.getMessage());
  }

  @Test
  void printExampleRejectsMissingAndDuplicateLookupFlags() {
    CliArgumentsException missing =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-example"}));
    assertEquals("--lookup", missing.argument());
    assertEquals(
        "--print-example requires --lookup and one example id value", missing.getMessage());

    CliArgumentsException duplicate =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {
                      "--print-example", "--lookup", "BUDGET", "--lookup", "DASHBOARD"
                    }));
    assertEquals("--lookup", duplicate.argument());
    assertEquals("Duplicate argument: --lookup", duplicate.getMessage());
  }

  @Test
  void printTaskPlanRejectsMissingAndDuplicateLookupFlags() {
    CliArgumentsException missing =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-task-plan"}));
    assertEquals("--lookup", missing.argument());
    assertEquals("--print-task-plan requires --lookup and one task id value", missing.getMessage());

    CliArgumentsException duplicate =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {
                      "--print-task-plan", "--lookup", "DASHBOARD", "--lookup", "TABULAR_REPORT"
                    }));
    assertEquals("--lookup", duplicate.argument());
    assertEquals("Duplicate argument: --lookup", duplicate.getMessage());
  }

  @Test
  void printTaskKeywordMatchRejectsMissingAndDuplicateQueryFlags() {
    CliArgumentsException missing =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-task-keyword-match"}));
    assertEquals("--query", missing.argument());
    assertEquals(
        "--print-task-keyword-match requires --query and one query value", missing.getMessage());

    CliArgumentsException duplicate =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {
                      "--print-task-keyword-match", "--query", "budget", "--query", "dashboard"
                    }));
    assertEquals("--query", duplicate.argument());
    assertEquals("Duplicate argument: --query", duplicate.getMessage());

    CliCommand.PrintTaskKeywordMatch blankQuery =
        assertInstanceOf(
            CliCommand.PrintTaskKeywordMatch.class,
            CliArguments.parse(new String[] {"--print-task-keyword-match", "--query", ""}));
    assertEquals("", blankQuery.query());
  }

  @Test
  void printTaskKeywordMatchRejectsTrailingPrimaryCommandsAfterItsOwnArguments() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {"--print-task-keyword-match", "--query", "budget", "--version"}));

    assertEquals("--version", exception.argument());
    assertEquals(
        "Only one primary command may be used per invocation; --print-task-keyword-match cannot be combined with --version",
        exception.getMessage());
  }

  @Test
  void primaryCommandsCannotFollowDoctorOrExecutionArguments() {
    CliArgumentsException afterDoctor =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--doctor-request", "--help"}));
    assertEquals("--help", afterDoctor.argument());
    assertEquals(
        "Only one primary command may be used per invocation; --doctor-request cannot be combined with --help",
        afterDoctor.getMessage());

    CliArgumentsException afterRequest =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--request", "request.json", "--version"}));
    assertEquals("--version", afterRequest.argument());
    assertEquals(
        "--version must be the primary command and cannot follow execution arguments",
        afterRequest.getMessage());

    CliArgumentsException duplicateDoctor =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--doctor-request", "--doctor-request"}));
    assertEquals("--doctor-request", duplicateDoctor.argument());
    assertEquals("Duplicate argument: --doctor-request", duplicateDoctor.getMessage());
  }

  @Test
  void printTaskKeywordMatchParsesIntoItsDedicatedCommand() {
    CliCommand.PrintTaskKeywordMatch command =
        assertInstanceOf(
            CliCommand.PrintTaskKeywordMatch.class,
            CliArguments.parse(
                new String[] {
                  "--print-task-keyword-match", "--query", "monthly sales dashboard with charts"
                }));

    assertEquals("monthly sales dashboard with charts", command.query());
  }

  @Test
  void helpVariantsParseIntoDedicatedTopics() {
    CliCommand.Help overview =
        assertInstanceOf(CliCommand.Help.class, CliArguments.parse(new String[] {"--help"}));
    assertEquals(CliCommand.HelpTopic.OVERVIEW, overview.topic());

    CliCommand.Help protocol =
        assertInstanceOf(
            CliCommand.Help.class, CliArguments.parse(new String[] {"--help-protocol"}));
    assertEquals(CliCommand.HelpTopic.PROTOCOL, protocol.topic());

    CliCommand.Help guidance =
        assertInstanceOf(
            CliCommand.Help.class, CliArguments.parse(new String[] {"--help-guidance"}));
    assertEquals(CliCommand.HelpTopic.GUIDANCE, guidance.topic());
  }

  @Test
  void helpVariantsReportTheirExactCommandTokenInConflictMessages() {
    CliArgumentsException protocolConflict =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--help-protocol", "--version"}));
    assertEquals("--version", protocolConflict.argument());
    assertEquals(
        "Only one primary command may be used per invocation; --help-protocol cannot be combined with --version",
        protocolConflict.getMessage());

    CliArgumentsException guidanceConflict =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--help-guidance", "--license"}));
    assertEquals("--license", guidanceConflict.argument());
    assertEquals(
        "Only one primary command may be used per invocation; --help-guidance cannot be combined with --license",
        guidanceConflict.getMessage());

    CliArgumentsException shortAliasConflict =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"-h", "--version"}));
    assertEquals("--version", shortAliasConflict.argument());
    assertEquals(
        "Only one primary command may be used per invocation; -h cannot be combined with --version",
        shortAliasConflict.getMessage());
  }

  @Test
  void bareHelpAliasIsRejected() {
    CliArgumentsException exception =
        assertThrows(CliArgumentsException.class, () -> CliArguments.parse(new String[] {"help"}));

    assertEquals("help", exception.argument());
    assertEquals("Unknown argument: help", exception.getMessage());
  }

  @Test
  void helpAcceptsResponsePath() {
    CliCommand.Help command =
        assertInstanceOf(
            CliCommand.Help.class,
            CliArguments.parse(new String[] {"--help", "--response", "help.txt"}));

    assertEquals(java.util.Optional.of(java.nio.file.Path.of("help.txt")), command.responsePath());
  }

  @Test
  void immediateCommandsAcceptResponsePathBeforeTheCommandFlag() {
    CliCommand.Help helpCommand =
        assertInstanceOf(
            CliCommand.Help.class,
            CliArguments.parse(new String[] {"--response", "help.txt", "--help"}));
    assertEquals(
        java.util.Optional.of(java.nio.file.Path.of("help.txt")), helpCommand.responsePath());

    CliCommand.Version versionCommand =
        assertInstanceOf(
            CliCommand.Version.class,
            CliArguments.parse(new String[] {"--response", "version.json", "--version"}));
    assertEquals(
        java.util.Optional.of(java.nio.file.Path.of("version.json")),
        versionCommand.responsePath());

    CliCommand.PrintProtocolCatalogSearch searchCommand =
        assertInstanceOf(
            CliCommand.PrintProtocolCatalogSearch.class,
            CliArguments.parse(
                new String[] {
                  "--response", "catalog.json", "--print-protocol-catalog", "--search", "sheet"
                }));
    assertEquals(
        java.util.Optional.of(java.nio.file.Path.of("catalog.json")), searchCommand.responsePath());
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

    assertEquals(java.util.Optional.empty(), command.requestPath());
    assertEquals(java.util.Optional.empty(), command.executionRootPath());
    assertEquals(java.util.Optional.empty(), command.tempRootPath());
    assertEquals(
        java.util.Optional.of(java.nio.file.Path.of("doctor-report.json")), command.responsePath());
  }

  @Test
  void doctorRequestParsesExecutionRootAndTempRootForStdinRequests() {
    CliCommand.DoctorRequest command =
        assertInstanceOf(
            CliCommand.DoctorRequest.class,
            CliArguments.parse(
                new String[] {
                  "--doctor-request",
                  "--execution-root",
                  "workspace",
                  "--temp-root",
                  "scratch",
                  "--response",
                  "doctor-report.json"
                }));

    assertEquals(java.util.Optional.empty(), command.requestPath());
    assertEquals(
        java.util.Optional.of(java.nio.file.Path.of("workspace")), command.executionRootPath());
    assertEquals(java.util.Optional.of(java.nio.file.Path.of("scratch")), command.tempRootPath());
    assertEquals(
        java.util.Optional.of(java.nio.file.Path.of("doctor-report.json")), command.responsePath());
  }

  @Test
  void executionFlagsParseIntoTheDefaultExecuteCommand() {
    CliCommand.Execute command =
        assertInstanceOf(
            CliCommand.Execute.class,
            CliArguments.parse(
                new String[] {
                  "--execution-root",
                  "workspace",
                  "--temp-root",
                  "scratch",
                  "--response",
                  "run.json"
                }));

    assertEquals(java.util.Optional.empty(), command.requestPath());
    assertEquals(
        java.util.Optional.of(java.nio.file.Path.of("workspace")), command.executionRootPath());
    assertEquals(java.util.Optional.of(java.nio.file.Path.of("scratch")), command.tempRootPath());
    assertEquals(java.util.Optional.of(java.nio.file.Path.of("run.json")), command.responsePath());
  }

  @Test
  void printTaskCatalogParsesLookupAndResponsePathTogether() {
    CliCommand.PrintTaskCatalog command =
        assertInstanceOf(
            CliCommand.PrintTaskCatalog.class,
            CliArguments.parse(
                new String[] {
                  "--print-task-catalog", "--lookup", "DASHBOARD", "--response", "task.json"
                }));

    assertEquals(java.util.Optional.of("DASHBOARD"), command.lookupId());
    assertEquals(java.util.Optional.of(java.nio.file.Path.of("task.json")), command.responsePath());
  }

  @Test
  void printTaskCatalogRejectsDuplicateLookupAndResponseArguments() {
    CliArgumentsException duplicateLookup =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {
                      "--print-task-catalog", "--lookup", "DASHBOARD", "--lookup", "AUDIT"
                    }));

    assertEquals("--lookup", duplicateLookup.argument());
    assertEquals("Duplicate argument: --lookup", duplicateLookup.getMessage());

    CliArgumentsException duplicateResponse =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {
                      "--print-task-catalog",
                      "--lookup",
                      "DASHBOARD",
                      "--response",
                      "a.json",
                      "--response",
                      "b.json"
                    }));

    assertEquals("--response", duplicateResponse.argument());
    assertEquals("Duplicate argument: --response", duplicateResponse.getMessage());
  }

  @Test
  void immediateCommandsRejectBlankOrMissingValues() {
    CliArgumentsException blankTask =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-task-catalog", "--lookup", ""}));
    assertEquals("--lookup", blankTask.argument());
    assertEquals("task lookup id must not be blank", blankTask.getMessage());

    CliArgumentsException blankTaskPlan =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-task-plan", "--lookup", ""}));
    assertEquals("--lookup", blankTaskPlan.argument());
    assertEquals("task lookup id must not be blank", blankTaskPlan.getMessage());

    CliArgumentsException missingTaskPlan =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-task-plan"}));
    assertEquals("--lookup", missingTaskPlan.argument());
    assertEquals(
        "--print-task-plan requires --lookup and one task id value", missingTaskPlan.getMessage());

    CliArgumentsException blankExample =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-example", "--lookup", ""}));
    assertEquals("--lookup", blankExample.argument());
    assertEquals("example lookup id must not be blank", blankExample.getMessage());

    CliArgumentsException missingExample =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-example"}));
    assertEquals("--lookup", missingExample.argument());
    assertEquals(
        "--print-example requires --lookup and one example id value", missingExample.getMessage());

    CliCommand.PrintTaskKeywordMatch blankQuery =
        assertInstanceOf(
            CliCommand.PrintTaskKeywordMatch.class,
            CliArguments.parse(new String[] {"--print-task-keyword-match", "--query", ""}));
    assertEquals("", blankQuery.query());
  }

  @Test
  void dependentFlagsExplainTheirRequiredParentCommand() {
    CliArgumentsException exampleException =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--example", "BUDGET"}));
    assertEquals("--example", exampleException.argument());
    assertEquals(
        "--example is no longer part of the CLI grammar; use --print-example --lookup <id>",
        exampleException.getMessage());

    CliArgumentsException taskException =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--task", "BUDGET"}));
    assertEquals("--task", taskException.argument());
    assertEquals(
        "--task is no longer part of the CLI grammar; use --lookup instead",
        taskException.getMessage());

    CliArgumentsException lookupException =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--lookup", "SET_CELL"}));
    assertEquals("--lookup", lookupException.argument());
    assertEquals(
        "--lookup requires --print-example, --print-task-catalog, --print-task-plan,"
            + " or --print-protocol-catalog and one lookup id value",
        lookupException.getMessage());

    CliArgumentsException searchException =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--search", "layout"}));
    assertEquals("--search", searchException.argument());
    assertEquals(
        "--search requires --print-protocol-catalog and one search text value",
        searchException.getMessage());

    CliArgumentsException queryException =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--query", "budget spreadsheet"}));
    assertEquals("--query", queryException.argument());
    assertEquals(
        "--query requires --print-task-keyword-match and one natural-language query value",
        queryException.getMessage());
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
    assertEquals("--version does not allow --request", versionFailure.getMessage());

    CliArgumentsException taskFailure =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {
                      "--print-task-plan", "--lookup", "DASHBOARD", "--request", "ignored.json"
                    }));
    assertEquals("--request", taskFailure.argument());
    assertEquals("--print-task-plan does not allow --request", taskFailure.getMessage());

    CliArgumentsException helpExecutionRootFailure =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--help", "--execution-root", "workspace"}));
    assertEquals("--execution-root", helpExecutionRootFailure.argument());
    assertEquals("--help does not allow --execution-root", helpExecutionRootFailure.getMessage());

    CliArgumentsException helpTempRootFailure =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--help", "--temp-root", "scratch"}));
    assertEquals("--temp-root", helpTempRootFailure.argument());
    assertEquals("--help does not allow --temp-root", helpTempRootFailure.getMessage());
  }

  @Test
  void immediateCommandsRejectTrailingDoctorFlagWithExplicitPrimaryCommandMessage() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-example-catalog", "--doctor-request"}));

    assertEquals("--doctor-request", exception.argument());
    assertEquals(
        "Only one primary command may be used per invocation; --print-example-catalog"
            + " cannot be combined with --doctor-request",
        exception.getMessage());
  }

  @Test
  void immediateCommandsRejectUnknownTrailingArguments() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--help", "--goal"}));

    assertEquals("--goal", exception.argument());
    assertEquals("Unknown argument: --goal", exception.getMessage());
  }

  @Test
  void immediateCommandsMustAppearFirstWhenPresent() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--request", "req.json", "--version"}));

    assertEquals("--version", exception.argument());
    assertEquals(
        "--version must be the primary command and cannot follow execution arguments",
        exception.getMessage());
  }

  @Test
  void executionRootCannotBeCombinedWithRequestFileExecution() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {"--request", "req.json", "--execution-root", "workspace"}));

    assertEquals("--execution-root", exception.argument());
    assertEquals(
        "--execution-root cannot be combined with --request because the request file directory already owns request-root resolution",
        exception.getMessage());
  }

  @Test
  void tempRootRejectsDuplicateArgumentsAcrossDoctorAndExecuteCommands() {
    CliArgumentsException duplicateDoctorTempRoot =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {
                      "--doctor-request", "--temp-root", "scratch-a", "--temp-root", "scratch-b"
                    }));
    assertEquals("--temp-root", duplicateDoctorTempRoot.argument());
    assertEquals("Duplicate argument: --temp-root", duplicateDoctorTempRoot.getMessage());

    CliArgumentsException duplicateExecuteTempRoot =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {
                      "--execution-root",
                      "workspace",
                      "--temp-root",
                      "scratch-a",
                      "--temp-root",
                      "scratch-b"
                    }));
    assertEquals("--temp-root", duplicateExecuteTempRoot.argument());
    assertEquals("Duplicate argument: --temp-root", duplicateExecuteTempRoot.getMessage());
  }
}
