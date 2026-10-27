package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.cli.render.HelpSurfaceReport;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.EnumMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;

/** Parses explicit subcommands directly into typed host operations. */
public final class CliArguments {
  private CliArguments() {}

  static CliCommand parse(String[] args) {
    return parseInvocation(args).command();
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap") // Each invocation owns its enum-keyed options.
  static CliInvocation parseInvocation(String[] args) {
    Objects.requireNonNull(args, "args must not be null");
    if (args.length == 0) {
      return new CliInvocation(
          new CliCommand.Help(CliCommand.HelpTopic.OVERVIEW, Optional.empty()),
          Optional.empty(),
          false);
    }
    String primary = args[0];
    CommandName operation =
        CommandName.forToken(primary)
            .orElseThrow(
                () ->
                    new CliArgumentsException(
                        primary, "Unknown command: " + primary + ". Use gridgrind help."));
    Map<Option, String> values = new EnumMap<>(Option.class);
    List<String> operands = new ArrayList<>();
    boolean pretty = false;
    int index = 1;
    while (index < args.length) {
      String token = args[index];
      if (!token.startsWith("-")) {
        if (token.isBlank())
          throw new CliArgumentsException(
              operandName(primary, operands),
              operandName(primary, operands) + " must not be blank");
        operands.add(token);
        index++;
        continue;
      }
      Option option = Option.forToken(token);
      if (option == Option.DETAILS) {
        if (values.containsKey(option))
          throw new CliArgumentsException(token, "Duplicate argument: " + token);
        values.put(option, "true");
        index++;
        continue;
      }
      if (option == Option.PRETTY) {
        if (pretty) throw new CliArgumentsException(token, "Duplicate argument: " + token);
        pretty = true;
        index++;
        continue;
      }
      if (values.containsKey(option)) {
        throw new CliArgumentsException(token, "Duplicate argument: " + token);
      }
      int valueIndex = CliPathArguments.nextValueIndex(args, index, token);
      String value = args[valueIndex];
      if (value.startsWith("--"))
        throw new CliArgumentsException(token, "Missing value for " + token);
      values.put(option, CliPathArguments.requireNonBlankValue(token, value, token.substring(2)));
      index = valueIndex + 1;
    }
    Optional<Path> response = path(values, Option.RESPONSE);
    Optional<CliOutputFormat> format =
        Optional.ofNullable(values.get(Option.FORMAT)).map(CliOutputFormat::parse);
    CliCommand command = command(operation, primary, operands, values, response);
    CliInvocation invocation = new CliInvocation(command, format, pretty);
    CliRenderOptionValidation.validate(command, format);
    return invocation;
  }

  private static CliCommand command(
      CommandName operation,
      String primary,
      List<String> operands,
      Map<Option, String> values,
      Optional<Path> response) {
    return switch (operation) {
      case RUN, DOCTOR -> execution(primary, operands, values, response);
      case HELP -> help(operands, values, response);
      case VERSION -> {
        noOperands(primary, operands);
        allow(primary, values, Option.RESPONSE, Option.FORMAT);
        yield new CliCommand.Version(response);
      }
      case LICENSE -> {
        noOperands(primary, operands);
        allow(primary, values, Option.RESPONSE, Option.FORMAT);
        yield new CliCommand.License(response);
      }
      case TEMPLATE -> {
        allow(primary, values, Option.RESPONSE);
        oneOperand(primary, operands);
        yield switch (operands.getFirst()) {
          case "request" -> new CliCommand.TemplateRequest(response);
          case "grant" -> new CliCommand.TemplateGrant(response);
          default ->
              throw new CliArgumentsException("document", "template requires request or grant");
        };
      }
      case SCHEMA -> {
        allow(primary, values, Option.RESPONSE);
        oneOperand(primary, operands);
        dev.erst.gridgrind.json.schema.JsonSchemas.Document document;
        try {
          document =
              dev.erst.gridgrind.json.schema.JsonSchemas.Document.valueOf(
                  operands.getFirst().toUpperCase(Locale.ROOT));
        } catch (IllegalArgumentException failure) {
          throw new CliArgumentsException(
              "document", "schema requires request, grant, or response", failure);
        }
        yield new CliCommand.Schema(document, response);
      }
      case CATALOG -> catalog(operands, values, response);
      case RECIPE -> recipe(operands, values, response);
    };
  }

  private static CliCommand help(
      List<String> operands, Map<Option, String> values, Optional<Path> response) {
    allow("help", values, Option.RESPONSE, Option.FORMAT);
    if (operands.size() > 1) throw unexpected("help", operands);
    CliCommand.HelpTopic topic;
    try {
      topic =
          operands.isEmpty()
              ? CliCommand.HelpTopic.OVERVIEW
              : CliCommand.HelpTopic.valueOf(operands.getFirst().toUpperCase(Locale.ROOT));
    } catch (IllegalArgumentException failure) {
      throw new CliArgumentsException(
          "topic", "help topic must be overview, protocol, or guidance", failure);
    }
    return new CliCommand.Help(topic, response);
  }

  private static CliCommand execution(
      String primary, List<String> operands, Map<Option, String> values, Optional<Path> response) {
    noOperands(primary, operands);
    allow(
        primary,
        values,
        Option.REQUEST,
        Option.GRANT,
        Option.EXECUTION_ROOT,
        Option.TEMP_ROOT,
        Option.SECRETS_PROVIDER,
        Option.RESPONSE);
    require(values, Option.REQUEST, primary);
    if ("run".equals(primary)) require(values, Option.GRANT, primary);
    Optional<Path> request = path(values, Option.REQUEST);
    Optional<Path> executionRoot = path(values, Option.EXECUTION_ROOT);
    Optional<Path> grant = path(values, Option.GRANT);
    CliExecutionArgumentValidation.validateTerminalArguments(
        request, executionRoot, grant, response);
    return "run".equals(primary)
        ? new CliCommand.Run(
            request,
            executionRoot,
            path(values, Option.TEMP_ROOT),
            grant,
            path(values, Option.SECRETS_PROVIDER),
            response)
        : new CliCommand.Doctor(
            request,
            executionRoot,
            path(values, Option.TEMP_ROOT),
            grant,
            path(values, Option.SECRETS_PROVIDER),
            response);
  }

  private static CliCommand catalog(
      List<String> operands, Map<Option, String> values, Optional<Path> response) {
    allow("catalog", values, Option.RESPONSE);
    if (operands.isEmpty()) return new CliCommand.CatalogIndex(response);
    if (operands.size() != 2) throw unexpected("catalog", operands);
    return switch (operands.getFirst()) {
      case "search" -> new CliCommand.CatalogSearch(operands.get(1), response);
      case "show" -> new CliCommand.CatalogShow(operands.get(1), response);
      default ->
          throw new CliArgumentsException(
              operands.getFirst(), "catalog requires search <text> or show <group>:<id>");
    };
  }

  private static CliCommand recipe(
      List<String> operands, Map<Option, String> values, Optional<Path> response) {
    if (operands.isEmpty()) {
      throw new CliArgumentsException(
          "recipe", "recipe requires list, show, match, or materialize");
    }
    String action = operands.getFirst();
    if ("list".equals(action)) {
      allow("recipe list", values, Option.RESPONSE);
      if (operands.size() != 1) throw unexpected("recipe list", operands);
      return new CliCommand.RecipeList(response);
    }
    if (operands.size() != 2) throw unexpected("recipe " + action, operands);
    if ("materialize".equals(action)) {
      allow("recipe materialize", values, Option.WORKSPACE);
      require(values, Option.WORKSPACE, "recipe materialize");
      return new CliCommand.RecipeMaterialize(
          operands.get(1), path(values, Option.WORKSPACE).orElseThrow());
    }
    if ("show".equals(action)) allow("recipe show", values, Option.RESPONSE, Option.DETAILS);
    else allow("recipe " + action, values, Option.RESPONSE);
    return switch (action) {
      case "show" ->
          new CliCommand.RecipeShow(operands.get(1), values.containsKey(Option.DETAILS), response);
      case "match" -> new CliCommand.RecipeMatch(operands.get(1), response);
      default ->
          throw new CliArgumentsException(
              action, "recipe requires list, show, match, or materialize");
    };
  }

  /** Recognizes pretty output before parsing, for failure reporting. */
  public static boolean prettyJsonHint(String[] args) {
    return java.util.Arrays.stream(args).anyMatch(Option.PRETTY.token::equals);
  }

  private static String operandName(String command, List<String> preceding) {
    if (preceding.isEmpty()) return command;
    return switch (preceding.getFirst()) {
      case "show", "materialize" -> "id";
      case "search", "match" -> "text";
      default -> command;
    };
  }

  private static Optional<Path> path(Map<Option, String> values, Option option) {
    return Optional.ofNullable(values.get(option))
        .map(
            value ->
                CliPathArguments.requirePathValue(
                    option.token,
                    value,
                    option.token.substring(2) + " path",
                    option == Option.REQUEST));
  }

  private static void require(Map<Option, String> values, Option option, String command) {
    if (!values.containsKey(option)) {
      throw new CliArgumentsException(option.token, command + " requires " + option.token);
    }
  }

  private static void allow(String command, Map<Option, String> values, Option... allowed) {
    Set<Option> supported = Set.of(allowed);
    for (Option option : values.keySet()) {
      if (!supported.contains(option)) {
        throw new CliArgumentsException(option.token, command + " does not allow " + option.token);
      }
    }
  }

  private static void noOperands(String command, List<String> operands) {
    if (!operands.isEmpty()) throw unexpected(command, operands);
  }

  private static void oneOperand(String command, List<String> operands) {
    if (operands.size() != 1) throw unexpected(command, operands);
  }

  private static CliArgumentsException unexpected(String command, List<String> operands) {
    return new CliArgumentsException(
        command, "Unexpected operands for " + command + ": " + String.join(" ", operands));
  }

  /** Returns help entries from the exact parser option vocabulary. */
  public static List<HelpSurfaceReport.DefinitionEntry> optionHelp() {
    return java.util.Arrays.stream(Option.values())
        .map(option -> new HelpSurfaceReport.DefinitionEntry(option.token, option.description))
        .toList();
  }

  /** Exact option vocabulary shared by parsing and help. */
  private enum Option {
    REQUEST(
        "--request", "Request JSON path; '-' reads stdin and requires an explicit execution root."),
    GRANT(
        "--grant",
        "Explicit authority document; required for run and optional for static doctor analysis."),
    RESPONSE("--response", "Primary output path instead of stdout; cannot overwrite the request."),
    EXECUTION_ROOT(
        "--execution-root", "Resolution root for stdin requests; cannot accompany a file request."),
    TEMP_ROOT("--temp-root", "Parent directory for a private per-run scratch directory."),
    SECRETS_PROVIDER(
        "--secrets-provider", "Secrets-provider descriptor path, resolved by the host."),
    WORKSPACE("--workspace", "New directory atomically populated by recipe materialize."),
    DETAILS("--details", "Return recipe metadata and request profile with recipe show."),
    FORMAT("--format", "text or structured output for help, version and license."),
    PRETTY("--pretty", "Indent JSON output.");
    private final String token;
    private final String description;

    Option(String token, String description) {
      this.token = token;
      this.description = description;
    }

    private static Option forToken(String token) {
      return java.util.Arrays.stream(values())
          .filter(option -> option.token.equals(token))
          .findFirst()
          .orElseThrow(() -> new CliArgumentsException(token, "Unknown option: " + token));
    }
  }
}
