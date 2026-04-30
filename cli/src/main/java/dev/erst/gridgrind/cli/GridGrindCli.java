package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.executor.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.executor.GridGrindProblems;
import dev.erst.gridgrind.executor.GridGrindRequestExecutor;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Objects;
import java.util.Optional;
import java.util.function.BooleanSupplier;

/** Thin command-line transport and dispatcher for the GridGrind protocol. */
public final class GridGrindCli {
  private final CliResponseWriter responseWriter;
  private final GridGrindCliExecutionCommands executionCommands;

  /** Creates the production CLI backed by the default request executor and transport helpers. */
  public GridGrindCli() {
    this(
        new DefaultGridGrindRequestExecutor(),
        new CliRequestReader(),
        new CliResponseWriter(),
        new CliJournalWriter(),
        StandardInputInteractivity.currentProcess());
  }

  GridGrindCli(GridGrindRequestExecutor requestExecutor) {
    this(
        requestExecutor,
        new CliRequestReader(),
        new CliResponseWriter(),
        new CliJournalWriter(),
        StandardInputInteractivity.never());
  }

  GridGrindCli(
      GridGrindRequestExecutor requestExecutor,
      CliRequestReader requestReader,
      CliResponseWriter responseWriter) {
    this(
        requestExecutor,
        requestReader,
        responseWriter,
        new CliJournalWriter(),
        StandardInputInteractivity.never());
  }

  GridGrindCli(
      GridGrindRequestExecutor requestExecutor,
      CliRequestReader requestReader,
      CliResponseWriter responseWriter,
      CliJournalWriter journalWriter) {
    this(
        requestExecutor,
        requestReader,
        responseWriter,
        journalWriter,
        StandardInputInteractivity.never());
  }

  GridGrindCli(
      GridGrindRequestExecutor requestExecutor,
      CliRequestReader requestReader,
      CliResponseWriter responseWriter,
      CliJournalWriter journalWriter,
      BooleanSupplier standardInputIsInteractive) {
    this.responseWriter = Objects.requireNonNull(responseWriter, "responseWriter must not be null");
    this.executionCommands =
        new GridGrindCliExecutionCommands(
            GridGrindRequestExecutor.requireNonNull(requestExecutor),
            requestReader,
            this.responseWriter,
            journalWriter,
            standardInputIsInteractive);
  }

  /** Runs one CLI invocation against stdin/stdout or explicit request/response file paths. */
  public int run(String[] args, InputStream stdin, OutputStream stdout) throws IOException {
    return run(args, stdin, stdout, OutputStream.nullOutputStream());
  }

  /**
   * Runs one CLI invocation against stdin/stdout/stderr or explicit request/response file paths.
   */
  public int run(String[] args, InputStream stdin, OutputStream stdout, OutputStream stderr)
      throws IOException {
    Objects.requireNonNull(args, "args must not be null");
    Objects.requireNonNull(stdin, "stdin must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");

    CliCommand command;
    try {
      command = CliArguments.parse(args);
    } catch (CliArgumentsException exception) {
      responseWriter.write(
          stdout,
          failure(
              GridGrindProblemCode.INVALID_ARGUMENTS,
              exception.getMessage(),
              new dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments(
                  dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument.named(
                      exception.argument())),
              exception));
      return 2;
    } catch (IllegalArgumentException exception) {
      responseWriter.write(
          stdout,
          failure(
              GridGrindProblemCode.INVALID_ARGUMENTS,
              exception.getMessage(),
              new dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments(
                  dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument
                      .unknown()),
              exception));
      return 2;
    }

    return switch (command) {
      case CliCommand.Help cmd -> GridGrindCliCatalogCommands.help(cmd, stdout, responseWriter);
      case CliCommand.Version cmd ->
          GridGrindCliCatalogCommands.version(cmd, stdout, responseWriter);
      case CliCommand.License cmd ->
          GridGrindCliCatalogCommands.license(cmd, stdout, responseWriter);
      case CliCommand.PrintRequestTemplate cmd ->
          GridGrindCliCatalogCommands.requestTemplate(cmd, stdout, responseWriter);
      case CliCommand.PrintExample cmd ->
          GridGrindCliCatalogCommands.example(cmd, stdout, responseWriter);
      case CliCommand.PrintTaskCatalog cmd ->
          GridGrindCliCatalogCommands.taskCatalog(cmd, stdout, responseWriter);
      case CliCommand.PrintTaskPlan cmd ->
          GridGrindCliCatalogCommands.taskPlan(cmd, stdout, responseWriter);
      case CliCommand.PrintGoalPlan cmd ->
          GridGrindCliCatalogCommands.goalPlan(cmd, stdout, responseWriter);
      case CliCommand.DoctorRequest doctor ->
          executionCommands.doctorRequest(doctor, stdin, stdout);
      case CliCommand.PrintProtocolCatalogAll cmd ->
          GridGrindCliCatalogCommands.protocolCatalogAll(cmd, stdout, responseWriter);
      case CliCommand.PrintProtocolCatalogSearch cmd ->
          GridGrindCliCatalogCommands.protocolCatalogSearch(cmd, stdout, responseWriter);
      case CliCommand.PrintProtocolCatalogLookup cmd ->
          GridGrindCliCatalogCommands.protocolCatalogLookup(cmd, stdout, responseWriter);
      case CliCommand.Execute execute -> {
        Optional<InputStream> requestInput =
            executionCommands.standardInputOrNullForImplicitHelp(execute, args, stdin, stdout);
        if (requestInput.isEmpty()) {
          yield 0;
        }
        yield executionCommands.executeCommand(execute, requestInput.orElseThrow(), stdout, stderr);
      }
    };
  }

  /**
   * Returns the full help text rendered for the given implementation version string.
   *
   * <p>Tests call this directly so doc routing and discovery guidance can be validated without a
   * packaged JAR manifest.
   */
  static String helpText(String implementationVersion) {
    return GridGrindCliProductInfo.helpText(implementationVersion);
  }

  /**
   * Returns the two-line product header shared by {@code --help} and {@code --version}:
   *
   * <pre>
   * GridGrind {version}
   * {description}
   * </pre>
   *
   * <p>This is the single source of truth for the product identity line printed to the user.
   */
  static String productHeader(String version, String description) {
    return GridGrindCliProductInfo.productHeader(version, description);
  }

  /**
   * Returns the given implementation version string, or {@code "unknown"} when the JAR manifest
   * attribute is absent (e.g. when running from the test classpath without a packaged JAR).
   */
  static String versionFrom(String implementationVersion) {
    return GridGrindCliProductInfo.versionFrom(implementationVersion);
  }

  /**
   * Loads the product description from the {@code gridgrind.properties} classpath resource bundled
   * by the build, falling back to {@code "GridGrind"} when the resource is absent (e.g. when
   * running from the test classpath before resources are processed).
   */
  static String descriptionFrom(Class<?> anchor) {
    return GridGrindCliProductInfo.descriptionFrom(anchor);
  }

  /**
   * Loads the product description from the supplied stream, falling back to {@code "GridGrind"}
   * when the stream is null, blank, or unreadable.
   */
  static String descriptionFrom(InputStream stream) {
    return GridGrindCliProductInfo.descriptionFrom(stream);
  }

  /**
   * Reads the bundled license texts from classpath resources and returns them as a single string.
   *
   * <p>Falls back to a brief notice when the resource is absent (e.g. test classpath without a
   * packaged JAR).
   */
  static String licenseText(Class<?> anchor) {
    return GridGrindCliProductInfo.licenseText(anchor);
  }

  /**
   * Assembles the license output from the supplied streams. Exposed for testing.
   *
   * <p>Any null or unreadable stream is silently skipped. Returns a fallback notice when all
   * streams are absent.
   */
  static String licenseText(
      InputStream own, InputStream notice, InputStream apache, InputStream bsd) {
    return GridGrindCliProductInfo.licenseText(own, notice, apache, bsd);
  }

  /**
   * Returns the built-in request template rendered as UTF-8 text via the supplied byte producer.
   *
   * <p>Tests call this directly to assert the failure path without mocking static codec methods.
   */
  static String requestTemplateText(RequestTemplateBytesSupplier supplier) {
    return GridGrindCliProductInfo.requestTemplateText(supplier);
  }

  private GridGrindResponse.Failure failure(
      GridGrindProblemCode code,
      String message,
      dev.erst.gridgrind.contract.dto.ProblemContext context,
      Throwable cause) {
    return GridGrindResponses.failure(
        GridGrindProtocolVersion.current(),
        GridGrindProblems.problem(code, message, context, cause));
  }

  /** Supplies request-template bytes for help rendering. */
  @FunctionalInterface
  interface RequestTemplateBytesSupplier {
    /** Returns the UTF-8 bytes that should be embedded into the help text. */
    byte[] get() throws IOException;
  }
}
