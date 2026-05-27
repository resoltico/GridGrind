package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.engine.api.GridGrindEngine;
import dev.erst.gridgrind.engine.api.GridGrindRequestDoctor;
import dev.erst.gridgrind.engine.api.GridGrindRequestExecutor;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
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
        GridGrindEngine.requestExecutor(),
        GridGrindEngine.requestDoctor(),
        new CliRequestReader(),
        new CliResponseWriter(),
        new CliJournalWriter(),
        StandardInputInteractivity.currentProcess());
  }

  static GridGrindCli forTesting(GridGrindRequestExecutor executor) {
    return new GridGrindCli(
        executor,
        GridGrindEngine.requestDoctor(),
        new CliRequestReader(),
        new CliResponseWriter(),
        new CliJournalWriter(),
        StandardInputInteractivity.never());
  }

  static GridGrindCli forTesting(
      GridGrindRequestExecutor executor, BooleanSupplier standardInputIsInteractive) {
    return new GridGrindCli(
        executor,
        GridGrindEngine.requestDoctor(),
        new CliRequestReader(),
        new CliResponseWriter(),
        new CliJournalWriter(),
        standardInputIsInteractive);
  }

  GridGrindCli(
      GridGrindRequestExecutor requestExecutor,
      GridGrindRequestDoctor requestDoctor,
      CliRequestReader requestReader,
      CliResponseWriter responseWriter,
      CliJournalWriter journalWriter,
      BooleanSupplier standardInputIsInteractive) {
    this.responseWriter = Objects.requireNonNull(responseWriter, "responseWriter must not be null");
    this.executionCommands =
        new GridGrindCliExecutionCommands(
            GridGrindRequestExecutor.requireNonNull(requestExecutor),
            GridGrindRequestDoctor.requireNonNull(requestDoctor),
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

    Optional<java.nio.file.Path> responsePathHint = responsePathHint(args);
    CliCommand command;
    try {
      command = CliArguments.parse(args);
    } catch (CliArgumentsException exception) {
      return responseWriter.writeCliFailureReport(
          responsePathHint, stdout, stderr, CliArgumentFailureSupport.reportFor(exception));
    } catch (IllegalArgumentException exception) {
      return responseWriter.writeCliFailureReport(
          responsePathHint, stdout, stderr, CliArgumentFailureSupport.reportFor(exception));
    }

    return switch (command) {
      case CliCommand.Help cmd ->
          GridGrindCliCatalogCommands.help(cmd, stdout, stderr, responseWriter);
      case CliCommand.Version cmd ->
          GridGrindCliCatalogCommands.version(cmd, stdout, stderr, responseWriter);
      case CliCommand.License cmd ->
          GridGrindCliCatalogCommands.license(cmd, stdout, stderr, responseWriter);
      case CliCommand.PrintRequestTemplate cmd ->
          GridGrindCliCatalogCommands.requestTemplate(cmd, stdout, stderr, responseWriter);
      case CliCommand.PrintExample cmd ->
          GridGrindCliCatalogCommands.example(cmd, stdout, stderr, responseWriter);
      case CliCommand.PrintExampleCatalog cmd ->
          GridGrindCliCatalogCommands.exampleCatalog(cmd, stdout, stderr, responseWriter);
      case CliCommand.PrintTaskCatalog cmd ->
          GridGrindCliCatalogCommands.taskCatalog(cmd, stdout, stderr, responseWriter);
      case CliCommand.PrintTaskPlan cmd ->
          GridGrindCliCatalogCommands.taskPlan(cmd, stdout, stderr, responseWriter);
      case CliCommand.PrintTaskKeywordMatch cmd ->
          GridGrindCliCatalogCommands.taskKeywordMatch(cmd, stdout, stderr, responseWriter);
      case CliCommand.DoctorRequest doctor ->
          executionCommands.doctorRequest(doctor, stdin, stdout, stderr);
      case CliCommand.PrintProtocolCatalogAll cmd ->
          GridGrindCliCatalogCommands.protocolCatalogAll(cmd, stdout, stderr, responseWriter);
      case CliCommand.PrintProtocolCatalogSearch cmd ->
          GridGrindCliCatalogCommands.protocolCatalogSearch(cmd, stdout, stderr, responseWriter);
      case CliCommand.PrintProtocolCatalogLookup cmd ->
          GridGrindCliCatalogCommands.protocolCatalogLookup(cmd, stdout, stderr, responseWriter);
      case CliCommand.Execute execute -> {
        Optional<InputStream> requestInput =
            executionCommands.standardInputIfPresent(execute, stdin);
        if (requestInput.isEmpty()) {
          yield responseWriter.writeCliFailureReport(
              execute.responsePath(),
              stdout,
              stderr,
              CliFailureReports.invalidArguments(
                  2,
                  "execute",
                  Optional.of("--request"),
                  "No request JSON was provided. Pass --request <path> or pipe one request"
                      + " document on standard input.",
                  List.of(
                      "gridgrind --print-request-template --response request.json",
                      "gridgrind --help",
                      "gridgrind --help-protocol"),
                  Optional.of(
                      "Use explicit --help for documentation; a bare gridgrind invocation"
                          + " now expects a real request document.")));
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
    return GridGrindCliProductInfo.helpText(CliCommand.HelpTopic.OVERVIEW, implementationVersion);
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
   * Returns the given implementation version string, or the bundled product-resource version when
   * the JAR manifest attribute is absent.
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
      InputStream own,
      InputStream notice,
      InputStream apache,
      InputStream bsd2,
      InputStream bsd3,
      InputStream edl) {
    return GridGrindCliProductInfo.licenseText(own, notice, apache, bsd2, bsd3, edl);
  }

  /**
   * Returns the built-in request template rendered as UTF-8 text via the supplied byte producer.
   *
   * <p>Tests call this directly to assert the failure path without mocking static codec methods.
   */
  static String requestTemplateText(RequestTemplateBytesSupplier supplier) {
    return GridGrindCliProductInfo.requestTemplateText(supplier);
  }

  private static Optional<java.nio.file.Path> responsePathHint(String[] args) {
    try {
      Optional<java.nio.file.Path> responsePath = CliArguments.responsePath(args);
      Optional<java.nio.file.Path> requestPath = CliArguments.requestPath(args);
      if (responsePath.isPresent()
          && requestPath.isPresent()
          && responsePath
              .orElseThrow()
              .toAbsolutePath()
              .equals(requestPath.orElseThrow().toAbsolutePath())) {
        return Optional.empty();
      }
      return responsePath;
    } catch (IllegalArgumentException exception) {
      return Optional.empty();
    }
  }

  /** Supplies request-template bytes for help rendering. */
  @FunctionalInterface
  interface RequestTemplateBytesSupplier {
    /** Returns the UTF-8 bytes that should be embedded into the help text. */
    byte[] get() throws IOException;
  }
}
