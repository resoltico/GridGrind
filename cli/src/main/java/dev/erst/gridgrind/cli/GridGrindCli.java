package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.catalog.GridGrindCliHelp;
import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.catalog.GridGrindTaskCatalog;
import dev.erst.gridgrind.contract.catalog.GridGrindTaskPlanner;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.InvalidJsonException;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.executor.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.executor.ExecutionInputBindings;
import dev.erst.gridgrind.executor.GridGrindProblems;
import dev.erst.gridgrind.executor.GridGrindRequestDoctor;
import dev.erst.gridgrind.executor.GridGrindRequestExecutor;
import dev.erst.gridgrind.executor.SourceBackedPlanResolver;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PushbackInputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.Properties;
import java.util.function.BooleanSupplier;

/** Thin command-line transport for the GridGrind protocol. */
@SuppressWarnings("PMD.ExcessiveImports")
public final class GridGrindCli {
  private final GridGrindRequestExecutor requestExecutor;
  private final CliRequestReader requestReader;
  private final CliResponseWriter responseWriter;
  private final CliJournalWriter journalWriter;
  private final BooleanSupplier standardInputIsInteractive;

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
    this.requestExecutor = GridGrindRequestExecutor.requireNonNull(requestExecutor);
    this.requestReader = Objects.requireNonNull(requestReader, "requestReader must not be null");
    this.responseWriter = Objects.requireNonNull(responseWriter, "responseWriter must not be null");
    this.journalWriter = Objects.requireNonNull(journalWriter, "journalWriter must not be null");
    this.standardInputIsInteractive =
        Objects.requireNonNull(
            standardInputIsInteractive, "standardInputIsInteractive must not be null");
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
              new GridGrindResponse.ProblemContext.ParseArguments(exception.argument()),
              exception));
      return 2;
    } catch (IllegalArgumentException exception) {
      responseWriter.write(
          stdout,
          failure(
              GridGrindProblemCode.INVALID_ARGUMENTS,
              message(exception),
              new GridGrindResponse.ProblemContext.ParseArguments(null),
              exception));
      return 2;
    }

    return switch (command) {
      case CliCommand.Help cmd ->
          writePayload(
              cmd.responsePath(), stdout, helpText(version()).getBytes(StandardCharsets.UTF_8));
      case CliCommand.Version cmd -> {
        String ver = version();
        String desc = descriptionFrom(GridGrindCli.class);
        yield writePayload(
            cmd.responsePath(), stdout, productHeader(ver, desc).getBytes(StandardCharsets.UTF_8));
      }
      case CliCommand.License cmd ->
          writePayload(
              cmd.responsePath(),
              stdout,
              licenseText(GridGrindCli.class).getBytes(StandardCharsets.UTF_8));
      case CliCommand.PrintRequestTemplate cmd ->
          writePayload(
              cmd.responsePath(),
              stdout,
              GridGrindJson.writeRequestBytes(GridGrindProtocolCatalog.requestTemplate()));
      case CliCommand.PrintExample cmd -> {
        var example = GridGrindProtocolCatalog.exampleFor(cmd.exampleId());
        if (example.isEmpty()) {
          String message = unknownExampleMessage(cmd.exampleId());
          yield responseWriter.write(
              cmd.responsePath(),
              stdout,
              failure(
                  GridGrindProblemCode.INVALID_ARGUMENTS,
                  message,
                  new GridGrindResponse.ProblemContext.ParseArguments("--print-example"),
                  new IllegalArgumentException(message)),
              2);
        }
        yield writePayload(
            cmd.responsePath(), stdout, GridGrindJson.writeRequestBytes(example.get().plan()));
      }
      case CliCommand.PrintTaskCatalog cmd -> {
        if (cmd.taskFilter() == null) {
          yield writePayload(
              cmd.responsePath(),
              stdout,
              GridGrindJson.writeTaskCatalogBytes(GridGrindTaskCatalog.catalog()));
        }
        var entry = GridGrindTaskCatalog.entryFor(cmd.taskFilter());
        if (entry.isEmpty()) {
          yield responseWriter.write(
              cmd.responsePath(),
              stdout,
              failure(
                  GridGrindProblemCode.INVALID_ARGUMENTS,
                  "Unknown task: " + cmd.taskFilter(),
                  new GridGrindResponse.ProblemContext.ParseArguments("--task"),
                  new IllegalArgumentException("Unknown task: " + cmd.taskFilter())),
              2);
        }
        yield writePayload(
            cmd.responsePath(),
            stdout,
            output -> GridGrindJson.writeTaskEntry(output, entry.get()));
      }
      case CliCommand.PrintTaskPlan cmd -> {
        var task = GridGrindTaskCatalog.entryFor(cmd.taskId());
        if (task.isEmpty()) {
          yield responseWriter.write(
              cmd.responsePath(),
              stdout,
              failure(
                  GridGrindProblemCode.INVALID_ARGUMENTS,
                  "Unknown task: " + cmd.taskId(),
                  new GridGrindResponse.ProblemContext.ParseArguments("--print-task-plan"),
                  new IllegalArgumentException("Unknown task: " + cmd.taskId())),
              2);
        }
        yield writePayload(
            cmd.responsePath(),
            stdout,
            GridGrindJson.writeTaskPlanTemplateBytes(GridGrindTaskPlanner.planFor(task.get())));
      }
      case CliCommand.PrintGoalPlan cmd -> {
        try {
          yield writePayload(
              cmd.responsePath(),
              stdout,
              GridGrindJson.writeGoalPlanReportBytes(
                  dev.erst.gridgrind.contract.catalog.GridGrindGoalPlanner.reportFor(cmd.goal())));
        } catch (IllegalArgumentException exception) {
          yield responseWriter.write(
              cmd.responsePath(),
              stdout,
              failure(
                  GridGrindProblemCode.INVALID_ARGUMENTS,
                  message(exception),
                  new GridGrindResponse.ProblemContext.ParseArguments("--print-goal-plan"),
                  exception),
              2);
        }
      }
      case CliCommand.DoctorRequest doctor -> doctorRequest(doctor, stdin, stdout);
      case CliCommand.PrintProtocolCatalog cmd -> {
        if (cmd.searchQuery() != null) {
          yield writePayload(
              cmd.responsePath(),
              stdout,
              output ->
                  GridGrindJson.writeCatalogLookupValue(
                      output, GridGrindProtocolCatalog.searchCatalog(cmd.searchQuery())));
        }
        if (cmd.operationFilter() == null) {
          yield writePayload(
              cmd.responsePath(),
              stdout,
              GridGrindJson.writeProtocolCatalogBytes(GridGrindProtocolCatalog.catalog()));
        }
        List<String> matches = GridGrindProtocolCatalog.matchingLookupIds(cmd.operationFilter());
        if (matches.size() > 1) {
          String message =
              "Ambiguous operation: "
                  + cmd.operationFilter()
                  + ". Use one of: "
                  + String.join(", ", matches);
          yield responseWriter.write(
              cmd.responsePath(),
              stdout,
              failure(
                  GridGrindProblemCode.INVALID_ARGUMENTS,
                  message,
                  new GridGrindResponse.ProblemContext.ParseArguments("--operation"),
                  new IllegalArgumentException(message)),
              2);
        }
        var lookupValue = GridGrindProtocolCatalog.lookupValueFor(cmd.operationFilter());
        if (lookupValue.isEmpty()) {
          yield responseWriter.write(
              cmd.responsePath(),
              stdout,
              failure(
                  GridGrindProblemCode.INVALID_ARGUMENTS,
                  "Unknown operation: " + cmd.operationFilter(),
                  new GridGrindResponse.ProblemContext.ParseArguments("--operation"),
                  new IllegalArgumentException("Unknown operation: " + cmd.operationFilter())),
              2);
        }
        yield writePayload(
            cmd.responsePath(),
            stdout,
            output -> GridGrindJson.writeCatalogLookupValue(output, lookupValue.get()));
      }
      case CliCommand.Execute execute -> {
        Optional<InputStream> requestInput =
            standardInputOrNullForImplicitHelp(execute, args, stdin, stdout);
        if (requestInput.isEmpty()) {
          yield 0;
        }
        yield executeCommand(execute, requestInput.orElseThrow(), stdout, stderr);
      }
    };
  }

  private Optional<InputStream> standardInputOrNullForImplicitHelp(
      CliCommand.Execute command, String[] args, InputStream stdin, OutputStream stdout)
      throws IOException {
    if (command.requestPath() != null) {
      return Optional.of(stdin);
    }
    if (args.length != 0) {
      return Optional.of(stdin);
    }
    if (standardInputIsInteractive.getAsBoolean()) {
      writeHelp(stdout);
      return Optional.empty();
    }
    PushbackInputStream peekable = new PushbackInputStream(stdin, 1);
    int firstByte = peekable.read();
    if (firstByte < 0) {
      writeHelp(stdout);
      return Optional.empty();
    }
    peekable.unread(firstByte);
    return Optional.of(peekable);
  }

  private static String unknownExampleMessage(String exampleId) {
    return suggestedExampleId(exampleId)
        .map(
            suggestion ->
                "Unknown example: "
                    + exampleId
                    + ". Example ids use stable upper-case tokens; did you mean "
                    + suggestion
                    + "?")
        .orElse("Unknown example: " + exampleId);
  }

  private static Optional<String> suggestedExampleId(String exampleId) {
    String normalizedExampleId = normalizeLookupToken(exampleId);
    return GridGrindProtocolCatalog.shippedExamples().stream()
        .filter(
            example ->
                example.id().equalsIgnoreCase(exampleId)
                    || example.fileName().equalsIgnoreCase(exampleId)
                    || exampleStem(example.fileName()).equalsIgnoreCase(exampleId)
                    || normalizeLookupToken(exampleStem(example.fileName()))
                        .equals(normalizedExampleId))
        .map(dev.erst.gridgrind.contract.catalog.GridGrindShippedExamples.ShippedExample::id)
        .findFirst();
  }

  private static String normalizeLookupToken(String value) {
    return value.toUpperCase(Locale.ROOT).replace('-', '_').replace(' ', '_');
  }

  private static String exampleStem(String fileName) {
    return fileName.substring(0, fileName.length() - 5);
  }

  private int executeCommand(
      CliCommand.Execute command, InputStream stdin, OutputStream stdout, OutputStream stderr)
      throws IOException {
    WorkbookPlan request;
    try {
      request = requestReader.read(command.requestPath(), stdin);
    } catch (InvalidJsonException
        | InvalidRequestShapeException
        | InvalidRequestException exception) {
      return responseWriter.write(
          command.responsePath(),
          stdout,
          GridGrindResponse.failure(
              GridGrindProtocolVersion.current(),
              GridGrindProblems.fromException(
                  exception,
                  new GridGrindResponse.ProblemContext.ReadRequest(
                      pathString(command.requestPath()), null, null, null))));
    } catch (IOException exception) {
      return responseWriter.write(
          command.responsePath(),
          stdout,
          GridGrindResponse.failure(
              GridGrindProtocolVersion.current(),
              GridGrindProblems.fromException(
                  exception,
                  new GridGrindResponse.ProblemContext.ReadRequest(
                      pathString(command.requestPath()), null, null, null))));
    }

    GridGrindResponse response;
    if (command.requestPath() == null && SourceBackedPlanResolver.requiresStandardInput(request)) {
      String message = GridGrindContractText.standardInputRequiresRequestMessage();
      response =
          failure(
              GridGrindProblemCode.INVALID_ARGUMENTS,
              message,
              new GridGrindResponse.ProblemContext.ParseArguments("--request"),
              new IllegalArgumentException(message));
      return responseWriter.write(command.responsePath(), stdout, response);
    }

    ExecutionInputBindings bindings =
        CliExecutionBindingsFactory.create(command.requestPath(), request, stdin);
    try {
      response = requestExecutor.execute(request, bindings, journalWriter.sinkFor(request, stderr));
    } catch (Exception exception) {
      response =
          GridGrindResponse.failure(
              GridGrindProtocolVersion.current(),
              GridGrindProblems.fromException(
                  exception,
                  new GridGrindResponse.ProblemContext.ExecuteRequest(
                      requestSourceType(request), requestPersistenceType(request))));
    }

    return responseWriter.write(command.responsePath(), stdout, response);
  }

  private int doctorRequest(
      CliCommand.DoctorRequest command, InputStream stdin, OutputStream stdout) throws IOException {
    RequestDoctorReport report;
    WorkbookPlan request;
    try {
      request = requestReader.read(command.requestPath(), stdin);
    } catch (InvalidJsonException
        | InvalidRequestShapeException
        | InvalidRequestException exception) {
      report =
          RequestDoctorReport.invalid(
              null,
              List.of(),
              GridGrindProblems.fromException(
                  exception,
                  new GridGrindResponse.ProblemContext.ReadRequest(
                      pathString(command.requestPath()), null, null, null)));
      return responseWriter.writeDoctorReport(command.responsePath(), stdout, report);
    } catch (IOException exception) {
      report =
          RequestDoctorReport.invalid(
              null,
              List.of(),
              GridGrindProblems.fromException(
                  exception,
                  new GridGrindResponse.ProblemContext.ReadRequest(
                      pathString(command.requestPath()), null, null, null)));
      return responseWriter.writeDoctorReport(command.responsePath(), stdout, report);
    }

    if (command.requestPath() == null && SourceBackedPlanResolver.requiresStandardInput(request)) {
      report = new GridGrindRequestDoctor().diagnose(request);
      String message = GridGrindContractText.standardInputRequiresRequestMessage();
      report =
          RequestDoctorReport.invalid(
              report.summary(),
              report.warnings(),
              GridGrindProblems.problem(
                  GridGrindProblemCode.INVALID_ARGUMENTS,
                  message,
                  new GridGrindResponse.ProblemContext.ParseArguments("--request"),
                  new IllegalArgumentException(message)));
      return responseWriter.writeDoctorReport(command.responsePath(), stdout, report);
    }

    ExecutionInputBindings bindings =
        command.requestPath() == null
            ? ExecutionInputBindings.processDefault()
            : CliExecutionBindingsFactory.create(command.requestPath(), request, stdin);
    report = new GridGrindRequestDoctor().diagnose(request, bindings);
    return responseWriter.writeDoctorReport(command.responsePath(), stdout, report);
  }

  private static String version() {
    return versionFrom(GridGrindCli.class.getPackage().getImplementationVersion());
  }

  /**
   * Returns the full help text rendered for the given implementation version string.
   *
   * <p>Tests call this directly so doc routing and discovery guidance can be validated without a
   * packaged JAR manifest.
   */
  static String helpText(String implementationVersion) {
    String version = versionFrom(implementationVersion);
    return GridGrindCliHelp.helpText(
        version,
        descriptionFrom(GridGrindCli.class),
        documentRef(version),
        containerImageRef(version));
  }

  private void writeHelp(OutputStream stdout) throws IOException {
    stdout.write(helpText(version()).getBytes(StandardCharsets.UTF_8));
    stdout.flush();
  }

  private int writePayload(Path responsePath, OutputStream stdout, byte[] payload)
      throws IOException {
    return responseWriter.writePayload(responsePath, stdout, payload, 0);
  }

  private int writePayload(Path responsePath, OutputStream stdout, OutputRenderer renderer)
      throws IOException {
    ByteArrayOutputStream buffer = new ByteArrayOutputStream();
    renderer.write(buffer);
    return responseWriter.writePayload(responsePath, stdout, buffer.toByteArray(), 0);
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
    return "GridGrind " + version + "\n" + description;
  }

  /**
   * Returns the given implementation version string, or {@code "unknown"} when the JAR manifest
   * attribute is absent (e.g. when running from the test classpath without a packaged JAR).
   */
  static String versionFrom(String implementationVersion) {
    if (implementationVersion == null) {
      return "unknown";
    }
    return implementationVersion;
  }

  /**
   * Loads the product description from the {@code gridgrind.properties} classpath resource bundled
   * by the build, falling back to {@code "GridGrind"} when the resource is absent (e.g. when
   * running from the test classpath before resources are processed).
   */
  static String descriptionFrom(Class<?> anchor) {
    return descriptionFrom(anchor.getResourceAsStream("/gridgrind.properties"));
  }

  /**
   * Loads the product description from the supplied stream, falling back to {@code "GridGrind"}
   * when the stream is null, blank, or unreadable.
   */
  static String descriptionFrom(InputStream stream) {
    if (stream == null) {
      return "GridGrind";
    }
    try (stream) {
      Properties properties = new Properties();
      properties.load(stream);
      String description = properties.getProperty("description", "");
      return description.isBlank() ? "GridGrind" : description;
    } catch (IOException exception) {
      return "GridGrind";
    }
  }

  /**
   * Reads the bundled license texts from classpath resources and returns them as a single string.
   *
   * <p>Falls back to a brief notice when the resource is absent (e.g. test classpath without a
   * packaged JAR).
   */
  static String licenseText(Class<?> anchor) {
    return licenseText(
        anchor.getResourceAsStream("/licenses/LICENSE"),
        anchor.getResourceAsStream("/licenses/NOTICE"),
        anchor.getResourceAsStream("/licenses/LICENSE-APACHE-2.0"),
        anchor.getResourceAsStream("/licenses/LICENSE-BSD-3-CLAUSE"));
  }

  /**
   * Assembles the license output from the supplied streams. Exposed for testing.
   *
   * <p>Any null or unreadable stream is silently skipped. Returns a fallback notice when all
   * streams are absent.
   */
  static String licenseText(
      InputStream own, InputStream notice, InputStream apache, InputStream bsd) {
    String ownText = readLicenseStream(own);
    String thirdParty = buildThirdParty(notice, apache, bsd);
    if (ownText.isEmpty() && thirdParty.isEmpty()) {
      return "License information not available in this distribution.\n";
    }
    StringBuilder result = new StringBuilder(ownText.length() + thirdParty.length() + 64);
    result.append(ownText);
    if (!thirdParty.isEmpty()) {
      if (!ownText.isEmpty()) {
        result.append("\n---\n\nThird-party notices and licenses:\n\n");
      }
      result.append(thirdParty);
    }
    String text = result.toString();
    return text.endsWith("\n") ? text : text + '\n';
  }

  private static String buildThirdParty(InputStream notice, InputStream apache, InputStream bsd) {
    String noticeText = readLicenseStream(notice);
    String apacheText = readLicenseStream(apache);
    String bsdText = readLicenseStream(bsd);
    int capacity = noticeText.length() + apacheText.length() + bsdText.length() + 2;
    StringBuilder result = new StringBuilder(capacity);
    String sep = "";
    if (!noticeText.isEmpty()) {
      result.append(noticeText);
      sep = "\n";
    }
    if (!apacheText.isEmpty()) {
      result.append(sep).append(apacheText);
      sep = "\n";
    }
    if (!bsdText.isEmpty()) {
      result.append(sep).append(bsdText);
    }
    return result.toString();
  }

  private static String readLicenseStream(InputStream stream) {
    if (stream == null) {
      return "";
    }
    try (stream) {
      return new String(stream.readAllBytes(), StandardCharsets.UTF_8);
    } catch (IOException ignored) {
      return "";
    }
  }

  /**
   * Returns the built-in request template rendered as UTF-8 text via the supplied byte producer.
   *
   * <p>Tests call this directly to assert the failure path without mocking static codec methods.
   */
  static String requestTemplateText(RequestTemplateBytesSupplier supplier) {
    Objects.requireNonNull(supplier, "supplier must not be null");
    try {
      return new String(supplier.get(), StandardCharsets.UTF_8);
    } catch (IOException exception) {
      throw new IllegalStateException("Failed to render the built-in request template", exception);
    }
  }

  private static String containerImageRef(String version) {
    String tag = "unknown".equals(version) ? "latest" : version;
    return "ghcr.io/resoltico/gridgrind:" + tag;
  }

  private static String documentRef(String version) {
    String gitRef = "unknown".equals(version) ? "main" : "v" + version;
    return "https://github.com/resoltico/GridGrind/blob/" + gitRef;
  }

  private GridGrindResponse.Failure failure(
      GridGrindProblemCode code,
      String message,
      GridGrindResponse.ProblemContext context,
      Throwable cause) {
    return GridGrindResponse.failure(
        GridGrindProtocolVersion.current(),
        GridGrindProblems.problem(code, message, context, cause));
  }

  private String message(Exception exception) {
    return exception.getMessage();
  }

  private String pathString(Path path) {
    return path == null ? null : path.toAbsolutePath().toString();
  }

  private String requestSourceType(WorkbookPlan request) {
    return switch (request.source()) {
      case WorkbookPlan.WorkbookSource.New _ -> "NEW";
      case WorkbookPlan.WorkbookSource.ExistingFile _ -> "EXISTING";
    };
  }

  private String requestPersistenceType(WorkbookPlan request) {
    return switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.None _ -> "NONE";
      case WorkbookPlan.WorkbookPersistence.OverwriteSource _ -> "OVERWRITE";
      case WorkbookPlan.WorkbookPersistence.SaveAs _ -> "SAVE_AS";
    };
  }

  /** Supplies request-template bytes for help rendering. */
  @FunctionalInterface
  private interface OutputRenderer {
    /** Writes one command payload into the supplied caller-owned output stream. */
    void write(OutputStream outputStream) throws IOException;
  }

  /** Supplies request-template bytes for help rendering. */
  @FunctionalInterface
  interface RequestTemplateBytesSupplier {
    /** Returns the UTF-8 bytes that should be embedded into the help text. */
    byte[] get() throws IOException;
  }
}
