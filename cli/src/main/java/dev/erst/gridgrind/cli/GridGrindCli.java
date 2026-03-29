package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.protocol.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.protocol.GridGrindJson;
import dev.erst.gridgrind.protocol.GridGrindProblemCode;
import dev.erst.gridgrind.protocol.GridGrindProblems;
import dev.erst.gridgrind.protocol.GridGrindProtocolCatalog;
import dev.erst.gridgrind.protocol.GridGrindProtocolVersion;
import dev.erst.gridgrind.protocol.GridGrindRequest;
import dev.erst.gridgrind.protocol.GridGrindRequestExecutor;
import dev.erst.gridgrind.protocol.GridGrindResponse;
import dev.erst.gridgrind.protocol.InvalidJsonException;
import dev.erst.gridgrind.protocol.InvalidRequestException;
import dev.erst.gridgrind.protocol.InvalidRequestShapeException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Properties;

/** Thin command-line transport for the GridGrind protocol. */
public final class GridGrindCli {
  private final GridGrindRequestExecutor requestExecutor;
  private final CliRequestReader requestReader;
  private final CliResponseWriter responseWriter;

  /** Creates the production CLI backed by the default request executor and transport helpers. */
  public GridGrindCli() {
    this(new DefaultGridGrindRequestExecutor(), new CliRequestReader(), new CliResponseWriter());
  }

  GridGrindCli(GridGrindRequestExecutor requestExecutor) {
    this(requestExecutor, new CliRequestReader(), new CliResponseWriter());
  }

  GridGrindCli(
      GridGrindRequestExecutor requestExecutor,
      CliRequestReader requestReader,
      CliResponseWriter responseWriter) {
    this.requestExecutor = GridGrindRequestExecutor.requireNonNull(requestExecutor);
    this.requestReader = Objects.requireNonNull(requestReader, "requestReader must not be null");
    this.responseWriter = Objects.requireNonNull(responseWriter, "responseWriter must not be null");
  }

  /** Runs one CLI invocation against stdin/stdout or explicit request/response file paths. */
  public int run(String[] args, InputStream stdin, OutputStream stdout) throws IOException {
    Objects.requireNonNull(args, "args must not be null");
    Objects.requireNonNull(stdin, "stdin must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");

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
      case CliCommand.Help _ -> {
        stdout.write(helpText(version()).getBytes(StandardCharsets.UTF_8));
        stdout.flush();
        yield 0;
      }
      case CliCommand.Version _ -> {
        stdout.write(("gridgrind " + version() + "\n").getBytes(StandardCharsets.UTF_8));
        stdout.flush();
        yield 0;
      }
      case CliCommand.PrintRequestTemplate _ -> {
        GridGrindJson.writeRequest(stdout, GridGrindProtocolCatalog.requestTemplate());
        stdout.write('\n');
        stdout.flush();
        yield 0;
      }
      case CliCommand.PrintProtocolCatalog _ -> {
        GridGrindJson.writeProtocolCatalog(stdout, GridGrindProtocolCatalog.catalog());
        stdout.write('\n');
        stdout.flush();
        yield 0;
      }
      case CliCommand.Execute execute -> executeCommand(execute, stdin, stdout);
    };
  }

  private int executeCommand(CliCommand.Execute command, InputStream stdin, OutputStream stdout)
      throws IOException {
    GridGrindRequest request;
    try {
      request = requestReader.read(command.requestPath(), stdin);
    } catch (InvalidJsonException
        | InvalidRequestShapeException
        | InvalidRequestException exception) {
      return responseWriter.write(
          command.responsePath(),
          stdout,
          new GridGrindResponse.Failure(
              GridGrindProtocolVersion.current(),
              GridGrindProblems.fromException(
                  exception,
                  new GridGrindResponse.ProblemContext.ReadRequest(
                      pathString(command.requestPath()), null, null, null))));
    } catch (IOException exception) {
      return responseWriter.write(
          command.responsePath(),
          stdout,
          new GridGrindResponse.Failure(
              GridGrindProtocolVersion.current(),
              GridGrindProblems.fromException(
                  exception,
                  new GridGrindResponse.ProblemContext.ReadRequest(
                      pathString(command.requestPath()), null, null, null))));
    }

    GridGrindResponse response;
    try {
      response = requestExecutor.execute(request);
    } catch (Exception exception) {
      response =
          new GridGrindResponse.Failure(
              GridGrindProtocolVersion.current(),
              GridGrindProblems.fromException(
                  exception, new GridGrindResponse.ProblemContext.ExecuteRequest(null, null)));
    }

    return responseWriter.write(command.responsePath(), stdout, response);
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
    String description = descriptionFrom(GridGrindCli.class);
    String requestTemplate = requestTemplateText();
    String documentRef = documentRef(version);
    return """
        GridGrind CLI %s
        %s

        Usage:
          gridgrind [--request <path>] [--response <path>]
          gridgrind --print-request-template
          gridgrind --print-protocol-catalog
          gridgrind --help | -h
          gridgrind --version

        Execution:
          GridGrind runs operations first, then reads, then saves the workbook (unless persistence is NONE); if any step fails, no workbook is written.
          A NEW workbook starts with zero sheets; use ENSURE_SHEET to create the first sheet.

        Limits:
          File format:              .xlsx only; .xls, .xlsm, and .xlsb are rejected.
          Sheet names:              1 to 31 characters (Excel limit).
          GET_WINDOW cell count:    rowCount * columnCount must not exceed 250,000.
          GET_SHEET_SCHEMA cells:   rowCount * columnCount must not exceed 250,000.
          Column widthCharacters:   > 0 and <= 255 (Excel limit).
          Row heightPoints:         > 0 and <= 1638.35 (Excel limit: 32767 twips).
          DATE / DATE_TIME inputs:  stored as numeric serial; GET_CELLS returns declaredType=NUMERIC.

        Minimal Valid Request:
        %s

        stdin Example:
          gridgrind --print-request-template | gridgrind

        Docker File Example:
          docker run --rm -i \\
            -v "$(pwd)":/workdir \\
            ghcr.io/resoltico/gridgrind:%s \\
            --request /workdir/request.json \\
            --response /workdir/response.json

          Paths passed in --request, --response, source.path, and persistence.path are resolved in
          the current execution environment. In Docker, use absolute mounted container paths.

        Discovery:
          gridgrind --print-request-template
          gridgrind --print-protocol-catalog

        Docs:
          Quick reference: %s/docs/QUICK_REFERENCE.md
          Operations reference: %s/docs/OPERATIONS.md
          Error reference: %s/docs/ERRORS.md

        Flags:
          --request <path>           Read the JSON request from a file instead of stdin.
          --response <path>          Write the JSON response to a file instead of stdout.
          --print-request-template   Print a minimal valid request JSON document.
          --print-protocol-catalog   Print the machine-readable protocol catalog.
          --help, -h                 Print this help text.
          --version                  Print the GridGrind version.
        """
        .formatted(
            version,
            description,
            indentBlock(requestTemplate),
            containerTag(version),
            documentRef,
            documentRef,
            documentRef);
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

  private static String requestTemplateText() {
    return requestTemplateText(
        () -> GridGrindJson.writeRequestBytes(GridGrindProtocolCatalog.requestTemplate()));
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

  private static String indentBlock(String value) {
    return value.indent(2).stripTrailing();
  }

  private static String containerTag(String version) {
    return "unknown".equals(version) ? "latest" : version;
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
    return new GridGrindResponse.Failure(
        GridGrindProtocolVersion.current(),
        GridGrindProblems.problem(code, message, context, cause));
  }

  private String message(Exception exception) {
    return exception.getMessage();
  }

  private String pathString(Path path) {
    return path == null ? null : path.toAbsolutePath().toString();
  }

  /** Supplies request-template bytes for help rendering. */
  @FunctionalInterface
  interface RequestTemplateBytesSupplier {
    /** Returns the UTF-8 bytes that should be embedded into the help text. */
    byte[] get() throws IOException;
  }
}
