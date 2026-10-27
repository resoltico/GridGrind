package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.cli.io.CliPrimaryOutputException;
import dev.erst.gridgrind.cli.io.CliRequestReader;
import dev.erst.gridgrind.cli.io.CliResponseWriter;
import dev.erst.gridgrind.engine.api.GridGrindEngine;
import dev.erst.gridgrind.engine.api.GridGrindRequestDoctor;
import dev.erst.gridgrind.engine.api.GridGrindRequestExecutor;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Objects;
import java.util.Optional;

/** Thin command-line transport and dispatcher for the GridGrind protocol. */
public final class GridGrindCli {
  private final CliResponseWriter responseWriter;
  private final RunCommand executionCommands;
  private final DoctorCommand doctorCommand;

  /** Creates the production CLI backed by the default request executor and transport helpers. */
  public GridGrindCli() {
    this(
        GridGrindEngine.requestExecutor(),
        GridGrindEngine.requestDoctor(),
        new CliRequestReader(),
        new CliResponseWriter());
  }

  static GridGrindCli forTesting(GridGrindRequestExecutor executor) {
    return new GridGrindCli(
        executor, GridGrindEngine.requestDoctor(), new CliRequestReader(), new CliResponseWriter());
  }

  GridGrindCli(
      GridGrindRequestExecutor requestExecutor,
      GridGrindRequestDoctor requestDoctor,
      CliRequestReader requestReader,
      CliResponseWriter responseWriter) {
    this.responseWriter = Objects.requireNonNull(responseWriter, "responseWriter must not be null");
    this.executionCommands =
        new RunCommand(
            GridGrindRequestExecutor.requireNonNull(requestExecutor),
            GridGrindRequestDoctor.requireNonNull(requestDoctor),
            requestReader,
            this.responseWriter);
    this.doctorCommand = new DoctorCommand(requestDoctor, requestReader, this.responseWriter);
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
    Optional<java.nio.file.Path> responsePathHint = CliPathArguments.responsePathHint(args);
    boolean prettyJsonHint = CliArguments.prettyJsonHint(args);
    try {
      return runInternal(args, stdin, stdout, stderr, responsePathHint, prettyJsonHint);
    } catch (CliPrimaryOutputException ignored) {
      return 1;
    } catch (Throwable exception) {
      return CliUnexpectedFailureSupport.emit(
          args, responsePathHint, prettyJsonHint, stdout, stderr, exception);
    }
  }

  private int runInternal(
      String[] args,
      InputStream stdin,
      OutputStream stdout,
      OutputStream stderr,
      Optional<java.nio.file.Path> responsePathHint,
      boolean prettyJsonHint)
      throws IOException {
    CliInvocation invocation;
    try {
      invocation = CliArguments.parseInvocation(args);
    } catch (CliArgumentsException exception) {
      return responseWriter.writeCommandError(
          responsePathHint,
          stdout,
          stderr,
          CommandErrors.reportFor(args, exception),
          prettyJsonHint);
    } catch (IllegalArgumentException exception) {
      return responseWriter.writeCommandError(
          responsePathHint,
          stdout,
          stderr,
          CommandErrors.reportFor(args, exception),
          prettyJsonHint);
    }
    CliCommand command = invocation.command();
    Optional<CliOutputFormat> outputFormat = invocation.outputFormat();
    boolean prettyJson = invocation.prettyJson();

    return switch (command) {
      case CliCommand.Schema cmd ->
          SchemaCommand.execute(cmd, prettyJson, stdout, stderr, responseWriter);
      case CliCommand.Help cmd ->
          HelpCommand.execute(cmd, outputFormat, prettyJson, stdout, stderr, responseWriter);
      case CliCommand.Version cmd ->
          VersionCommand.execute(cmd, outputFormat, prettyJson, stdout, stderr, responseWriter);
      case CliCommand.License cmd ->
          LicenseCommand.execute(cmd, outputFormat, prettyJson, stdout, stderr, responseWriter);
      case CliCommand.TemplateRequest cmd ->
          TemplateCommand.request(cmd, prettyJson, stdout, stderr, responseWriter);
      case CliCommand.TemplateGrant cmd ->
          TemplateCommand.grant(cmd, prettyJson, stdout, stderr, responseWriter);
      case CliCommand.RecipeList cmd ->
          RecipeCommand.recipeCatalog(cmd, prettyJson, stdout, stderr, responseWriter);
      case CliCommand.RecipeShow cmd ->
          RecipeCommand.recipe(cmd, prettyJson, stdout, stderr, responseWriter);
      case CliCommand.RecipeMaterialize cmd ->
          RecipeCommand.materializeRecipe(cmd, prettyJson, stdout, stderr, responseWriter);
      case CliCommand.RecipeMatch cmd ->
          RecipeCommand.recipeKeywordMatch(cmd, prettyJson, stdout, stderr, responseWriter);
      case CliCommand.Doctor doctor -> doctorCommand.run(doctor, stdin, stdout, stderr, prettyJson);
      case CliCommand.CatalogIndex cmd ->
          CatalogCommand.protocolCatalogIndex(cmd, prettyJson, stdout, stderr, responseWriter);
      case CliCommand.CatalogSearch cmd ->
          CatalogCommand.protocolCatalogSearch(cmd, prettyJson, stdout, stderr, responseWriter);
      case CliCommand.CatalogShow cmd ->
          CatalogCommand.protocolCatalogLookup(cmd, prettyJson, stdout, stderr, responseWriter);
      case CliCommand.Run execute ->
          executionCommands.executeCommand(execute, stdin, stdout, stderr, prettyJson);
    };
  }
}
