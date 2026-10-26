package dev.erst.gridgrind.core.problem;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.CliArgument;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestInput;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.ResponseOutput;
import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.InputReference;
import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.ProblemLocation;
import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.StepReference;
import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.WorkbookReference;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.Objects;
import java.util.Optional;

/** Structured execution metadata that pinpoints where and why a failure happened. */
@Union(discriminator = "type")
public sealed interface ProblemContext
    permits ProblemContext.ParseArguments,
        CliRuntimeContext,
        RequestInputContext,
        ProblemContext.RequestShapeContext,
        ProblemContext.ExecuteStep,
        ProblemContext.PersistWorkbook,
        ProblemContext.ExecuteRequest,
        ProblemContext.WriteResponse {
  /** Metadata-owned type identifying the failure context. */
  default String type() {
    return dev.erst.gridgrind.core.catalog.ProtocolTypeMetadataSupport.requiredTypeId(getClass());
  }

  /** Context for failures that occur while parsing CLI arguments. */
  @Variant(
      id = "PARSE_ARGUMENTS",
      summary = "Context for failures that occur while parsing CLI arguments.")
  record ParseArguments(CliArgument argument) implements ProblemContext {
    public ParseArguments {
      Objects.requireNonNull(argument, "argument must not be null");
    }

    /** Returns the concrete CLI argument when parsing failure was attributable to one option. */
    public Optional<String> argumentName() {
      return argument.argumentValue();
    }
  }

  /** Context for failures that occur while reading or structurally checking the JSON request. */
  @Variant(
      id = "READ_REQUEST",
      summary =
          "Context for failures that occur while reading or structurally checking the JSON request.")
  record ReadRequest(RequestInput request, JsonLocation json) implements RequestInputContext {
    public ReadRequest {
      Objects.requireNonNull(request, "request must not be null");
      json = Objects.requireNonNullElseGet(json, JsonLocation.Unavailable::new);
    }

    /** Returns one copy enriched with JSON cursor details when none were already present. */
    public ReadRequest withJson(JsonLocation discovered) {
      Objects.requireNonNull(discovered, "discovered must not be null");
      return new ReadRequest(request, json instanceof JsonLocation.Unavailable ? discovered : json);
    }
  }

  /**
   * Context for failures that occur while binding structurally valid JSON into the request model.
   */
  @Variant(
      id = "BIND_REQUEST",
      summary =
          "Context for failures that occur while binding structurally valid JSON into the request model.")
  record BindRequest(RequestInput request, JsonLocation json) implements RequestInputContext {
    public BindRequest {
      Objects.requireNonNull(request, "request must not be null");
      json = Objects.requireNonNullElseGet(json, JsonLocation.Unavailable::new);
    }

    /** Returns one copy enriched with JSON cursor details when none were already present. */
    public BindRequest withJson(JsonLocation discovered) {
      Objects.requireNonNull(discovered, "discovered must not be null");
      return new BindRequest(request, json instanceof JsonLocation.Unavailable ? discovered : json);
    }
  }

  /** Common request-shape context for stages that need source and persistence families. */
  sealed interface RequestShapeContext extends ProblemContext
      permits ValidateRequest, ResolveInputs, OpenWorkbook, ExecuteCalculation {
    /** Returns the authored request shape active for the failure context. */
    RequestShape request();

    /** Returns the decoded request source family when known. */
    default Optional<String> sourceType() {
      return request().sourceTypeValue();
    }

    /** Returns the decoded request persistence family when known. */
    default Optional<String> persistenceType() {
      return request().persistenceTypeValue();
    }
  }

  /** Context for failures that occur while statically validating bound request fragments. */
  @Variant(
      id = "VALIDATE_REQUEST",
      summary =
          "Context for failures that occur while statically validating bound request fragments.")
  record ValidateRequest(RequestShape request, Optional<JsonLocation> json)
      implements RequestShapeContext {
    public ValidateRequest {
      Objects.requireNonNull(request, "request must not be null");
      json = Objects.requireNonNullElseGet(json, Optional::empty);
    }

    /** Creates an unlocated static-validation context for programmatically assembled requests. */
    public ValidateRequest(RequestShape request) {
      this(request, Optional.empty());
    }

    /** Returns one copy carrying the authored JSON location of the static violation. */
    public ValidateRequest withJson(JsonLocation location) {
      return new ValidateRequest(
          request, Optional.of(Objects.requireNonNull(location, "location must not be null")));
    }
  }

  /** Context for failures that occur while resolving source-backed authored inputs. */
  @Variant(
      id = "RESOLVE_INPUTS",
      summary = "Context for failures that occur while resolving source-backed authored inputs.")
  record ResolveInputs(RequestShape request, InputReference input, Optional<JsonLocation> json)
      implements RequestShapeContext {
    public ResolveInputs {
      Objects.requireNonNull(request, "request must not be null");
      Objects.requireNonNull(input, "input must not be null");
      json = Objects.requireNonNullElseGet(json, Optional::empty);
    }

    /** Creates an input-resolution context without a retained authored request location. */
    public ResolveInputs(RequestShape request, InputReference input) {
      this(request, input, Optional.empty());
    }

    /** Returns the authored input kind when one failing binding was identified. */
    public Optional<String> inputKind() {
      return input.inputKindValue();
    }

    /** Returns the authored input path when one failing binding resolved to one file. */
    public Optional<String> inputPath() {
      return input.inputPathValue();
    }
  }

  /** Context for failures that occur while opening the source workbook. */
  @Variant(
      id = "OPEN_WORKBOOK",
      summary = "Context for failures that occur while opening the source workbook.")
  record OpenWorkbook(RequestShape request, WorkbookReference workbook)
      implements RequestShapeContext {
    public OpenWorkbook {
      Objects.requireNonNull(request, "request must not be null");
      Objects.requireNonNull(workbook, "workbook must not be null");
    }

    /** Returns the opened workbook path when execution targeted one existing file. */
    public Optional<String> sourceWorkbookPath() {
      return workbook.sourceWorkbookPathValue();
    }
  }

  /** Common calculation-failure context for the two top-level calculation stages. */
  sealed interface ExecuteCalculation extends RequestShapeContext
      permits ExecuteCalculation.Preflight, ExecuteCalculation.Execution {
    /** Returns the best-known workbook location derived for the calculation failure. */
    ProblemLocation location();

    /** Returns one copy enriched with exception-derived location facts. */
    default ExecuteCalculation withLocation(ProblemLocation discovered) {
      Objects.requireNonNull(discovered, "discovered must not be null");
      ProblemLocation merged = mergeLocation(location(), discovered);
      return switch (this) {
        case Preflight context -> new Preflight(context.request(), merged);
        case Execution context -> new Execution(context.request(), merged);
      };
    }

    /** Returns the localized sheet name when available. */
    default Optional<SheetLabel> sheetName() {
      return location().sheetNameValue();
    }

    /** Returns the localized address when available. */
    default Optional<String> address() {
      return location().addressValue();
    }

    /** Returns the localized range when available. */
    default Optional<String> range() {
      return location().rangeValue();
    }

    /** Returns the localized named range when available. */
    default Optional<String> namedRangeName() {
      return location().namedRangeNameValue();
    }

    /** Returns the localized formula when available. */
    default Optional<String> formula() {
      return location().formulaValue();
    }

    /** Context for failures that occur during calculation preflight analysis. */
    @Variant(
        id = "CALCULATION_PREFLIGHT",
        summary = "Context for failures that occur during calculation preflight analysis.")
    record Preflight(RequestShape request, ProblemLocation location) implements ExecuteCalculation {
      public Preflight {
        Objects.requireNonNull(request, "request must not be null");
        Objects.requireNonNull(location, "location must not be null");
      }
    }

    /** Context for failures that occur during immediate calculation execution. */
    @Variant(
        id = "CALCULATION_EXECUTION",
        summary = "Context for failures that occur during immediate calculation execution.")
    record Execution(RequestShape request, ProblemLocation location) implements ExecuteCalculation {
      public Execution {
        Objects.requireNonNull(request, "request must not be null");
        Objects.requireNonNull(location, "location must not be null");
      }
    }
  }

  /** Context for failures that occur while executing one ordered workbook step. */
  @Variant(
      id = "EXECUTE_STEP",
      summary = "Context for failures that occur while executing one ordered workbook step.")
  record ExecuteStep(
      RequestShape request,
      StepReference step,
      ProblemLocation location,
      Optional<StepReference> surfacedAtStep)
      implements ProblemContext {
    public ExecuteStep {
      Objects.requireNonNull(request, "request must not be null");
      Objects.requireNonNull(step, "step must not be null");
      Objects.requireNonNull(location, "location must not be null");
      surfacedAtStep = Objects.requireNonNullElseGet(surfacedAtStep, Optional::empty);
    }

    /** Creates a step context whose failure surfaced while that same step executed. */
    public ExecuteStep(RequestShape request, StepReference step, ProblemLocation location) {
      this(request, step, location, Optional.empty());
    }

    /** Returns the decoded request source family when step execution reached that point. */
    public Optional<String> sourceType() {
      return request.sourceTypeValue();
    }

    /** Returns the decoded request persistence family when step execution reached that point. */
    public Optional<String> persistenceType() {
      return request.persistenceTypeValue();
    }

    /** Returns the zero-based step index for the failing workbook step. */
    public int stepIndex() {
      return step.stepIndex();
    }

    /** Returns the authored step id for the failing workbook step. */
    public StepId stepId() {
      return step.stepId();
    }

    /** Returns the concrete workbook step type for the failing workbook step. */
    public String stepType() {
      return step.stepType();
    }

    /** Returns the localized sheet name when the failing step resolved to one sheet target. */
    public Optional<SheetLabel> sheetName() {
      return location.sheetNameValue();
    }

    /** Returns the localized cell address when the failing step resolved to one cell target. */
    public Optional<String> address() {
      return location.addressValue();
    }

    /** Returns the localized range when the failing step resolved to one rectangular range. */
    public Optional<String> range() {
      return location.rangeValue();
    }

    /** Returns the localized named range when the failing step resolved to one named target. */
    public Optional<String> namedRangeName() {
      return location.namedRangeNameValue();
    }

    /** Returns the localized formula when the failing step resolved to one formula cell. */
    public Optional<String> formula() {
      return location.formulaValue();
    }

    /** Returns one copy enriched with exception-derived location facts. */
    public ExecuteStep withLocation(ProblemLocation discovered) {
      Objects.requireNonNull(discovered, "discovered must not be null");
      return new ExecuteStep(request, step, mergeLocation(location, discovered), surfacedAtStep);
    }
  }

  /** Context for failures that occur while persisting the workbook to its destination path. */
  @Variant(
      id = "PERSIST_WORKBOOK",
      summary =
          "Context for failures that occur while persisting the workbook to its destination path.")
  record PersistWorkbook(
      RequestShape request, Optional<String> sourceWorkbookPath, Optional<String> persistencePath)
      implements ProblemContext {
    public PersistWorkbook {
      Objects.requireNonNull(request, "request must not be null");
      sourceWorkbookPath = normalizeOptionalPath(sourceWorkbookPath, "sourceWorkbookPath");
      persistencePath = normalizeOptionalPath(persistencePath, "persistencePath");
      if (sourceWorkbookPath.isPresent() == persistencePath.isPresent()) {
        throw new IllegalArgumentException(
            "PersistWorkbook requires exactly one of sourceWorkbookPath or persistencePath");
      }
    }

    /** Creates one persist-workbook context from the typed persistence-target helper. */
    public PersistWorkbook(
        RequestShape request, ProblemContextWorkbookSurfaces.PersistenceReference persistence) {
      this(
          request,
          requiredPersistence(persistence).sourceWorkbookPathValue(),
          requiredPersistence(persistence).persistencePathValue());
    }

    /** Returns the decoded request source family when persistence reached that point. */
    public Optional<String> sourceType() {
      return request.sourceTypeValue();
    }

    /** Returns the decoded request persistence family when persistence reached that point. */
    public Optional<String> persistenceType() {
      return request.persistenceTypeValue();
    }

    private static Optional<String> normalizeOptionalPath(Optional<String> path, String fieldName) {
      Optional<String> normalized = Objects.requireNonNullElseGet(path, Optional::empty);
      if (normalized.isEmpty()) {
        return Optional.empty();
      }
      return Optional.of(WorkbookPlan.requireNonBlank(normalized.orElseThrow(), fieldName));
    }

    private static ProblemContextWorkbookSurfaces.PersistenceReference requiredPersistence(
        ProblemContextWorkbookSurfaces.PersistenceReference persistence) {
      return Objects.requireNonNull(persistence, "persistence must not be null");
    }
  }

  /** Context for top-level execution failures not attributed to a specific pipeline stage. */
  @Variant(
      id = "EXECUTE_REQUEST",
      summary =
          "Context for top-level execution failures not attributed to a specific pipeline stage.")
  record ExecuteRequest(RequestShape request) implements ProblemContext {
    public ExecuteRequest {
      Objects.requireNonNull(request, "request must not be null");
    }

    /** Returns the decoded request source family when top-level execution failed. */
    public Optional<String> sourceType() {
      return request.sourceTypeValue();
    }

    /** Returns the decoded request persistence family when top-level execution failed. */
    public Optional<String> persistenceType() {
      return request.persistenceTypeValue();
    }
  }

  /** Context for failures that occur while writing the JSON response to its destination. */
  @Variant(
      id = "WRITE_RESPONSE",
      summary =
          "Context for failures that occur while writing the JSON response to its destination.")
  record WriteResponse(ResponseOutput output) implements ProblemContext {
    public WriteResponse {
      Objects.requireNonNull(output, "output must not be null");
    }

    /** Returns the authored response path when output targeted one file instead of stdout. */
    public Optional<String> responsePath() {
      return output.responsePathValue();
    }
  }

  /** Merges one discovered workbook location into one existing context location. */
  static ProblemLocation mergeLocation(ProblemLocation current, ProblemLocation discovered) {
    return ProblemContextSupport.mergeLocation(current, discovered);
  }
}
