package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.ResponseOutput;
import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.InputReference;
import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.PersistenceReference;
import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.ProblemLocation;
import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.StepReference;
import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.WorkbookReference;
import java.util.Objects;
import java.util.Optional;

/** Structured execution metadata that pinpoints where and why a failure happened. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "stage")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ProblemContext.ParseArguments.class, name = "PARSE_ARGUMENTS"),
  @JsonSubTypes.Type(value = ProblemContext.ReadRequest.class, name = "READ_REQUEST"),
  @JsonSubTypes.Type(value = ProblemContext.ValidateRequest.class, name = "VALIDATE_REQUEST"),
  @JsonSubTypes.Type(value = ProblemContext.ResolveInputs.class, name = "RESOLVE_INPUTS"),
  @JsonSubTypes.Type(value = ProblemContext.OpenWorkbook.class, name = "OPEN_WORKBOOK"),
  @JsonSubTypes.Type(
      value = ProblemContext.ExecuteCalculation.Preflight.class,
      name = "CALCULATION_PREFLIGHT"),
  @JsonSubTypes.Type(
      value = ProblemContext.ExecuteCalculation.Execution.class,
      name = "CALCULATION_EXECUTION"),
  @JsonSubTypes.Type(value = ProblemContext.ExecuteStep.class, name = "EXECUTE_STEP"),
  @JsonSubTypes.Type(value = ProblemContext.PersistWorkbook.class, name = "PERSIST_WORKBOOK"),
  @JsonSubTypes.Type(value = ProblemContext.ExecuteRequest.class, name = "EXECUTE_REQUEST"),
  @JsonSubTypes.Type(value = ProblemContext.WriteResponse.class, name = "WRITE_RESPONSE")
})
public sealed interface ProblemContext {
  /** Pipeline stage in which the failure occurred. */
  String stage();

  /** Context for failures that occur while parsing CLI arguments. */
  record ParseArguments(CliArgument argument) implements ProblemContext {
    public ParseArguments {
      Objects.requireNonNull(argument, "argument must not be null");
    }

    /** Returns the concrete CLI argument when parsing failure was attributable to one option. */
    public Optional<String> argumentName() {
      return argument.argumentValue();
    }

    @Override
    public String stage() {
      return "PARSE_ARGUMENTS";
    }
  }

  /** Context for failures that occur while reading and parsing the JSON request. */
  record ReadRequest(RequestInput request, JsonLocation json) implements ProblemContext {
    public ReadRequest {
      Objects.requireNonNull(request, "request must not be null");
      json = Objects.requireNonNullElseGet(json, JsonLocation.Unavailable::new);
    }

    /** Returns the authored request file path when the request did not come from standard input. */
    public Optional<String> requestPath() {
      return request.requestPathValue();
    }

    /** Returns the JSON path when parsing located one precise failing request field. */
    public Optional<String> jsonPath() {
      return json.jsonPathValue();
    }

    /** Returns the request JSON line when the parser exposed one concrete cursor. */
    public Optional<Integer> jsonLine() {
      return json.jsonLineValue();
    }

    /** Returns the request JSON column when the parser exposed one concrete cursor. */
    public Optional<Integer> jsonColumn() {
      return json.jsonColumnValue();
    }

    @Override
    public String stage() {
      return "READ_REQUEST";
    }

    /** Returns one copy enriched with JSON cursor details when none were already present. */
    public ReadRequest withJson(JsonLocation discovered) {
      Objects.requireNonNull(discovered, "discovered must not be null");
      return new ReadRequest(request, json instanceof JsonLocation.Unavailable ? discovered : json);
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

  /** Context for failures that occur while validating request fields before execution. */
  record ValidateRequest(RequestShape request) implements RequestShapeContext {
    public ValidateRequest {
      Objects.requireNonNull(request, "request must not be null");
    }

    @Override
    public String stage() {
      return "VALIDATE_REQUEST";
    }
  }

  /** Context for failures that occur while resolving source-backed authored inputs. */
  record ResolveInputs(RequestShape request, InputReference input) implements RequestShapeContext {
    public ResolveInputs {
      Objects.requireNonNull(request, "request must not be null");
      Objects.requireNonNull(input, "input must not be null");
    }

    /** Returns the authored input kind when one failing binding was identified. */
    public Optional<String> inputKind() {
      return input.inputKindValue();
    }

    /** Returns the authored input path when one failing binding resolved to one file. */
    public Optional<String> inputPath() {
      return input.inputPathValue();
    }

    @Override
    public String stage() {
      return "RESOLVE_INPUTS";
    }
  }

  /** Context for failures that occur while opening the source workbook. */
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

    @Override
    public String stage() {
      return "OPEN_WORKBOOK";
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
    default Optional<String> sheetName() {
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
    record Preflight(RequestShape request, ProblemLocation location) implements ExecuteCalculation {
      public Preflight {
        Objects.requireNonNull(request, "request must not be null");
        Objects.requireNonNull(location, "location must not be null");
      }

      @Override
      public String stage() {
        return "CALCULATION_PREFLIGHT";
      }
    }

    /** Context for failures that occur during immediate calculation execution. */
    record Execution(RequestShape request, ProblemLocation location) implements ExecuteCalculation {
      public Execution {
        Objects.requireNonNull(request, "request must not be null");
        Objects.requireNonNull(location, "location must not be null");
      }

      @Override
      public String stage() {
        return "CALCULATION_EXECUTION";
      }
    }
  }

  /** Context for failures that occur while executing one ordered workbook step. */
  record ExecuteStep(RequestShape request, StepReference step, ProblemLocation location)
      implements ProblemContext {
    public ExecuteStep {
      Objects.requireNonNull(request, "request must not be null");
      Objects.requireNonNull(step, "step must not be null");
      Objects.requireNonNull(location, "location must not be null");
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
    public String stepId() {
      return step.stepId();
    }

    /** Returns the authored step family for the failing workbook step. */
    public String stepKind() {
      return step.stepKind();
    }

    /** Returns the concrete workbook step type for the failing workbook step. */
    public String stepType() {
      return step.stepType();
    }

    /** Returns the localized sheet name when the failing step resolved to one sheet target. */
    public Optional<String> sheetName() {
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

    @Override
    public String stage() {
      return "EXECUTE_STEP";
    }

    /** Returns one copy enriched with exception-derived location facts. */
    public ExecuteStep withLocation(ProblemLocation discovered) {
      Objects.requireNonNull(discovered, "discovered must not be null");
      return new ExecuteStep(request, step, mergeLocation(location, discovered));
    }
  }

  /** Context for failures that occur while persisting the workbook to its destination path. */
  record PersistWorkbook(RequestShape request, PersistenceReference persistence)
      implements ProblemContext {
    public PersistWorkbook {
      Objects.requireNonNull(request, "request must not be null");
      Objects.requireNonNull(persistence, "persistence must not be null");
    }

    /** Returns the decoded request source family when persistence reached that point. */
    public Optional<String> sourceType() {
      return request.sourceTypeValue();
    }

    /** Returns the decoded request persistence family when persistence reached that point. */
    public Optional<String> persistenceType() {
      return request.persistenceTypeValue();
    }

    /** Returns the overwritten source path when persistence targeted the opened workbook. */
    public Optional<String> sourceWorkbookPath() {
      return persistence.sourceWorkbookPathValue();
    }

    /** Returns the explicit save-as path when persistence targeted one new file. */
    public Optional<String> persistencePath() {
      return persistence.persistencePathValue();
    }

    @Override
    public String stage() {
      return "PERSIST_WORKBOOK";
    }
  }

  /** Context for top-level execution failures not attributed to a specific pipeline stage. */
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

    @Override
    public String stage() {
      return "EXECUTE_REQUEST";
    }
  }

  /** Context for failures that occur while writing the JSON response to its destination. */
  record WriteResponse(ResponseOutput output) implements ProblemContext {
    public WriteResponse {
      Objects.requireNonNull(output, "output must not be null");
    }

    /** Returns the authored response path when output targeted one file instead of stdout. */
    public Optional<String> responsePath() {
      return output.responsePathValue();
    }

    @Override
    public String stage() {
      return "WRITE_RESPONSE";
    }
  }

  /** Merges one discovered workbook location into one existing context location. */
  static ProblemLocation mergeLocation(ProblemLocation current, ProblemLocation discovered) {
    return ProblemContextSupport.mergeLocation(current, discovered);
  }
}
