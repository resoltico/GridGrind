package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
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

  /** Null-free request-shape summary reused across request-execution failure contexts. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = RequestShape.Unknown.class, name = "UNKNOWN"),
    @JsonSubTypes.Type(value = RequestShape.Known.class, name = "KNOWN")
  })
  sealed interface RequestShape permits RequestShape.Unknown, RequestShape.Known {
    /** Returns the explicit unknown request-shape variant. */
    static RequestShape unknown() {
      return new RequestShape.Unknown();
    }

    /** Returns the decoded request-shape variant with source and persistence families. */
    static RequestShape known(String sourceType, String persistenceType) {
      return new Known(sourceType, persistenceType);
    }

    /** Returns the known request-shape facts when decoding reached that point. */
    default Optional<Known> known() {
      return switch (this) {
        case Known known -> Optional.of(known);
        case RequestShape.Unknown _ -> Optional.empty();
      };
    }

    /** Returns the decoded request source family when known. */
    default Optional<String> sourceTypeValue() {
      return known().map(Known::sourceType);
    }

    /** Returns the decoded request persistence family when known. */
    default Optional<String> persistenceTypeValue() {
      return known().map(Known::persistenceType);
    }

    /** Request details were unavailable because request decoding or validation never completed. */
    record Unknown() implements RequestShape {}

    /** Source and persistence families were known from the decoded request. */
    record Known(String sourceType, String persistenceType) implements RequestShape {
      public Known {
        sourceType = requireNonBlank(sourceType, "sourceType");
        persistenceType = requireNonBlank(persistenceType, "persistenceType");
      }
    }
  }

  /** Concrete request input source, never encoded with a nullable path. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = RequestInput.StandardInput.class, name = "STANDARD_INPUT"),
    @JsonSubTypes.Type(value = RequestInput.RequestFile.class, name = "FILE")
  })
  sealed interface RequestInput permits RequestInput.StandardInput, RequestInput.RequestFile {
    /** Returns the request-input variant for standard input. */
    static RequestInput standardInput() {
      return new StandardInput();
    }

    /** Returns the request-input variant for one concrete file path. */
    static RequestInput requestFile(String requestPath) {
      return new RequestFile(requireNonBlank(requestPath, "requestPath"));
    }

    /** Returns the request file path when input came from one file. */
    default Optional<String> requestPathValue() {
      return switch (this) {
        case RequestFile requestFile -> Optional.of(requestFile.requestPath());
        case StandardInput _ -> Optional.empty();
      };
    }

    /** JSON request was read from standard input. */
    record StandardInput() implements RequestInput {}

    /** JSON request was read from one file path. */
    record RequestFile(String requestPath) implements RequestInput {
      public RequestFile {
        requestPath = requireNonBlank(requestPath, "requestPath");
      }
    }
  }

  /** CLI response destination, never encoded with a nullable path. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = ResponseOutput.StandardOutput.class, name = "STANDARD_OUTPUT"),
    @JsonSubTypes.Type(value = ResponseOutput.ResponseFile.class, name = "FILE")
  })
  sealed interface ResponseOutput
      permits ResponseOutput.StandardOutput, ResponseOutput.ResponseFile {
    /** Returns the response-output variant for standard output. */
    static ResponseOutput standardOutput() {
      return new StandardOutput();
    }

    /** Returns the response-output variant for one concrete file path. */
    static ResponseOutput responseFile(String responsePath) {
      return new ResponseFile(requireNonBlank(responsePath, "responsePath"));
    }

    /** Returns the response file path when output targets one file. */
    default Optional<String> responsePathValue() {
      return switch (this) {
        case ResponseFile responseFile -> Optional.of(responseFile.responsePath());
        case StandardOutput _ -> Optional.empty();
      };
    }

    /** JSON response was written to standard output. */
    record StandardOutput() implements ResponseOutput {}

    /** JSON response was written to one file path. */
    record ResponseFile(String responsePath) implements ResponseOutput {
      public ResponseFile {
        responsePath = requireNonBlank(responsePath, "responsePath");
      }
    }
  }

  /** Request JSON cursor for malformed payloads, replacing nullable path/line/column slots. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = JsonLocation.Unavailable.class, name = "UNAVAILABLE"),
    @JsonSubTypes.Type(value = JsonLocation.LineColumn.class, name = "LINE_COLUMN"),
    @JsonSubTypes.Type(value = JsonLocation.Located.class, name = "LOCATED")
  })
  sealed interface JsonLocation
      permits JsonLocation.Unavailable, JsonLocation.LineColumn, JsonLocation.Located {
    /** Returns the explicit unavailable JSON-location variant. */
    static JsonLocation unavailable() {
      return new Unavailable();
    }

    /** Returns a JSON-location variant with parser line and column only. */
    static JsonLocation lineColumn(Integer jsonLine, Integer jsonColumn) {
      Objects.requireNonNull(jsonLine, "jsonLine must not be null");
      Objects.requireNonNull(jsonColumn, "jsonColumn must not be null");
      return new LineColumn(jsonLine, jsonColumn);
    }

    /** Returns a JSON-location variant with path, line, and column. */
    static JsonLocation located(String jsonPath, Integer jsonLine, Integer jsonColumn) {
      Objects.requireNonNull(jsonLine, "jsonLine must not be null");
      Objects.requireNonNull(jsonColumn, "jsonColumn must not be null");
      return new Located(requireNonBlank(jsonPath, "jsonPath"), jsonLine, jsonColumn);
    }

    /** Returns the JSON path when one precise request cursor was captured. */
    default Optional<String> jsonPathValue() {
      return switch (this) {
        case Located located -> Optional.of(located.jsonPath());
        case LineColumn _ -> Optional.empty();
        case Unavailable _ -> Optional.empty();
      };
    }

    /** Returns the JSON line when one precise request cursor was captured. */
    default Optional<Integer> jsonLineValue() {
      return switch (this) {
        case LineColumn lineColumn -> Optional.of(lineColumn.jsonLine());
        case Located located -> Optional.of(located.jsonLine());
        case Unavailable _ -> Optional.empty();
      };
    }

    /** Returns the JSON column when one precise request cursor was captured. */
    default Optional<Integer> jsonColumnValue() {
      return switch (this) {
        case LineColumn lineColumn -> Optional.of(lineColumn.jsonColumn());
        case Located located -> Optional.of(located.jsonColumn());
        case Unavailable _ -> Optional.empty();
      };
    }

    /** No JSON cursor could be derived for the request failure. */
    record Unavailable() implements JsonLocation {}

    /** Only the request line and column were available from the parser. */
    record LineColumn(int jsonLine, int jsonColumn) implements JsonLocation {
      public LineColumn {
        if (jsonLine < 1) {
          throw new IllegalArgumentException("jsonLine must be greater than 0");
        }
        if (jsonColumn < 1) {
          throw new IllegalArgumentException("jsonColumn must be greater than 0");
        }
      }
    }

    /** One precise JSON cursor extracted from a payload failure. */
    record Located(String jsonPath, int jsonLine, int jsonColumn) implements JsonLocation {
      public Located {
        jsonPath = requireNonBlank(jsonPath, "jsonPath");
        if (jsonLine < 1) {
          throw new IllegalArgumentException("jsonLine must be greater than 0");
        }
        if (jsonColumn < 1) {
          throw new IllegalArgumentException("jsonColumn must be greater than 0");
        }
      }
    }
  }

  /** Optional CLI argument reference without nullable string padding. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = CliArgument.Unknown.class, name = "UNKNOWN"),
    @JsonSubTypes.Type(value = CliArgument.Named.class, name = "NAMED")
  })
  sealed interface CliArgument permits CliArgument.Unknown, CliArgument.Named {
    /** Returns the explicit unknown CLI-argument variant. */
    static CliArgument unknown() {
      return new CliArgument.Unknown();
    }

    /** Returns the CLI-argument variant for one named flag or operand. */
    static CliArgument named(String argument) {
      return new Named(requireNonBlank(argument, "argument"));
    }

    /** Returns the concrete CLI argument that triggered parsing failure when known. */
    default Optional<String> argumentValue() {
      return switch (this) {
        case Named named -> Optional.of(named.argument());
        case CliArgument.Unknown _ -> Optional.empty();
      };
    }

    /** The parse failure was not attributable to one specific argument. */
    record Unknown() implements CliArgument {}

    /** The parse failure was attributable to one named option or operand. */
    record Named(String argument) implements CliArgument {
      public Named {
        argument = requireNonBlank(argument, "argument");
      }
    }
  }

  /** Source-backed authored input reference without nullable kind/path slots. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = InputReference.Unknown.class, name = "UNKNOWN"),
    @JsonSubTypes.Type(value = InputReference.KindOnly.class, name = "KIND"),
    @JsonSubTypes.Type(value = InputReference.PathReference.class, name = "PATH")
  })
  sealed interface InputReference
      permits InputReference.Unknown, InputReference.KindOnly, InputReference.PathReference {
    /** Returns the explicit unknown input-reference variant. */
    static InputReference unknown() {
      return new InputReference.Unknown();
    }

    /** Returns the input-reference variant with one authored input kind and no path. */
    static InputReference kind(String inputKind) {
      return new KindOnly(requireNonBlank(inputKind, "inputKind"));
    }

    /** Returns the input-reference variant with one authored input kind and concrete path. */
    static InputReference path(String inputKind, String inputPath) {
      return new PathReference(
          requireNonBlank(inputKind, "inputKind"), requireNonBlank(inputPath, "inputPath"));
    }

    /** Returns the authored input kind when the failing binding was identified. */
    default Optional<String> inputKindValue() {
      return switch (this) {
        case KindOnly reference -> Optional.of(reference.inputKind());
        case PathReference reference -> Optional.of(reference.inputKind());
        case InputReference.Unknown _ -> Optional.empty();
      };
    }

    /** Returns the authored input path when the failing binding was identified. */
    default Optional<String> inputPathValue() {
      return switch (this) {
        case PathReference reference -> Optional.of(reference.inputPath());
        case KindOnly _ -> Optional.empty();
        case InputReference.Unknown _ -> Optional.empty();
      };
    }

    /** No input-binding path could be derived for the resolution failure. */
    record Unknown() implements InputReference {}

    /** The failing authored input kind was known, but no concrete path applied. */
    record KindOnly(String inputKind) implements InputReference {
      public KindOnly {
        inputKind = requireNonBlank(inputKind, "inputKind");
      }
    }

    /** One concrete authored input binding resolved to one path. */
    record PathReference(String inputKind, String inputPath) implements InputReference {
      public PathReference {
        inputKind = requireNonBlank(inputKind, "inputKind");
        inputPath = requireNonBlank(inputPath, "inputPath");
      }
    }
  }

  /** Workbook open target without nullable source-path fields. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = WorkbookReference.NewWorkbook.class, name = "NEW"),
    @JsonSubTypes.Type(value = WorkbookReference.ExistingFile.class, name = "EXISTING")
  })
  sealed interface WorkbookReference
      permits WorkbookReference.NewWorkbook, WorkbookReference.ExistingFile {
    /** Returns the workbook-reference variant for a brand-new in-memory workbook. */
    static WorkbookReference newWorkbook() {
      return new NewWorkbook();
    }

    /** Returns the workbook-reference variant for one existing `.xlsx` file path. */
    static WorkbookReference existingFile(String path) {
      return new ExistingFile(requireNonBlank(path, "path"));
    }

    /** Returns the source workbook path when execution opened one existing file. */
    default Optional<String> sourceWorkbookPathValue() {
      return switch (this) {
        case ExistingFile existingFile -> Optional.of(existingFile.path());
        case NewWorkbook _ -> Optional.empty();
      };
    }

    /** Workbook creation started from a brand-new in-memory workbook. */
    record NewWorkbook() implements WorkbookReference {}

    /** Workbook opening targeted one existing `.xlsx` file. */
    record ExistingFile(String path) implements WorkbookReference {
      public ExistingFile {
        path = requireNonBlank(path, "path");
      }
    }
  }

  /** Persistence attempt reference without nullable source/output path padding. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = PersistenceReference.OverwriteSource.class, name = "OVERWRITE"),
    @JsonSubTypes.Type(value = PersistenceReference.SaveAs.class, name = "SAVE_AS")
  })
  sealed interface PersistenceReference
      permits PersistenceReference.OverwriteSource, PersistenceReference.SaveAs {
    /** Returns the persistence-reference variant that targets the opened source workbook. */
    static PersistenceReference overwriteSource(String sourceWorkbookPath) {
      return new OverwriteSource(requireNonBlank(sourceWorkbookPath, "sourceWorkbookPath"));
    }

    /** Returns the persistence-reference variant that targets one explicit save-as path. */
    static PersistenceReference saveAs(String persistencePath) {
      return new SaveAs(requireNonBlank(persistencePath, "persistencePath"));
    }

    /** Returns the overwritten source path when persistence targeted the opened workbook. */
    default Optional<String> sourceWorkbookPathValue() {
      return switch (this) {
        case OverwriteSource overwriteSource -> Optional.of(overwriteSource.sourceWorkbookPath());
        case SaveAs _ -> Optional.empty();
      };
    }

    /** Returns the explicit SAVE_AS destination path when persistence targeted one new file. */
    default Optional<String> persistencePathValue() {
      return switch (this) {
        case SaveAs saveAs -> Optional.of(saveAs.persistencePath());
        case OverwriteSource _ -> Optional.empty();
      };
    }

    /** Workbook persistence attempted to overwrite the opened source file. */
    record OverwriteSource(String sourceWorkbookPath) implements PersistenceReference {
      public OverwriteSource {
        sourceWorkbookPath = requireNonBlank(sourceWorkbookPath, "sourceWorkbookPath");
      }
    }

    /** Workbook persistence attempted to write one explicit SAVE_AS path. */
    record SaveAs(String persistencePath) implements PersistenceReference {
      public SaveAs {
        persistencePath = requireNonBlank(persistencePath, "persistencePath");
      }
    }
  }

  /** Workbook location reference used by calculation and step failures without null padding. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = ProblemLocation.Unknown.class, name = "UNKNOWN"),
    @JsonSubTypes.Type(value = ProblemLocation.Sheet.class, name = "SHEET"),
    @JsonSubTypes.Type(value = ProblemLocation.Address.class, name = "ADDRESS"),
    @JsonSubTypes.Type(value = ProblemLocation.Cell.class, name = "CELL"),
    @JsonSubTypes.Type(value = ProblemLocation.RangeOnly.class, name = "RANGE_ONLY"),
    @JsonSubTypes.Type(value = ProblemLocation.Range.class, name = "RANGE"),
    @JsonSubTypes.Type(value = ProblemLocation.NamedRange.class, name = "NAMED_RANGE"),
    @JsonSubTypes.Type(value = ProblemLocation.SheetNamedRange.class, name = "SHEET_NAMED_RANGE"),
    @JsonSubTypes.Type(value = ProblemLocation.FormulaCell.class, name = "FORMULA_CELL")
  })
  sealed interface ProblemLocation
      permits ProblemLocation.Unknown,
          ProblemLocation.Sheet,
          ProblemLocation.Address,
          ProblemLocation.Cell,
          ProblemLocation.RangeOnly,
          ProblemLocation.Range,
          ProblemLocation.NamedRange,
          ProblemLocation.SheetNamedRange,
          ProblemLocation.FormulaCell {
    /** Returns the explicit unknown workbook-location variant. */
    static ProblemLocation unknown() {
      return new ProblemLocation.Unknown();
    }

    /** Returns the workbook-location variant for one whole sheet. */
    static ProblemLocation sheet(String sheetName) {
      return new Sheet(requireNonBlank(sheetName, "sheetName"));
    }

    /** Returns the workbook-location variant for one address without a resolved sheet. */
    static ProblemLocation address(String address) {
      return new Address(requireNonBlank(address, "address"));
    }

    /** Returns the workbook-location variant for one concrete sheet-backed cell. */
    static ProblemLocation cell(String sheetName, String address) {
      return new Cell(requireNonBlank(sheetName, "sheetName"), requireNonBlank(address, "address"));
    }

    /** Returns the workbook-location variant for one range without a resolved sheet. */
    static ProblemLocation range(String range) {
      return new RangeOnly(requireNonBlank(range, "range"));
    }

    /** Returns the workbook-location variant for one concrete sheet-backed range. */
    static ProblemLocation range(String sheetName, String range) {
      return new Range(requireNonBlank(sheetName, "sheetName"), requireNonBlank(range, "range"));
    }

    /** Returns the workbook-location variant for one workbook-scoped named range. */
    static ProblemLocation namedRange(String name) {
      return new NamedRange(requireNonBlank(name, "name"));
    }

    /** Returns the workbook-location variant for one sheet-scoped named range. */
    static ProblemLocation namedRange(String sheetName, String name) {
      return new SheetNamedRange(
          requireNonBlank(sheetName, "sheetName"), requireNonBlank(name, "name"));
    }

    /** Returns the workbook-location variant for one formula-bearing cell. */
    static ProblemLocation formulaCell(String sheetName, String address, String formula) {
      return new FormulaCell(
          requireNonBlank(sheetName, "sheetName"),
          requireNonBlank(address, "address"),
          requireNonBlank(formula, "formula"));
    }

    /** Returns the sheet name when the failure was localized to one sheet-backed target. */
    default Optional<String> sheetNameValue() {
      return switch (this) {
        case Sheet sheet -> Optional.of(sheet.sheetName());
        case Address _ -> Optional.empty();
        case Cell cell -> Optional.of(cell.sheetName());
        case RangeOnly _ -> Optional.empty();
        case Range range -> Optional.of(range.sheetName());
        case NamedRange _ -> Optional.empty();
        case SheetNamedRange namedRange -> Optional.of(namedRange.sheetName());
        case FormulaCell formulaCell -> Optional.of(formulaCell.sheetName());
        case ProblemLocation.Unknown _ -> Optional.empty();
      };
    }

    /** Returns the cell address when the failure was localized to one concrete cell. */
    default Optional<String> addressValue() {
      return switch (this) {
        case Address address -> Optional.of(address.address());
        case Cell cell -> Optional.of(cell.address());
        case FormulaCell formulaCell -> Optional.of(formulaCell.address());
        default -> Optional.empty();
      };
    }

    /** Returns the range when the failure was localized to one concrete rectangular range. */
    default Optional<String> rangeValue() {
      return switch (this) {
        case RangeOnly range -> Optional.of(range.range());
        case Range range -> Optional.of(range.range());
        default -> Optional.empty();
      };
    }

    /** Returns the named-range identifier when the failure was localized to one named range. */
    default Optional<String> namedRangeNameValue() {
      return switch (this) {
        case NamedRange namedRange -> Optional.of(namedRange.name());
        case SheetNamedRange namedRange -> Optional.of(namedRange.name());
        default -> Optional.empty();
      };
    }

    /** Returns the formula when the failure was localized to one formula-bearing cell. */
    default Optional<String> formulaValue() {
      if (this instanceof FormulaCell formulaCell) {
        return Optional.of(formulaCell.formula());
      }
      return Optional.empty();
    }

    /** The failure could not be localized to one workbook target. */
    record Unknown() implements ProblemLocation {}

    /** The failure was localized to one whole sheet. */
    record Sheet(String sheetName) implements ProblemLocation {
      public Sheet {
        sheetName = requireNonBlank(sheetName, "sheetName");
      }
    }

    /** The failure was localized to one concrete address without a resolved sheet. */
    record Address(String address) implements ProblemLocation {
      public Address {
        address = requireNonBlank(address, "address");
      }
    }

    /** The failure was localized to one concrete cell. */
    record Cell(String sheetName, String address) implements ProblemLocation {
      public Cell {
        sheetName = requireNonBlank(sheetName, "sheetName");
        address = requireNonBlank(address, "address");
      }
    }

    /** The failure was localized to one concrete range without a resolved sheet. */
    record RangeOnly(String range) implements ProblemLocation {
      public RangeOnly {
        range = requireNonBlank(range, "range");
      }
    }

    /** The failure was localized to one concrete rectangular range. */
    record Range(String sheetName, String range) implements ProblemLocation {
      public Range {
        sheetName = requireNonBlank(sheetName, "sheetName");
        range = requireNonBlank(range, "range");
      }
    }

    /** The failure was localized to one named range. */
    record NamedRange(String name) implements ProblemLocation {
      public NamedRange {
        name = requireNonBlank(name, "name");
      }
    }

    /** The failure was localized to one named range scoped to one sheet. */
    record SheetNamedRange(String sheetName, String name) implements ProblemLocation {
      public SheetNamedRange {
        sheetName = requireNonBlank(sheetName, "sheetName");
        name = requireNonBlank(name, "name");
      }
    }

    /** The failure was localized to one formula-bearing cell. */
    record FormulaCell(String sheetName, String address, String formula)
        implements ProblemLocation {
      public FormulaCell {
        sheetName = requireNonBlank(sheetName, "sheetName");
        address = requireNonBlank(address, "address");
        formula = requireNonBlank(formula, "formula");
      }
    }
  }

  /** Stable step identity captured for step execution failures. */
  record StepReference(int stepIndex, String stepId, String stepKind, String stepType) {
    public StepReference {
      if (stepIndex < 0) {
        throw new IllegalArgumentException("stepIndex must be greater than or equal to 0");
      }
      stepId = requireNonBlank(stepId, "stepId");
      stepKind = requireNonBlank(stepKind, "stepKind");
      stepType = requireNonBlank(stepType, "stepType");
    }
  }

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

  /** Context for failures that occur while validating request fields before execution. */
  record ValidateRequest(RequestShape request) implements ProblemContext {
    public ValidateRequest {
      Objects.requireNonNull(request, "request must not be null");
    }

    /** Returns the decoded request source family when validation reached that point. */
    public Optional<String> sourceType() {
      return request.sourceTypeValue();
    }

    /** Returns the decoded request persistence family when validation reached that point. */
    public Optional<String> persistenceType() {
      return request.persistenceTypeValue();
    }

    @Override
    public String stage() {
      return "VALIDATE_REQUEST";
    }
  }

  /** Context for failures that occur while resolving source-backed authored inputs. */
  record ResolveInputs(RequestShape request, InputReference input) implements ProblemContext {
    public ResolveInputs {
      Objects.requireNonNull(request, "request must not be null");
      Objects.requireNonNull(input, "input must not be null");
    }

    /** Returns the decoded request source family when input resolution reached that point. */
    public Optional<String> sourceType() {
      return request.sourceTypeValue();
    }

    /** Returns the decoded request persistence family when input resolution reached that point. */
    public Optional<String> persistenceType() {
      return request.persistenceTypeValue();
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
  record OpenWorkbook(RequestShape request, WorkbookReference workbook) implements ProblemContext {
    public OpenWorkbook {
      Objects.requireNonNull(request, "request must not be null");
      Objects.requireNonNull(workbook, "workbook must not be null");
    }

    /** Returns the decoded request source family when workbook opening reached that point. */
    public Optional<String> sourceType() {
      return request.sourceTypeValue();
    }

    /** Returns the decoded request persistence family when workbook opening reached that point. */
    public Optional<String> persistenceType() {
      return request.persistenceTypeValue();
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
  sealed interface ExecuteCalculation extends ProblemContext
      permits ExecuteCalculation.Preflight, ExecuteCalculation.Execution {
    /** Returns the authored request shape active for the calculation failure. */
    RequestShape request();

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

    /** Returns the decoded request source family when known. */
    default Optional<String> sourceType() {
      return request().sourceTypeValue();
    }

    /** Returns the decoded request persistence family when known. */
    default Optional<String> persistenceType() {
      return request().persistenceTypeValue();
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

  private static String requireNonBlank(String value, String fieldName) {
    return ProblemContextSupport.requireNonBlank(value, fieldName);
  }
}
