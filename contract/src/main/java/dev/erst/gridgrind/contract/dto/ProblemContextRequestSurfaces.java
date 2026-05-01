package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;
import java.util.Optional;

/** Request and CLI surface facts reused across problem-context stages. */
public interface ProblemContextRequestSurfaces {
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
    @JsonSubTypes.Type(value = JsonLocation.PathOnly.class, name = "PATH_ONLY"),
    @JsonSubTypes.Type(value = JsonLocation.LineColumn.class, name = "LINE_COLUMN"),
    @JsonSubTypes.Type(value = JsonLocation.Located.class, name = "LOCATED")
  })
  sealed interface JsonLocation
      permits JsonLocation.Unavailable,
          JsonLocation.PathOnly,
          JsonLocation.LineColumn,
          JsonLocation.Located {
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

    /** Returns a JSON-location variant with path only. */
    static JsonLocation pathOnly(String jsonPath) {
      return new PathOnly(requireNonBlank(jsonPath, "jsonPath"));
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
        case PathOnly pathOnly -> Optional.of(pathOnly.jsonPath());
        case LineColumn _ -> Optional.empty();
        case Unavailable _ -> Optional.empty();
      };
    }

    /** Returns the JSON line when one precise request cursor was captured. */
    default Optional<Integer> jsonLineValue() {
      return switch (this) {
        case LineColumn lineColumn -> Optional.of(lineColumn.jsonLine());
        case Located located -> Optional.of(located.jsonLine());
        case PathOnly _ -> Optional.empty();
        case Unavailable _ -> Optional.empty();
      };
    }

    /** Returns the JSON column when one precise request cursor was captured. */
    default Optional<Integer> jsonColumnValue() {
      return switch (this) {
        case LineColumn lineColumn -> Optional.of(lineColumn.jsonColumn());
        case Located located -> Optional.of(located.jsonColumn());
        case PathOnly _ -> Optional.empty();
        case Unavailable _ -> Optional.empty();
      };
    }

    /** No JSON cursor could be derived for the request failure. */
    record Unavailable() implements JsonLocation {}

    /** Only the JSON path was available from request validation. */
    record PathOnly(String jsonPath) implements JsonLocation {
      public PathOnly {
        jsonPath = requireNonBlank(jsonPath, "jsonPath");
      }
    }

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

  private static String requireNonBlank(String value, String fieldName) {
    return ProblemContextSupport.requireNonBlank(value, fieldName);
  }
}
