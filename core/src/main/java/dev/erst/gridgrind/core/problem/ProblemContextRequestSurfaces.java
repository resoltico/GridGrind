package dev.erst.gridgrind.core.problem;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;
import java.util.Optional;

/** Request and host surface facts reused across problem-context stages. */
public interface ProblemContextRequestSurfaces {
  /** Null-free request-shape summary reused across request-execution failure contexts. */
  @Union(discriminator = "type")
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
    @Variant(
        id = "UNKNOWN",
        summary =
            "Request details were unavailable because request decoding or validation never completed.")
    record Unknown() implements RequestShape {}

    /** Source and persistence families were known from the decoded request. */
    @Variant(
        id = "KNOWN",
        summary = "Source and persistence families were known from the decoded request.")
    record Known(String sourceType, String persistenceType) implements RequestShape {
      public Known {
        sourceType = requireNonBlank(sourceType, "sourceType");
        persistenceType = requireNonBlank(persistenceType, "persistenceType");
      }
    }
  }

  /** Concrete request input source, never encoded with a nullable path. */
  @Union(discriminator = "type")
  sealed interface RequestInput permits RequestInput.Stdin, RequestInput.RequestFile {
    /** Returns the request-input variant for standard input. */
    static RequestInput standardInput() {
      return new Stdin();
    }

    /** Returns the request-input variant for one concrete file path. */
    static RequestInput requestFile(String requestPath) {
      return new RequestFile(requireNonBlank(requestPath, "requestPath"));
    }

    /** Returns the request file path when input came from one file. */
    default Optional<String> requestPathValue() {
      return switch (this) {
        case RequestFile requestFile -> Optional.of(requestFile.requestPath());
        case Stdin _ -> Optional.empty();
      };
    }

    /** JSON request was read from standard input. */
    @Variant(id = "STANDARD_INPUT", summary = "JSON request was read from standard input.")
    record Stdin() implements RequestInput {}

    /** JSON request was read from one file path. */
    @Variant(id = "FILE", summary = "JSON request was read from one file path.")
    record RequestFile(String requestPath) implements RequestInput {
      public RequestFile {
        requestPath = requireNonBlank(requestPath, "requestPath");
      }
    }
  }

  /** Host response destination, never encoded with a nullable path. */
  @Union(discriminator = "type")
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
    @Variant(id = "STANDARD_OUTPUT", summary = "JSON response was written to standard output.")
    record StandardOutput() implements ResponseOutput {}

    /** JSON response was written to one file path. */
    @Variant(id = "FILE", summary = "JSON response was written to one file path.")
    record ResponseFile(String responsePath) implements ResponseOutput {
      public ResponseFile {
        responsePath = requireNonBlank(responsePath, "responsePath");
      }
    }
  }

  /** Request JSON cursor for malformed payloads, replacing nullable path/line/column slots. */
  @Union(discriminator = "type")
  sealed interface JsonLocation
      permits JsonLocation.Unavailable,
          JsonLocation.PathOnly,
          JsonLocation.ByteOffset,
          JsonLocation.PathAtByteOffset,
          JsonLocation.ByteOffsetLineColumn,
          JsonLocation.DuplicateKey,
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

    /** Returns a JSON-location variant with one exact UTF-8 byte offset. */
    static JsonLocation byteOffset(long byteOffset) {
      return new ByteOffset(byteOffset);
    }

    /** Returns a syntax location with machine-precise offset and human-editable coordinates. */
    static JsonLocation byteOffsetLineColumn(long byteOffset, int jsonLine, int jsonColumn) {
      return new ByteOffsetLineColumn(byteOffset, jsonLine, jsonColumn);
    }

    /** Returns a JSON-location variant with a path and exact UTF-8 byte offset. */
    static JsonLocation pathAtByteOffset(String jsonPath, long byteOffset) {
      return new PathAtByteOffset(requireNonBlank(jsonPath, "jsonPath"), byteOffset);
    }

    /** Returns a duplicate-property location that remains distinguishable without a JSON path. */
    static JsonLocation duplicateKey(
        String containingObjectPath, String key, int occurrenceOrdinal, long byteOffset) {
      return new DuplicateKey(containingObjectPath, key, occurrenceOrdinal, byteOffset);
    }

    /** Returns a JSON-location variant with path, line, and column. */
    static JsonLocation located(String jsonPath, Integer jsonLine, Integer jsonColumn) {
      Objects.requireNonNull(jsonLine, "jsonLine must not be null");
      Objects.requireNonNull(jsonColumn, "jsonColumn must not be null");
      return new Located(requireNonBlank(jsonPath, "jsonPath"), jsonLine, jsonColumn);
    }

    /** Returns the JSON path when one precise request cursor was captured. */
    default Optional<String> jsonPathValue() {
      if (this instanceof Located located) {
        return Optional.of(located.jsonPath());
      }
      if (this instanceof PathOnly pathOnly) {
        return Optional.of(pathOnly.jsonPath());
      }
      if (this instanceof PathAtByteOffset located) {
        return Optional.of(located.jsonPath());
      }
      return Optional.empty();
    }

    /** Returns the exact UTF-8 byte offset when a structural parser located one token. */
    default Optional<Long> byteOffsetValue() {
      if (this instanceof ByteOffset located) {
        return Optional.of(located.byteOffset());
      }
      if (this instanceof PathAtByteOffset located) {
        return Optional.of(located.byteOffset());
      }
      if (this instanceof ByteOffsetLineColumn located) {
        return Optional.of(located.byteOffset());
      }
      if (this instanceof DuplicateKey duplicate) {
        return Optional.of(duplicate.byteOffset());
      }
      return Optional.empty();
    }

    /** Returns duplicate-key identity when a property occurrence cannot have one JSON path. */
    default Optional<DuplicateKey> duplicateKeyValue() {
      return this instanceof DuplicateKey duplicate ? Optional.of(duplicate) : Optional.empty();
    }

    /** Returns the JSON line when one precise request cursor was captured. */
    default Optional<Integer> jsonLineValue() {
      if (this instanceof LineColumn lineColumn) {
        return Optional.of(lineColumn.jsonLine());
      }
      if (this instanceof Located located) {
        return Optional.of(located.jsonLine());
      }
      if (this instanceof ByteOffsetLineColumn located) {
        return Optional.of(located.jsonLine());
      }
      return Optional.empty();
    }

    /** Returns the JSON column when one precise request cursor was captured. */
    default Optional<Integer> jsonColumnValue() {
      if (this instanceof LineColumn lineColumn) {
        return Optional.of(lineColumn.jsonColumn());
      }
      if (this instanceof Located located) {
        return Optional.of(located.jsonColumn());
      }
      if (this instanceof ByteOffsetLineColumn located) {
        return Optional.of(located.jsonColumn());
      }
      return Optional.empty();
    }

    /** No JSON cursor could be derived for the request failure. */
    @Variant(
        id = "UNAVAILABLE",
        summary = "No JSON cursor could be derived for the request failure.")
    record Unavailable() implements JsonLocation {}

    /** Only the JSON path was available from request validation. */
    @Variant(
        id = "PATH_ONLY",
        summary = "Only the JSON path was available from request validation.")
    record PathOnly(String jsonPath) implements JsonLocation {
      public PathOnly {
        jsonPath = requireNonBlank(jsonPath, "jsonPath");
      }
    }

    /** A structural parser located an offending token without a JSON path. */
    @Variant(
        id = "BYTE_OFFSET",
        summary = "A structural parser located an offending token without a JSON path.")
    record ByteOffset(long byteOffset) implements JsonLocation {
      public ByteOffset {
        requireByteOffset(byteOffset);
      }
    }

    /** A structural parser located an offending object member or scalar token precisely. */
    @Variant(
        id = "PATH_BYTE_OFFSET",
        summary =
            "A structural parser located an offending object member or scalar token precisely.")
    record PathAtByteOffset(String jsonPath, long byteOffset) implements JsonLocation {
      public PathAtByteOffset {
        jsonPath = requireNonBlank(jsonPath, "jsonPath");
        requireByteOffset(byteOffset);
      }
    }

    /** A syntax location with both exact UTF-8 offset and one-based line/column coordinates. */
    @Variant(
        id = "BYTE_OFFSET_LINE_COLUMN",
        summary =
            "A syntax location with both exact UTF-8 offset and one-based line/column coordinates.")
    record ByteOffsetLineColumn(long byteOffset, int jsonLine, int jsonColumn)
        implements JsonLocation {
      public ByteOffsetLineColumn {
        requireByteOffset(byteOffset);
        if (jsonLine < 1 || jsonColumn < 1) {
          throw new IllegalArgumentException("jsonLine and jsonColumn must be greater than 0");
        }
      }
    }

    /** One duplicate property occurrence, which has no unique JSON path of its own. */
    @Variant(
        id = "DUPLICATE_KEY",
        summary = "One duplicate property occurrence, which has no unique JSON path of its own.")
    record DuplicateKey(
        String containingObjectPath, String key, int occurrenceOrdinal, long byteOffset)
        implements JsonLocation {
      public DuplicateKey {
        Objects.requireNonNull(containingObjectPath, "containingObjectPath must not be null");
        key = requireNonBlank(key, "key");
        if (occurrenceOrdinal < 0) {
          throw new IllegalArgumentException("occurrenceOrdinal must not be negative");
        }
        requireByteOffset(byteOffset);
      }
    }

    /** Only the request line and column were available from the parser. */
    @Variant(
        id = "LINE_COLUMN",
        summary = "Only the request line and column were available from the parser.")
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
    @Variant(id = "LOCATED", summary = "One precise JSON cursor extracted from a payload failure.")
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

    private static void requireByteOffset(long byteOffset) {
      if (byteOffset < 0) {
        throw new IllegalArgumentException("byteOffset must not be negative");
      }
    }
  }

  /** Optional host argument reference without nullable string padding. */
  @Union(discriminator = "type")
  sealed interface HostArgument permits HostArgument.Unknown, HostArgument.Named {
    /** Returns the explicit unknown host-argument variant. */
    static HostArgument unknown() {
      return new HostArgument.Unknown();
    }

    /** Returns the host-argument variant for one named flag or operand. */
    static HostArgument named(String argument) {
      return new Named(requireNonBlank(argument, "argument"));
    }

    /** Returns the concrete host argument that triggered parsing failure when known. */
    default Optional<String> argumentValue() {
      return switch (this) {
        case Named named -> Optional.of(named.argument());
        case HostArgument.Unknown _ -> Optional.empty();
      };
    }

    /** The parse failure was not attributable to one specific argument. */
    @Variant(
        id = "UNKNOWN",
        summary = "The parse failure was not attributable to one specific argument.")
    record Unknown() implements HostArgument {}

    /** The parse failure was attributable to one named option or operand. */
    @Variant(
        id = "NAMED",
        summary = "The parse failure was attributable to one named option or operand.")
    record Named(String argument) implements HostArgument {
      public Named {
        argument = requireNonBlank(argument, "argument");
      }
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    return TextValidation.requireNonBlank(value, fieldName);
  }
}
