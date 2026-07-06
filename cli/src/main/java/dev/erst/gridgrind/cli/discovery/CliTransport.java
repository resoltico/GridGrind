package dev.erst.gridgrind.cli.discovery;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Optional;

/** Optional transport metadata for CLI diagnostics written somewhere other than stderr. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "wroteTo")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CliTransport.StandardOutput.class, name = "STDOUT"),
  @JsonSubTypes.Type(value = CliTransport.ResponseFile.class, name = "FILE")
})
public sealed interface CliTransport
    permits CliTransport.StandardOutput, CliTransport.ResponseFile {
  /** Returns the transport variant for stdout fallback payloads. */
  static CliTransport standardOutput() {
    return new StandardOutput();
  }

  /** Returns the transport variant for one concrete response file path. */
  static CliTransport responseFile(String responsePath) {
    return new ResponseFile(CliDiscoveryValidation.requireNonBlank(responsePath, "responsePath"));
  }

  /** Returns the response-file path only when the diagnostic was persisted to one file. */
  default Optional<String> responsePathValue() {
    return switch (this) {
      case ResponseFile responseFile -> Optional.of(responseFile.responsePath());
      case StandardOutput _ -> Optional.empty();
    };
  }

  /** The diagnostic payload was written to stdout as the fallback channel. */
  record StandardOutput() implements CliTransport {}

  /** The diagnostic payload was written to one response file path. */
  record ResponseFile(String responsePath) implements CliTransport {
    public ResponseFile {
      responsePath = CliDiscoveryValidation.requireNonBlank(responsePath, "responsePath");
    }
  }
}
