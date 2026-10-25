package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput;
import java.util.Optional;

/** Shared request transport and JSON-location facts for request-intake diagnostics. */
public sealed interface RequestInputContext extends ProblemContext
    permits ProblemContext.ReadRequest, ProblemContext.BindRequest {
  /** Returns the carrier identifying standard input or one authored request file. */
  RequestInput request();

  /** Returns the precise JSON location available for the request failure. */
  JsonLocation json();

  /** Returns the authored request file path when the request did not come from standard input. */
  default Optional<String> requestPath() {
    return request().requestPathValue();
  }

  /** Returns the JSON path when intake located one precise failing request field. */
  default Optional<String> jsonPath() {
    return json().jsonPathValue();
  }

  /** Returns the exact UTF-8 request byte offset when intake located one token. */
  default Optional<Long> byteOffset() {
    return json().byteOffsetValue();
  }

  /** Returns duplicate-key identity when one property occurrence cannot have a unique path. */
  default Optional<JsonLocation.DuplicateKey> duplicateKey() {
    return json().duplicateKeyValue();
  }

  /** Returns the request JSON line when the parser exposed one concrete cursor. */
  default Optional<Integer> jsonLine() {
    return json().jsonLineValue();
  }

  /** Returns the request JSON column when the parser exposed one concrete cursor. */
  default Optional<Integer> jsonColumn() {
    return json().jsonColumnValue();
  }
}
