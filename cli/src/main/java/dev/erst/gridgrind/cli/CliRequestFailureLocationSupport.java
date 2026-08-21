package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.contract.json.PayloadException;
import dev.erst.gridgrind.contract.json.PayloadLocation;
import java.util.Optional;

/** Projects a payload exception's owned location into the public CLI diagnostic context. */
final class CliRequestFailureLocationSupport {
  private CliRequestFailureLocationSupport() {}

  static JsonLocation locationFor(Throwable failure) {
    return payloadException(failure)
        .map(CliRequestFailureLocationSupport::locationForPayload)
        .orElseGet(JsonLocation::unavailable);
  }

  private static Optional<PayloadException> payloadException(Throwable failure) {
    Throwable current = failure;
    while (current != null) {
      if (current instanceof PayloadException payloadException) {
        return Optional.of(payloadException);
      }
      current = current.getCause();
    }
    return Optional.empty();
  }

  private static JsonLocation locationForPayload(PayloadException failure) {
    return switch (failure.jsonLocation()) {
      case PayloadLocation.Located located ->
          JsonLocation.located(
              located.jsonPathValue(), located.jsonLineValue(), located.jsonColumnValue());
      case PayloadLocation.PathOnly pathOnly -> JsonLocation.pathOnly(pathOnly.jsonPathValue());
      case PayloadLocation.LineColumn lineColumn ->
          JsonLocation.lineColumn(lineColumn.jsonLineValue(), lineColumn.jsonColumnValue());
      case PayloadLocation.Unavailable _ -> JsonLocation.unavailable();
    };
  }
}
