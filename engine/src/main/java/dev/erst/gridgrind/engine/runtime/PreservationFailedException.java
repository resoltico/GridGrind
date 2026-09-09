package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.Preservation;
import java.io.IOException;
import java.util.List;

/** Signals that a host-required before-and-after preservation comparison did not hold. */
final class PreservationFailedException extends IOException {
  private static final long serialVersionUID = 1L;
  private final Preservation.Kind kind;
  private final List<String> identifiers;

  PreservationFailedException(Preservation.Kind kind, List<String> identifiers) {
    super(messageFor(kind));
    this.kind = kind;
    this.identifiers = List.copyOf(identifiers);
  }

  Preservation.Kind kind() {
    return kind;
  }

  List<String> identifiers() {
    return identifiers;
  }

  private static String messageFor(Preservation.Kind kind) {
    return switch (kind) {
      case SEMANTIC ->
          "host-required semantic preservation no longer matches its before-mutation facts";
      case BYTE ->
          "host-required opaque OOXML part preservation no longer matches its source bytes";
    };
  }
}
