package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** Captures whether a sheet is protected and, if so, with which supported lock flags. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
@JsonSubTypes({
  @JsonSubTypes.Type(value = SheetProtectionReport.Unprotected.class, name = "UNPROTECTED"),
  @JsonSubTypes.Type(value = SheetProtectionReport.Protected.class, name = "PROTECTED")
})
public sealed interface SheetProtectionReport
    permits SheetProtectionReport.Unprotected, SheetProtectionReport.Protected {
  /** Sheet protection is disabled. */
  record Unprotected() implements SheetProtectionReport {}

  /** Sheet protection is enabled with the reported supported lock flags. */
  record Protected(SheetProtectionSettings settings) implements SheetProtectionReport {
    public Protected {
      Objects.requireNonNull(settings, "settings must not be null");
    }
  }
}
