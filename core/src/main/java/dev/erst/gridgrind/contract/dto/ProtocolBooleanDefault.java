package dev.erst.gridgrind.contract.dto;

import java.util.Optional;

/** Declares the effective wire default for an optional boolean request field. */
public enum ProtocolBooleanDefault {
  /** No boolean default is declared. */
  UNSPECIFIED,
  /** Omission resolves to {@code false}. */
  FALSE,
  /** Omission resolves to {@code true}. */
  TRUE;

  /** Returns the declared boolean value when this annotation value represents one. */
  public Optional<Boolean> value() {
    return switch (this) {
      case UNSPECIFIED -> Optional.empty();
      case FALSE -> Optional.of(false);
      case TRUE -> Optional.of(true);
    };
  }

  /** Resolves one nullable boundary value through this declared default. */
  public boolean resolve(Boolean value) {
    if (value != null) {
      return value;
    }
    return switch (this) {
      case UNSPECIFIED ->
          throw new IllegalStateException(
              "An unspecified boolean default cannot normalize a value");
      case FALSE -> false;
      case TRUE -> true;
    };
  }
}
