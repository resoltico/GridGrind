package dev.erst.gridgrind.core.vocabulary;

/** A validated Excel defined-name identifier. */
public record DefinedName(String value) implements WireString {
  public DefinedName {
    value = ProtocolDefinedNameValidation.validateName(value);
  }

  @Override
  public String toString() {
    return value;
  }
}
