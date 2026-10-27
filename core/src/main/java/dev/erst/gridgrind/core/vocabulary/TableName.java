package dev.erst.gridgrind.core.vocabulary;

/** A workbook table identifier governed by Excel defined-name rules. */
public record TableName(String value) implements WireString {
  public TableName {
    value = ProtocolDefinedNameValidation.validateName(value);
  }

  @Override
  public String toString() {
    return value;
  }
}
