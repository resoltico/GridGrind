package dev.erst.gridgrind.core.vocabulary;

/** A nonblank defined-name label observed in a workbook, without authoring restrictions. */
public record DefinedNameLabel(String value) implements WireString {
  public DefinedNameLabel {
    value = TextValidation.requireNonBlank(value, "name");
  }

  @Override
  public String toString() {
    return value;
  }
}
