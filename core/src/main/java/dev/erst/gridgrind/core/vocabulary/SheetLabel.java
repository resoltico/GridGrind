package dev.erst.gridgrind.core.vocabulary;

/** Observed worksheet text, including labels that violate Excel's authoring-name rules. */
public record SheetLabel(String value) implements WireString {
  public SheetLabel {
    value = TextValidation.requireNonBlank(value, "sheetName");
  }

  @Override
  public String toString() {
    return value;
  }
}
