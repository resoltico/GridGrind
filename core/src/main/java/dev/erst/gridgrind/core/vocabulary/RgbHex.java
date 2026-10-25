package dev.erst.gridgrind.core.vocabulary;

/** An RGB literal in canonical uppercase #RRGGBB form. */
public record RgbHex(String value) implements WireString {
  public RgbHex {
    value = ProtocolRgbColorSupport.requireRgbHex(value, "rgb");
  }

  @Override
  public String toString() {
    return value;
  }
}
