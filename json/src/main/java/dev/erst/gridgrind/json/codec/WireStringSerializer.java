package dev.erst.gridgrind.json.codec;

import dev.erst.gridgrind.core.vocabulary.WireString;
import tools.jackson.core.JsonGenerator;
import tools.jackson.databind.SerializationContext;
import tools.jackson.databind.ValueSerializer;

/** Writes every validated domain string as its scalar wire value. */
final class WireStringSerializer extends ValueSerializer<WireString> {
  @Override
  public void serialize(WireString value, JsonGenerator generator, SerializationContext context) {
    generator.writeString(value.value());
  }
}
