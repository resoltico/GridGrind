package dev.erst.gridgrind.json.codec;

import dev.erst.gridgrind.core.vocabulary.WireScalar;
import tools.jackson.core.JsonGenerator;
import tools.jackson.databind.SerializationContext;
import tools.jackson.databind.ValueSerializer;

/** Writes every validated domain scalar as its scalar wire value. */
final class WireScalarSerializer extends ValueSerializer<WireScalar> {
  @Override
  public void serialize(WireScalar value, JsonGenerator generator, SerializationContext context) {
    context.writeValue(generator, value.scalarValue());
  }
}
