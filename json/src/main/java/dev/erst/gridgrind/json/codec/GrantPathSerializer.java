package dev.erst.gridgrind.json.codec;

import java.nio.file.Path;
import tools.jackson.core.JsonGenerator;
import tools.jackson.databind.SerializationContext;
import tools.jackson.databind.ValueSerializer;

/** Writes canonical grant paths as filesystem strings. */
final class GrantPathSerializer extends ValueSerializer<Path> {
  @Override
  public void serialize(Path value, JsonGenerator generator, SerializationContext context) {
    generator.writeString(value.toString());
  }
}
