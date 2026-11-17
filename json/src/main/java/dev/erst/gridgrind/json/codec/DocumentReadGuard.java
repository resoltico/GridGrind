package dev.erst.gridgrind.json.codec;

import tools.jackson.core.JsonParser;
import tools.jackson.databind.DeserializationContext;
import tools.jackson.databind.ValueDeserializer;
import tools.jackson.databind.util.AccessPattern;

/** Prevents complete request and authority documents from bypassing their owned intake APIs. */
final class DocumentReadGuard<T> extends ValueDeserializer<T> {
  private final Class<T> type;
  private final String reader;

  DocumentReadGuard(Class<T> type, String reader) {
    this.type = type;
    this.reader = reader;
  }

  @Override
  public T deserialize(JsonParser parser, DeserializationContext context) {
    throw unsupportedRead();
  }

  @Override
  public T getNullValue(DeserializationContext context) {
    throw unsupportedRead();
  }

  @Override
  public T getAbsentValue(DeserializationContext context) {
    throw unsupportedRead();
  }

  @Override
  public AccessPattern getNullAccessPattern() {
    return AccessPattern.DYNAMIC;
  }

  @Override
  public Class<?> handledType() {
    return type;
  }

  private IllegalStateException unsupportedRead() {
    return new IllegalStateException(
        "Direct "
            + type.getSimpleName()
            + " document deserialization is unsupported; use "
            + reader
            + " instead.");
  }
}
