package dev.erst.gridgrind.contract.selector;

import java.lang.reflect.RecordComponent;
import tools.jackson.core.JsonGenerator;
import tools.jackson.databind.SerializationContext;
import tools.jackson.databind.ValueSerializer;

/** Serializes selectors through the canonical leaf selector type id plus record fields. */
final class SelectorJsonSerializer extends ValueSerializer<Selector> {
  @Override
  public void serialize(Selector selector, JsonGenerator generator, SerializationContext context) {
    Class<?> selectorType = selector.getClass();
    generator.writeStartObject(selector);
    generator.writeStringProperty("type", SelectorJsonSupport.typeIdFor(selectorType));
    for (RecordComponent component : selectorType.getRecordComponents()) {
      generator.writePOJOProperty(component.getName(), readComponent(selector, component));
    }
    generator.writeEndObject();
  }

  static Object readComponent(Selector selector, RecordComponent component) {
    try {
      return component.getAccessor().invoke(selector);
    } catch (ReflectiveOperationException | IllegalArgumentException exception) {
      throw new IllegalStateException(
          "Unable to read selector component "
              + component.getName()
              + " from "
              + selector.getClass().getName(),
          exception);
    }
  }
}
