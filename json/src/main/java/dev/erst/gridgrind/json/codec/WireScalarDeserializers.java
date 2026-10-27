package dev.erst.gridgrind.json.codec;

import dev.erst.gridgrind.core.vocabulary.WireScalar;
import org.jspecify.annotations.Nullable;
import tools.jackson.databind.BeanDescription;
import tools.jackson.databind.DeserializationConfig;
import tools.jackson.databind.JavaType;
import tools.jackson.databind.ValueDeserializer;
import tools.jackson.databind.module.SimpleDeserializers;

/** Selects one constructor-backed scalar codec for each concrete primitive-domain value. */
final class WireScalarDeserializers extends SimpleDeserializers {
  private static final long serialVersionUID = 1L;

  @Override
  public @Nullable ValueDeserializer<?> findBeanDeserializer(
      JavaType type, DeserializationConfig config, BeanDescription.Supplier beanDescription) {
    Class<?> rawType = type.getRawClass();
    if (WireScalar.isValueType(rawType)) {
      return new WireScalarDeserializer(rawType.asSubclass(WireScalar.class));
    }
    return super.findBeanDeserializer(type, config, beanDescription);
  }
}
