package dev.erst.gridgrind.json.codec;

import dev.erst.gridgrind.core.metadata.ModelDefaults;
import dev.erst.gridgrind.core.metadata.WireDefault;
import java.lang.reflect.RecordComponent;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import tools.jackson.core.JsonGenerator;
import tools.jackson.databind.BeanDescription;
import tools.jackson.databind.SerializationConfig;
import tools.jackson.databind.SerializationContext;
import tools.jackson.databind.ser.BeanPropertyWriter;
import tools.jackson.databind.ser.ValueSerializerModifier;

/** Omits only component values explicitly declared as omitted wire defaults. */
final class RecordDefaultSerialization extends ValueSerializerModifier {
  private static final long serialVersionUID = 1L;

  @Override
  public List<BeanPropertyWriter> changeProperties(
      SerializationConfig config,
      BeanDescription.Supplier bean,
      List<BeanPropertyWriter> properties) {
    if (!bean.isRecordType()) {
      return properties;
    }
    List<BeanPropertyWriter> writers = new ArrayList<>(properties.size());
    for (BeanPropertyWriter property : properties) {
      RecordComponent component =
          Arrays.stream(bean.getBeanClass().getRecordComponents())
              .filter(candidate -> candidate.getName().equals(property.getName()))
              .findFirst()
              .orElse(null);
      WireDefault metadata = component == null ? null : component.getAnnotation(WireDefault.class);
      writers.add(
          component != null && metadata != null && metadata.omit()
              ? new DefaultWriter(property, component)
              : property);
    }
    return writers;
  }

  /** Retains ordinary property handling unless the declared default matches. */
  private static final class DefaultWriter extends BeanPropertyWriter {
    private static final long serialVersionUID = 1L;
    private final RecordComponent component;

    DefaultWriter(BeanPropertyWriter original, RecordComponent component) {
      super(original);
      this.component = component;
    }

    @Override
    @SuppressWarnings("PMD.UseConcurrentHashMap") // Each serialization owns its sibling values.
    public void serializeAsProperty(Object bean, JsonGenerator output, SerializationContext context)
        throws Exception {
      Map<String, Object> values = new LinkedHashMap<>();
      for (RecordComponent sibling : component.getDeclaringRecord().getRecordComponents()) {
        values.put(sibling.getName(), sibling.getAccessor().invoke(bean));
      }
      if (!Objects.equals(ModelDefaults.value(component, values), get(bean))) {
        super.serializeAsProperty(bean, output, context);
      }
    }
  }
}
