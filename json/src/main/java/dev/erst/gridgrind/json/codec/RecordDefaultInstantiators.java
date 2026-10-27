package dev.erst.gridgrind.json.codec;

import dev.erst.gridgrind.core.metadata.ModelDefaults;
import dev.erst.gridgrind.core.metadata.WireDefault;
import java.lang.reflect.RecordComponent;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.Map;
import tools.jackson.databind.BeanDescription;
import tools.jackson.databind.DeserializationConfig;
import tools.jackson.databind.DeserializationContext;
import tools.jackson.databind.deser.NullValueProvider;
import tools.jackson.databind.deser.SettableBeanProperty;
import tools.jackson.databind.deser.ValueInstantiator;
import tools.jackson.databind.deser.bean.PropertyValueBuffer;
import tools.jackson.databind.module.SimpleValueInstantiators;
import tools.jackson.databind.util.AccessPattern;

/** Supplies absent component defaults before the canonical record constructor is invoked. */
final class RecordDefaultInstantiators extends SimpleValueInstantiators {
  private static final long serialVersionUID = 1L;

  @Override
  public ValueInstantiator modifyValueInstantiator(
      DeserializationConfig config, BeanDescription.Supplier bean, ValueInstantiator original) {
    return bean.isRecordType()
            && Arrays.stream(bean.getBeanClass().getRecordComponents())
                .anyMatch(component -> component.isAnnotationPresent(WireDefault.class))
        ? new Defaults(original, bean.getBeanClass())
        : original;
  }

  private record Missing(RecordComponent component) {}

  /** Decorates the canonical record creator with absence resolution. */
  private static final class Defaults extends ValueInstantiator.Delegating {
    private static final long serialVersionUID = 1L;
    private final Class<?> recordType;

    Defaults(ValueInstantiator original, Class<?> type) {
      super(original);
      recordType = type;
    }

    @Override
    public SettableBeanProperty[] getFromObjectArguments(DeserializationConfig config) {
      SettableBeanProperty[] properties = super.getFromObjectArguments(config);
      if (properties == null) {
        throw new IllegalStateException(
            "Defaulted record has no canonical object creator: " + recordType);
      }
      properties = properties.clone();
      for (int index = 0; index < properties.length; index++) {
        SettableBeanProperty property = properties[index];
        RecordComponent component =
            recordType
                .getRecordComponents()[
                ModelDefaults.componentIndex(recordType, property.getName())];
        if (component.isAnnotationPresent(WireDefault.class)) {
          properties[index] = property.withNullProvider(new DefaultProvider(component));
        }
      }
      return properties;
    }

    @Override
    public Object createFromObjectWith(DeserializationContext context, Object[] original) {
      Object[] arguments = original.clone();
      boolean[] resolving = new boolean[arguments.length];
      for (int index = 0; index < arguments.length; index++) {
        resolve(arguments, index, resolving);
      }
      return super.createFromObjectWith(context, arguments);
    }

    @SuppressWarnings("PMD.UseConcurrentHashMap") // Each binding owns its ordered sibling values.
    private void resolve(Object[] values, int index, boolean[] resolving) {
      if (!(values[index] instanceof Missing missing)) {
        return;
      }
      if (resolving[index]) {
        throw new IllegalStateException("Cyclic component default in " + recordType);
      }
      resolving[index] = true;
      WireDefault metadata = missing.component().getAnnotation(WireDefault.class);
      if (!metadata.from().isEmpty()) {
        resolve(values, ModelDefaults.componentIndex(recordType, metadata.from()), resolving);
      }
      Map<String, Object> siblings = new LinkedHashMap<>();
      RecordComponent[] components = recordType.getRecordComponents();
      for (int position = 0; position < components.length; position++) {
        siblings.put(components[position].getName(), values[position]);
      }
      values[index] = ModelDefaults.value(missing.component(), siblings);
      resolving[index] = false;
    }

    @Override
    public Object createFromObjectWith(
        DeserializationContext context,
        SettableBeanProperty[] properties,
        PropertyValueBuffer buffer) {
      return createFromObjectWith(context, buffer.getParameters(context, properties));
    }
  }

  private record DefaultProvider(RecordComponent component) implements NullValueProvider {
    @Override
    public Object getNullValue(DeserializationContext context) {
      return context.reportInputMismatch(
          component.getType(), "Explicit null is not an omission default");
    }

    @Override
    public Object getAbsentValue(DeserializationContext context) {
      return new Missing(component);
    }

    @Override
    public AccessPattern getNullAccessPattern() {
      return AccessPattern.DYNAMIC;
    }
  }
}
