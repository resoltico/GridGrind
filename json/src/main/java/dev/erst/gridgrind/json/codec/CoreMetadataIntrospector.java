package dev.erst.gridgrind.json.codec;

import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.util.List;
import org.jspecify.annotations.Nullable;
import tools.jackson.databind.cfg.MapperConfig;
import tools.jackson.databind.introspect.Annotated;
import tools.jackson.databind.introspect.AnnotatedClass;
import tools.jackson.databind.introspect.JacksonAnnotationIntrospector;
import tools.jackson.databind.jsontype.NamedType;

/** Derives wire identity from core metadata while host DTOs retain their own annotations. */
final class CoreMetadataIntrospector extends JacksonAnnotationIntrospector {
  private static final long serialVersionUID = 1L;

  @Override
  public JsonTypeInfo.@Nullable Value findPolymorphicTypeInfo(
      MapperConfig<?> config, Annotated annotated) {
    Union union = annotated.getAnnotation(Union.class);
    if (union == null || union.discriminator().isBlank()) {
      return super.findPolymorphicTypeInfo(config, annotated);
    }
    return JsonTypeInfo.Value.construct(
        JsonTypeInfo.Id.NAME,
        JsonTypeInfo.As.PROPERTY,
        union.discriminator(),
        null,
        false,
        null,
        null);
  }

  @Override
  public @Nullable List<NamedType> findSubtypes(MapperConfig<?> config, Annotated annotated) {
    if (annotated.getAnnotation(Union.class) == null) {
      return super.findSubtypes(config, annotated);
    }
    return annotated.getRawType().isSealed()
        ? List.of(GridGrindJsonSubtypeSupport.namedLeafSubtypes(annotated.getRawType()))
        : List.of();
  }

  @Override
  public @Nullable String findTypeName(MapperConfig<?> config, AnnotatedClass annotated) {
    Variant variant = annotated.getAnnotation(Variant.class);
    return variant == null ? super.findTypeName(config, annotated) : variant.id();
  }
}
