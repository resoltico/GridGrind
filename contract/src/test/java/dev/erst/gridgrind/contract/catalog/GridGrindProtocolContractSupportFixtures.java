package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.dto.ProtocolBooleanDefault;
import dev.erst.gridgrind.contract.dto.ProtocolField;
import java.util.Optional;

/** Synthetic request shapes that isolate JSON-creator contract edge cases. */
final class GridGrindProtocolContractSupportFixtures {
  private GridGrindProtocolContractSupportFixtures() {}

  static java.util.List<Class<? extends Record>> recordContractTypes() {
    return java.util.List.of(
        OptionalFallbackRecord.class,
        ComponentOptionalRecord.class,
        WireNamedRecord.class,
        PlainRecord.class,
        AccessorNamedRecord.class,
        BlankPropertyRecord.class,
        NullableRecord.class,
        NonAbsentRecord.class,
        NonDefaultRecord.class);
  }

  record OptionalFallbackRecord(String required, Optional<String> optional) {}

  record ComponentOptionalRecord(
      String required, @ProtocolField(optional = true) String componentOptional) {}

  record WireNamedRecord(
      @JsonProperty("wireRequired") String required,
      @JsonProperty("wireOptional") Optional<String> optional) {}

  /** Type without a discriminator annotation. */
  interface Unannotated {}

  /** Type with an unusable blank discriminator annotation. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = " ")
  interface BlankDiscriminator {}

  /** Type with an explicitly named discriminator. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
  interface NamedDiscriminator {}

  record PlainRecord(String plain) {}

  record AccessorNamedRecord(String sourceName) {
    @Override
    @JsonProperty("accessorName")
    public String sourceName() {
      return sourceName;
    }
  }

  record BlankPropertyRecord(@JsonProperty("") String blankAnnotation) {}

  record NullableRecord(@org.jspecify.annotations.Nullable String nullable) {}

  record NonAbsentRecord(@JsonInclude(JsonInclude.Include.NON_ABSENT) String nonAbsent) {}

  record NonDefaultRecord(@JsonInclude(JsonInclude.Include.NON_DEFAULT) String nonDefault) {}

  record RequiredPrimitiveRecord(@ProtocolField boolean enabled) {}

  record OptionalPrimitiveRecord(
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean enabled) {}
}
