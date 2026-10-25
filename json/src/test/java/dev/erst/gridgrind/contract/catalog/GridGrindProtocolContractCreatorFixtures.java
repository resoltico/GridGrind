package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.contract.dto.ProtocolBooleanDefault;
import dev.erst.gridgrind.contract.dto.ProtocolField;

/** Synthetic creator signatures for primitive optionality boundary checks. */
final class GridGrindProtocolContractCreatorFixtures {
  private GridGrindProtocolContractCreatorFixtures() {}

  static java.util.List<Class<? extends Record>> creatorContractTypes() {
    return java.util.List.of(
        CreatorlessRecord.class,
        MultipleCreatorRecord.class,
        NullableCreatorRecord.class,
        UnnamedCreatorRecord.class);
  }

  record CreatorlessRecord(
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean enabled) {}

  record MultipleCreatorRecord(
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean enabled) {
    @JsonCreator
    private MultipleCreatorRecord(@JsonProperty("enabled") Boolean enabled) {
      this(ProtocolBooleanDefault.FALSE.resolve(enabled));
    }

    @JsonCreator
    static MultipleCreatorRecord create(@JsonProperty("enabled") Boolean enabled) {
      return new MultipleCreatorRecord(enabled);
    }
  }

  record NullableCreatorRecord(
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean enabled) {
    @JsonCreator
    static NullableCreatorRecord create(@JsonProperty("enabled") Boolean enabled) {
      return new NullableCreatorRecord(ProtocolBooleanDefault.FALSE.resolve(enabled));
    }
  }

  record UnnamedCreatorRecord(
      @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
          boolean enabled) {
    @JsonCreator
    static UnnamedCreatorRecord create(Boolean enabled) {
      return new UnnamedCreatorRecord(ProtocolBooleanDefault.FALSE.resolve(enabled));
    }
  }
}
