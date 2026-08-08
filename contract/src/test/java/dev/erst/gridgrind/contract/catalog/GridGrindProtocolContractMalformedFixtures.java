package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.contract.dto.ProtocolField;
import java.util.List;

/** Synthetic malformed request shapes that prove contract-definition failures remain explicit. */
final class GridGrindProtocolContractMalformedFixtures {
  private GridGrindProtocolContractMalformedFixtures() {}

  static List<Class<? extends Record>> malformedContractTypes() {
    return List.of(
        DuplicateWireNameRecord.class,
        OptionalDiscriminatorRecord.class,
        MismatchedCreatorRecord.class,
        DuplicateCreatorPropertyRecord.class);
  }

  record DuplicateWireNameRecord(
      @JsonProperty("same") String first, @JsonProperty("same") String second) {}

  record OptionalDiscriminatorRecord(@ProtocolField(optional = true) String type) {}

  record MismatchedCreatorRecord(String expected) {
    @JsonCreator
    static MismatchedCreatorRecord create(@JsonProperty("actual") String actual) {
      return new MismatchedCreatorRecord(actual);
    }
  }

  record DuplicateCreatorPropertyRecord(String first, String second) {
    @JsonCreator
    static DuplicateCreatorPropertyRecord create(
        @JsonProperty("value") String first, @JsonProperty("value") String second) {
      return new DuplicateCreatorPropertyRecord(first, second);
    }
  }
}
