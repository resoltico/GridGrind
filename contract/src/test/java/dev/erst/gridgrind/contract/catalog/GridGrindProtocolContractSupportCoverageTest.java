package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.ConcurrentHashMap;
import org.junit.jupiter.api.Test;

/** Covers fallback and invariant branches for protocol-contract required-field support. */
class GridGrindProtocolContractSupportCoverageTest {
  @Test
  void requiredFieldFallbacksHonorOptionalComponentsMetadataAndConflictProtection() {
    assertEquals(
        List.of("required"),
        GridGrindProtocolContractSupport.requiredFieldNames(OptionalFallbackRecord.class));
    assertEquals(
        List.of("required"),
        GridGrindProtocolContractSupport.requiredFieldNames(MetadataOptionalRecord.class));

    Map<Class<? extends Record>, List<String>> requiredFieldsByRecordType =
        new ConcurrentHashMap<>();
    GridGrindProtocolContractSupport.register(
        requiredFieldsByRecordType, OptionalFallbackRecord.class, List.of("required"));
    IllegalStateException conflict =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindProtocolContractSupport.register(
                    requiredFieldsByRecordType,
                    OptionalFallbackRecord.class,
                    List.of("different")));
    assertEquals(
        "Conflicting required-field definitions for "
            + OptionalFallbackRecord.class.getName()
            + ": [required] vs [different]",
        conflict.getMessage());
  }

  private record OptionalFallbackRecord(String required, Optional<String> optional) {}

  @ProtocolTypeMetadata(
      id = "METADATA_OPTIONAL_RECORD",
      summary = "Synthetic record used to cover metadata-owned optional fields",
      optionalFields = {"metadataOptional"})
  private record MetadataOptionalRecord(String required, String metadataOptional) {}
}
