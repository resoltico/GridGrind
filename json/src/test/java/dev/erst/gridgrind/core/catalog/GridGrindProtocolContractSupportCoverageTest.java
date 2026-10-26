package dev.erst.gridgrind.core.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.metadata.CatalogIgnored;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import java.util.HashSet;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;

/** Verifies canonical record presence, explicit defaults and catalog consumers. */
class GridGrindProtocolContractSupportCoverageTest {
  private record Presence(
      String required,
      Optional<String> optional,
      @WireDefault("fallback") String defaulted,
      @CatalogIgnored String internal) {}

  private record RequiredPrimitive(boolean enabled) {}

  private record DefaultPrimitive(@WireDefault("true") boolean enabled) {}

  private record UndeclaredPrimitive(@ProtocolField(optional = true) boolean enabled) {}

  @Test
  void canonicalComponentsOwnPresenceAndDefaults() {
    assertEquals(
        List.of("required"), GridGrindProtocolContractSupport.requiredFieldNames(Presence.class));
    assertEquals(
        List.of("optional", "defaulted"),
        GridGrindProtocolContractSupport.optionalFieldNames(Presence.class));
    assertEquals(
        List.of("required", "type"),
        GridGrindProtocolContractSupport.effectiveObjectContract(Presence.class, "type")
            .requiredFields());
    assertThrows(
        IllegalStateException.class,
        () -> GridGrindProtocolContractSupport.effectiveObjectContract(Presence.class, "optional"));
  }

  @Test
  void primitiveInventoryDistinguishesRequiredFieldsAndExplicitDefaults() {
    assertEquals(
        List.of(
            new GridGrindProtocolContractSupport.RequestPrimitiveField(
                RequiredPrimitive.class, "enabled", true, Optional.empty())),
        GridGrindProtocolContractSupport.requestPrimitiveFields(Set.of(RequiredPrimitive.class)));
    assertEquals(
        List.of(
            new GridGrindProtocolContractSupport.RequestPrimitiveField(
                DefaultPrimitive.class, "enabled", false, Optional.of(true))),
        GridGrindProtocolContractSupport.requestPrimitiveFields(Set.of(DefaultPrimitive.class)));
    assertThrows(
        IllegalStateException.class,
        () -> GridGrindProtocolContractSupport.effectiveObjectContract(UndeclaredPrimitive.class));
  }

  @Test
  void domainStringsStayScalarInTheRequestGraph() {
    Set<Class<? extends Record>> records = new HashSet<>();
    GridGrindProtocolContractSupport.collectRequestTypes(
        PlanHeader.class, records, new HashSet<>());
    assertEquals(Set.of(PlanHeader.class), records);
    assertTrue(
        GridGrindProtocolContractSupport.requestInputRecordTypes().contains(WorkbookPlan.class));
  }

  @Test
  void invalidObjectContractsAreRejected() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindProtocolContractSupport.EffectiveObjectContract(
                List.of("a"), List.of("a")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindProtocolContractSupport.EffectiveObjectContract(List.of(""), List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> GridGrindProtocolContractSupport.discriminatorContract(" "));
  }

  @Test
  void requestCatalogOptionalityComesFromRecordComponents() {
    assertCatalogOptionality(WorkbookPlan.class, GridGrindProtocolCatalog.catalog().requestType());
    requestDescriptors()
        .forEach(
            descriptor ->
                assertCatalogOptionality(descriptor.recordType(), descriptor.catalogEntry()));
  }

  private static void assertCatalogOptionality(
      Class<? extends Record> recordType, TypeEntry catalogEntry) {
    if (!GridGrindProtocolContractSupport.isRequestInputRecord(recordType)) {
      return;
    }
    Set<String> expected =
        Set.copyOf(GridGrindProtocolContractSupport.optionalFieldNames(recordType));
    Set<String> actual =
        catalogEntry.fields().stream()
            .filter(field -> field.requirement() == FieldRequirement.OPTIONAL)
            .map(FieldEntry::name)
            .collect(java.util.stream.Collectors.toUnmodifiableSet());
    assertEquals(expected, actual, () -> "Catalog optionality drifted for " + recordType.getName());
    GridGrindProtocolContractSupport.requestPrimitiveFields().stream()
        .filter(field -> field.recordType().equals(recordType))
        .forEach(
            primitive ->
                assertEquals(
                    primitive.defaultBoolean(),
                    catalogEntry
                        .field(primitive.fieldName())
                        .orElseThrow()
                        .defaultValue()
                        .filter(Boolean.class::isInstance)
                        .map(Boolean.class::cast),
                    () ->
                        "Catalog primitive default drifted for "
                            + recordType.getName()
                            + "."
                            + primitive.fieldName()));
  }

  private static Stream<RequestCatalogDescriptor> requestDescriptors() {
    return Stream.concat(
        CatalogTypes.ALL_TYPES.stream()
            .map(
                descriptor ->
                    new RequestCatalogDescriptor(descriptor.recordType(), descriptor.typeEntry())),
        Stream.concat(
            CatalogTypes.NESTED_TYPE_GROUPS.stream()
                .flatMap(group -> group.typeDescriptors().stream())
                .map(
                    descriptor ->
                        new RequestCatalogDescriptor(
                            descriptor.recordType(), descriptor.typeEntry())),
            CatalogTypes.PLAIN_TYPE_DESCRIPTORS.stream()
                .map(
                    descriptor ->
                        new RequestCatalogDescriptor(
                            descriptor.recordType(), descriptor.typeEntry()))));
  }

  private record RequestCatalogDescriptor(
      Class<? extends Record> recordType, TypeEntry catalogEntry) {}

  private record PlanHeader(Optional<dev.erst.gridgrind.core.vocabulary.PlanId> planId) {}
}
