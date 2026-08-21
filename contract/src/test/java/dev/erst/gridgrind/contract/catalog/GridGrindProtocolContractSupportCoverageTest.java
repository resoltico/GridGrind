package dev.erst.gridgrind.contract.catalog;

import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolContractCreatorFixtures.*;
import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolContractMalformedFixtures.*;
import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolContractSupportFixtures.*;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.fasterxml.jackson.annotation.JsonIgnore;
import dev.erst.gridgrind.contract.dto.ProtocolBooleanDefault;
import dev.erst.gridgrind.contract.dto.ProtocolField;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.lang.reflect.Type;
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;

/** Covers record-owned optionality for protocol-contract required-field support. */
class GridGrindProtocolContractSupportCoverageTest {
  @Test
  void requiredFieldsComeFromTheEffectiveCreatorContract() {
    assertTrue(recordContractTypes().contains(OptionalFallbackRecord.class));
    assertEquals(
        List.of("required"),
        GridGrindProtocolContractSupport.requiredFieldNames(OptionalFallbackRecord.class));
    assertEquals(
        List.of("required"),
        GridGrindProtocolContractSupport.requiredFieldNames(ComponentOptionalRecord.class));
    assertEquals(
        List.of("componentOptional"),
        GridGrindProtocolContractSupport.optionalFieldNames(ComponentOptionalRecord.class));
    assertEquals(
        List.of("wireRequired"),
        GridGrindProtocolContractSupport.requiredFieldNames(WireNamedRecord.class));
    assertEquals(
        List.of("wireOptional"),
        GridGrindProtocolContractSupport.optionalFieldNames(WireNamedRecord.class));
    assertEquals(
        List.of("plain", "type"),
        GridGrindProtocolContractSupport.effectiveObjectContract(PlainRecord.class, "type")
            .requiredFields());
    assertEquals(
        List.of("required", "type"),
        GridGrindProtocolContractSupport.effectiveObjectContract(
                OptionalFallbackRecord.class, "type")
            .requiredFields());
    assertEquals(
        List.of("optional"),
        GridGrindProtocolContractSupport.effectiveObjectContract(
                OptionalFallbackRecord.class, "type")
            .optionalFields());
    assertEquals(
        List.of("plain"),
        GridGrindProtocolContractSupport.effectiveObjectContract(PlainRecord.class, "plain")
            .requiredFields());
  }

  @Test
  void requestCatalogOptionalityComesFromRecordComponents() {
    assertCatalogOptionality(WorkbookPlan.class, GridGrindProtocolCatalog.catalog().requestType());
    requestDescriptors()
        .forEach(
            descriptor ->
                assertCatalogOptionality(descriptor.recordType(), descriptor.catalogEntry()));
  }

  @Test
  void primitiveRequestFieldInventoryRetainsTheEffectiveRequiredness() {
    record IgnoredPrimitiveRecord(@JsonIgnore boolean hidden) {}
    record UndeclaredOptionalPrimitive(@ProtocolField(optional = true) boolean enabled) {}
    record BooleanDefaultOnText(
        @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
            String enabled) {}
    record BooleanDefaultOnRequired(
        @ProtocolField(booleanDefault = ProtocolBooleanDefault.FALSE) boolean enabled) {}

    assertEquals(
        List.of(
            new GridGrindProtocolContractSupport.RequestPrimitiveField(
                RequiredPrimitiveRecord.class, "enabled", true, Optional.empty())),
        GridGrindProtocolContractSupport.requestPrimitiveFields(
            Set.of(RequiredPrimitiveRecord.class)));
    assertEquals(
        List.of(
            new GridGrindProtocolContractSupport.RequestPrimitiveField(
                OptionalPrimitiveRecord.class, "enabled", false, Optional.of(false))),
        GridGrindProtocolContractSupport.requestPrimitiveFields(
            Set.of(OptionalPrimitiveRecord.class)));
    assertEquals(
        List.of(),
        GridGrindProtocolContractSupport.requestPrimitiveFields(
            Set.of(IgnoredPrimitiveRecord.class)));
    assertEquals(
        "Optional primitive request fields must declare an explicit default: "
            + UndeclaredOptionalPrimitive.class.getName()
            + ".enabled",
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolContractSupport.effectiveObjectContract(
                        UndeclaredOptionalPrimitive.class))
            .getMessage());
    assertEquals(
        "Only boolean request fields may declare a boolean default: "
            + BooleanDefaultOnText.class.getName()
            + ".enabled",
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolContractSupport.effectiveObjectContract(
                        BooleanDefaultOnText.class))
            .getMessage());
    assertEquals(
        "Only optional request fields may declare a default: "
            + BooleanDefaultOnRequired.class.getName()
            + ".enabled",
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolContractSupport.effectiveObjectContract(
                        BooleanDefaultOnRequired.class))
            .getMessage());
    assertEquals(
        "recordType must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new GridGrindProtocolContractSupport.RequestPrimitiveField(
                        null, "enabled", true, Optional.empty()))
            .getMessage());
    assertEquals(
        "fieldName must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new GridGrindProtocolContractSupport.RequestPrimitiveField(
                        RequiredPrimitiveRecord.class, null, true, Optional.empty()))
            .getMessage());
    assertEquals(
        "fieldName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindProtocolContractSupport.RequestPrimitiveField(
                        RequiredPrimitiveRecord.class, " ", true, Optional.empty()))
            .getMessage());
    assertEquals(
        "required primitive request fields must not declare defaults",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindProtocolContractSupport.RequestPrimitiveField(
                        RequiredPrimitiveRecord.class, "enabled", true, Optional.of(false)))
            .getMessage());
    assertEquals(
        "optional primitive request fields must declare defaults",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindProtocolContractSupport.RequestPrimitiveField(
                        OptionalPrimitiveRecord.class, "enabled", false, Optional.empty()))
            .getMessage());
  }

  @Test
  void contractSupportHandlesCreatorAndTypeGraphEdgeCasesWithoutAssumingTheirAbsence() {
    assertTrue(creatorContractTypes().contains(NullableCreatorRecord.class));
    assertEquals(
        Optional.empty(), GridGrindProtocolContractSupport.discriminatorField(Unannotated.class));
    assertEquals(
        Optional.empty(),
        GridGrindProtocolContractSupport.discriminatorField(BlankDiscriminator.class));
    assertEquals(
        Optional.of("kind"),
        GridGrindProtocolContractSupport.discriminatorField(NamedDiscriminator.class));
    assertEquals(
        List.of("plain"), GridGrindProtocolContractSupport.requiredFieldNames(PlainRecord.class));
    assertEquals(
        List.of("accessorName"),
        GridGrindProtocolContractSupport.requiredFieldNames(AccessorNamedRecord.class));
    assertEquals(
        List.of("blankAnnotation"),
        GridGrindProtocolContractSupport.requiredFieldNames(BlankPropertyRecord.class));
    assertEquals(
        List.of("nullable"),
        GridGrindProtocolContractSupport.requiredFieldNames(NullableRecord.class));
    assertEquals(
        List.of("nonAbsent"),
        GridGrindProtocolContractSupport.optionalFieldNames(NonAbsentRecord.class));
    assertEquals(
        List.of("nonDefault"),
        GridGrindProtocolContractSupport.requiredFieldNames(NonDefaultRecord.class));
    assertFalse(GridGrindProtocolContractSupport.isRequestInputRecord(PlainRecord.class));
    assertTrue(
        GridGrindProtocolContractSupport.requestInputRecordTypes().contains(WorkbookPlan.class));

    assertEquals(
        Optional.empty(),
        GridGrindProtocolContractSupport.creatorParameterType(CreatorlessRecord.class, "enabled"));
    assertEquals(
        Optional.empty(),
        GridGrindProtocolContractSupport.creatorParameterType(
            MultipleCreatorRecord.class, "enabled"));
    assertEquals(
        Optional.of(Boolean.class),
        GridGrindProtocolContractSupport.creatorParameterType(
            NullableCreatorRecord.class, "enabled"));
    assertEquals(
        Optional.empty(),
        GridGrindProtocolContractSupport.creatorParameterType(
            NullableCreatorRecord.class, "missing"));
    assertEquals(
        Optional.empty(),
        GridGrindProtocolContractSupport.creatorParameterType(
            UnnamedCreatorRecord.class, "enabled"));

    Set<Class<? extends Record>> records = new LinkedHashSet<>();
    GridGrindProtocolContractSupport.collectRequestTypes(
        PlainRecord.class, records, new HashSet<>());
    GridGrindProtocolContractSupport.collectRequestTypes(new Type() {}, records, new HashSet<>());
    assertEquals(Set.of(PlainRecord.class), records);
  }

  @Test
  void rejectsMalformedCreatorAndObjectContractsBeforeRequestAnalysisCanUseThem() {
    assertEquals(
        Set.of(
            DuplicateWireNameRecord.class,
            OptionalDiscriminatorRecord.class,
            MismatchedCreatorRecord.class,
            DuplicateCreatorPropertyRecord.class),
        Set.copyOf(malformedContractTypes()));
    assertEquals(
        List.of("required", "optional"),
        new GridGrindProtocolContractSupport.EffectiveObjectContract(
                List.of("required"), List.of("optional"))
            .fields());
    assertEquals(
        "contract field names must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindProtocolContractSupport.EffectiveObjectContract(
                        List.of(" "), List.of()))
            .getMessage());
    assertEquals(
        "contract field names must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindProtocolContractSupport.EffectiveObjectContract(
                        List.of("required"), List.of(" ")))
            .getMessage());
    assertEquals(
        "contract fields must be unique and non-overlapping",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindProtocolContractSupport.EffectiveObjectContract(
                        List.of("duplicate", "duplicate"), List.of()))
            .getMessage());
    assertEquals(
        "contract fields must be unique and non-overlapping",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindProtocolContractSupport.EffectiveObjectContract(
                        List.of("shared"), List.of("shared")))
            .getMessage());
    assertEquals(
        "contract fields must be unique and non-overlapping",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindProtocolContractSupport.EffectiveObjectContract(
                        List.of("required"), List.of("optional", "optional")))
            .getMessage());
    assertEquals(
        "Record exposes duplicate JSON field names: " + DuplicateWireNameRecord.class.getName(),
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolContractSupport.effectiveObjectContract(
                        DuplicateWireNameRecord.class))
            .getMessage());
    assertEquals(
        "JSON creator fields must exactly match visible record fields for "
            + MismatchedCreatorRecord.class.getName()
            + ": creator=[actual], record=[expected]",
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolContractSupport.effectiveObjectContract(
                        MismatchedCreatorRecord.class))
            .getMessage());
    assertEquals(
        "Request record must declare at most one @JsonCreator: "
            + MultipleCreatorRecord.class.getName(),
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolContractSupport.effectiveObjectContract(
                        MultipleCreatorRecord.class))
            .getMessage());
    assertEquals(
        "Every @JsonCreator parameter must declare @JsonProperty: "
            + UnnamedCreatorRecord.class.getName(),
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolContractSupport.effectiveObjectContract(
                        UnnamedCreatorRecord.class))
            .getMessage());
    assertEquals(
        "JSON creator must not declare duplicate property names: "
            + DuplicateCreatorPropertyRecord.class.getName(),
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolContractSupport.effectiveObjectContract(
                        DuplicateCreatorPropertyRecord.class))
            .getMessage());
    assertEquals(
        "A discriminator cannot be optional in the selected request contract: type",
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolContractSupport.effectiveObjectContract(
                        OptionalDiscriminatorRecord.class, "type"))
            .getMessage());
    assertEquals(
        "discriminator must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> GridGrindProtocolContractSupport.discriminatorContract(" "))
            .getMessage());
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
                    catalogEntry.field(primitive.fieldName()).orElseThrow().defaultBoolean(),
                    () ->
                        "Catalog primitive default drifted for "
                            + recordType.getName()
                            + "."
                            + primitive.fieldName()));
  }

  private static Stream<RequestCatalogDescriptor> requestDescriptors() {
    return Stream.concat(
        GridGrindProtocolCatalogTypeDescriptors.ALL_TYPES.stream()
            .map(
                descriptor ->
                    new RequestCatalogDescriptor(descriptor.recordType(), descriptor.typeEntry())),
        Stream.concat(
            GridGrindProtocolCatalogFieldGroupSupport.NESTED_TYPE_GROUPS.stream()
                .flatMap(group -> group.typeDescriptors().stream())
                .map(
                    descriptor ->
                        new RequestCatalogDescriptor(
                            descriptor.recordType(), descriptor.typeEntry())),
            GridGrindProtocolCatalogFieldGroupSupport.PLAIN_TYPE_DESCRIPTORS.stream()
                .map(
                    descriptor ->
                        new RequestCatalogDescriptor(
                            descriptor.recordType(), descriptor.typeEntry()))));
  }

  private record RequestCatalogDescriptor(
      Class<? extends Record> recordType, TypeEntry catalogEntry) {}
}
