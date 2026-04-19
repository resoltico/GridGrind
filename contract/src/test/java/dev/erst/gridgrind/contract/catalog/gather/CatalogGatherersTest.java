package dev.erst.gridgrind.contract.catalog.gather;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.catalog.FieldEntry;
import dev.erst.gridgrind.contract.catalog.FieldRequirement;
import dev.erst.gridgrind.contract.catalog.FieldShape;
import dev.erst.gridgrind.contract.catalog.ScalarType;
import java.lang.reflect.RecordComponent;
import java.util.Arrays;
import java.util.List;
import java.util.Set;
import org.junit.jupiter.api.Test;

/** Direct tests for the small gatherer set used by protocol-catalog construction. */
class CatalogGatherersTest {
  @Test
  void toOrderedUniqueOrThrowPreservesEncounterOrder() {
    List<Entry> actual =
        List.of(new Entry("alpha", "A"), new Entry("beta", "B"), new Entry("gamma", "C")).stream()
            .gather(CatalogGatherers.toOrderedUniqueOrThrow(Entry::key, "test entry"))
            .toList();

    assertEquals(
        List.of(new Entry("alpha", "A"), new Entry("beta", "B"), new Entry("gamma", "C")), actual);
  }

  @Test
  void toOrderedUniqueOrThrowRejectsDuplicatesUsingProjectedValues() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                List.of(new Entry("alpha", "LEFT"), new Entry("alpha", "RIGHT")).stream()
                    .gather(
                        CatalogGatherers.toOrderedUniqueOrThrow(
                            Entry::key, Entry::value, "test entry"))
                    .toList());

    assertEquals(
        "Duplicate test entry detected while building the protocol catalog: LEFT / RIGHT",
        failure.getMessage());
  }

  @Test
  void toOrderedUniqueOrThrowIsDeterministicAcrossRepeatedTraversals() {
    List<Entry> values = List.of(new Entry("alpha", "A"), new Entry("beta", "B"));

    List<Entry> first =
        values.stream()
            .gather(CatalogGatherers.toOrderedUniqueOrThrow(Entry::key, "entry"))
            .toList();
    List<Entry> second =
        values.stream()
            .gather(CatalogGatherers.toOrderedUniqueOrThrow(Entry::key, "entry"))
            .toList();

    assertEquals(first, second);
  }

  @Test
  void toOrderedUniqueOrThrowRejectsBlankLabels() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> CatalogGatherers.toOrderedUniqueOrThrow(Entry::key, " "));

    assertEquals("label must not be blank", failure.getMessage());
  }

  @Test
  void expandFieldsWithMetadataPreservesRecordComponentOrderAndEnrichesMetadata() {
    List<FieldEntry> actual =
        Arrays.stream(CatalogFixture.class.getRecordComponents())
            .gather(CatalogGatherers.expandFieldsWithMetadata(Set.of("notes", "enabled")))
            .toList();

    assertEquals(
        List.of("title", "notes", "enabled"), actual.stream().map(FieldEntry::name).toList());
    assertEquals(FieldRequirement.REQUIRED, actual.get(0).requirement());
    assertEquals(FieldRequirement.OPTIONAL, actual.get(1).requirement());
    assertEquals(FieldRequirement.OPTIONAL, actual.get(2).requirement());
    assertEquals(new FieldShape.Scalar(ScalarType.STRING), actual.get(0).shape());
    assertEquals(
        new FieldShape.ListShape(new FieldShape.Scalar(ScalarType.STRING)), actual.get(1).shape());
    assertEquals(List.of("ON", "OFF"), actual.get(2).enumValues());
  }

  @Test
  void expandFieldsWithMetadataIsDeterministicAcrossRepeatedTraversals() {
    RecordComponent[] components = CatalogFixture.class.getRecordComponents();

    List<FieldEntry> first =
        Arrays.stream(components)
            .gather(CatalogGatherers.expandFieldsWithMetadata(Set.of("notes")))
            .toList();
    List<FieldEntry> second =
        Arrays.stream(components)
            .gather(CatalogGatherers.expandFieldsWithMetadata(Set.of("notes")))
            .toList();

    assertEquals(first, second);
  }

  @Test
  void expandFieldsWithMetadataRejectsNullOptionalFields() {
    NullPointerException failure =
        assertThrows(
            NullPointerException.class, () -> CatalogGatherers.expandFieldsWithMetadata(null));

    assertEquals("optionalFields must not be null", failure.getMessage());
  }

  @Test
  void duplicateEntryFailureIncludesLabelAndBothValues() {
    IllegalStateException failure =
        CatalogDuplicateFailures.duplicateEntryFailure("test label", "LEFT", "RIGHT");

    assertTrue(failure.getMessage().contains("test label"));
    assertTrue(failure.getMessage().contains("LEFT"));
    assertTrue(failure.getMessage().contains("RIGHT"));
  }

  @Test
  void duplicateEntryFailureRejectsBlankLabels() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> CatalogDuplicateFailures.duplicateEntryFailure(" ", "LEFT", "RIGHT"));

    assertEquals("label must not be blank", failure.getMessage());
  }

  record Entry(String key, String value) {}

  record CatalogFixture(String title, List<String> notes, Mode enabled) {}

  /** Sample enum used to assert catalog enum-value metadata enrichment. */
  enum Mode {
    ON,
    OFF
  }
}
