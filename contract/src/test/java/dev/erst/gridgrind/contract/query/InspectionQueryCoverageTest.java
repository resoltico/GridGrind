package dev.erst.gridgrind.contract.query;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator;
import java.nio.charset.StandardCharsets;
import org.junit.jupiter.api.Test;

/** Covers request-defaulting branches for inspection queries that deserialize through creators. */
class InspectionQueryCoverageTest {
  @Test
  void exportCustomXmlMappingCreatorSuppliesDefaultsAndPreservesExplicitValues() {
    InspectionQuery.ExportCustomXmlMapping defaulted =
        InspectionQuery.ExportCustomXmlMapping.create(
            new CustomXmlMappingLocator(1L, null), null, null);
    InspectionQuery.ExportCustomXmlMapping explicit =
        InspectionQuery.ExportCustomXmlMapping.create(
            new CustomXmlMappingLocator(null, "BudgetMapping"), true, "UTF-16");

    assertFalse(defaulted.validateSchema());
    assertEquals(StandardCharsets.UTF_8.name(), defaulted.encoding());
    assertTrue(explicit.validateSchema());
    assertEquals("UTF-16", explicit.encoding());
  }
}
