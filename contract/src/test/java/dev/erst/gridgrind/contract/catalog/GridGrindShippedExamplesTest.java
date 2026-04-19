package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

/** Tests for generated-example metadata validation and accessors. */
class GridGrindShippedExamplesTest {
  @Test
  void shippedExampleRejectsLowerCaseDiscoveryIds() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new GridGrindShippedExamples.ShippedExample(
                    "workbook_health",
                    "workbook-health-request.json",
                    "summary",
                    GridGrindProtocolCatalog.requestTemplate()));

    assertEquals("id must use upper-case discovery tokens", failure.getMessage());
  }

  @Test
  void shippedExampleRejectsNonJsonFileNames() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new GridGrindShippedExamples.ShippedExample(
                    "WORKBOOK_HEALTH",
                    "workbook-health-request.txt",
                    "summary",
                    GridGrindProtocolCatalog.requestTemplate()));

    assertEquals("fileName must end with .json", failure.getMessage());
  }

  @Test
  void shippedExampleEntryRejectsNonJsonFileNames() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new ShippedExampleEntry(
                    "WORKBOOK_HEALTH", "workbook-health-request.txt", "summary"));

    assertEquals("fileName must end with .json", failure.getMessage());
  }
}
